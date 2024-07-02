function Invoke-Sf {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $Command)
    Write-Verbose $Command
    return Invoke-Expression -Command $Command
}

function Show-SfResult {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][psobject] $Result)
    $result = $Result | ConvertFrom-Json
    if ($result.status -ne 0) {
        Write-Debug $result
        throw ($result.message)
    }
    return $result.result
}

function Install-SalesforceLwcDevServer {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "npm install -g node-gyp"
    Invoke-Sf -Command "sf plugins install @salesforce/lwc-dev-server"
    Invoke-Sf -Command "sf plugins update"
}

function Start-SalesforceLwcDevServer {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "sfdx force:lightning:lwc:start"
}

function Set-SalesforceDefaultDevHub {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $DevhubUsername
    )
    Invoke-Sf -Command "sf config set target-dev-hub=$DevhubUsername --global"
}

function Remove-SalesforceDefaultDevHub {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "sf config unset target-dev-hub --global"
}

function Get-SalesforceConfig {
    [CmdletBinding()]
    Param()
    $command = "sf config list --json"
    $result = Invoke-Sf -Command $command
    Show-SfResult -Result $result
}

function Get-SalesforceScratchOrgs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][switch] $SkipConnectionStatus,
        [Parameter(Mandatory = $false)][switch] $Last
    )

    $command = "sf org list --all"
    if ($SkipConnectionStatus) {
        $command += " --skip-connection-status"
    }
    $command += " --json"
    $result = Invoke-Sf -Command $command

    $result = $result | ConvertFrom-Json
    $result = $result.result.scratchOrgs
    $result = $result | Select-Object orgId, instanceUrl, username, connectedStatus, isDevHub, lastUsed, alias
    if ($Last) {
        $result = $result | Sort-Object lastUsed -Descending | Select-Object -First 1
    }
    return $result
}

function New-SalesforceScratchOrg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $DevhubUsername,
        [Parameter(Mandatory = $false)][switch] $Set,
        [Parameter(Mandatory = $false)][int] $DurationDays,
        [Parameter(Mandatory = $false)][string] $DefinitionFile = 'config/project-scratch-def.json',
        [Parameter(Mandatory = $false)][int] $WaitMinutes        
    )
    $command = "sf org create scratch"
    if ($DevhubUsername) {
        $command += " --target-dev-hub $DevhubUsername"
    }
    if ($DaysDuration) {
        $command += " --duration-days $DurationDays"
    }
    $command += " --definition-file $DefinitionFile"
    if ($WaitMinutes) {
        $command += " --wait $WaitMinutes"
    }
    $command += " --json"

    $result = Invoke-Sf -Command $command
    Show-SfResult -Result $result

    $scratchOrgUsername = $result.username
    if ($Set) {
        Set-SalesforceProjectUser -Username $scratchOrgUsername
    }
}

function Remove-SalesforceScratchOrg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $ScratchOrgUserName,
        [Parameter()][switch] $NoPrompt
    )
    $command = "sf org delete scratch --target-org $ScratchOrgUserName"
    if ($NoPrompt) {
        $command += " --no-prompt"
    }
    Invoke-Sf -Command $command
}

function Remove-SalesforceScratchOrgs {
    [CmdletBinding()]
    Param()
    
    $scratchOrgs = Get-SalesforceScratchOrgs
    foreach ($scratchOrg in $scratchOrgs) {
        Remove-SalesforceScratchOrg -ScratchOrgUserName ($scratchOrg.username) -NoPrompt
    }    
}

function New-SalesforceProject {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $false)][string][ValidateSet('standard', 'empty', 'analytics')] $Template = 'standard',
        [Parameter(Mandatory = $false)][string] $DefaultUserName = $null,

        [Parameter(Mandatory = $false)][string] $OutputDirectory,
        [Parameter(Mandatory = $false)][string] $DefaultPackageDirectory,
        [Parameter(Mandatory = $false)][string] $Namespace,
        [Parameter(Mandatory = $false)][switch] $GenerateManifest
    )
    $command = "sf force project create --name $Name"

    if ($OutputDirectory) {
        $command += " --output-dir $OutputDirectory"
    }
    if ($DefaultPackageDirectory) {
        $command += " --default-package-dir $DefaultPackageDirectory"
    }
    if ($Namespace) {
        $command += " --namespace $Namespace"
    }
    if ($GenerateManifest) {
        $command += " --manifest"
    }

    $command += " --template $Template"
    $command += " --json"

    $result = Invoke-Sf -Command $command
    $result = Show-SfResult -Result $result

    if (($null -ne $DefaultUserName) -and ($DefaultUserName -ne '')) {
        $projectFolder = Join-Path -Path $result.outputDir -ChildPath $Name
        New-Item -Path $projectFolder -Name ".sfdx" -ItemType Directory | Out-Null
        Set-SalesforceProject -DefaultUserName $DefaultUserName -ProjectFolder $projectFolder
    }
    return $result
}

function New-SalesforceProjectAndScratchOrg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $true)][string] $DevhubUsername
    )
    New-SalesforceProject -Name $Name
    Push-Location -Path $Name    
    Remove-SalesforceScratchOrgs 
    $scratchOrg = New-SalesforceScratchOrg -DevhubUsername $DevhubUsername    
    Set-SalesforceProjectUser -Username ($scratchOrg.username)
}

function Set-SalesforceProject {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $DefaultUserName,
        [Parameter(Mandatory = $false)][string] $ProjectFolder
    )

    if (($null -eq $ProjectFolder) -or ($ProjectFolder -eq '')) {
        $sfdxFolder = (Get-Location).Path
    }
    else {
        $sfdxFolder = $ProjectFolder
    }

    if ($sfdxFolder.EndsWith(".sfdx") -eq $false) {
        $sfdxFolder = Join-Path -Path $sfdxFolder -ChildPath ".sfdx"
    }

    if ((Test-Path -Path $sfdxFolder) -eq $false) {
        throw ".sfdx folder does not exist ing $sfdxFolder"
    }

    $sfdxFile = Join-Path -Path $sfdxFolder -ChildPath "sfdx-config.json"
    if (Test-Path -Path $sfdxFile) {
        throw "File already exists $sfdxFile"
    }

    New-Item -Path $sfdxFile | Out-Null
    $json = "{ `"defaultusername`": `"$DefaultUserName`" }"
    Set-Content -Path $sfdxFile -Value $json
}

function Get-IsSalesforceProject {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $ProjectFolder)

    $sfdxProjectFile = Join-Path -Path $ProjectFolder -ChildPath "sfdx-project.json"
    if (Test-Path -Path $sfdxProjectFile) {
        return $true
    }
    return $false
}

function Get-SalesforceDefaultUserName {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $false)][string] $ProjectFolder)

    if (($null -eq $ProjectFolder) -or ($ProjectFolder -eq '')) {
        $sfdxFolder = (Get-Location).Path
    } else {
        $sfdxFolder = $ProjectFolder
    }

    $sfdxConfigFile = ""
    $files = Get-ChildItem -Recurse -Filter "sfdx-config.json"
    foreach ($file in $files) {
        if ($file.FullName -like "*.sfdx*") {
            $sfdxConfigFile = $file
            break
        }
    }

    if (!(Test-Path -Path $sfdxConfigFile)) {
        throw "Missing Salesforce Project File (sfdx-config.json)"
    }
    Write-Verbose "Found sfdx config ($sfdxConfigFile)"

    $salesforceSettings = Get-Content -Raw -Path $sfdxConfigFile | ConvertFrom-Json
    return $salesforceSettings.defaultusername
}

function Get-SalesforceProjectConfig {
    [CmdletBinding()]
    Param()
    $sfdxConfigFile = ""
    $files = Get-ChildItem -Recurse -Filter "sfdx-config.json"
    foreach ($file in $files) {
        if ($file.FullName -like "*.sfdx*") {
            $sfdxConfigFile = $file
            break
        }
    }

    if (!(Test-Path -Path $sfdxConfigFile)) {
        throw "Missing Salesforce Project File (sfdx-config.json)"
    }
    Write-Verbose "Found sfdx config ($sfdxConfigFile)"
    return $sfdxConfigFile
}

function Get-SalesforceProjectUser {
    [CmdletBinding()]
    Param()
    $sfdxConfigFile = Get-SalesforceProjectConfig
    $salesforceSettings = Get-Content -Raw -Path $sfdxConfigFile | ConvertFrom-Json
    return $salesforceSettings.defaultusername
}

function Set-SalesforceProjectUser {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $Username)
    Invoke-Sf -Command "sf config set target-org $Username"
}

function DeployAndTest-SalesforceApex {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $ClassName,
        [Parameter(Mandatory = $false)][string] $TestClassName
    )

    $command = "sf project deploy start"
    $command += " --metadata ApexClass:$ClassName*"
    $command += " --ignore-conflicts"
    $command += " --test-level RunSpecifiedTests"
    $command += " --tests $TestClassName"

    Invoke-Sf -Command $command
}

function Test-Salesforce {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $ClassName,
        [Parameter(Mandatory = $false)][string] $TestName,
        [Parameter(Mandatory = $false)][string] $Username,

        [Parameter(Mandatory = $false)][string][ValidateSet('human', 'tap', 'junit', 'json')] $ResultFormat = 'json',

        [Parameter(Mandatory = $false)][switch] $RunAsynchronously,
        [Parameter(Mandatory = $false)][switch] $CodeCoverage,
        [Parameter(Mandatory = $false)][int] $WaitMinutes = 10,

        [Parameter(Mandatory = $false)][string] $OutputDirectory
    )

    $command = "sf apex run test"
    if ($ClassName -and $TestName) {
        # Run specific Test in a Class
        $command += " --tests $ClassName.$TestName"
        if ($RunAsynchronously) { $command += "" }
        else { $command += " --synchronous" }

    } elseif ((-not $TestName) -and ($ClassName)) {
        # Run Test Class
        $command += " --class-names $ClassName"
        if ($RunAsynchronously) { $command += "" }
        else { $command += " --synchronous" }
    } else {
        $command += " --test-level RunLocalTests" # Run all Tests
    }

    if ($OutputDirectory) {
        $command += " --output-dir $OutputDirectory"
    } else {
        $command += " --output-dir $PSScriptRoot"
    }

    if ($CodeCoverage) { $command += " --detailed-coverage" }
    if ($Username) { $command += " --target-org $Username" }
    $command += " --result-format $ResultFormat"

    $result = Invoke-Sf -Command $command
    $result = $result | ConvertFrom-Json

    Write-Verbose $result

    $result.result.tests
    if ($result.result.summary.outcome -ne 'Passed') {
        throw ($result.result.summary.failing.tostring() + " Tests Failed")
    }

    if (!$IncludeCodeCoverage) {
        return
    }

    [int]$codeCoverage = ($result.result.summary.testRunCoverage -replace '%')
    if ($codeCoverage -lt 75) {
        $result.result.coverage.coverage
        throw 'Insufficent code coverage '
    }
}

function Get-SalesforceCodeCoverage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $ApexClassOrTrigger = $null,
        [Parameter(Mandatory = $true)][string] $Username
    )
    $query = "SELECT ApexTestClass.Name, TestMethodName, ApexClassOrTrigger.Name, NumLinesUncovered, NumLinesCovered, Coverage "
    $query += "FROM ApexCodeCoverage "
    if (($null -ne $ApexClassOrTrigger) -and ($ApexClassOrTrigger -ne '')) {
        $apexClass = Get-SalesforceApexClass -Name $ApexClassOrTrigger -Username $Username
        $apexClassId = $apexClass.Id
        $query += "WHERE ApexClassOrTriggerId = '$apexClassId' "
    }

    $result = Invoke-Sf -Command "sf data query --query `"$query`" --use-tooling-api --target-org $Username --json"
    $result = $result | ConvertFrom-Json
    if ($result.status -ne 0) {
        throw ($result.message)
    }
    $result = $result.result.records

    $values = @()
    foreach ($item in $result) {
        $value = New-Object -TypeName PSObject
        $value | Add-Member -MemberType NoteProperty -Name 'ApexClassOrTrigger' -Value $item.ApexClassOrTrigger.Name
        $value | Add-Member -MemberType NoteProperty -Name 'ApexTestClass' -Value $item.ApexTestClass.Name
        $value | Add-Member -MemberType NoteProperty -Name 'TestMethodName' -Value $item.TestMethodName

        $codeCoverage = 0
        $codeLength = $item.NumLinesCovered + $item.NumLinesUncovered
        if ($codeLength -gt 0) {
            $codeCoverage = $item.NumLinesCovered / $codeLength
        }
        $value | Add-Member -MemberType NoteProperty -Name 'CodeCoverage' -Value $codeCoverage.toString("P")
        $codeCoverageOK = $false
        if ($codeCoverage -ge 0.75) { $codeCoverageOK = $true }

        $value | Add-Member -MemberType NoteProperty -Name 'CodeCoverageOK' -Value $codeCoverageOK
        $value | Add-Member -MemberType NoteProperty -Name 'NumLinesCovered' -Value $item.NumLinesCovered
        $value | Add-Member -MemberType NoteProperty -Name 'NumLinesUncovered' -Value $item.NumLinesUncovered
        $values += $value
    }

    return $values
}

function Install-SalesforceJest {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "yarn add -D @salesforce/sfdx-lwc-jest"
}

function New-SalesforceJestTest {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $LwcName)
    $filePath = "force-app/main/default/lwc/$LwcName/$LwcName.js"
    $command = "sf force lightning lwc test create --filepath $filePath --json"
    $result = Invoke-Sf -Command $command
    return Show-SfResult -Result $result
}

function Test-SalesforceJest {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "npm run test:unit"
}

function Debug-SalesforceJest {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "npm run test:unit:debug"
}

function Watch-SalesforceJest {
    [CmdletBinding()]
    Param()
    Invoke-Sf -Command "npm run test:unit:watch"
}

function Deploy-SalesforceComponent {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string][ValidateSet('ApexClass', 'ApexTrigger')] $Type = 'ApexClass',
        [Parameter(Mandatory = $false)][string] $Name,
        [Parameter(Mandatory = $true)][string] $Username
    )
    $command = "sfdx force:source:deploy --metadata $Type"
    if ($Name) {
        $command += ":$Name"
    }
    $command += " --targetusername $Username"
    $command += " --json"

    $response = Invoke-Sf -Command $command | ConvertFrom-Json
    if ($response.result.success -ne $true) {
        Write-Verbose $result
        throw ("Failed to Deploy ")
    }
}

function Get-SalesforceType {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $false)][string] $FileName)

    if ($FileName.EndsWith(".cls")) {
        return "ApexClass"
    }
    if ($FileName.EndsWith(".cls")) {
        return "ApexTrigger"
    }
    return ""
}

function Get-SalesforceName {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $false)][string] $FileName)

    $name = (Get-Item $FileName).Basename
    Write-Verbose ("Apex Name: " + $name)
    return $name
}

function Get-SalesforceTestResultsApexFolder {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $ProjectFolder)

    $folder = Join-Path -Path $ProjectFolder -ChildPath ".sfdx\tools\testresults\apex"
    Write-Verbose ("Apex Test Results Folder: " + $folder)
    # TODO: Check Folder Exists
    return $folder
}

function Get-SalesforceApexTestsClasses {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $ProjectFolder)

    $classesFolder = Join-Path -Path $ProjectFolder -ChildPath "force-app\main\default\classes"
    $classes = Get-ChildItem -Path $classesFolder -Filter *.cls
    $testClasses = @()
    foreach ($class in $classes) {
        if (Select-String -Path $class -Pattern "@isTest") {
            Write-Verbose ("Found Apex Test Class: " + $class)
            $testClasses += Get-SalesforceName -FileName $class
        }
    }
    $testClassNames = $testClasses -join ","
    Write-Verbose ("Apex Test Class Names: " + $testClassNames)
    return $testClassNames
}

function Watch-SalesforceApex {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $ProjectFolder,
        [Parameter(Mandatory = $true)][string] $FileName
    )

    if ((Get-IsSalesforceProject -ProjectFolder $ProjectFolder) -eq $false) {
        Write-Verbose "Not a Salesforce Project"
        return
    }
    $username = Get-SalesforceDefaultUserName -ProjectFolder $ProjectFolder

    $type = Get-SalesforceType -FileName $FileName
    if (($type -eq "ApexClass") -or ($type -eq "ApexTrigger")) {
        $name = Get-SalesforceName -FileName $FileName
        Deploy-SalesforceComponent -Type $type -Name $name -Username $username

        $outputDir = Get-SalesforceTestResultsApexFolder -ProjectFolder $ProjectFolder
        $testClassNames = Get-SalesforceApexTestsClasses -ProjectFolder $ProjectFolder
        Test-Salesforce -Username $username -ClassName $testClassNames -IncludeCodeCoverage:$false -OutputDirectory $outputDir
    }
}

function Push-Salesforce {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $Username,
        [Parameter(Mandatory = $false)][switch] $IgnoreErrors,
        [Parameter(Mandatory = $false)][switch] $IgnoreConflicts,
        [Parameter(Mandatory = $false)][switch] $IgnoreWarnings,

        [Parameter(Mandatory = $false)][switch] $Async,
        [Parameter(Mandatory = $false)][switch] $Concise,
        [Parameter(Mandatory = $false)][switch] $DryRun,

        [Parameter(Mandatory = $false)][switch] $Test
    )

    $command = "sf project deploy start"
    if ($Username) { $command += " --target-org $Username" }
    if ($IgnoreErrors) { $command += " --ignore-errors" }
    if ($IgnoreConflicts) { $command += " --ignore-conflicts" }
    if ($IgnoreWarnings) { $command += " --ignore-warnings" }

    if ($Async) { $command += " --async" }
    if ($Concise) { $command += " --concise" }
    if ($DryRun) { $command += " --dry-run" }

    if ($Test) { $command += " --test-level RunLocalTests" }

    Invoke-Sf -Command $command
}

function Pull-Salesforce {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][string] $Username,
        [Parameter(Mandatory = $false)][string] $PackageNames,
        [Parameter(Mandatory = $false)][switch] $IgnoreConflicts,
        [Parameter(Mandatory = $false)][switch] $IgnoreWarnings
    )

    $command = "sf project retrieve start"
    if ($Username) { $command += " --target-org $Username"}
    if ($PackageNames) { $command += " --package-name $PackageNames"}
    if ($IgnoreConflicts) { $command += " --ignore-conflicts"}
    Invoke-Sf -Command $command
}

function New-SalesforceApexClass {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $false)][string]
            [ValidateSet('DefaultApexClass', 'ApexUnitTest', 'ApexUnitTest', 'InboundEmailService')]
            $Template = 'DefaultApexClass',
        [Parameter(Mandatory = $false)][string] $OutputDirectory = 'ForceAppDefault'
    )

    $command = "sf apex generate class"
    $command += " --name $Name"
    $command += " --template $Template"

    if ($OutputDirectory = 'ForceAppDefault') {
        $OutputDirectory = "force-app/main/default/classes"
    }
    if ($OutputDirectory) {
        $command += " --output-dir $OutputDirectory"
    }
    Invoke-Sf -Command $command
}

function New-SalesforceApexTrigger {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $false)][string]
            [ValidateSet('before insert', 'before update', 'before delete', 'after insert', 'after update', 'after delete', 'after undelete')]
            $Event = 'before insert',
        [Parameter(Mandatory = $false)][string] $SObject,
        [Parameter(Mandatory = $false)][string] $OutputDirectory = 'ForceAppDefault'
    )

    $command = "sf apex generate trigger"
    $command += " --name $Name"
    $command += " --event $Event"
    if ($SObject) {
        $command += " --sobject $SObject"
    }
    if ($OutputDirectory = 'ForceAppDefault') {
        $OutputDirectory = "force-app/main/default/triggers"
    }
    if ($OutputDirectory) {
        $command += " --output-dir $OutputDirectory"
    }
    Invoke-Sf -Command $command
}

Export-ModuleMember Install-SalesforceLwcDevServer
Export-ModuleMember Start-SalesforceLwcDevServer

Export-ModuleMember Set-SalesforceDefaultDevHub
Export-ModuleMember Remove-SalesforceDefaultDevHub
Export-ModuleMember Get-SalesforceConfig

Export-ModuleMember Get-SalesforceScratchOrgs
Export-ModuleMember New-SalesforceScratchOrg
Export-ModuleMember Remove-SalesforceScratchOrg
Export-ModuleMember Remove-SalesforceScratchOrgs

Export-ModuleMember New-SalesforceProject
Export-ModuleMember Set-SalesforceProject
Export-ModuleMember Get-SalesforceDefaultUserName
Export-ModuleMember Get-SalesforceProjectUser
Export-ModuleMember Set-SalesforceProjectUser
Export-ModuleMember New-SalesforceProjectAndScratchOrg

Export-ModuleMember Test-Salesforce
Export-ModuleMember DeployAndTest-SalesforceApex
Export-ModuleMember Get-SalesforceCodeCoverage

Export-ModuleMember Install-SalesforceJest
Export-ModuleMember New-SalesforceJestTest
Export-ModuleMember Test-SalesforceJest
Export-ModuleMember Debug-SalesforceJest
Export-ModuleMember Watch-SalesforceJest

Export-ModuleMember Watch-SalesforceApex

Export-ModuleMember Push-Salesforce
Export-ModuleMember Pull-Salesforce

Export-ModuleMember New-SalesforceApexClass
Export-ModuleMember New-SalesforceApexTrigger