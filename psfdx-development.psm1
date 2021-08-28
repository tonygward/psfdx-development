function Invoke-Sfdx {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $Command)        
    Write-Verbose $Command
    return Invoke-Expression -Command $Command
}

function Show-SfdxResult {
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
    Invoke-Sfdx -Command "npm install -g node-gyp"
    Invoke-Sfdx -Command "sfdx plugins:install @salesforce/lwc-dev-server"
    Invoke-Sfdx -Command "sfdx plugins:update"
}

function Start-SalesforceLwcDevServer {
    [CmdletBinding()]
    Param()       
    Invoke-Sfdx -Command "sfdx force:lightning:lwc:start"    
}

function Get-SalesforceScratchOrgs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][switch] $SkipConnectionStatus,
        [Parameter(Mandatory = $false)][switch] $Last
    )

    $command = "sfdx force:org:list --all"
    if ($SkipConnectionStatus) {
        $command += " --skipconnectionstatus"
    }
    $command += " --json"
    $result = Invoke-Sfdx -Command "sfdx force:org:list --all --json"
    
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
        [Parameter(Mandatory = $true)][string] $Username,
        [Parameter(Mandatory = $false)][string] $DevhubUsername,
        [Parameter(Mandatory = $false)][int] $DurationDays,
        [Parameter(Mandatory = $false)][string] $DefinitionFile = 'config/project-scratch-def.json',
        [Parameter(Mandatory = $false)][int] $WaitMinutes
    )     
    $command = "sfdx force:org:create"   
    $command += " --type scratch"
    if ($DevhubUsername) {
        $command += " --targetdevhubusername $DevhubUsername"
    }
    if ($DaysDuration) {
        $command += " --durationdays $DurationDays"
    }
    $command += " --definitionfile $DefinitionFile"
    $command += " --targetdevhubusername $Username"
    if ($WaitMinutes) { 
        $command += " --wait $WaitMinutes"
    }
    $command += " --json"

    $result = Invoke-Sfdx -Command $command
    return Show-SfdxResult -Result $result
}

function Remove-SalesforceScratchOrg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $ScratchOrgUserName,
        [Parameter()][switch] $NoPrompt
    )    
    # TODO: Check is Scratch Org   
    $command = "sfdx force:org:delete --targetusername $ScratchOrgUserName"
    if ($NoPrompt) {
        $command += " --noprompt"
    }
    Invoke-Sfdx -Command $command
}

function New-SalesforceProject {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $false)][string][ValidateSet('standard', 'empty')] $Template = 'standard',
        [Parameter(Mandatory = $false)][string] $DefaultUserName = $null,

        [Parameter(Mandatory = $false)][string] $OutputDirectory,
        [Parameter(Mandatory = $false)][string] $DefaultPackageDirectory,
        [Parameter(Mandatory = $false)][string] $Namespace,
        [Parameter(Mandatory = $false)][switch] $GenerateManifest
    )      
    # "sfdx force:project:create --projectname $Name --template $Template --json"
    $command = "sfdx force:project:create --projectname $Name"

    if ($OutputDirectory) {
        $command += " --outputdir $OutputDirectory"
    }
    if ($DefaultPackageDirectory) {
        $command += " --defaultpacakgedir $DefaultPackageDirectory"
    }
    if ($Namespace) {
        $command += " --namespace $Namespace"
    }
    if ($GenerateManifest) {
        $command += " --manifest"
    }
    
    $command += " --template $Template"
    $command += " --json"

    $result = Invoke-Sfdx -Command "sfdx force:project:create --projectname $Name --template $Template --json" 
    $result = Show-SfdxResult -Result $result
    
    if (($null -ne $DefaultUserName) -and ($DefaultUserName -ne '')) {                
        $projectFolder = Join-Path -Path $result.outputDir -ChildPath $Name
        New-Item -Path $projectFolder -Name ".sfdx" -ItemType Directory | Out-Null
        Set-SalesforceProject -DefaultUserName $DefaultUserName -ProjectFolder $projectFolder 
    }
    return $result
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

function Test-Salesforce {
    [CmdletBinding()]
    Param(        
        [Parameter(Mandatory = $false)][string] $ClassName,       
        [Parameter(Mandatory = $false)][string] $TestName, 
        [Parameter(Mandatory = $true)][string] $Username,

        [Parameter(Mandatory = $false)][string][ValidateSet('human', 'tap', 'junit', 'json')] $ResultFormat = 'json',

        [Parameter(Mandatory = $false)][switch] $RunAsynchronously,
        [Parameter(Mandatory = $false)][switch] $DetailedCoverage,

        [Parameter(Mandatory = $false)][int] $WaitMinutes = 10
    )   
    
    $command = "sfdx force:apex:test:run"
    if ($ClassName -and $TestName) {
        # Run specific Test in a Class
        $command += " --tests $ClassName.$TestName" 
        if ($RunAsynchronously) { $command += "" }
        else { $command += " --synchronous" }

    }     
    elseif (-not $TestName) {
        # Run Test Class
        $command += " --classnames $ClassName" 
        if ($RunAsynchronously) { $command += "" }
        else { $command += " --synchronous" }
    }     
    else {
        # Run all Tests
        $command += " --testlevel RunLocalTests"           
    }

    $command += " --wait:$WaitMinutes"
    $command += " --outputdir $PSScriptRoot"

    if ($DetailedCoverage) {
        $command += " --detailedcoverage"
    }    
    $command += " --codecoverage"
    $command += " --targetusername $Username"
    $command += " --resultformat $ResultFormat"
    $command += " --json"

    $result = Invoke-Sfdx -Command $command
    $result = $result | ConvertFrom-Json
    
    [int]$codeCoverage = ($result.result.summary.testRunCoverage -replace '%')
    if ($codeCoverage -lt 75) { 
        $result.result.coverage.coverage                
        throw 'Insufficent code coverage '
    }

    $result.result.tests
    if ($result.result.summary.outcome -ne 'Passed') { 
        throw ($result.result.summary.failing.tostring() + " Tests Failed") 
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

    $result = Invoke-Sfdx -Command "sfdx force:data:soql:query --query `"$query`" --usetoolingapi --targetusername $Username --json"
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
    Invoke-Sfdx -Command "sfdx force:lightning:lwc:test:setup"
}

function New-SalesforceJestTest {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $true)][string] $LwcName)       
    $filePath = "force-app/main/default/lwc/$LwcName/$LwcName.js"
    $command = "sfdx force:lightning:lwc:test:create --filepath $filePath --json"
    $result = Invoke-Sfdx -Command $command    
    return Show-SfdxResult -Result $result 
}

function Test-SalesforceJest {
    [CmdletBinding()]
    Param()       
    Invoke-Sfdx -Command "npm run test:unit"
}

function Debug-SalesforceJest {
    [CmdletBinding()]
    Param()       
    Invoke-Sfdx -Command "npm run test:unit:debug"
}

function Watch-SalesforceJest {
    [CmdletBinding()]
    Param()       
    Invoke-Sfdx -Command "npm run test:unit:watch"
}

Export-ModuleMember Install-SalesforceLwcDevServer
Export-ModuleMember Start-SalesforceLwcDevServer

Export-ModuleMember Get-SalesforceScratchOrgs
Export-ModuleMember New-SalesforceScratchOrg
Export-ModuleMember Remove-SalesforceScratchOrg

Export-ModuleMember New-SalesforceProject
Export-ModuleMember Set-SalesforceProject

Export-ModuleMember Test-Salesforce
Export-ModuleMember Get-SalesforceCodeCoverage

Export-ModuleMember Install-SalesforceJest
Export-ModuleMember New-SalesforceJestTest
Export-ModuleMember Test-SalesforceJest
Export-ModuleMember Debug-SalesforceJest
Export-ModuleMember Watch-SalesforceJest