[CmdletBinding()]
Param([Parameter(Mandatory = $false)][switch] $LocalFile) 

$name = "psfdx-development"

# Get Module Path
$modulespath = ($env:psmodulepath -split ";")[0]
$path = Join-Path -Path $modulespath -ChildPath $name
Write-Verbose "Module Path: $path"

# Get Source Path
$filename = "$name.psm1"
$gitUrl = "https://raw.githubusercontent.com/tonygward/$name/master/$name.psm1"

# Install
Write-Host "Creating module directory"
New-Item -Type Container -Force -path $path | Out-Null

Write-Host "Downloading and installing"
$destination = Join-Path -Path $path -ChildPath $filename
if ($LocalFile) {
    $currentFile = Join-Path -Path $PSScriptRoot -ChildPath $filename        
    Write-Verbose "Source File: $currentFile"
    Copy-Item -Path $currentFile -Destination $destination -Force
}
else {
    Write-Verbose "Source File: $gitUrl"
    $response = Invoke-WebRequest -Uri $gitUrl -Verbose:$false
    # TODO: Check WebRequest Status    
    $response.Content | Out-File -FilePath (Join-Path -Path $path -ChildPath $filename) -Force
}

Write-Host "Installed."
Write-Host ('Use "Import-Module ' + $name + '" and then "Get-Command -Module ' + $name + '"')

Import-Module $name -Force