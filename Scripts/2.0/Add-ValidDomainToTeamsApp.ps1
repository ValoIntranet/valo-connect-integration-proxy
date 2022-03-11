<#
 .SYNOPSIS
    Add-ValidDomainToTeamsApp.ps1

 .DESCRIPTION
    Adds a valid domain to a custom Microsoft Teams app manifest file

 .PARAMETER AppManifestPath
    Path to the app manifest file to add the valid domain to

 .PARAMETER ValidDomain
    The domain to add as a ValidDomain to the Microsoft Teams app manifest file

 .PARAMETER NoBackup
    (switch) Don't take a backup of the app manifest file

 .PARAMETER KeepTempFolder
    (switch) Keep the temp folder where the script updates the manifest

#>

param(    
    [Parameter(Mandatory = $false)]
    [string]$AppManifestPath = ".\Valo.Connect.app.zip",

    [Parameter(Mandatory = $true)]
    [string]$ValidDomain,

    [Parameter(Mandatory = $false)]
    [switch]$NoBackup,

    [Parameter(Mandatory = $false)]
    [switch]$KeepTempFolder

)

if ($true -ne $AppManifestPath.ToLower().EndsWith(".zip")) {
    Write-Error "AppManifestPath doesn't have a .zip extension"
    exit
}

$AppManifestFile = Get-Item $AppManifestPath -ErrorAction SilentlyContinue

if ($null -eq $AppManifestFile) {
    Write-Error "Couldn't find file from AppManifestPath paramater: $AppManifestPath"
    exit
}

$DestinationFolderName = $AppManifestFile.Name.Replace($AppManifestFile.Extension, "")
$DestinationPath = Join-Path -Path $AppManifestFile.Directory.FullName -ChildPath "Temp" -AdditionalChildPath $DestinationFolderName

Expand-Archive $AppManifestFile.FullName -DestinationPath $DestinationPath -Force

$ManifestPath = Join-Path -Path $DestinationPath -ChildPath "manifest.json"
if ($true -ne (Test-Path $ManifestPath)) {
    Write-Error "Unable to find manifest.json in app manifest. Is $($AppManifestFile.Name) a valid Teams app?"
    exit
}

$ManifestContentsJson = ConvertFrom-Json (Get-Content $ManifestPath -Raw)
if ($true -eq $ManifestContentsJson.validDomains.Contains($ValidDomain)) {
    Write-Warning "Manifest file in $($AppManifestFile.Name) already contains validDomain $ValidDomain, no change required"
    exit
}

$ManifestContentsJson.validDomains += $ValidDomain
New-Item -Path $DestinationPath -Name "manifest.json" -ItemType File -Value (ConvertTo-Json -InputObject $ManifestContentsJson -Depth 99) -Force | Out-Null

$BackupPath = Join-Path -Path $AppManifestFile.Directory.FullName -ChildPath "Temp" -AdditionalChildPath "Backup"
New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null

if ($true -ne $NoBackup) {
    $BackupFileDestination = Join-Path -Path $BackupPath -ChildPath $AppManifestFile.Name
    Write-Output "Backing up $($AppManifestFile.Name) to $($BackupFileDestination.Replace($AppManifestFile.Directory.FullName, "."))"
    Move-Item $AppManifestFile.FullName -Destination $BackupFileDestination -Force
}

Write-Output "Updating $($AppManifestFile.Name) with additional ValidDomain $ValidDomain"
Compress-Archive -DestinationPath $AppManifestFile.FullName -LiteralPath (Get-ChildItem $DestinationPath -Depth 1) -Force

if ($true -ne $KeepTempFolder) {
    Remove-Item -Path $BackupPath -Recurse
}
