# Run this ON THE TARGET MACHINE to remove the add-in(s) installed by
# install.ps1.
#
# Usage:
#   .\uninstall.ps1                        # removes Word, Excel, and PowerPoint
#   .\uninstall.ps1 -App Word              # removes just Word
#   .\uninstall.ps1 -RemoveTrustedCert     # also untrusts the signing certificate

param(
    [ValidateSet('Word', 'Excel', 'PowerPoint', 'All')]
    [string]$App = 'All',
    [string]$InstallRoot = "$env:LOCALAPPDATA\AirchatOffice",
    [switch]$RemoveTrustedCert
)

$Apps = if ($App -eq 'All') { @('Word', 'Excel', 'PowerPoint') } else { @($App) }
$OfficeApps = @{ Word = 'Word'; Excel = 'Excel'; PowerPoint = 'PowerPoint' }

foreach ($appName in $Apps) {
    $projName = "${appName}AiAddIn"
    $regKey = "HKCU:\Software\Microsoft\Office\$($OfficeApps[$appName])\Addins\$projName"
    if (Test-Path $regKey) {
        Remove-Item $regKey -Recurse -Force
        Write-Host "Removed registration for $projName."
    }
    $destDir = Join-Path $InstallRoot $projName
    if (Test-Path $destDir) {
        Remove-Item $destDir -Recurse -Force
        Write-Host "Removed installed files at $destDir."
    }
}

if ($RemoveTrustedCert) {
    $Thumbprint = 'D7C0D2DECAE0E7D967DE0A2C1B5DBFF185A932FD'
    $cert = Get-ChildItem "Cert:\CurrentUser\TrustedPublisher\$Thumbprint" -ErrorAction SilentlyContinue
    if ($cert) {
        Remove-Item $cert.PSPath -Force
        Write-Host "Removed trusted certificate."
    }
}

Write-Host "Uninstall complete. Restart Office applications for changes to take effect."
