# Run this ON THE TARGET MACHINE to remove the add-in(s) installed by
# install.ps1.
#
# Usage (default: removes the machine-wide install made by install.ps1; needs an ELEVATED PowerShell):
#   .\uninstall.ps1                        # removes Word, Excel, PowerPoint, and Outlook for all users
#   .\uninstall.ps1 -App Word              # removes just Word
#   .\uninstall.ps1 -RemoveTrustedCert     # also untrusts the signing certificate
#   .\uninstall.ps1 -CurrentUser           # removes an install made with install.ps1 -CurrentUser (no admin needed)
#   .\uninstall.ps1 -WhatIf                # list what it would remove and remove nothing

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Word', 'Excel', 'PowerPoint', 'Outlook', 'All')]
    [string]$App = 'All',
    [string]$InstallRoot,
    [string]$PackageDirectory = $PSScriptRoot,
    [switch]$RemoveTrustedCert,
    [switch]$CurrentUser
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'paths.ps1')
$AllUsers = -not $CurrentUser
$Apps = if ($App -eq 'All') { @('Word', 'Excel', 'PowerPoint', 'Outlook') } else { @($App) }
$OfficeApps = @{ Word = 'Word'; Excel = 'Excel'; PowerPoint = 'PowerPoint'; Outlook = 'Outlook' }

if ($AllUsers) {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $needAdmin = "Removing a machine-wide install touches Program Files and HKLM, so it needs an elevated PowerShell (Run as administrator). Use -CurrentUser to remove a per-user install."
        if ($WhatIfPreference) { Write-Host "NOTE: $needAdmin (continuing only because of -WhatIf; nothing will be changed)." -ForegroundColor Yellow }
        else { throw $needAdmin }
    }
}

if (-not $InstallRoot) {
    if ($AllUsers) {
        $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
        $InstallRoot = Join-Path $programFiles 'OpenDocs'
    }
    else {
        $InstallRoot = "$env:LOCALAPPDATA\OpenDocs"
    }
}
# Refuse removal while Office holds this product's files open.
$processNames = @{Word='WINWORD';Excel='EXCEL';PowerPoint='POWERPNT';Outlook='OUTLOOK'}
$session = (Get-Process -Id $PID).SessionId
foreach ($appName in $Apps) {
    $running = @(Get-Process -Name $processNames[$appName] -ErrorAction SilentlyContinue)
    if (-not $AllUsers) { $running = @($running | Where-Object { $_.SessionId -eq $session }) }
    if ($running.Count -gt 0 -and -not $WhatIfPreference) { Write-Host "Close $appName before uninstalling."; exit 1618 }
    [void](Assert-OpenDocsProductDirectory -Root $InstallRoot -ProductName "${appName}AiAddIn")
}
$registryRoots = if ($AllUsers) { @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office') } else { @('HKCU:\Software\Microsoft\Office') }
$removed = if ($WhatIfPreference) { 'Would remove' } else { 'Removed' }

foreach ($appName in $Apps) {
    $projName = "${appName}AiAddIn"
    foreach ($officeRoot in $registryRoots) {
        $regKey = "$officeRoot\$($OfficeApps[$appName])\Addins\$projName"
        if (Test-Path $regKey) {
            Remove-Item $regKey -Recurse -Force
            Write-Host "$removed registration for $projName ($regKey)."
        }
    }
    $destDir = Assert-OpenDocsProductDirectory -Root $InstallRoot -ProductName $projName
    if (Test-Path $destDir) {
        Remove-Item $destDir -Recurse -Force
        Write-Host "$removed installed files at $destDir."
    }

    if ($appName -eq 'Outlook' -and -not $AllUsers) {
        foreach ($ver in @('16.0', '15.0')) {
            $dna = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Resiliency\DoNotDisableAddinList"
            if (Test-Path $dna) { Remove-ItemProperty -Path $dna -Name $projName -ErrorAction SilentlyContinue }
        }
    }
}

if ($RemoveTrustedCert) {
    $certificateFile = Join-Path $PackageDirectory 'OpenDocsDevCert.cer'
    if (-not (Test-Path -LiteralPath $certificateFile)) { throw 'The package public certificate is required for explicit certificate removal.' }
    $Thumbprint = (New-Object Security.Cryptography.X509Certificates.X509Certificate2($certificateFile)).Thumbprint
    $location = if ($AllUsers) { 'LocalMachine' } else { 'CurrentUser' }
    $cert = Get-ChildItem "Cert:\$location\TrustedPublisher\$Thumbprint" -ErrorAction SilentlyContinue
    if ($cert) {
        Remove-Item $cert.PSPath -Force
        Write-Host "$removed trusted certificate from $location."
    }
}

Write-Host "Uninstall complete. Restart Office applications for changes to take effect."

exit 0
