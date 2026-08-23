# Run this ON THE TARGET MACHINE, from inside the package folder produced
# by package.ps1 (it must sit alongside AirchatOfficeDevCert.cer and the
# per-app WordAiAddIn/ExcelAiAddIn/PowerPointAiAddIn build-output folders).
#
# Usage:
#   .\install.ps1                  # installs Word, Excel, and PowerPoint
#   .\install.ps1 -App Word        # installs just Word
#
# No internet access is required or used by this script.

param(
    [ValidateSet('Word', 'Excel', 'PowerPoint', 'All')]
    [string]$App = 'All',
    [string]$InstallRoot = "$env:LOCALAPPDATA\AirchatOffice"
)

$ErrorActionPreference = 'Stop'
$PackageDir = $PSScriptRoot
$Apps = if ($App -eq 'All') { @('Word', 'Excel', 'PowerPoint') } else { @($App) }
$OfficeApps = @{ Word = 'Word'; Excel = 'Excel'; PowerPoint = 'PowerPoint' }

function Test-Prerequisites {
    $problems = @()

    # .NET Framework 4.8 = release key >= 528040 (528449 for 4.8.1, also fine).
    $ndpKey = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    $release = (Get-ItemProperty -Path $ndpKey -Name Release -ErrorAction SilentlyContinue).Release
    if (-not $release -or $release -lt 528040) {
        $problems += ".NET Framework 4.8 (or later) does not appear to be installed (found release key: $release). Install the offline .NET Framework 4.8 redistributable before continuing."
    }

    # VSTO Runtime - presence of its registry uninstall/setup key is a reasonable proxy.
    $vstoKeyPaths = @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
    )
    $vstoFound = $false
    foreach ($p in $vstoKeyPaths) { if (Test-Path $p) { $vstoFound = $true } }
    if (-not $vstoFound) {
        $problems += "VSTO 2010 Runtime does not appear to be installed. Install the offline vstor_redist.exe before continuing - without it, Office will not load any VSTO add-in."
    }

    # WebView2 Runtime - the well-known Edge WebView2 client GUID's registry key.
    $webview2KeyPaths = @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
        'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
        'HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
    )
    $webview2Found = $false
    foreach ($p in $webview2KeyPaths) { if (Test-Path $p) { $webview2Found = $true } }
    if (-not $webview2Found) {
        $problems += "WebView2 Runtime does not appear to be installed. Install the WebView2 Evergreen Standalone (or Fixed Version) installer before continuing - the add-in's whole chat panel is a WebView2 control and will fail to load without it."
    }

    return $problems
}

Write-Host "Checking prerequisites..."
$problems = Test-Prerequisites
if ($problems.Count -gt 0) {
    Write-Host ""
    Write-Host "WARNING: potential missing prerequisites detected:" -ForegroundColor Yellow
    foreach ($p in $problems) { Write-Host "  - $p" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "Continuing with installation anyway - these checks can have false positives on some Windows/Office builds - but if the add-in fails to load afterward, install the flagged prerequisite(s) first." -ForegroundColor Yellow
    Write-Host ""
}

$certPath = Join-Path $PackageDir 'AirchatOfficeDevCert.cer'
if (-not (Test-Path $certPath)) { throw "Certificate file not found next to install.ps1: $certPath" }

Write-Host "Trusting the add-in's signing certificate (current user only, no admin needed)..."
Import-Certificate -FilePath $certPath -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null

foreach ($appName in $Apps) {
    $projName = "${appName}AiAddIn"
    $srcDir = Join-Path $PackageDir $projName
    if (-not (Test-Path $srcDir)) { throw "Package is missing $projName - re-run package.ps1 with -App $appName or -App All." }

    $destDir = Join-Path $InstallRoot $projName
    Write-Host "Installing $projName to $destDir ..."
    if (Test-Path $destDir) { Remove-Item $destDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item "$srcDir\*" $destDir -Recurse -Force

    $vstoPath = Join-Path $destDir "$projName.vsto"
    if (-not (Test-Path $vstoPath)) { throw "$vstoPath not found after copy - package looks incomplete." }
    $manifestUri = "file:///" + ($vstoPath -replace '\\', '/') + "|vstolocal"

    $regKey = "HKCU:\Software\Microsoft\Office\$($OfficeApps[$appName])\Addins\$projName"
    Write-Host "Registering $projName at $regKey ..."
    New-Item -Path $regKey -Force | Out-Null
    Set-ItemProperty -Path $regKey -Name 'Description' -Value $projName
    Set-ItemProperty -Path $regKey -Name 'FriendlyName' -Value $projName
    Set-ItemProperty -Path $regKey -Name 'LoadBehavior' -Value 3 -Type DWord
    Set-ItemProperty -Path $regKey -Name 'Manifest' -Value $manifestUri

    Write-Host "$projName installed and registered."
}

Write-Host ""
Write-Host "Done. Start (or restart) the corresponding Office application(s) to load the add-in(s)."
Write-Host "If an app doesn't show the panel: File > Options > Add-ins > Manage: COM Add-ins > Go,"
Write-Host "and confirm the add-in is checked and not listed under Disabled Items."
