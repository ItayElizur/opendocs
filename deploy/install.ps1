# Run this ON THE TARGET MACHINE, from inside the package folder produced
# by package.ps1 (it must sit alongside OpenDocsDevCert.cer and the
# per-app WordAiAddIn/ExcelAiAddIn/PowerPointAiAddIn build-output folders).
#
# Usage (default: for every user of this computer, from an ELEVATED PowerShell):
#   .\install.ps1                      # installs Word, Excel, PowerPoint, and Outlook for all users
#   .\install.ps1 -App Word            # installs just Word
#   .\install.ps1 -CurrentUser         # this user only; no admin needed
#   .\install.ps1 -WhatIf              # list what it would change and change nothing (works without elevation)
#
# All users (default): files in Program Files\OpenDocs (everyone can run them, only administrators can change them),
#   registered under HKLM.
# -CurrentUser: files in %LOCALAPPDATA%\OpenDocs, registered under HKCU.
# -TrustCertificate: also trust OpenDocsDevCert.cer in Trusted Publishers (computer, or current user with -CurrentUser). Not needed in
#   the configurations tested - the add-ins load without it - so it is off by default; use it only if a policy blocks the add-ins.
# Run uninstall.ps1 with the same switch to remove it.
#
# No internet access is required or used by this script.

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Word', 'Excel', 'PowerPoint', 'Outlook', 'All')]
    [string]$App = 'All',
    [string]$InstallRoot,
    [switch]$CurrentUser,
    [switch]$TrustCertificate
)

$ErrorActionPreference = 'Stop'
$AllUsers = -not $CurrentUser
$PackageDir = $PSScriptRoot
$Apps = if ($App -eq 'All') { @('Word', 'Excel', 'PowerPoint', 'Outlook') } else { @($App) }
$OfficeApps = @{ Word = 'Word'; Excel = 'Excel'; PowerPoint = 'PowerPoint'; Outlook = 'Outlook' }
$ProcessNames = @{ Word = 'WINWORD'; Excel = 'EXCEL'; PowerPoint = 'POWERPNT'; Outlook = 'OUTLOOK' }

function Test-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 'x64' or 'x86' for the installed Office, or $null when it cannot be told. A 32-bit Office on 64-bit Windows reads
# add-in registrations from HKLM\SOFTWARE\WOW6432Node, a 64-bit Office from HKLM\SOFTWARE.
function Get-OfficeBitness {
    $c2r = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name Platform -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.Platform) { return $c2r.Platform }
    foreach ($ver in @('16.0', '15.0')) {
        foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
            $outlook = Get-ItemProperty -Path "$root\$ver\Outlook" -Name Bitness -ErrorAction SilentlyContinue
            if ($outlook -and $outlook.Bitness) { return $outlook.Bitness }
        }
    }
    return $null
}

# The registry roots the add-ins are registered under: HKCU for a per-user install; for all users HKLM, in the view(s)
# Office reads (both when the bitness cannot be determined - the unused one is harmless).
function Get-OfficeRegistryRoots {
    if (-not $AllUsers) { return @('HKCU:\Software\Microsoft\Office') }
    if (-not [Environment]::Is64BitOperatingSystem) { return @('HKLM:\SOFTWARE\Microsoft\Office') }
    $bitness = Get-OfficeBitness
    if ($bitness -eq 'x64') { return @('HKLM:\SOFTWARE\Microsoft\Office') }
    if ($bitness -eq 'x86') { return @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office') }
    return @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')
}

function Test-CertTrusted {
    param([string]$Thumbprint, [string[]]$Locations)
    foreach ($location in $Locations) {
        if (Test-Path "Cert:\$location\TrustedPublisher\$Thumbprint") { return $true }
    }
    return $false
}

if ($AllUsers -and -not (Test-Administrator)) {
    $needAdmin = "A machine-wide install writes to Program Files, HKLM and the computer's certificate store, so it needs an elevated PowerShell (right-click PowerShell, Run as administrator). Use -CurrentUser to install for the current user only, without admin."
    if ($WhatIfPreference) { Write-Host "NOTE: $needAdmin (continuing only because of -WhatIf; nothing will be changed)." -ForegroundColor Yellow }
    else { throw $needAdmin }
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

# Office keeps an installed add-in's files locked while it runs, so replacing them would fail half way through. For an
# all-users install any user's session counts; otherwise only this user's.
$mySession = (Get-Process -Id $PID).SessionId
$running = @()
foreach ($appName in $Apps) {
    $procs = @(Get-Process -Name $ProcessNames[$appName] -ErrorAction SilentlyContinue)
    if (-not $AllUsers) { $procs = @($procs | Where-Object { $_.SessionId -eq $mySession }) }
    if ($procs.Count -gt 0) { $running += $appName }
}
if ($running.Count -gt 0) {
    $closeFirst = "Close " + ($running -join ', ') + " first: Office keeps the add-in's files locked while it is running."
    if ($WhatIfPreference) { Write-Host "NOTE: $closeFirst" -ForegroundColor Yellow }
    else { throw $closeFirst }
}

# Files that arrive via a zip transfer from another machine are commonly
# flagged with a "Mark of the Web" (Zone.Identifier) by Windows, which can
# make .NET/Office refuse to load the assemblies with a generic "runtime
# error occurred during the loading of the COM add-in" - unblock everything
# in the package up front so this never bites silently.
Write-Host "Unblocking package files (clears any Mark-of-the-Web from a zip transfer)..."
Get-ChildItem -Path $PackageDir -Recurse | Unblock-File -ErrorAction SilentlyContinue

if ($TrustCertificate) {
    $certPath = Join-Path $PackageDir 'OpenDocsDevCert.cer'
    if (-not (Test-Path $certPath)) { throw "Certificate file not found next to install.ps1: $certPath" }
    $thumbprint = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)).Thumbprint

    # A certificate trusted for the whole computer also covers a per-user install; the reverse is not true, because
    # CurrentUser belongs to whichever account ran this script.
    $certLocation = if ($AllUsers) { 'LocalMachine' } else { 'CurrentUser' }
    $coveredBy = if ($AllUsers) { @('LocalMachine') } else { @('CurrentUser', 'LocalMachine') }
    if (Test-CertTrusted -Thumbprint $thumbprint -Locations $coveredBy) {
        Write-Host "The add-in's signing certificate is already trusted - skipping the import."
    }
    else {
        $who = if ($AllUsers) { "this computer, all users" } else { "the current user only, no admin needed" }
        Write-Host "Trusting the add-in's signing certificate ($who)..."
        try {
            Import-Certificate -FilePath $certPath -CertStoreLocation "Cert:\$certLocation\TrustedPublisher" | Out-Null
        }
        catch {
            if (Test-CertTrusted -Thumbprint $thumbprint -Locations $coveredBy) {
                Write-Host "WARNING: importing the certificate reported '$($_.Exception.Message)', but it is trusted now - continuing." -ForegroundColor Yellow
            }
            else {
                # Not fatal: in the configurations tested (default Office, and Office set to require signed add-ins) the add-in
                # loads without the certificate being trusted, so a refused import should not stop the install.
                Write-Host ("WARNING: could not add the signing certificate to Cert:\$certLocation\TrustedPublisher: $($_.Exception.Message)") -ForegroundColor Yellow
                Write-Host "Continuing - the add-in normally loads without it. If a security policy on this machine does block the add-in, trust OpenDocsDevCert.cer:" -ForegroundColor Yellow
                Write-Host "  - 'Access denied' on one machine while others work usually means a Group Policy or permission restriction on Trusted Publishers" -ForegroundColor Yellow
                Write-Host "    (run 'gpresult /h gp.html' to see which policies apply); IT can deploy the certificate through Group Policy," -ForegroundColor Yellow
                Write-Host "  - or run this script from an elevated PowerShell without -CurrentUser to trust it for the whole computer." -ForegroundColor Yellow
                Write-Host "  - A certificate imported by hand from an elevated shell only lands in the store of the account running that shell." -ForegroundColor Yellow
            }
        }
    }
}
else {
    Write-Host "Not trusting the signing certificate (the add-ins load without it; pass -TrustCertificate if a policy blocks them)."
}

$registryRoots = Get-OfficeRegistryRoots

foreach ($appName in $Apps) {
    $projName = "${appName}AiAddIn"
    $srcDir = Join-Path $PackageDir $projName
    if (-not (Test-Path $srcDir)) { throw "Package is missing $projName - re-run package.ps1 with -App $appName or -App All." }
    if (-not (Test-Path (Join-Path $srcDir "$projName.vsto"))) { throw "$projName.vsto not found in the package - package looks incomplete." }

    $destDir = Join-Path $InstallRoot $projName
    Write-Host "Installing $projName to $destDir ..."
    if (Test-Path $destDir) { Remove-Item $destDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item "$srcDir\*" $destDir -Recurse -Force

    $vstoPath = Join-Path $destDir "$projName.vsto"
    $manifestUri = "file:///" + ($vstoPath -replace '\\', '/') + "|vstolocal"

    foreach ($officeRoot in $registryRoots) {
        $regKey = "$officeRoot\$($OfficeApps[$appName])\Addins\$projName"
        Write-Host "Registering $projName at $regKey ..."
        New-Item -Path $regKey -Force | Out-Null
        if ($WhatIfPreference) {
            # the key was not really created, so setting values on it would fail
            Write-Host "What if: setting Description, FriendlyName, LoadBehavior=3 and Manifest=$manifestUri on $regKey"
        }
        else {
            Set-ItemProperty -Path $regKey -Name 'Description' -Value $projName
            Set-ItemProperty -Path $regKey -Name 'FriendlyName' -Value $projName
            Set-ItemProperty -Path $regKey -Name 'LoadBehavior' -Value 3 -Type DWord
            Set-ItemProperty -Path $regKey -Name 'Manifest' -Value $manifestUri
        }
    }

    # Two registrations of the same add-in (this account's own plus the machine-wide one) would load it twice.
    if ($AllUsers -and (Test-Path "HKCU:\Software\Microsoft\Office\$($OfficeApps[$appName])\Addins\$projName")) {
        Write-Host "NOTE: this account also has a per-user registration of $projName. Run uninstall.ps1 -CurrentUser so it is not registered twice." -ForegroundColor Yellow
    }

    # Outlook (2013+) auto-disables add-ins it deems slow to start or that
    # threw during a previous load. Whitelist this one and clear any stale
    # disabled/crashing entries so a single bad dev launch doesn't stick.
    # That list lives in each user's own HKCU hive, which an elevated machine-wide
    # install cannot reach, so it is only done for -CurrentUser.
    if ($appName -eq 'Outlook' -and -not $AllUsers) {
        foreach ($ver in @('16.0', '15.0')) {
            $res = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Resiliency"
            $dna = "$res\DoNotDisableAddinList"
            New-Item -Path $dna -Force | Out-Null
            if ($WhatIfPreference) { Write-Host "What if: whitelisting $projName in $dna" }
            else { Set-ItemProperty -Path $dna -Name $projName -Value 1 -Type DWord }
            foreach ($bucket in @('DisabledItems', 'CrashingAddinList')) {
                $bpath = "$res\$bucket"
                if (Test-Path $bpath) {
                    Get-Item $bpath | Select-Object -ExpandProperty Property | ForEach-Object {
                        Remove-ItemProperty -Path $bpath -Name $_ -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        Write-Host "Whitelisted $projName against Outlook's slow/crashing add-in resiliency."
    }
    elseif ($appName -eq 'Outlook') {
        Write-Host "NOTE: Outlook's per-user 'slow add-in' resiliency list is not set by a machine-wide install. If Outlook disables $projName after a slow start or crash, re-enable it under File > Options > Add-ins > Manage: Disabled Items."
    }

    Write-Host "$projName installed and registered."
}

Write-Host ""
Write-Host "Done. Start (or restart) the corresponding Office application(s) to load the add-in(s)."
if ($AllUsers) {
    Write-Host "Installed for every user of this computer; each user gets the add-in the next time they start Office."
}
Write-Host "If an app doesn't show the panel: File > Options > Add-ins > Manage: COM Add-ins > Go,"
Write-Host "and confirm the add-in is checked and not listed under Disabled Items."
Write-Host ""
Write-Host "If LoadBehavior shows 'Not Loaded' with a runtime error: this script already"
Write-Host "unblocked the files, so if it still fails, get the real exception from"
Write-Host "Event Viewer > Windows Logs > Application (look for a '.NET Runtime' or"
Write-Host "'Application Error' entry timestamped right when you opened the Office app)."
