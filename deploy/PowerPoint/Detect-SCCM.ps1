# Standalone SCCM detection script. Paste the entire file into the PowerShell detection rule.
[CmdletBinding()]
param([switch]$CurrentUser, [string]$InstallRoot, [version]$MinimumVersion = '0.0.0.0')
$ErrorActionPreference = 'Stop'
# Shared by install.ps1, package.ps1 and Get-Prerequisites.ps1. Loading this file makes no changes.
function Get-OpenDocsArchitecture {
    $native = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    if ($native -eq 'ARM64') { throw 'This prerequisite package supports x86/x64 Windows, not ARM64.' }
    if ([Environment]::Is64BitOperatingSystem) { return 'x64' }
    return 'x86'
}

function Get-OpenDocsRegistryValue {
    param([string]$Hive = 'LocalMachine', [string]$View = 'Registry64', [string]$Key, [string]$Name)
    $base = $null; $subkey = $null
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$Hive, [Microsoft.Win32.RegistryView]::$View)
        $subkey = $base.OpenSubKey($Key)
        if ($subkey) { return $subkey.GetValue($Name) }
    } finally {
        if ($subkey) { $subkey.Dispose() }; if ($base) { $base.Dispose() }
    }
}

function Get-OpenDocsPrerequisiteSpecs {
    param([ValidateSet('x64','x86')][string]$Architecture = (Get-OpenDocsArchitecture))
    @(
        [pscustomobject]@{ Id='Net48'; Name='.NET Framework 4.8'; Installer='ndp48-x86-x64-allos-enu.exe'; Arguments='/q /norestart'; Url='https://go.microsoft.com/fwlink/?linkid=2088631' }
        [pscustomobject]@{ Id='Vsto'; Name='VSTO Runtime'; Installer='vstor_redist.exe'; Arguments='/q /norestart'; Url='https://go.microsoft.com/fwlink/?LinkId=158918' }
        [pscustomobject]@{ Id='WebView2'; Name='WebView2 Runtime'; Installer="MicrosoftEdgeWebView2RuntimeInstaller$($Architecture.ToUpperInvariant()).exe"; Arguments='/silent /install'; Url=$(if($Architecture -eq 'x64'){'https://go.microsoft.com/fwlink/?linkid=2124701'}else{'https://go.microsoft.com/fwlink/?linkid=2099617'}) }
    )
}

function Test-OpenDocsPrerequisite {
    param([ValidateSet('Net48','Vsto','WebView2')][string]$Id, [bool]$AllUsers = $true)
    $views = if ([Environment]::Is64BitOperatingSystem) { @('Registry64','Registry32') } else { @('Registry32') }
    foreach ($view in $views) {
        switch ($Id) {
            'Net48' {
                $release = Get-OpenDocsRegistryValue -View $view -Key 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release
                if ($release -and [int]$release -ge 528040) { return $true }
            }
            'Vsto' {
                $version = Get-OpenDocsRegistryValue -View $view -Key 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R' -Name Version
                $parsed = $null
                if ([version]::TryParse([string]$version, [ref]$parsed) -and $parsed -ge [version]'10.0') { return $true }
            }
            'WebView2' {
                $hives = if ($AllUsers) { @('LocalMachine') } else { @('LocalMachine','CurrentUser') }
                foreach ($hive in $hives) {
                    $version = Get-OpenDocsRegistryValue -Hive $hive -View $view -Key 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' -Name pv
                    $parsed = $null
                    if ([version]::TryParse([string]$version, [ref]$parsed) -and $parsed -ge [version]'131.0.2903.40') { return $true }
                }
            }
        }
    }
    return $false
}

function Assert-OpenDocsMicrosoftInstaller {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Offline installer missing: $Path. Run Get-Prerequisites.ps1 inside the product folder on the build machine and repackage." }
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or -not $signature.SignerCertificate -or $signature.SignerCertificate.Subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        throw "The prerequisite does not have a valid Microsoft signature: $Path"
    }
}

function Assert-OpenDocsPrerequisitePayload {
    param([string]$Directory, [object[]]$Specs)
    $manifestPath = Join-Path $Directory 'manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw "Offline prerequisite manifest missing: $manifestPath. Run Get-Prerequisites.ps1 and repackage." }
    $manifest = @(Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json)
    foreach ($spec in $Specs) {
        $path = Join-Path $Directory $spec.Installer
        Assert-OpenDocsMicrosoftInstaller $path
        $entry = @($manifest | Where-Object { $_.Name -eq $spec.Installer })
        if ($entry.Count -ne 1 -or $entry[0].Sha256 -notmatch '^[A-Fa-f0-9]{64}$' -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $entry[0].Sha256) {
            throw "Prerequisite SHA256 mismatch or missing manifest entry: $($spec.Installer)"
        }
    }
}

function Invoke-OpenDocsPrerequisiteProcess {
    param([string]$Path, [string]$Arguments)
    $process = Start-Process -FilePath $Path -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
    return $process.ExitCode
}

function Install-OpenDocsPrerequisites {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$Directory, [bool]$AllUsers = $true, [bool]$IsAdministrator, [ValidateSet('x64','x86')][string]$Architecture = (Get-OpenDocsArchitecture))
    $missing = @(Get-OpenDocsPrerequisiteSpecs $Architecture | Where-Object { -not (Test-OpenDocsPrerequisite -Id $_.Id -AllUsers $AllUsers) })
    if ($missing.Count -eq 0) { Write-Host 'All required runtimes are installed.'; return 0 }
    if ($WhatIfPreference) {
        foreach ($spec in $missing) { [void]$PSCmdlet.ShouldProcess($spec.Name, "Install local prerequisite $($spec.Installer)") }
        return 0
    }
    if (-not $IsAdministrator) { throw 'Missing runtimes require administrator rights, including with -CurrentUser. Install prerequisites from an elevated shell first, then run the per-user add-in installation as the intended user.' }
    # Validate ALL missing payloads before launching any installer.
    Assert-OpenDocsPrerequisitePayload -Directory $Directory -Specs $missing
    foreach ($spec in $missing) {
        if (-not $PSCmdlet.ShouldProcess($spec.Name, 'Install prerequisite silently without automatic restart')) { throw "Prerequisite installation declined: $($spec.Name)" }
        Write-Host "Installing $($spec.Name)..."
        $code = Invoke-OpenDocsPrerequisiteProcess -Path (Join-Path $Directory $spec.Installer) -Arguments $spec.Arguments
        if ($code -eq 3010 -or $code -eq 1641) {
            Write-Host 'Restart required. Add-in installation has not started. Restart the computer, then rerun install.ps1.' -ForegroundColor Yellow
            return 3010
        }
        if ($code -ne 0) { throw "$($spec.Name) installer failed with exit code $code. Add-in installation stopped." }
        # All prerequisites installed here are machine-wide, even for a per-user add-in.
        if (-not (Test-OpenDocsPrerequisite -Id $spec.Id -AllUsers $true)) { throw "$($spec.Name) was not detected after installation. Add-in installation stopped." }
    }
    return 0
}

# Read-only deployment checks. No Office application is launched.

function Get-OpenDocsInstallRoot {
    param([string]$InstallRoot, [switch]$CurrentUser)
    if ($InstallRoot) { return [IO.Path]::GetFullPath($InstallRoot) }
    if ($CurrentUser) { return (Join-Path $env:LOCALAPPDATA 'OpenDocs') }
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    return (Join-Path $programFiles 'OpenDocs')
}

function Get-OpenDocsProductState {
    param([ValidateSet('Word','Excel','PowerPoint','Outlook')][string]$App, [string]$InstallRoot, [switch]$CurrentUser, [version]$MinimumVersion)
    $root = Get-OpenDocsInstallRoot -InstallRoot $InstallRoot -CurrentUser:$CurrentUser
    $name = "${App}AiAddIn"
    $directory = Join-Path $root $name
    $manifestPath = [IO.Path]::GetFullPath((Join-Path $directory "$name.vsto"))
    $assemblyPath = Join-Path $directory "$name.dll"
    $problems = @()
    $manifestExists = Test-Path -LiteralPath $manifestPath -PathType Leaf
    $assemblyExists = Test-Path -LiteralPath $assemblyPath -PathType Leaf
    if (-not $manifestExists) { $problems += 'Deployment manifest missing' }
    if (-not $assemblyExists) { $problems += 'Add-in assembly missing' }
    $version = $null
    if ($assemblyExists) {
        $rawVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($assemblyPath).FileVersion
        [void][version]::TryParse([string]$rawVersion, [ref]$version)
    }
    if ($MinimumVersion -and (-not $version -or $version -lt $MinimumVersion)) { $problems += 'Installed version is older than required' }
    $hive = if ($CurrentUser) { 'CurrentUser' } else { 'LocalMachine' }
    $views = if ([Environment]::Is64BitOperatingSystem) { @('Registry64','Registry32') } else { @('Registry32') }
    $platform = Get-OpenDocsRegistryValue -Hive LocalMachine -View Registry64 -Key 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name Platform
    if (-not $CurrentUser) {
        if ($platform -eq 'x64') { $views = @('Registry64') }
        elseif ($platform -eq 'x86') { $views = @('Registry32') }
    }
    $registrations = @()
    foreach ($view in $views) {
        $key = "Software\Microsoft\Office\$App\Addins\$name"
        $value = Get-OpenDocsRegistryValue -Hive $hive -View $view -Key $key -Name Manifest
        $load = Get-OpenDocsRegistryValue -Hive $hive -View $view -Key $key -Name LoadBehavior
        $matches = $false
        if ($value) {
            try {
                $uriText = [string]$value -replace '\|vstolocal$', ''
                $uri = [uri]$uriText
                $matches = $uri.IsFile -and [string]::Equals([IO.Path]::GetFullPath($uri.LocalPath), $manifestPath, [StringComparison]::OrdinalIgnoreCase)
            } catch { $matches = $false }
            $registrations += [pscustomobject]@{ Hive=$hive; View=$view; Manifest=[string]$value; LoadBehavior=$load; ExpectedPath=$matches }
        }
    }
    $registered = @($registrations | Where-Object { $_.ExpectedPath -and $_.LoadBehavior -eq 3 }).Count -gt 0
    if (-not $registered) { $problems += 'Correct enabled Office registration missing' }
    if (-not $CurrentUser) {
        foreach ($view in $views) {
            $userValue = Get-OpenDocsRegistryValue -Hive CurrentUser -View $view -Key "Software\Microsoft\Office\$App\Addins\$name" -Name Manifest
            if ($userValue) { $problems += 'This account has a per-user registration that may override the machine registration'; break }
        }
    }
    $runtimes = @(Get-OpenDocsPrerequisiteSpecs | ForEach-Object {
        [pscustomobject]@{ Name=$_.Name; Installed=(Test-OpenDocsPrerequisite -Id $_.Id -AllUsers (-not $CurrentUser)) }
    })
    $runtimeReady = @($runtimes | Where-Object { -not $_.Installed }).Count -eq 0
    if (-not $runtimeReady) { $problems += 'Required runtime missing or outdated' }
    $installed = $manifestExists -and $assemblyExists -and $registered -and (-not $MinimumVersion -or ($version -and $version -ge $MinimumVersion))
    [pscustomobject]@{
        Product=$App; Scope=$(if($CurrentUser){'CurrentUser'}else{'AllUsers'}); Installed=[bool]$installed
        Ready=([bool]$installed -and $runtimeReady -and $problems.Count -eq 0)
        Version=$(if($version){$version.ToString()}else{$null}); InstallDirectory=$directory
        Registrations=$registrations; Prerequisites=$runtimes; Problems=$problems
        Note='File/registry/runtime checks only; actual loading inside Office and certificate trust are not verified.'
    }
}

try {
    $state = Get-OpenDocsProductState -App 'PowerPoint' -CurrentUser:$CurrentUser -InstallRoot $InstallRoot -MinimumVersion $MinimumVersion
    if ($state.Ready) { Write-Output ('OpenDocs PowerPoint installed: '+$state.Version) }
    exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }