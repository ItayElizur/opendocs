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
