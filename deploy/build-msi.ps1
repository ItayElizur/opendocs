# Run this ON THIS DEV MACHINE. Builds OpenDocs.msi: a machine-wide installer for all four add-ins, with a
# policy.json of your choice built in. Needs everything package.ps1 needs, plus WiX Toolset 5:
#   dotnet tool install --global wix --version 5.0.2
#
# Usage:
#   .\build-msi.ps1                                   # version from the latest vX.Y.Z git tag, shipped policy.json
#   .\build-msi.ps1 -Version 0.3.0                    # explicit version
#   .\build-msi.ps1 -PolicyFile C:\it\policy.json     # build in your own policy
#   .\build-msi.ps1 -SkipPackage                      # reuse deploy\dist from an earlier package.ps1 run
#
# Output: deploy\msi\OpenDocs-<version>.msi. See build-msi.ps1.md for how the MSI is put together.

param(
    [string]$Version,
    [string]$PolicyFile = "$PSScriptRoot\policy.json",
    [switch]$SkipPackage,
    [string]$OutDir = "$PSScriptRoot\msi"
)

$ErrorActionPreference = 'Stop'
$DistDir = "$PSScriptRoot\dist"
$Apps = @('Word', 'Excel', 'PowerPoint', 'Outlook')

$wix = Get-Command wix -ErrorAction SilentlyContinue
if (-not $wix -and (Test-Path "$env:USERPROFILE\.dotnet\tools\wix.exe")) { $wix = Get-Item "$env:USERPROFILE\.dotnet\tools\wix.exe" }
if (-not $wix) { throw "WiX Toolset not found. Install it with: dotnet tool install --global wix --version 5.0.2" }
$wixExe = if ($wix -is [System.Management.Automation.CommandInfo]) { $wix.Source } else { $wix.FullName }

if (-not $Version) {
    $tag = git -C $PSScriptRoot describe --tags --abbrev=0 --match 'v[0-9]*.[0-9]*.[0-9]*' 2>$null
    if (-not $tag) { throw "No vX.Y.Z git tag found - pass -Version X.Y.Z." }
    $Version = $tag.TrimStart('v')
}
if ($Version -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,5}$') { throw "Version must be X.Y.Z (MSI limits: X and Y up to 255, Z up to 65535), got '$Version'." }

$PolicyFile = (Resolve-Path $PolicyFile).Path

if (-not $SkipPackage) {
    & "$PSScriptRoot\package.ps1" -OutDir $DistDir
    if (-not $?) { throw "package.ps1 failed." }
}
foreach ($app in $Apps) {
    if (-not (Test-Path "$DistDir\${app}AiAddIn\${app}AiAddIn.vsto")) { throw "$DistDir\${app}AiAddIn is missing or incomplete - run without -SkipPackage." }
}

# Read the policy with the add-ins' own parser (OfficeAi.Shared.dll from the package), so a typo is caught here and
# not on every user's machine: an unknown app or mode name, or broken JSON, switches every mode off in every app.
# The resolve handler stands in for the binding redirects Office applies from the add-in's .dll.config.
$libDir = "$DistDir\ExcelAiAddIn"
[AppDomain]::CurrentDomain.add_AssemblyResolve({
    param($s, $e)
    $p = Join-Path $libDir ((New-Object Reflection.AssemblyName $e.Name).Name + '.dll')
    if (Test-Path $p) { [Reflection.Assembly]::LoadFrom($p) } else { $null }
})
Add-Type -Path "$libDir\OfficeAi.Shared.dll"
$rules = [OfficeAi.Shared.ModePolicy]::Load($PolicyFile)
$allOff = @($Apps | Where-Object { $rules.Disabled[$_].Count -eq 4 })
if ($allOff.Count -eq $Apps.Count) {
    throw "$PolicyFile switches every editing mode off in every app - a misspelled app or mode name, or invalid JSON (comments and trailing commas are fine). Fix it and run again."
}
Write-Host "Policy $PolicyFile :"
foreach ($app in $Apps) {
    $off = @($rules.Disabled[$app]) -join ', '
    $ask = @($rules.Approvals[$app].GetEnumerator() | ForEach-Object { "$($_.Key) in [" + (@($_.Value) -join ', ') + "]" }) -join '; '
    Write-Host ("  {0,-10} modes off: {1,-40} asks: {2}" -f $app, $(if ($off) { $off } else { '-' }), $(if ($ask) { $ask } else { '-' }))
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$msi = Join-Path $OutDir "OpenDocs-$Version.msi"
& $wixExe build "$PSScriptRoot\OpenDocs.wxs" -arch x64 -o $msi `
    -d "Version=$Version" -d "DistDir=$DistDir" -d "PolicyFile=$PolicyFile"
if ($LASTEXITCODE -ne 0) { throw "wix build failed (exit code $LASTEXITCODE)." }

Write-Host ""
Write-Host "MSI ready: $msi"
Write-Host "Install silently (elevated):  msiexec /i `"$msi`" /qn /l*v install.log"
Write-Host "See docs/it-deployment.md for rollout, policy and uninstall."
