# Run this ON THIS DEV MACHINE. Builds the selected add-in(s) in Release
# configuration, exports the public half of the manifest-signing certificate
# (never the private key), and stages everything into deploy\dist\ ready to
# zip up and copy to another computer.
#
# Usage:
#   .\package.ps1                  # packages Word, Excel, and PowerPoint
#   .\package.ps1 -App Word        # packages just Word
#   .\package.ps1 -OutDir C:\tmp\airchat-package

param(
    [ValidateSet('Word', 'Excel', 'PowerPoint', 'All')]
    [string]$App = 'All',
    [string]$OutDir = "$PSScriptRoot\dist"
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
$MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$Thumbprint = 'D7C0D2DECAE0E7D967DE0A2C1B5DBFF185A932FD'

if (-not (Test-Path $MSBuild)) {
    throw "MSBuild not found at $MSBuild - update the path in this script if Visual Studio is installed elsewhere."
}

$Apps = if ($App -eq 'All') { @('Word', 'Excel', 'PowerPoint') } else { @($App) }

if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# Export the public certificate only - the private key stays on this machine.
$certPath = Join-Path $OutDir 'AirchatOfficeDevCert.cer'
$cert = Get-ChildItem "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
if (-not $cert) { $cert = Get-ChildItem "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction SilentlyContinue }
if (-not $cert) {
    throw "Signing certificate $Thumbprint not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My on this machine."
}
Export-Certificate -Cert $cert -FilePath $certPath | Out-Null
Write-Host "Exported signing certificate (public key only) to $certPath"

foreach ($appName in $Apps) {
    $projName = "${appName}AiAddIn"
    $csproj = Join-Path $RepoRoot "$projName\$projName.csproj"
    if (-not (Test-Path $csproj)) { throw "Project file not found: $csproj" }

    Write-Host ""
    Write-Host "Building $projName (Release)..."
    & $MSBuild $csproj -t:Build -p:Configuration=Release -v:minimal
    if ($LASTEXITCODE -ne 0) { throw "$projName build failed (exit code $LASTEXITCODE)." }

    $srcBin = Join-Path $RepoRoot "$projName\bin\Release"
    $vstoFile = Join-Path $srcBin "$projName.vsto"
    if (-not (Test-Path $vstoFile)) { throw "$vstoFile missing after build - Release config may not be signing manifests correctly." }

    $destBin = Join-Path $OutDir $projName
    Copy-Item $srcBin $destBin -Recurse
    Write-Host "Staged $projName build output to $destBin"
}

Copy-Item "$PSScriptRoot\install.ps1" $OutDir -Force
Copy-Item "$PSScriptRoot\uninstall.ps1" $OutDir -Force

Write-Host ""
Write-Host "Package ready at: $OutDir"
Write-Host "Zip this folder, copy it to the target machine, then from inside it run:"
Write-Host "    powershell -ExecutionPolicy Bypass -File .\install.ps1"
