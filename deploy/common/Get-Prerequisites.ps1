# Run on the BUILD machine. Only downloads/verifies installers; does not execute them.
[CmdletBinding()]
param(
    [ValidateSet('x64','x86')][string]$Architecture = 'x64',
    [string]$OutDir = (Join-Path $PSScriptRoot 'prerequisites')
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
. (Join-Path $PSScriptRoot 'prerequisites.ps1')
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$entries = @()
foreach ($spec in (Get-OpenDocsPrerequisiteSpecs $Architecture)) {
    $path = Join-Path $OutDir $spec.Installer
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host "Downloading $($spec.Name) from Microsoft..."
        # Download to a temporary name, so an interrupted download cannot be reused as a complete installer.
        $download = $path + '.download'
        Invoke-WebRequest -Uri $spec.Url -OutFile $download -UseBasicParsing
        Assert-OpenDocsMicrosoftInstaller $download
        Move-Item -LiteralPath $download -Destination $path -Force
    }
    Assert-OpenDocsMicrosoftInstaller $path
    $entries += [pscustomobject]@{ Name=$spec.Installer; Sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; Source=$spec.Url }
}
$entries | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutDir 'manifest.json') -Encoding UTF8
Write-Host "Offline prerequisites ready for $Architecture in $OutDir. They have NOT been installed."
