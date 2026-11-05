[CmdletBinding()]
param([switch]$CurrentUser, [string]$InstallRoot, [version]$MinimumVersion)
$ErrorActionPreference = 'Stop'
try {
    . (Join-Path $PSScriptRoot 'common\status.ps1')
    if (-not $MinimumVersion) {
        $metadata = Get-Content (Join-Path $PSScriptRoot 'product.json') -Raw | ConvertFrom-Json
        if ($metadata.Version) { $MinimumVersion = [version]$metadata.Version }
    }
    $state = Get-OpenDocsProductState -App 'Excel' -CurrentUser:$CurrentUser -InstallRoot $InstallRoot -MinimumVersion $MinimumVersion
    $state | ConvertTo-Json -Depth 6
    if ($state.Ready) { exit 0 }; exit 1
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 2 }

