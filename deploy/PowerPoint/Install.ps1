[CmdletBinding(SupportsShouldProcess)]
param([switch]$CurrentUser, [string]$InstallRoot)
$ErrorActionPreference = 'Stop'
try {
    & (Join-Path $PSScriptRoot 'common\install-engine.ps1') -App 'PowerPoint' -PackageDirectory $PSScriptRoot -CurrentUser:$CurrentUser -InstallRoot $InstallRoot -WhatIf:$WhatIfPreference
    exit $LASTEXITCODE
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1603 }

