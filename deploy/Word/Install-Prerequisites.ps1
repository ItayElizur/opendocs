[CmdletBinding(SupportsShouldProcess)]
param()
$ErrorActionPreference = 'Stop'
try {
    . (Join-Path $PSScriptRoot 'common\prerequisites.ps1')
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $admin = ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $code = Install-OpenDocsPrerequisites -Directory (Join-Path $PSScriptRoot 'prerequisites') -AllUsers $true -IsAdministrator $admin -WhatIf:$WhatIfPreference
    exit $code
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1603 }

