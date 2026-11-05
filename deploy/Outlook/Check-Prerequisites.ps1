[CmdletBinding()]
param([switch]$CurrentUser)
$ErrorActionPreference = 'Stop'
try {
    . (Join-Path $PSScriptRoot 'common\prerequisites.ps1')
    $state = @(Get-OpenDocsPrerequisiteSpecs | ForEach-Object { [pscustomobject]@{Name=$_.Name;Installed=(Test-OpenDocsPrerequisite -Id $_.Id -AllUsers (-not $CurrentUser))} })
    $state | ConvertTo-Json
    if (@($state | Where-Object { -not $_.Installed }).Count) { exit 1 }; exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 2 }

