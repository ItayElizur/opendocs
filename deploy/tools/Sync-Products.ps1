# Regenerate the self-contained support copies and standalone SCCM detectors after editing deploy/common.
$ErrorActionPreference = 'Stop'
$deployRoot = Split-Path $PSScriptRoot
$common = Join-Path $deployRoot 'common'
$prereq = [IO.File]::ReadAllText((Join-Path $common 'prerequisites.ps1'))
$status = [IO.File]::ReadAllText((Join-Path $common 'status.ps1'))
$status = [regex]::Replace($status, '(?m)^\. \(Join-Path \$PSScriptRoot ''prerequisites.ps1''\)\r?\n', '')
foreach ($app in @('Word','Excel','PowerPoint','Outlook')) {
    $directory = Join-Path $deployRoot $app
    Copy-Item -LiteralPath $common -Destination $directory -Recurse -Force
    $header = @'
# Standalone SCCM detection script. Paste the entire file into the PowerShell detection rule.
[CmdletBinding()]
param([switch]$CurrentUser, [string]$InstallRoot, [version]$MinimumVersion = '0.0.0.0')
$ErrorActionPreference = 'Stop'
'@
    $footer = @'
try {
    $state = Get-OpenDocsProductState -App '__APP__' -CurrentUser:$CurrentUser -InstallRoot $InstallRoot -MinimumVersion $MinimumVersion
    if ($state.Ready) { Write-Output ('OpenDocs __APP__ installed: '+$state.Version) }
    exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
'@
    [IO.File]::WriteAllText((Join-Path $directory 'Detect-SCCM.ps1'), $header+"`r`n"+$prereq+"`r`n"+$status+"`r`n"+$footer.Replace('__APP__',$app))
}
Write-Host 'Regenerated Word, Excel, PowerPoint and Outlook support files.'
