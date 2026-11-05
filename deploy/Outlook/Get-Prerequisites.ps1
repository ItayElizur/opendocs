[CmdletBinding()]
param([ValidateSet('x64','x86')][string]$Architecture='x64')
& (Join-Path $PSScriptRoot 'common\Get-Prerequisites.ps1') -Architecture $Architecture -OutDir (Join-Path $PSScriptRoot 'prerequisites')

