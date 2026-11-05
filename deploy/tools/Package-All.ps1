[CmdletBinding()]
param(
    [ValidateSet('Word','Excel','PowerPoint','Outlook','All')][string]$App='All',
    [string]$OutDir=(Join-Path (Split-Path $PSScriptRoot) ('dist\'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))),
    [ValidateSet('x64','x86')][string]$PrerequisiteArchitecture='x64',
    [string]$PrerequisiteSource, [string]$CertificateThumbprint, [string]$MSBuildPath, [switch]$ScriptsOnly
)
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'Sync-Products.ps1')
$products = if($App -eq 'All'){@('Word','Excel','PowerPoint','Outlook')}else{@($App)}
foreach($product in $products) {
    & (Join-Path (Split-Path $PSScriptRoot) "$product\Package.ps1") -OutDir (Join-Path $OutDir $product) -PrerequisiteArchitecture $PrerequisiteArchitecture -PrerequisiteSource $PrerequisiteSource -CertificateThumbprint $CertificateThumbprint -MSBuildPath $MSBuildPath -ScriptsOnly:$ScriptsOnly
}
