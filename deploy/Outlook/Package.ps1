[CmdletBinding()]
param(
    [string]$OutDir = (Join-Path (Split-Path $PSScriptRoot) ('dist\Outlook\'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))),
    [ValidateSet('x64','x86')][string]$PrerequisiteArchitecture='x64',
    [string]$PrerequisiteSource, [string]$CertificateThumbprint, [string]$MSBuildPath,
    [switch]$ScriptsOnly, [switch]$SkipBuild, [string]$PayloadDirectory, [string]$PublicCertificatePath
)
$ErrorActionPreference='Stop'
$synchronizer = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'deploy\tools\Sync-Products.ps1'
if (Test-Path -LiteralPath $synchronizer) { & $synchronizer }
& (Join-Path $PSScriptRoot 'common\package-engine.ps1') -App 'Outlook' -TemplateDirectory $PSScriptRoot -RepositoryRoot (Split-Path (Split-Path $PSScriptRoot)) -OutDir $OutDir -PrerequisiteArchitecture $PrerequisiteArchitecture -PrerequisiteSource $PrerequisiteSource -CertificateThumbprint $CertificateThumbprint -MSBuildPath $MSBuildPath -ScriptsOnly:$ScriptsOnly -SkipBuild:$SkipBuild -PayloadDirectory $PayloadDirectory -PublicCertificatePath $PublicCertificatePath

