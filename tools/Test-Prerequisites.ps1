# No installers, certificate imports, downloads or registry writes are performed by these tests.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot) 'deploy\common\prerequisites.ps1')
$script:passed = 0
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message }; $script:passed++ }
function Assert-Throws([scriptblock]$Action, [string]$Match) {
    $caught = $false
    try { & $Action | Out-Null } catch { if ($_.Exception.Message -notlike "*$Match*") { throw }; $caught = $true }
    Assert $caught "Expected failure containing $Match"
}
$script:registry = @{}
function Get-OpenDocsRegistryValue {
    param($Hive='LocalMachine',$View='Registry64',$Key,$Name)
    return $script:registry["$Hive|$Name"]
}
Assert (-not (Test-OpenDocsPrerequisite Net48)) 'Missing .NET incorrectly detected.'
$script:registry['LocalMachine|Release'] = 533320
Assert (Test-OpenDocsPrerequisite Net48) '.NET 4.8.1 not accepted.'
$script:registry['LocalMachine|Release'] = 461808
Assert (-not (Test-OpenDocsPrerequisite Net48)) '.NET 4.7.2 incorrectly accepted.'
$script:registry['LocalMachine|Version'] = '10.0.60910'
Assert (Test-OpenDocsPrerequisite Vsto) 'VSTO version not accepted.'
$script:registry['LocalMachine|Version'] = ''
Assert (-not (Test-OpenDocsPrerequisite Vsto)) 'Empty VSTO version accepted.'
$script:registry['CurrentUser|pv'] = '153.0.0.0'
Assert (Test-OpenDocsPrerequisite WebView2 -AllUsers $false) 'Per-user WebView2 missing.'
Assert (-not (Test-OpenDocsPrerequisite WebView2 -AllUsers $true)) 'Per-user WebView2 accepted for all users.'
$script:registry['LocalMachine|pv'] = '0.0.0.0'
Assert (-not (Test-OpenDocsPrerequisite WebView2)) 'Zero WebView2 version accepted.'
$script:registry['LocalMachine|pv'] = '120.0.0.0'
Assert (-not (Test-OpenDocsPrerequisite WebView2)) 'Old WebView2 version accepted.'
$script:registry['LocalMachine|pv'] = '153.0.0.0'
Assert (Test-OpenDocsPrerequisite WebView2) 'Machine WebView2 version missing.'
Assert ((Get-OpenDocsPrerequisiteSpecs x86)[2].Installer -eq 'MicrosoftEdgeWebView2RuntimeInstallerX86.exe') 'Wrong x86 installer.'
Assert ((Get-OpenDocsPrerequisiteSpecs x64)[2].Installer -eq 'MicrosoftEdgeWebView2RuntimeInstallerX64.exe') 'Wrong x64 installer.'

# Replace external boundaries only: execute the real prerequisite orchestration with simulated machine state.
$script:installed = @{}
$script:launches = @()
$script:code = 0
$script:validateCount = 0
$script:validationError = $false
$script:detectAfterInstall = $true
function Test-OpenDocsPrerequisite { param($Id,$AllUsers) return [bool]$script:installed[$Id] }
function Assert-OpenDocsPrerequisitePayload {
    param($Directory,$Specs)
    $script:validateCount++
    if ($script:validationError) { throw 'Missing payload' }
}
function Invoke-OpenDocsPrerequisiteProcess {
    param($Path,$Arguments)
    $spec = Get-OpenDocsPrerequisiteSpecs x64 | Where-Object { $_.Installer -eq (Split-Path $Path -Leaf) }
    $script:launches += $spec.Id
    if ($script:detectAfterInstall) { $script:installed[$spec.Id] = $true }
    return $script:code
}
Assert ((Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $false -WhatIf) -eq 0) 'WhatIf failed.'
Assert ($script:launches.Count -eq 0 -and $script:validateCount -eq 0) 'WhatIf performed side effects.'
Assert-Throws { Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $false } 'administrator rights'
$script:validationError = $true
Assert-Throws { Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $true } 'Missing payload'
Assert ($script:launches.Count -eq 0) 'Installer ran before payload validation.'
$script:validationError = $false
Assert ((Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $true) -eq 0) 'Runtime installation failed.'
Assert (($script:launches -join ',') -eq 'Net48,Vsto,WebView2') 'Wrong dependency install order.'
$script:launches = @(); $script:validateCount = 0
Assert ((Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $false) -eq 0) 'Installed runtimes required elevation.'
Assert ($script:launches.Count -eq 0 -and $script:validateCount -eq 0) 'Installed runtimes were reinstalled.'
foreach ($code in @(3010,1641)) {
    $script:installed = @{}; $script:launches = @(); $script:code = $code
    Assert ((Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $true) -eq 3010) 'Wrong reboot code.'
    Assert ($script:launches.Count -eq 1) 'Continued installing after restart was required.'
}
$script:installed = @{}; $script:launches = @(); $script:code = 1603
Assert-Throws { Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $true } 'exit code 1603'
Assert ($script:launches.Count -eq 1) 'Continued after failed installer.'
$script:installed = @{}; $script:launches = @(); $script:code = 0; $script:detectAfterInstall = $false
Assert-Throws { Install-OpenDocsPrerequisites -Directory C:\Unused -Architecture x64 -IsAdministrator $true } 'not detected after installation'
# Exercise the real package-integrity gate with a mocked signature provider only.
. (Join-Path (Split-Path $PSScriptRoot) 'deploy\common\prerequisites.ps1')
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('OpenDocsPrerequisiteTests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture | Out-Null
$file = Join-Path $fixture 'runtime.exe'
$manifestPath = Join-Path $fixture 'manifest.json'
$script:signer = 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US'
$script:signatureStatus = 'Valid'
function Get-AuthenticodeSignature { param($LiteralPath) return [pscustomobject]@{Status=$script:signatureStatus;SignerCertificate=[pscustomobject]@{Subject=$script:signer}} }
try {
    $spec = @([pscustomobject]@{Installer='runtime.exe'})
    Assert-Throws { Assert-OpenDocsPrerequisitePayload -Directory $fixture -Specs $spec } 'manifest missing'
    Set-Content -LiteralPath $file -Value 'test fixture, never executed'
    @([pscustomobject]@{Name='runtime.exe';Sha256=(Get-FileHash $file).Hash}) | ConvertTo-Json | Set-Content -LiteralPath $manifestPath
    Assert-OpenDocsPrerequisitePayload -Directory $fixture -Specs $spec
    Assert $true 'Valid fixture rejected.'
    Add-Content -LiteralPath $file -Value 'changed'
    Assert-Throws { Assert-OpenDocsPrerequisitePayload -Directory $fixture -Specs $spec } 'SHA256 mismatch'
    $script:signer = 'CN=Other Vendor, O=Other Vendor'
    Assert-Throws { Assert-OpenDocsMicrosoftInstaller $file } 'valid Microsoft signature'
    $script:signer = 'CN=Microsoft Corporation, O=Microsoft Corporation'; $script:signatureStatus = 'HashMismatch'
    Assert-Throws { Assert-OpenDocsMicrosoftInstaller $file } 'valid Microsoft signature'
} finally {
    foreach ($path in @($file,$manifestPath)) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
    Remove-Item -LiteralPath $fixture -Force
}
Write-Host "Passed $script:passed prerequisite tests. No real installers were run."
