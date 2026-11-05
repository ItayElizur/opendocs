$ErrorActionPreference = 'Stop'
$deploy = Join-Path (Split-Path $PSScriptRoot) 'deploy'
. (Join-Path $deploy 'common\status.ps1')
$script:passed = 0
function Assert($condition, [string]$message) { if (-not $condition) { throw $message }; $script:passed++ }
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('OpenDocsProductTests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture | Out-Null
$script:registration = $null
$script:loadBehavior = 3
$script:runtimeReady = $true
$script:userOverride = $false
function Get-OpenDocsRegistryValue {
    param($Hive='LocalMachine',$View='Registry64',$Key,$Name)
    if ($Name -eq 'Platform') { return 'x64' }
    if ($Hive -eq 'CurrentUser' -and -not $script:userOverride) { return $null }
    if ($Name -eq 'Manifest') { return $script:registration }
    if ($Name -eq 'LoadBehavior') { return $script:loadBehavior }
}
function Test-OpenDocsPrerequisite { param($Id,$AllUsers) return $script:runtimeReady }
foreach ($app in @('Word','Excel','PowerPoint','Outlook')) {
    $directory = Join-Path $fixture "${app}AiAddIn"
    New-Item -ItemType Directory -Path $directory | Out-Null
    $dll = Join-Path $directory "${app}AiAddIn.dll"
    $manifest = Join-Path $directory "${app}AiAddIn.vsto"
    # Use a real versioned DLL only as a file-version fixture; it is never loaded as an add-in.
    Copy-Item -LiteralPath ([string].Assembly.Location) -Destination $dll
    Set-Content -LiteralPath $manifest -Value '<fixture />'
    $script:registration = 'file:///' + ($manifest -replace '\\','/') + '|vstolocal'
    $script:runtimeReady = $true; $script:loadBehavior = 3; $script:userOverride = $false
    $state = Get-OpenDocsProductState -App $app -InstallRoot $fixture
    Assert $state.Ready "$app complete fixture was not ready."
    Assert ($state.Product -eq $app) 'Wrong product returned.'
    $script:loadBehavior = 2
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture).Installed) 'Disabled add-in detected as enabled.'
    $script:loadBehavior = 3; $script:runtimeReady = $false
    $state = Get-OpenDocsProductState -App $app -InstallRoot $fixture
    Assert ($state.Installed -and -not $state.Ready) 'Runtime readiness was not separated from installation.'
    $script:runtimeReady = $true
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture -MinimumVersion '999.0').Installed) 'Older version accepted.'
    $script:userOverride = $true
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture).Ready) 'Per-user registration conflict missed.'
    $script:userOverride = $false; $script:registration = 'file:///C:/wrong/place.vsto|vstolocal'
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture).Installed) 'Wrong manifest path accepted.'
    $script:registration = $null
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture).Installed) 'Absent registration accepted.'
    Remove-Item -LiteralPath $manifest -Force
    Assert (-not (Get-OpenDocsProductState -App $app -InstallRoot $fixture).Installed) 'Missing manifest accepted.'
    Remove-Item -LiteralPath $dll -Force
    Remove-Item -LiteralPath $directory -Force

    $source = Join-Path $deploy $app
    foreach ($required in @('Install.ps1','Uninstall.ps1','Test-Installed.ps1','Detect-SCCM.ps1','Check-Prerequisites.ps1','Install-Prerequisites.ps1','Get-Prerequisites.ps1','Package.ps1','product.json','README.md','common\status.ps1','common\prerequisites.ps1','common\install-engine.ps1','common\uninstall-engine.ps1')) {
        Assert (Test-Path -LiteralPath (Join-Path $source $required)) "$app missing $required"
    }
    $detector = Get-Content -LiteralPath (Join-Path $source 'Detect-SCCM.ps1') -Raw
    Assert (-not $detector.Contains('$PSScriptRoot')) 'SCCM detection depends on neighboring files.'
    $original = (Get-FileHash (Join-Path $deploy 'common\status.ps1')).Hash
    Assert ((Get-FileHash (Join-Path $source 'common\status.ps1')).Hash -eq $original) 'Support copy is stale.'
}
. (Join-Path $deploy 'common\paths.ps1')
Assert ((Assert-OpenDocsProductDirectory -Root $fixture -ProductName WordAiAddIn) -eq (Join-Path $fixture 'WordAiAddIn')) 'Safe product path rejected.'
$rejected = $false
try { Assert-OpenDocsProductDirectory -Root $fixture -ProductName '..\escape' | Out-Null } catch { $rejected = $true }
Assert $rejected 'Path traversal was not rejected.'
Remove-Item -LiteralPath $fixture -Force
Write-Host "Passed $script:passed product deployment tests. No Office, registry, certificate or runtime changes were made."
