# Read-only deployment checks. No Office application is launched.
. (Join-Path $PSScriptRoot 'prerequisites.ps1')

function Get-OpenDocsInstallRoot {
    param([string]$InstallRoot, [switch]$CurrentUser)
    if ($InstallRoot) { return [IO.Path]::GetFullPath($InstallRoot) }
    if ($CurrentUser) { return (Join-Path $env:LOCALAPPDATA 'OpenDocs') }
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    return (Join-Path $programFiles 'OpenDocs')
}

function Get-OpenDocsProductState {
    param([ValidateSet('Word','Excel','PowerPoint','Outlook')][string]$App, [string]$InstallRoot, [switch]$CurrentUser, [version]$MinimumVersion)
    $root = Get-OpenDocsInstallRoot -InstallRoot $InstallRoot -CurrentUser:$CurrentUser
    $name = "${App}AiAddIn"
    $directory = Join-Path $root $name
    $manifestPath = [IO.Path]::GetFullPath((Join-Path $directory "$name.vsto"))
    $assemblyPath = Join-Path $directory "$name.dll"
    $problems = @()
    $manifestExists = Test-Path -LiteralPath $manifestPath -PathType Leaf
    $assemblyExists = Test-Path -LiteralPath $assemblyPath -PathType Leaf
    if (-not $manifestExists) { $problems += 'Deployment manifest missing' }
    if (-not $assemblyExists) { $problems += 'Add-in assembly missing' }
    $version = $null
    if ($assemblyExists) {
        $rawVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($assemblyPath).FileVersion
        [void][version]::TryParse([string]$rawVersion, [ref]$version)
    }
    if ($MinimumVersion -and (-not $version -or $version -lt $MinimumVersion)) { $problems += 'Installed version is older than required' }
    $hive = if ($CurrentUser) { 'CurrentUser' } else { 'LocalMachine' }
    $views = if ([Environment]::Is64BitOperatingSystem) { @('Registry64','Registry32') } else { @('Registry32') }
    $platform = Get-OpenDocsRegistryValue -Hive LocalMachine -View Registry64 -Key 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name Platform
    if (-not $CurrentUser) {
        if ($platform -eq 'x64') { $views = @('Registry64') }
        elseif ($platform -eq 'x86') { $views = @('Registry32') }
    }
    $registrations = @()
    foreach ($view in $views) {
        $key = "Software\Microsoft\Office\$App\Addins\$name"
        $value = Get-OpenDocsRegistryValue -Hive $hive -View $view -Key $key -Name Manifest
        $load = Get-OpenDocsRegistryValue -Hive $hive -View $view -Key $key -Name LoadBehavior
        $matches = $false
        if ($value) {
            try {
                $uriText = [string]$value -replace '\|vstolocal$', ''
                $uri = [uri]$uriText
                $matches = $uri.IsFile -and [string]::Equals([IO.Path]::GetFullPath($uri.LocalPath), $manifestPath, [StringComparison]::OrdinalIgnoreCase)
            } catch { $matches = $false }
            $registrations += [pscustomobject]@{ Hive=$hive; View=$view; Manifest=[string]$value; LoadBehavior=$load; ExpectedPath=$matches }
        }
    }
    $registered = @($registrations | Where-Object { $_.ExpectedPath -and $_.LoadBehavior -eq 3 }).Count -gt 0
    if (-not $registered) { $problems += 'Correct enabled Office registration missing' }
    if (-not $CurrentUser) {
        foreach ($view in $views) {
            $userValue = Get-OpenDocsRegistryValue -Hive CurrentUser -View $view -Key "Software\Microsoft\Office\$App\Addins\$name" -Name Manifest
            if ($userValue) { $problems += 'This account has a per-user registration that may override the machine registration'; break }
        }
    }
    $runtimes = @(Get-OpenDocsPrerequisiteSpecs | ForEach-Object {
        [pscustomobject]@{ Name=$_.Name; Installed=(Test-OpenDocsPrerequisite -Id $_.Id -AllUsers (-not $CurrentUser)) }
    })
    $runtimeReady = @($runtimes | Where-Object { -not $_.Installed }).Count -eq 0
    if (-not $runtimeReady) { $problems += 'Required runtime missing or outdated' }
    $installed = $manifestExists -and $assemblyExists -and $registered -and (-not $MinimumVersion -or ($version -and $version -ge $MinimumVersion))
    [pscustomobject]@{
        Product=$App; Scope=$(if($CurrentUser){'CurrentUser'}else{'AllUsers'}); Installed=[bool]$installed
        Ready=([bool]$installed -and $runtimeReady -and $problems.Count -eq 0)
        Version=$(if($version){$version.ToString()}else{$null}); InstallDirectory=$directory
        Registrations=$registrations; Prerequisites=$runtimes; Problems=$problems
        Note='File/registry/runtime checks only; actual loading inside Office and certificate trust are not verified.'
    }
}
