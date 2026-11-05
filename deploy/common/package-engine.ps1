[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Word','Excel','PowerPoint','Outlook')][string]$App,
    [Parameter(Mandatory)][string]$TemplateDirectory,
    [Parameter(Mandatory)][string]$RepositoryRoot,
    [Parameter(Mandatory)][string]$OutDir,
    [string]$PrerequisiteSource,
    [ValidateSet('x64','x86')][string]$PrerequisiteArchitecture = 'x64',
    [string]$CertificateThumbprint,
    [string]$MSBuildPath,
    [switch]$ScriptsOnly,
    [switch]$SkipBuild,
    [string]$PayloadDirectory,
    [string]$PublicCertificatePath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'prerequisites.ps1')
if ($ScriptsOnly -and $SkipBuild) { throw 'Choose ScriptsOnly or SkipBuild, not both.' }
$name = "${App}AiAddIn"
$repo = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\')
$template = [IO.Path]::GetFullPath($TemplateDirectory).TrimEnd('\')
$output = [IO.Path]::GetFullPath($OutDir).TrimEnd('\')
if (-not $PrerequisiteSource) {
    $PrerequisiteSource = Join-Path $template 'prerequisites'
}
$prereq = [IO.Path]::GetFullPath($PrerequisiteSource).TrimEnd('\')
if ($output -eq [IO.Path]::GetPathRoot($output).TrimEnd('\')) { throw 'Output cannot be a drive root.' }
foreach ($protected in @($repo,$template,$prereq)) {
    if ($protected -eq $output -or $protected.StartsWith($output+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Output would overwrite source files.' }
}
if (Test-Path -LiteralPath $output) { throw 'Output already exists. Choose a new -OutDir; existing packages are never deleted automatically.' }
$specs = @(Get-OpenDocsPrerequisiteSpecs $PrerequisiteArchitecture)
Assert-OpenDocsPrerequisitePayload -Directory $prereq -Specs $specs
$publicCert = $null
if (-not $ScriptsOnly) {
    $projectFile = Join-Path $repo "$name\$name.csproj"
    if (-not $SkipBuild) {
        if (-not (Test-Path -LiteralPath $projectFile)) { throw 'Source project missing. Package.ps1 runs on the build machine, not the target.' }
        if (-not $CertificateThumbprint) {
            [xml]$xml = Get-Content -LiteralPath $projectFile
            $CertificateThumbprint = @($xml.Project.PropertyGroup.ManifestCertificateThumbprint | Where-Object { $_ })[0]
        }
        $publicCert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
        if (-not $publicCert -or -not $publicCert.HasPrivateKey) { throw 'A signing certificate with private key is required on the build machine. Pass -CertificateThumbprint.' }
        if (-not $MSBuildPath) {
            $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
            if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Installer/vswhere.exe not found.' }
            $vs = & $vswhere -latest -version '[17.0,18.0)' -property installationPath
            if (-not $vs) { throw 'Visual Studio 2022 not found.' }
            $MSBuildPath = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
        }
        $appDirectory = Split-Path $projectFile
        Push-Location $appDirectory
        try {
            & npm.cmd ci
            if ($LASTEXITCODE) { throw 'npm ci failed.' }
            & .\node_modules\.bin\esbuild.cmd web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@officeai/agent-core=../shared/web-src/agent-core/index.ts --alias:@officeai/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap
            if ($LASTEXITCODE) { throw 'Web bundle build failed.' }
        } finally { Pop-Location }
        & $MSBuildPath $projectFile /restore /t:Build /p:Configuration=Release "/p:ManifestCertificateThumbprint=$CertificateThumbprint" /v:minimal
        if ($LASTEXITCODE) { throw 'Release build failed.' }
        $PayloadDirectory = Join-Path $repo "$name\bin\Release"
    }
    if (-not $PayloadDirectory) { throw '-SkipBuild requires -PayloadDirectory and -PublicCertificatePath.' }
    if (-not $publicCert) {
        if (-not $PublicCertificatePath) { throw '-PublicCertificatePath is required for a prebuilt payload.' }
        $publicCert = New-Object Security.Cryptography.X509Certificates.X509Certificate2($PublicCertificatePath)
    }
    if ($publicCert.NotAfter -le (Get-Date)) { throw 'Signing certificate has expired.' }
    foreach ($required in @("$name.dll","$name.vsto","$name.dll.manifest",'web\index.html','web\bundle.js')) {
        if (-not (Test-Path -LiteralPath (Join-Path $PayloadDirectory $required))) { throw "Payload file missing: $required" }
    }
    [xml]$manifest = Get-Content -LiteralPath (Join-Path $PayloadDirectory "$name.vsto")
    $certNode = $manifest.SelectSingleNode('//*[local-name()="X509Certificate"]')
    if (-not $certNode) { throw 'Deployment manifest has no signing certificate.' }
    $manifestCert = New-Object Security.Cryptography.X509Certificates.X509Certificate2(,[Convert]::FromBase64String($certNode.InnerText))
    if ($manifestCert.Thumbprint -ne $publicCert.Thumbprint) { throw 'Payload signing certificate does not match the supplied public certificate.' }
}
New-Item -ItemType Directory -Path $output | Out-Null
# Every product package carries its own support files; no references to sibling products.
foreach ($item in (Get-ChildItem -LiteralPath $template -File)) { Copy-Item -LiteralPath $item.FullName -Destination $output -Force }
$support = Join-Path $output 'common'
New-Item -ItemType Directory -Path $support | Out-Null
Get-ChildItem -LiteralPath $PSScriptRoot -File | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $support -Force }
$destPrereq = Join-Path $output 'prerequisites'
New-Item -ItemType Directory -Path $destPrereq | Out-Null
foreach ($spec in $specs) { Copy-Item -LiteralPath (Join-Path $prereq $spec.Installer) -Destination $destPrereq -Force }
Copy-Item -LiteralPath (Join-Path $prereq 'manifest.json') -Destination $destPrereq -Force
$product = Get-Content -LiteralPath (Join-Path $template 'product.json') -Raw | ConvertFrom-Json
$product.PackageReady = -not $ScriptsOnly
if (-not $ScriptsOnly) {
    Copy-Item -LiteralPath $PayloadDirectory -Destination (Join-Path $output $name) -Recurse
    [IO.File]::WriteAllBytes((Join-Path $output 'OpenDocsDevCert.cer'), $publicCert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    $product.Version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $PayloadDirectory "$name.dll")).FileVersion
    $parsedVersion = $null
    if (-not [version]::TryParse([string]$product.Version, [ref]$parsedVersion)) { throw 'Payload assembly has no valid FileVersion.' }
    $detectionPath = Join-Path $output 'Detect-SCCM.ps1'
    $detection = [IO.File]::ReadAllText($detectionPath).Replace("MinimumVersion = '0.0.0.0'", "MinimumVersion = '$($parsedVersion.ToString())'")
    [IO.File]::WriteAllText($detectionPath, $detection)
}
$product | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'product.json') -Encoding UTF8
Write-Host "Prepared $App at $output. PackageReady=$($product.PackageReady)."
if ($ScriptsOnly) { Write-Warning 'Scripts/dependencies only: no add-in binaries or certificate. Build a full package before deployment.' }
