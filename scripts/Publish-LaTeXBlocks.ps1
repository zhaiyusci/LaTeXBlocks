param(
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version = '0.2.93',
    [string]$StemTeXSourceDir
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'src\LaTeXBlocks.Word.AddIn\LaTeXBlocks.Word.AddIn.csproj'
$powerPointProject = Join-Path $root 'src\LaTeXBlocks.PowerPoint.AddIn\LaTeXBlocks.PowerPoint.AddIn.csproj'
$publishDir = Join-Path $root 'src\LaTeXBlocks.Word.AddIn\bin\Release\app.publish'
$powerPointPublishDir = Join-Path $root 'src\LaTeXBlocks.PowerPoint.AddIn\bin\Release\app.publish'
$stagingDir = Join-Path $root 'dist\staging'
$outputDir = Join-Path $root 'dist\release'
$installer = Join-Path $outputDir "LaTeXBlocks-Setup-$Version.exe"
$checksumPath = $installer + '.sha256'
$certificatePath = Join-Path $stagingDir 'publisher.cer'
$msbuild = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
$iscc = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
$vcRedist = Get-ChildItem 'C:\Program Files\Microsoft Visual Studio\18\Community\VC\Redist\MSVC' `
    -Filter vc_redist.x64.exe -File -Recurse -ErrorAction SilentlyContinue |
    Sort-Object { [Version]$_.VersionInfo.FileVersion } -Descending |
    Select-Object -First 1

if ([string]::IsNullOrWhiteSpace($StemTeXSourceDir)) {
    $StemTeXSourceDir = Join-Path $root '..\xetex\stemtex\dist\stemtex-installer\StemTeX'
}
$StemTeXSourceDir = [IO.Path]::GetFullPath($StemTeXSourceDir)

foreach ($required in @(
    $msbuild,
    $iscc,
    $project,
    $powerPointProject,
    (Join-Path $StemTeXSourceDir 'runtime\VERSION'),
    (Join-Path $StemTeXSourceDir 'runtime\bin\sdk\stemtex-renderer.dll'),
    (Join-Path $StemTeXSourceDir 'runtime\bin\windows\dvisvgmdaemon.dll'),
    (Join-Path $StemTeXSourceDir 'gui\profiles\xits_cjk\preamble.tex'),
    (Join-Path $StemTeXSourceDir 'gui\profiles\arial_lete_simhei\preamble.tex')
)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Required build input was not found: $required" }
}
if (-not $vcRedist) { throw 'The VC++ 2015-2022 x64 redistributable was not found in Visual Studio.' }

$stemTeXVersion = [Version]((Get-Content (Join-Path $StemTeXSourceDir 'runtime\VERSION') -Raw).Trim())
if ($stemTeXVersion -lt [Version]'0.12.0') {
    throw "LaTeX Blocks requires StemTeX 0.12.0 or newer; found $stemTeXVersion."
}

$rootPath = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
foreach ($path in @($stagingDir, $publishDir, $powerPointPublishDir)) {
    $resolved = [IO.Path]::GetFullPath($path)
    if (-not $resolved.StartsWith($rootPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean a path outside the repository: $resolved"
    }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
New-Item -ItemType Directory -Path $stagingDir, $outputDir -Force | Out-Null
if ((Test-Path -LiteralPath $installer) -or
    (Test-Path -LiteralPath $checksumPath)) {
    throw "Refusing to overwrite an existing release artifact for version $Version."
}

$certificate = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq 'CN=LaTeX Blocks Development' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
if (-not $certificate) {
    throw 'Run Initialize-LaTeXBlocks.ps1 before publishing.'
}

$versionParts = @($Version.Split('.'))
$applicationVersion = ($versionParts + @('0') | Select-Object -First 4) -join '.'

& $msbuild $project /t:Publish /p:Configuration=Release /p:VisualStudioVersion=18.0 `
    /p:IsWebBootstrapper=false /p:BootstrapperComponentsLocation=HomeSite `
    /p:EnableVstoProjectRegistration=false '/p:PublishUrl=bin\Release\app.publish\' `
    "/p:ApplicationVersion=$applicationVersion" /v:minimal
if ($LASTEXITCODE -ne 0) { throw "VSTO publish failed with exit code $LASTEXITCODE." }

& $msbuild $powerPointProject /t:Publish /p:Configuration=Release /p:VisualStudioVersion=18.0 `
    /p:IsWebBootstrapper=false /p:BootstrapperComponentsLocation=HomeSite `
    /p:EnableVstoProjectRegistration=false '/p:PublishUrl=bin\Release\app.publish\' `
    "/p:ApplicationVersion=$applicationVersion" /v:minimal
if ($LASTEXITCODE -ne 0) { throw "PowerPoint VSTO publish failed with exit code $LASTEXITCODE." }

foreach ($requiredPublishOutput in @(
    (Join-Path $publishDir 'setup.exe'),
    (Join-Path $publishDir 'LaTeXBlocks.Word.AddIn.vsto'),
    (Join-Path $powerPointPublishDir 'LaTeXBlocks.PowerPoint.AddIn.vsto')
)) {
    if (-not (Test-Path -LiteralPath $requiredPublishOutput)) {
        throw "VSTO publish did not produce: $requiredPublishOutput"
    }
}

[void](Export-Certificate -Cert $certificate -FilePath $certificatePath -Force)

$iss = Join-Path $root 'installer\LaTeXBlocks.iss'
& $iscc "/DMyAppVersion=$Version" "/DSourceDir=$publishDir" `
    "/DPowerPointSourceDir=$powerPointPublishDir" "/DStemTeXSourceDir=$StemTeXSourceDir" `
    "/DCertPath=$certificatePath" "/DVcRedistPath=$($vcRedist.FullName)" `
    "/DVcMajor=$($vcRedist.VersionInfo.FileMajorPart)" "/DVcMinor=$($vcRedist.VersionInfo.FileMinorPart)" `
    "/DVcBuild=$($vcRedist.VersionInfo.FileBuildPart)" "/DVcRevision=$($vcRedist.VersionInfo.FilePrivatePart)" `
    "/DOutputDir=$outputDir" $iss
if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed with exit code $LASTEXITCODE." }

if (-not (Test-Path -LiteralPath $installer)) { throw "Installer was not produced: $installer" }

$signTool = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin' -Filter signtool.exe -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
    Sort-Object FullName -Descending |
    Select-Object -First 1
if ($signTool) {
    & $signTool.FullName sign /sha1 $certificate.Thumbprint /fd SHA256 $installer
    if ($LASTEXITCODE -ne 0) { throw "Signing the installer failed with exit code $LASTEXITCODE." }
}

$hash = Get-FileHash -LiteralPath $installer -Algorithm SHA256
[IO.File]::WriteAllText(
    $checksumPath,
    $hash.Hash + '  ' + [IO.Path]::GetFileName($installer) + [Environment]::NewLine,
    [Text.UTF8Encoding]::new($false))
Write-Output "Installer: $installer"
Write-Output "StemTeX: $StemTeXSourceDir ($stemTeXVersion)"
Write-Output "SHA256: $($hash.Hash)"
Write-Output "Checksum: $checksumPath"
