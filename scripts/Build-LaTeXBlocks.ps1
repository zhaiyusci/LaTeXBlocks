param([ValidateSet('Debug','Release')][string]$Configuration='Debug')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$msbuild='C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
if(-not (Test-Path $msbuild)){throw 'Visual Studio MSBuild was not found.'}
& $msbuild (Join-Path $root 'LaTeXBlocks.sln') /m /t:Build "/p:Configuration=$Configuration" `
    /p:VisualStudioVersion=18.0 /p:EnableVstoProjectRegistration=false /v:minimal
exit $LASTEXITCODE
