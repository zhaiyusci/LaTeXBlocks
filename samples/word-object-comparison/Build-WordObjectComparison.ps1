[CmdletBinding()]
param(
    [string]$OutputDirectory
)

# The compiled probe keeps Word COM calls on a single STA thread.  This is
# deliberately separate from the add-in: it measures Word's native layout of
# a text run and three InlineShape payloads.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = $here
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$project = Join-Path $here 'WordObjectComparison.csproj'
$msbuild = Join-Path ${env:ProgramFiles} 'Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path $msbuild)) {
    throw "MSBuild was not found at $msbuild. Build $project with Visual Studio's MSBuild."
}

& $msbuild $project '/t:Build' '/p:Configuration=Release' '/p:Platform=AnyCPU' '/nologo' '/v:minimal'
if ($LASTEXITCODE -ne 0) { throw "MSBuild failed with exit code $LASTEXITCODE." }

$exe = Join-Path $here 'bin\Release\WordObjectComparison.exe'
& $exe $OutputDirectory
if ($LASTEXITCODE -ne 0) { throw "The Word comparison probe failed with exit code $LASTEXITCODE." }
