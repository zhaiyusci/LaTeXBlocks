param([ValidateSet('Debug','Release')][string]$Configuration='Debug')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$test=Join-Path $root "tests\LaTeXBlocks.WordSmoke\bin\$Configuration\LaTeXBlocks.WordSmoke.exe"
if(-not (Test-Path $test)){throw 'Build LaTeX Blocks before running its smoke test.'}
& $test
exit $LASTEXITCODE
