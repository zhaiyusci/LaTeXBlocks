param(
    [ValidateSet('Debug','Release')][string]$Configuration = 'Debug',
    [Alias('Host')][ValidateSet('Word','PowerPoint','Both')][string]$TargetHost = 'Both',
    [switch]$Interaction
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$tests = @()
if ($TargetHost -in @('Word', 'Both')) {
    $tests += Join-Path $root "tests\LaTeXBlocks.WordSmoke\bin\$Configuration\LaTeXBlocks.WordSmoke.exe"
}
if ($TargetHost -in @('PowerPoint', 'Both')) {
    $tests += Join-Path $root "tests\LaTeXBlocks.PowerPointSmoke\bin\$Configuration\LaTeXBlocks.PowerPointSmoke.exe"
}
foreach ($test in $tests) {
    if (-not (Test-Path -LiteralPath $test)) { throw "Build LaTeX Blocks before running its smoke tests: $test" }
    $isWord = [IO.Path]::GetFileName($test) -eq 'LaTeXBlocks.WordSmoke.exe'
    $previousInteraction = $env:LATEXBLOCKS_UIA_FONT_COLOR_SMOKE
    try {
        if ($Interaction -and $isWord) {
            $env:LATEXBLOCKS_UIA_FONT_COLOR_SMOKE = '1'
        }
        & $test
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }
    finally {
        $env:LATEXBLOCKS_UIA_FONT_COLOR_SMOKE = $previousInteraction
    }
}
