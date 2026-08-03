$ErrorActionPreference = 'Stop'

$registryPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
$registryName = 'TeX Formula Glyph Spike (TrueType)'
$target = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts\TeXFormulaGlyphSpike.ttf'
$current = (Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction Stop).$registryName

if ($current -ne $target) {
    throw "Refusing to remove an unexpected font registration: $current"
}
if (-not ($target.StartsWith((Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'), [System.StringComparison]::OrdinalIgnoreCase))) {
    throw "Refusing to remove a font outside the per-user Fonts directory: $target"
}

Remove-ItemProperty -Path $registryPath -Name $registryName
if (Test-Path -LiteralPath $target) {
    Remove-Item -LiteralPath $target -Force
}
Write-Host 'Removed only the original experiment font registration and file.'
