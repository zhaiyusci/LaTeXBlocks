<# Remove only the current-user Formula Glyph Character Spike test font. #>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$family = 'Formula Glyph Character Spike'
$valueName = "$family (TrueType)"
$fontDirectory = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$destination = Join-Path $fontDirectory 'FormulaGlyphCharacterSpike.ttf'
$legacyValueName = 'Formula Glyph Spike (TrueType)'
$legacyDestination = Join-Path $fontDirectory 'FormulaGlyphSpike.ttf'
$fontRegistry = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'

$currentProperties = Get-ItemProperty -Path $fontRegistry -ErrorAction Stop
$registeredProperty = $currentProperties.PSObject.Properties[$valueName]
$registered = if ($null -eq $registeredProperty) { $null } else { $registeredProperty.Value }
if ($registered -eq $destination) {
    Remove-ItemProperty -Path $fontRegistry -Name $valueName -ErrorAction Stop
}
Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue

# Remove only the earlier family name created during the same spike, if it is
# still registered to this exact per-user file. Do not touch any TeX Formula
# Glyph Spike registrations created by other experiments.
$legacyRegisteredProperty = $currentProperties.PSObject.Properties[$legacyValueName]
$legacyRegistered = if ($null -eq $legacyRegisteredProperty) { $null } else { $legacyRegisteredProperty.Value }
if ($legacyRegistered -eq $legacyDestination) {
    Remove-ItemProperty -Path $fontRegistry -Name $legacyValueName -ErrorAction Stop
}
Remove-Item -LiteralPath $legacyDestination -Force -ErrorAction SilentlyContinue

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FormulaGlyphSpikeUninstallBroadcast {
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam,
        uint flags, uint timeout, out UIntPtr result);
}
'@
[UIntPtr]$ignored = [UIntPtr]::Zero
[void][FormulaGlyphSpikeUninstallBroadcast]::SendMessageTimeout(
    [IntPtr]0xffff, 0x001D, [UIntPtr]::Zero, [IntPtr]::Zero, 0x0002, 1000, [ref]$ignored)

Write-Host "Removed '$family' from the current user. Restart Word to clear its font cache."
