<#
  Install the experiment font for the current Windows user only.

  Close Word before running it, then start Word again. This is necessary for
  Word to rebuild its font list. Run Uninstall-TestFont.ps1 when the sample is
  no longer needed.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$family = 'Formula Glyph Character Spike'
$valueName = "$family (TrueType)"
$source = Join-Path $PSScriptRoot 'FormulaGlyphSpike.ttf'
$fontDirectory = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$destination = Join-Path $fontDirectory 'FormulaGlyphCharacterSpike.ttf'
$fontRegistry = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'

if (-not (Test-Path -LiteralPath $source)) {
    throw "Generated font was not found: $source. Run generate_font.py first."
}

New-Item -ItemType Directory -Force -Path $fontDirectory | Out-Null
Copy-Item -LiteralPath $source -Destination $destination -Force
New-ItemProperty -Path $fontRegistry -Name $valueName -PropertyType String -Value $destination -Force | Out-Null

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FormulaGlyphSpikeBroadcast {
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam,
        uint flags, uint timeout, out UIntPtr result);
}
'@
[UIntPtr]$ignored = [UIntPtr]::Zero
[void][FormulaGlyphSpikeBroadcast]::SendMessageTimeout(
    [IntPtr]0xffff, 0x001D, [UIntPtr]::Zero, [IntPtr]::Zero, 0x0002, 1000, [ref]$ignored)

Write-Host "Installed '$family' for the current user. Close and reopen Word before opening the sample."
Write-Host "Remove it later with .\Uninstall-TestFont.ps1"
