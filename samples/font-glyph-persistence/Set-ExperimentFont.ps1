[CmdletBinding()]
param(
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$source = Join-Path $root 'TeXFormulaGlyphSpike2.ttf'
$fontDirectory = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$target = Join-Path $fontDirectory 'TeXFormulaGlyphSpike2.ttf'
$registryPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
$registryName = 'TeX Formula Glyph Spike 2 (TrueType)'

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class FontChangeBroadcast {
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint msg, UIntPtr wParam, string lParam,
        uint flags, uint timeout, out UIntPtr result);
}
'@

if ($Remove) {
    $existing = Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Force
    }
    Write-Host "Removed experimental per-user font registration."
}
else {
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Expected generated font was not found: $source"
    }
    New-Item -ItemType Directory -Force -Path $fontDirectory | Out-Null
    Copy-Item -LiteralPath $source -Destination $target -Force
    New-ItemProperty -Path $registryPath -Name $registryName -PropertyType String -Value $target -Force | Out-Null
    Write-Host "Installed experimental font for this user only: $target"
}

[UIntPtr]$ignored = [UIntPtr]::Zero
[void][FontChangeBroadcast]::SendMessageTimeout(
    [IntPtr]0xffff, 0x001d, [UIntPtr]::Zero, $null, 0x0002, 1000, [ref]$ignored)
