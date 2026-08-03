<#
  Transiently expose FormulaGlyphSpike.ttf to Windows, build a Word sample, and
  save both DOCX and PDF.  The font resource is removed in finally, so this does
  not install a user/system font.

  Prerequisite: run generate_font.py first.
#>
[CmdletBinding()]
param(
    [string]$FontPath,
    [string]$DocumentPath,
    [string]$PdfPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($FontPath)) {
    $FontPath = Join-Path $scriptRoot 'FormulaGlyphSpike.ttf'
}
if ([string]::IsNullOrWhiteSpace($DocumentPath)) {
    $DocumentPath = Join-Path $scriptRoot 'FormulaGlyphSpike-Word.docx'
}
if ([string]::IsNullOrWhiteSpace($PdfPath)) {
    $PdfPath = Join-Path $scriptRoot 'FormulaGlyphSpike-Word.pdf'
}

if (-not (Test-Path -LiteralPath $FontPath)) {
    throw "Font was not found: $FontPath. Run generate_font.py first."
}

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class FormulaGlyphSpikeNative {
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int AddFontResourceEx(string name, uint flags, IntPtr reserved);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RemoveFontResourceEx(string name, uint flags, IntPtr reserved);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam,
        uint flags, uint timeout, out UIntPtr result);
}
'@

$WM_FONTCHANGE = 0x001D
$HWND_BROADCAST = [IntPtr]0xffff
$SMTO_ABORTIFHUNG = 0x0002

function Invoke-FontChangeBroadcast {
    [UIntPtr]$ignored = [UIntPtr]::Zero
    [void][FormulaGlyphSpikeNative]::SendMessageTimeout(
        $HWND_BROADCAST, $WM_FONTCHANGE, [UIntPtr]::Zero, [IntPtr]::Zero,
        $SMTO_ABORTIFHUNG, 1000, [ref]$ignored)
}

$fontAdded = $false
$word = $null
$document = $null
$selection = $null
try {
    $count = [FormulaGlyphSpikeNative]::AddFontResourceEx(
        (Resolve-Path -LiteralPath $FontPath), 0, [IntPtr]::Zero)
    if ($count -le 0) {
        throw "AddFontResourceEx failed (Win32 error $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))."
    }
    $fontAdded = $true
    Invoke-FontChangeBroadcast

    Add-Type -AssemblyName System.Drawing
    $visibleToGdi = [System.Drawing.Text.InstalledFontCollection]::new().Families |
        Where-Object { $_.Name -eq 'Formula Glyph Character Spike' }
    if ($null -eq $visibleToGdi) {
        throw 'The transient font resource was accepted but is not visible through GDI font enumeration.'
    }
    Write-Host 'Formula Glyph Character Spike is visible through GDI font enumeration.'

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Add()
    $document.EmbedTrueTypeFonts = $true
    $document.SaveSubsetFonts = $true

    $selection = $word.Selection
    $selection.Font.Name = 'Times New Roman'
    $selection.Font.Size = 20
    $selection.TypeText('Normal text on each side: What does ')
    Write-Host 'Inserted leading text.'

    # This is one actual Unicode character in a run using the spike font.
    $selection.Font.Name = 'Formula Glyph Character Spike'
    Write-Host 'Selected Formula Glyph Character Spike for the Latin font slot.'
    $selection.Font.Size = 20
    $selection.TypeText([char]0xE000)
    Write-Host 'Inserted U+E000.'

    $selection.Font.Name = 'Times New Roman'
    $selection.Font.Size = 20
    $selection.TypeText(' stand for?')
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Font.Name = 'Times New Roman'
    $selection.Font.Size = 20
    $selection.TypeText('ASCII control: What does ')
    $selection.Font.Name = 'Formula Glyph Character Spike'
    $selection.Font.Size = 20
    $selection.TypeText('X')
    $selection.Font.Name = 'Times New Roman'
    $selection.Font.Size = 20
    $selection.TypeText(' stand for?')
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Font.Name = 'Times New Roman'
    $selection.Font.Size = 11
    $selection.TypeText('The first formula is one PUA character (U+E000), not an InlineShape. The second line maps diagnostic ASCII X to the same glyph. The spaces are ordinary U+0020 text spaces.')

    # Explicitly apply the character font to only the U+E000 run so the DOCX
    # records a normal w:t character plus w:rFonts, rather than a drawing.
    $document.SaveAs2($DocumentPath, 16) # wdFormatDocumentDefault
    $document.Close(0) # Close before the OOXML post-step changes the font hint.
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null

    & python.exe (Join-Path $scriptRoot 'patch_pua_font_slot.py') $DocumentPath
    if ($LASTEXITCODE -ne 0) {
        throw "patch_pua_font_slot.py failed with exit code $LASTEXITCODE."
    }

    # Reopen it in Word with the transient font resource still registered.  A
    # save at this point asks Word to embed the character font, if it accepts
    # it, and PDF export confirms that Word can actually paint the glyph.
    $document = $word.Documents.Open($DocumentPath)
    $document.EmbedTrueTypeFonts = $true
    $document.SaveSubsetFonts = $true
    $document.Save()
    $document.ExportAsFixedFormat($PdfPath, 17) # wdExportFormatPDF

    Write-Host "Created $DocumentPath"
    Write-Host "Created $PdfPath"
    Write-Host "Embedded fonts requested: $($document.EmbedTrueTypeFonts)"
}
finally {
    if ($selection -ne $null) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($selection)
    }
    if ($document -ne $null) {
        $document.Close(0) # wdDoNotSaveChanges
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    if ($fontAdded) {
        [void][FormulaGlyphSpikeNative]::RemoveFontResourceEx(
            (Resolve-Path -LiteralPath $FontPath), 0, [IntPtr]::Zero)
        Invoke-FontChangeBroadcast
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
