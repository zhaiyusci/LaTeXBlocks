<# Verify that the already-created DOCX paints after the test font is removed. #>
[CmdletBinding()]
param(
    [string]$DocumentPath,
    [string]$PdfPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($DocumentPath)) {
    $DocumentPath = Join-Path $scriptRoot 'FormulaGlyphSpike-Word.docx'
}
if ([string]::IsNullOrWhiteSpace($PdfPath)) {
    $PdfPath = Join-Path $scriptRoot 'FormulaGlyphSpike-Word-embedded-reopen.pdf'
}
if (-not (Test-Path -LiteralPath $DocumentPath)) {
    throw "DOCX was not found: $DocumentPath. Run verify_word.ps1 first."
}

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open((Resolve-Path -LiteralPath $DocumentPath).Path)
    $document.ExportAsFixedFormat($PdfPath, 17) # wdExportFormatPDF
    Write-Host "Created $PdfPath from the embedded-font DOCX."
}
finally {
    if ($document -ne $null) {
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
