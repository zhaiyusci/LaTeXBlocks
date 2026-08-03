$ErrorActionPreference = 'Stop'

# This is an isolated Word-layout experiment, not add-in code.
# It deliberately leaves wp:effectExtent at Word's zero default.

$root = Split-Path -Parent $PSCommandPath
$svgPath = Join-Path $root 'formula.svg'
$docxPath = Join-Path $root 'U2060-Inline-Image-Experiment.docx'
$pdfPath = Join-Path $root 'U2060-Inline-Image-Experiment.pdf'

$wdFormatDocumentDefault = 16
$wdExportFormatPdf = 17
$wdHorizontalPositionRelativeToTextBoundary = 7
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3

function Release-Com([object]$value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function Add-Text([object]$document, [string]$text, [string]$fontName = 'Arial', [double]$fontSize = 11,
                  [bool]$bold = $false, [int]$style = 0, [double]$afterPt = 6) {
    $start = $document.Content.End - 1
    $range = $document.Range($start, $start)
    try {
        $range.Text = $text + "`r"
        $paragraph = $document.Range($start, $document.Content.End - 1)
        try {
            if ($style -ne 0) { $paragraph.Style = $style }
            $paragraph.Font.Name = $fontName
            $paragraph.Font.Size = $fontSize
            $paragraph.Font.Bold = if ($bold) { -1 } else { 0 }
            $paragraph.ParagraphFormat.SpaceBefore = 0
            $paragraph.ParagraphFormat.SpaceAfter = $afterPt
            $paragraph.ParagraphFormat.LineSpacingRule = 0
        }
        finally { Release-Com $paragraph }
    }
    finally { Release-Com $range }
}

function Add-FormulaLine([object]$document, [string]$svg, [string]$fontName, [string]$boundary, [string]$alternativeText) {
    # The actual text sequence is: A U+0020 [boundary] <InlineShape> [boundary] U+0020 B.
    $start = $document.Content.End - 1
    $prefix = 'A ' + $boundary
    $prefixRange = $document.Range($start, $start)
    try { $prefixRange.Text = $prefix } finally { Release-Com $prefixRange }

    $picturePosition = $start + $prefix.Length
    $pictureRange = $document.Range($picturePosition, $picturePosition)
    $shape = $null
    try {
        $shape = $document.InlineShapes.AddPicture($svg, $false, $true, $pictureRange)
        $shape.LockAspectRatio = -1
        $shape.Width = 72
        $shape.AlternativeText = $alternativeText

        $shapeStart = $shape.Range.Start
        $shapeEnd = $shape.Range.End
        $suffixRange = $document.Range($shapeEnd, $shapeEnd)
        try { $suffixRange.Text = $boundary + ' B' + "`r" } finally { Release-Com $suffixRange }

        $line = $document.Range($start, $document.Content.End - 1)
        try {
            $line.Font.Name = $fontName
            $line.Font.Size = 16
            $line.Font.Bold = 0
            $line.ParagraphFormat.SpaceBefore = 0
            $line.ParagraphFormat.SpaceAfter = 4
            $line.ParagraphFormat.LineSpacingRule = 0
        }
        finally { Release-Com $line }

        return [pscustomobject]@{
            Shape = $shape
            PrefixSpaceStart = $start + 1
            PrefixSpaceEnd = $start + 2
            SuffixSpaceStart = $shapeEnd + $boundary.Length
            SuffixSpaceEnd = $shapeEnd + $boundary.Length + 1
        }
    }
    catch {
        Release-Com $shape
        throw
    }
    finally { Release-Com $pictureRange }
}

function Measure-Space([object]$document, [int]$start, [int]$end) {
    $left = $null
    $right = $null
    try {
        $left = $document.Range($start, $start)
        $right = $document.Range($end, $end)
        return [Math]::Round(
            [double]$right.Information($wdHorizontalPositionRelativeToTextBoundary) -
            [double]$left.Information($wdHorizontalPositionRelativeToTextBoundary), 2)
    }
    finally {
        Release-Com $right
        Release-Com $left
    }
}

if (Test-Path -LiteralPath $docxPath) {
    throw "Refusing to overwrite an existing experiment: $docxPath"
}

$word = $null
$document = $null
$tnrBare = $null
$tnrWj = $null
$simSunBare = $null
$simSunWj = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.PageSetup.TopMargin = 72
    $document.PageSetup.BottomMargin = 72
    $document.PageSetup.LeftMargin = 72
    $document.PageSetup.RightMargin = 72

    $normal = $document.Styles.Item($wdStyleNormal)
    try {
        $normal.Font.Name = 'Arial'
        $normal.Font.Size = 11
        $normal.ParagraphFormat.SpaceAfter = 6
    }
    finally { Release-Com $normal }

    Add-Text $document 'U+2060 around an inline SVG image' 'Arial' 20 $true 0 2
    Add-Text $document 'Real Microsoft Word experiment — no add-in code, no effect-extent correction.' 'Arial' 10 $false 0 12
    Add-Text $document 'Each formula is the same 72 pt SVG InlineShape. The ordinary U+0020 space remains outside the image; the second line adds U+2060 WORD JOINER immediately inside each of those spaces.' 'Arial' 10 $false 0 12

    Add-Text $document 'Times New Roman, 16 pt' 'Arial' 13 $true $wdStyleHeading2 4
    Add-Text $document 'Bare image — text is literally A ␠ [image] ␠ B.' 'Arial' 10 $false 0 1
    $tnrBare = Add-FormulaLine $document $svgPath 'Times New Roman' '' 'LaTeX source: $E = mc^2$'
    Add-Text $document 'U+2060 wrapper — text is literally A ␠ U+2060 [image] U+2060 ␠ B.' 'Arial' 10 $false 0 1
    $tnrWj = Add-FormulaLine $document $svgPath 'Times New Roman' ([string][char]0x2060) 'LaTeX source: $E = mc^2$'

    Add-Text $document 'SimSun, 16 pt' 'Arial' 13 $true $wdStyleHeading2 4
    Add-Text $document 'The same experiment with SimSun as the host font. This is included because a correct workaround must survive font changes.' 'Arial' 10 $false 0 1
    $simSunBare = Add-FormulaLine $document $svgPath 'SimSun' '' 'LaTeX source: $E = mc^2$'
    $simSunWj = Add-FormulaLine $document $svgPath 'SimSun' ([string][char]0x2060) 'LaTeX source: $E = mc^2$'

    $document.Repaginate()
    $word.ScreenRefresh()
    Start-Sleep -Milliseconds 300

    $tnrBareLeft = Measure-Space $document $tnrBare.PrefixSpaceStart $tnrBare.PrefixSpaceEnd
    $tnrBareRight = Measure-Space $document $tnrBare.SuffixSpaceStart $tnrBare.SuffixSpaceEnd
    $tnrWjLeft = Measure-Space $document $tnrWj.PrefixSpaceStart $tnrWj.PrefixSpaceEnd
    $tnrWjRight = Measure-Space $document $tnrWj.SuffixSpaceStart $tnrWj.SuffixSpaceEnd
    $simSunBareLeft = Measure-Space $document $simSunBare.PrefixSpaceStart $simSunBare.PrefixSpaceEnd
    $simSunBareRight = Measure-Space $document $simSunBare.SuffixSpaceStart $simSunBare.SuffixSpaceEnd
    $simSunWjLeft = Measure-Space $document $simSunWj.PrefixSpaceStart $simSunWj.PrefixSpaceEnd
    $simSunWjRight = Measure-Space $document $simSunWj.SuffixSpaceStart $simSunWj.SuffixSpaceEnd

    Add-Text $document 'Measured result' 'Arial' 13 $true $wdStyleHeading2 4
    Add-Text $document ("Times New Roman: bare = {0}/{1} pt; U+2060 = {2}/{3} pt (left/right ordinary spaces)." -f $tnrBareLeft,$tnrBareRight,$tnrWjLeft,$tnrWjRight) 'Arial' 10 $false 0 2
    Add-Text $document ("SimSun: bare = {0}/{1} pt; U+2060 = {2}/{3} pt." -f $simSunBareLeft,$simSunBareRight,$simSunWjLeft,$simSunWjRight) 'Arial' 10 $false 0 8
    Add-Text $document 'This experiment records the host-font dependence of U+2060 around inline images. The production add-in uses U+2060 to avoid Word''s direct image-adjacency path, but does not claim one identical measured advance for every host font.' 'Arial' 10 $false 0 0

    $document.SaveAs2($docxPath, $wdFormatDocumentDefault)
    $document.ExportAsFixedFormat($pdfPath, $wdExportFormatPdf)
}
finally {
    Release-Com $tnrBare.Shape
    Release-Com $tnrWj.Shape
    Release-Com $simSunBare.Shape
    Release-Com $simSunWj.Shape
    if ($document) { try { $document.Close($false) } catch {}; Release-Com $document }
    if ($word) { try { $word.Quit() } catch {}; Release-Com $word }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Created $docxPath"
