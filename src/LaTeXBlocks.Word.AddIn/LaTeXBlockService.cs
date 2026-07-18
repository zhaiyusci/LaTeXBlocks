using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    internal sealed class LaTeXBlockService
    {
        internal const string EquationSequenceIdentifier = "LaTeXEquation";
        internal const string EquationBookmarkPrefix = "LTXEQ_";
        private const float EquationSideColumnPercent = 10.0f;
        private const double EquationSvgWidthSafetyPt = 2.0;
        private readonly WordInterop.Application application;
        private readonly StemTeXBackend renderers;
        private readonly string cacheDirectory;

        internal LaTeXBlockService(WordInterop.Application application, StemTeXBackend renderers)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.renderers = renderers ?? throw new ArgumentNullException(nameof(renderers));
            cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LaTeXBlocks", "cache");
            Directory.CreateDirectory(cacheDirectory);
        }

        internal string[] Profiles => renderers.Profiles;

        internal LaTeXBlockRender RenderPreview(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile,
            double fontSizePt = 10)
        {
            return RenderPreviewAsync(source, widthPt, mode, profile, fontSizePt).GetAwaiter().GetResult();
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double fontSizePt = 10)
        {
            var result = await renderers.RenderLatestAsync(profile, source, widthPt,
                mode == LaTeXBlockLayoutMode.Auto, fontSizePt);
            return new LaTeXBlockRender(WriteSvg(result.Bytes), result.Bytes, result.DepthPt, fontSizePt);
        }

        internal WordInterop.InlineShape InsertBlock(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile)
        {
            EnsureDocument();
            var fontSizePt = ResolveFontSize(application.Selection.Range, mode, 10);
            var render = RenderPreview(source, widthPt, mode, profile, fontSizePt);
            return InsertRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertRendered(string source, double widthPt, LaTeXBlockLayoutMode mode,
            LaTeXBlockRender render)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));
            var target = application.Selection.Range.Duplicate;
            var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt, mode, render.FontSizePt,
                LaTeXBlockRole.Content);
            return InsertRenderedAt(target, source, mode, render, metadata, true);
        }

        internal WordInterop.InlineShape InsertNumberedBlock(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile)
        {
            EnsureDocument();
            var fontSizePt = ResolveFontSize(application.Selection.Range, mode, 10);
            var render = RenderPreview(source, widthPt, mode, profile, fontSizePt);
            return InsertNumberedRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertNumberedRendered(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));

            var document = application.ActiveDocument;
            var target = application.Selection.Range.Duplicate;
            ValidateNumberedEquationTarget(target);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            WordInterop.Table table = null;
            var undoStarted = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Insert Numbered Equation");
                undoStarted = true;
                target.ParagraphFormat.LeftIndent = 0;
                target.ParagraphFormat.RightIndent = 0;
                target.ParagraphFormat.FirstLineIndent = 0;
                table = document.Tables.Add(target, 1, 3,
                    WordInterop.WdDefaultTableBehavior.wdWord9TableBehavior,
                    WordInterop.WdAutoFitBehavior.wdAutoFitFixed);
                ConfigureNumberedEquationTable(table, render.FontSizePt);

                var centerWidth = (double)table.Cell(1, 2).Width;
                var renderedWidth = ReadSvgWidthPt(render.SvgBytes);
                if (renderedWidth > centerWidth + 0.5)
                    throw new InvalidOperationException("The rendered formula width (" +
                        renderedWidth.ToString("0.#", CultureInfo.InvariantCulture) +
                        " pt) is wider than the numbered equation area (" +
                        centerWidth.ToString("0.#", CultureInfo.InvariantCulture) + " pt). Reduce Block width.");

                var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt, mode, render.FontSizePt,
                    LaTeXBlockRole.NumberedEquation);
                var formulaTarget = table.Cell(1, 2).Range.Duplicate;
                formulaTarget.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                var shape = InsertRenderedAt(formulaTarget, source, mode, render, metadata, false);
                ValidateNumberedEquationPlacement(shape, centerWidth);
                AddEquationNumber(table.Cell(1, 3), metadata.Id);
                InsertTableSeparatorAfter(table);
                UpdateEquationNumbers(document);
                shape.Range.Select();
                return shape;
            }
            catch
            {
                try { table?.Delete(); } catch { }
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        private WordInterop.InlineShape InsertRenderedAt(WordInterop.Range requestedTarget, string source,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, LaTeXBlockMetadata metadata, bool select)
        {
            var target = requestedTarget.Duplicate;
            target.Text = string.Empty;
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var insertionPath = PrepareInsertionSvg(render, mode);
            var shape = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
            ApplyContract(shape, source, metadata);
            shape = RemoveWordInlineEffectExtent(shape);
            ApplyBaselinePosition(shape, metadata);
            if (select) shape.Range.Select();
            return shape;
        }

        internal WordInterop.InlineShape UpdateBlock(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double? fontSizePt = null, bool selectReplacement = true)
        {
            var size = fontSizePt ?? ResolveFontSize(oldShape.Range, mode, 10);
            var render = RenderPreview(source, widthPt, mode, profile, size);
            return UpdateRendered(oldShape, source, widthPt, mode, render, selectReplacement);
        }

        internal WordInterop.InlineShape UpdateRendered(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, bool selectReplacement = true)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!TryReadContract(oldShape, out var previous, out _))
                throw new InvalidOperationException("The selected image is not a LaTeX Block.");
            var numberedCellWidth = previous.Role == LaTeXBlockRole.NumberedEquation
                ? (double?)NumberedEquationCellWidth(oldShape)
                : null;
            if (numberedCellWidth.HasValue && ReadSvgWidthPt(render.SvgBytes) > numberedCellWidth.Value + 0.5)
                throw new InvalidOperationException("The rendered formula is wider than the numbered equation area. Reduce Block width or shorten the formula.");

            var target = oldShape.Range.Duplicate;
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt, mode, render.FontSizePt,
                previous.Role);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            WordInterop.InlineShape replacement = null;
            try
            {
                var insertionPath = PrepareInsertionSvg(render, mode);
                replacement = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
                ApplyContract(replacement, source, metadata);
                replacement = RemoveWordInlineEffectExtent(replacement);
                ApplyBaselinePosition(replacement, metadata);
                if (numberedCellWidth.HasValue)
                    ValidateNumberedEquationPlacement(replacement, numberedCellWidth.Value);
                oldShape.Delete();
                if (selectReplacement) replacement.Range.Select();
                return replacement;
            }
            catch
            {
                try { replacement?.Delete(); } catch { }
                throw;
            }
        }

        internal bool TryGetSelectedBlock(out WordInterop.InlineShape shape, out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            if (application.Documents.Count == 0 || application.Selection == null) return false;
            var selection = application.Selection;
            if (selection.InlineShapes.Count == 1)
            {
                var candidate = selection.InlineShapes[1];
                if (TryReadContract(candidate, out metadata, out _))
                {
                    shape = candidate;
                    return true;
                }
            }
            return false;
        }

        internal static bool TryReadContract(WordInterop.InlineShape shape, out LaTeXBlockMetadata metadata,
            out string source)
        {
            metadata = null;
            source = null;
            if (shape == null) return false;
            try
            {
                // MathType and other embedded OLE objects are also InlineShapes, but
                // several picture metadata properties return E_NOTIMPL on those types.
                // Never probe Title/AlternativeText until the host object is a picture.
                if (!IsSupportedInlineShapeType(shape.Type)) return false;
                if (!LaTeXBlockMetadata.TryParse(shape.Title, out metadata)) return false;
                source = shape.AlternativeText;
                if (string.IsNullOrWhiteSpace(source)) { metadata = null; source = null; return false; }
                return true;
            }
            catch (COMException) { metadata = null; source = null; return false; }
            catch (NotImplementedException) { metadata = null; source = null; return false; }
        }

        internal static bool IsSupportedInlineShapeType(WordInterop.WdInlineShapeType type)
        {
            const int WordSvgInlineShapeType = 17; // Current Word value; absent from the shipped Office PIA enum.
            return type == WordInterop.WdInlineShapeType.wdInlineShapePicture ||
                   type == WordInterop.WdInlineShapeType.wdInlineShapeLinkedPicture ||
                   (int)type == WordSvgInlineShapeType;
        }

        internal int UpdateEquationNumbers(WordInterop.Document document = null)
        {
            document = document ?? application.ActiveDocument;
            if (document == null) return 0;

            var fields = document.StoryRanges[WordInterop.WdStoryType.wdMainTextStory].Fields;
            var equationFields = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                var field = fields[index];
                if (!IsEquationSequenceField(field)) continue;
                if (!field.Update())
                    throw new InvalidOperationException("Word could not update equation number " + (equationFields + 1) + ".");
                equationFields++;
            }
            return equationFields;
        }

        internal static bool IsEquationSequenceField(WordInterop.Field field)
        {
            if (field == null || field.Type != WordInterop.WdFieldType.wdFieldSequence) return false;
            var code = field.Code.Text ?? string.Empty;
            return Regex.IsMatch(code, "^\\s*SEQ\\s+" + EquationSequenceIdentifier + "(?:\\s|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        internal static string EquationBookmarkName(Guid id)
        {
            return EquationBookmarkPrefix + id.ToString("N");
        }

        internal static double SuggestedNumberedEquationWidth(WordInterop.Range target, double preferredWidthPt = 360)
        {
            if (target == null || !(preferredWidthPt > 0)) return preferredWidthPt;
            try
            {
                var page = target.Sections[1].PageSetup;
                var available = (double)page.PageWidth - page.LeftMargin - page.RightMargin;
                var columns = page.TextColumns;
                if (columns.Count > 1)
                    available = (available - columns.Spacing * (columns.Count - 1)) / columns.Count;
                available = Math.Max(36, available);
                var centerWidth = available * (100 - 2 * EquationSideColumnPercent) / 100.0;
                return Math.Max(36, Math.Min(preferredWidthPt, centerWidth - EquationSvgWidthSafetyPt));
            }
            catch
            {
                return preferredWidthPt;
            }
        }

        internal static double ReadSvgWidthPt(byte[] svgBytes)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new InvalidDataException("StemTeX returned an empty SVG.");
            var svg = Encoding.UTF8.GetString(svgBytes);
            var match = Regex.Match(svg,
                "<svg\\b[^>]*\\bwidth=(?<q>['\"])(?<value>[-+0-9.eE]+)\\s*(?<unit>[A-Za-z]*)\\k<q>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success || !double.TryParse(match.Groups["value"].Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var value) || !(value > 0))
                throw new InvalidDataException("StemTeX SVG has no positive physical width.");
            switch (match.Groups["unit"].Value.ToLowerInvariant())
            {
                case "pt": return value;
                case "bp": return value;
                case "px": return value * 72.0 / 96.0;
                case "in": return value * 72.0;
                case "cm": return value * 72.0 / 2.54;
                case "mm": return value * 72.0 / 25.4;
                case "pc": return value * 12.0;
                case "": return value * 72.0 / 96.0;
                default: throw new InvalidDataException("StemTeX SVG width uses an unsupported unit: " +
                    match.Groups["unit"].Value);
            }
        }

        internal static void ValidateNumberedEquationTarget(WordInterop.Range target)
        {
            if (target.Start != target.End)
                throw new InvalidOperationException("Place the insertion point on an empty equation line; numbered equations do not replace a selection.");
            if (target.StoryType != WordInterop.WdStoryType.wdMainTextStory)
                throw new InvalidOperationException("Numbered equations are currently supported only in the main document body.");
            if (Convert.ToBoolean(target.Information[WordInterop.WdInformation.wdWithInTable]))
                throw new InvalidOperationException("Nested numbered-equation tables are not supported in this version.");
            var paragraphText = (target.Paragraphs[1].Range.Text ?? string.Empty)
                .Replace("\r", string.Empty).Replace("\a", string.Empty);
            if (!string.IsNullOrWhiteSpace(paragraphText))
                throw new InvalidOperationException("Place the insertion point on an empty equation line before inserting a numbered equation.");
        }

        private static double NumberedEquationCellWidth(WordInterop.InlineShape shape)
        {
            var range = shape.Range;
            if (!Convert.ToBoolean(range.Information[WordInterop.WdInformation.wdWithInTable]) ||
                range.Cells.Count != 1)
                throw new InvalidOperationException("A numbered equation must remain in its equation table.");
            var cell = range.Cells[1];
            if (cell.ColumnIndex != 2 || cell.RowIndex != 1 || cell.Range.Tables.Count != 1 ||
                cell.Range.Tables[1].Rows.Count != 1 || cell.Range.Tables[1].Columns.Count != 3)
                throw new InvalidOperationException("The numbered equation table structure is no longer valid.");
            return cell.Width;
        }

        private static void ValidateNumberedEquationPlacement(WordInterop.InlineShape shape, double maximumWidthPt)
        {
            NumberedEquationCellWidth(shape);
            if (shape.Width > maximumWidthPt + 0.5)
                throw new InvalidOperationException("The rendered formula is wider than the numbered equation area. Reduce Block width or shorten the formula.");
        }

        private static void ConfigureNumberedEquationTable(WordInterop.Table table, double fontSizePt)
        {
            table.Borders.Enable = 0;
            table.AllowAutoFit = false;
            table.PreferredWidthType = WordInterop.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100;
            table.Rows.Alignment = WordInterop.WdRowAlignment.wdAlignRowCenter;
            table.Rows.AllowBreakAcrossPages = 0;
            table.Rows.SetLeftIndent(0, WordInterop.WdRulerStyle.wdAdjustNone);
            table.TopPadding = 0;
            table.BottomPadding = 0;
            table.LeftPadding = 0;
            table.RightPadding = 0;
            table.Spacing = 0;

            table.Columns[1].PreferredWidthType = WordInterop.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[1].PreferredWidth = EquationSideColumnPercent;
            table.Columns[2].PreferredWidthType = WordInterop.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[2].PreferredWidth = 100 - 2 * EquationSideColumnPercent;
            table.Columns[3].PreferredWidthType = WordInterop.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[3].PreferredWidth = EquationSideColumnPercent;

            table.Cell(1, 1).VerticalAlignment = WordInterop.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, 2).VerticalAlignment = WordInterop.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, 3).VerticalAlignment = WordInterop.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, 1).Range.ParagraphFormat.Alignment = WordInterop.WdParagraphAlignment.wdAlignParagraphLeft;
            table.Cell(1, 2).Range.ParagraphFormat.Alignment = WordInterop.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(1, 3).Range.ParagraphFormat.Alignment = WordInterop.WdParagraphAlignment.wdAlignParagraphRight;
            for (var column = 1; column <= 3; column++)
            {
                var paragraph = table.Cell(1, column).Range.ParagraphFormat;
                paragraph.LeftIndent = 0;
                paragraph.RightIndent = 0;
                paragraph.FirstLineIndent = 0;
                paragraph.SpaceBefore = 0;
                paragraph.SpaceAfter = 0;
                paragraph.LineSpacingRule = WordInterop.WdLineSpacing.wdLineSpaceSingle;
            }
            if (fontSizePt >= 1 && fontSizePt <= 200)
                table.Cell(1, 3).Range.Font.Size = (float)fontSizePt;
        }

        private static void AddEquationNumber(WordInterop.Cell numberCell, Guid blockId)
        {
            var document = numberCell.Range.Document;
            var start = numberCell.Range.Start;
            document.Range(start, start).Text = "(";
            var fieldRange = document.Range(start + 1, start + 1);
            var field = document.Fields.Add(fieldRange, WordInterop.WdFieldType.wdFieldSequence,
                EquationSequenceIdentifier + " \\* ARABIC", false);
            if (!field.Update())
                throw new InvalidOperationException("Word could not create the equation number field.");
            document.Bookmarks.Add(EquationBookmarkName(blockId), field.Result);
            document.Range(field.Result.End + 1, field.Result.End + 1).Text = ")";
        }

        private static void InsertTableSeparatorAfter(WordInterop.Table table)
        {
            var document = table.Range.Document;
            var separatorStart = table.Range.End;
            document.Range(separatorStart, separatorStart).Text = "\r";
            var separator = document.Range(separatorStart, separatorStart + 1);
            separator.Font.Size = 1;
            separator.ParagraphFormat.SpaceBefore = 0;
            separator.ParagraphFormat.SpaceAfter = 0;
            separator.ParagraphFormat.LineSpacingRule = WordInterop.WdLineSpacing.wdLineSpaceExactly;
            separator.ParagraphFormat.LineSpacing = 1;
        }

        private static void ApplyContract(WordInterop.InlineShape shape, string source, LaTeXBlockMetadata metadata)
        {
            shape.AlternativeText = source;
            shape.Title = metadata.ToString();
            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            ApplyBaselinePosition(shape, metadata);
        }

        private static void ApplyBaselinePosition(WordInterop.InlineShape shape, LaTeXBlockMetadata metadata)
        {
            // Word aligns the bottom of an InlineShape to the text baseline. Move the image
            // character down by the TeX box depth. This is always the TeX/Western baseline:
            // CJK glyph extents inside the SVG do not define a second alignment reference.
            // Word persists this API value as whole points.
            shape.Range.Font.Position = -(int)Math.Round(metadata.DepthPt, MidpointRounding.AwayFromZero);
        }

        private static WordInterop.InlineShape RemoveWordInlineEffectExtent(WordInterop.InlineShape shape)
        {
            // AddPicture wraps an SVG in wp:inline and gives it a bottom effect extent
            // (typically one CSS pixel). Word includes that host-only extent in inline
            // baseline layout even though it is not part of the SVG or the TeX box.
            // Reinsert the same Flat OPC object with b=0, then remove InsertXML's
            // temporary paragraph boundary. The SVG, metadata and TeX depth are unchanged.
            var flatOpc = shape.Range.WordOpenXML;
            var effect = Regex.Match(flatOpc,
                "<wp:effectExtent\\b(?=[^>]*\\bb=\"(?<bottom>[0-9]+)\")[^>]*/>",
                RegexOptions.CultureInvariant);
            if (!effect.Success || effect.Groups["bottom"].Value == "0") return shape;

            var patched = Regex.Replace(flatOpc,
                "(<wp:effectExtent\\b[^>]*\\bb=\")[^\"]+(\"[^>]*/>)", "${1}0${2}",
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            var originalStart = shape.Range.Start;
            var document = shape.Range.Document;
            var insertion = document.Range(originalStart, originalStart);
            shape.Delete();
            try
            {
                insertion.InsertXML(patched);
            }
            catch
            {
                // Keep the user's formula recoverable if Word rejects the normalized
                // package in an unusual story or protected range.
                insertion.InsertXML(flatOpc);
                throw;
            }

            var replacementRange = document.Range(originalStart, originalStart + 1);
            if (replacementRange.InlineShapes.Count != 1)
                throw new InvalidDataException("Word did not reinsert the normalized inline SVG.");
            var replacement = replacementRange.InlineShapes[1];

            var separator = document.Range(replacement.Range.End, replacement.Range.End + 1);
            if (separator.Text == "\r") separator.Delete();
            return replacement;
        }

        internal static double ResolveFontSize(WordInterop.Range target, LaTeXBlockLayoutMode mode, double fallback)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return fallback;
            var fontSize = (double)target.Font.Size;
            if (fontSize < 1 || fontSize > 200) return fallback;
            return fontSize;
        }

        internal static bool ShouldRefreshForHostFontSizeChange(double selectedSizePt,
            double currentSizePt, double renderedSizePt)
        {
            return currentSizePt >= 1 && currentSizePt <= 200 &&
                   Math.Abs(selectedSizePt - currentSizePt) > 0.001 &&
                   Math.Abs(renderedSizePt - currentSizePt) > 0.001;
        }

        private string PrepareInsertionSvg(LaTeXBlockRender render, LaTeXBlockLayoutMode mode)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return render.SvgPath;
            return WriteSvg(ApplyFractionalBaselineCompensation(render.SvgBytes, render.DepthPt, 1));
        }

        internal static byte[] ApplyFractionalBaselineCompensation(byte[] svgBytes, double depthPt, double scale)
        {
            if (svgBytes == null) throw new ArgumentNullException(nameof(svgBytes));
            if (!(scale > 0)) throw new ArgumentOutOfRangeException(nameof(scale));

            var scaledDepth = depthPt * scale;
            var wordDepth = Math.Round(scaledDepth, MidpointRounding.AwayFromZero);
            // Word supplies the whole-point component. Shift the SVG viewport so the
            // remaining fraction moves the TeX baseline by exactly the residual amount.
            // wp:effectExtent is deliberately excluded: Word derives it from the image
            // and it is not a stable part of the inline character's baseline metric.
            var residualInSvgUnits = (scaledDepth - wordDepth) / scale;
            var svg = Encoding.UTF8.GetString(svgBytes);
            if (Math.Abs(residualInSvgUnits) < 0.000001) return svgBytes;

            var viewBox = Regex.Match(svg,
                "\\bviewBox=(?<q>['\"])(?<x>[-+0-9.eE]+)\\s+(?<y>[-+0-9.eE]+)\\s+(?<w>[-+0-9.eE]+)\\s+(?<h>[-+0-9.eE]+)\\k<q>",
                RegexOptions.CultureInvariant);
            if (!viewBox.Success) throw new InvalidDataException("StemTeX SVG has no numeric viewBox for fractional baseline compensation.");

            var number = CultureInfo.InvariantCulture;
            var top = double.Parse(viewBox.Groups["y"].Value, number);
            var correctedTop = top - residualInSvgUnits;
            var quote = viewBox.Groups["q"].Value;
            var correctedViewBox = "viewBox=" + quote + viewBox.Groups["x"].Value + " " +
                correctedTop.ToString("0.######", number) + " " + viewBox.Groups["w"].Value + " " +
                viewBox.Groups["h"].Value + quote;
            svg = svg.Remove(viewBox.Index, viewBox.Length).Insert(viewBox.Index, correctedViewBox);
            svg = Regex.Replace(svg, "<svg\\b", "<svg data-latexblocks-baseline-residual='" +
                (residualInSvgUnits * scale).ToString("0.######", number) + "pt'", RegexOptions.CultureInvariant,
                TimeSpan.FromSeconds(1));
            return Encoding.UTF8.GetBytes(svg);
        }

        private string WriteSvg(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) throw new InvalidDataException("StemTeX returned an empty SVG.");
            var path = Path.Combine(cacheDirectory, Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        private void EnsureDocument()
        {
            if (application.Documents.Count == 0)
                throw new InvalidOperationException("Open a Word document before inserting a LaTeX Block.");
        }
    }

    internal sealed class LaTeXBlockRender
    {
        internal LaTeXBlockRender(string svgPath, byte[] svgBytes, double depthPt, double fontSizePt)
        {
            SvgPath = svgPath;
            SvgBytes = svgBytes ?? throw new ArgumentNullException(nameof(svgBytes));
            DepthPt = depthPt;
            FontSizePt = fontSizePt;
        }
        internal string SvgPath { get; }
        internal byte[] SvgBytes { get; }
        internal double DepthPt { get; }
        internal double FontSizePt { get; }
    }
}
