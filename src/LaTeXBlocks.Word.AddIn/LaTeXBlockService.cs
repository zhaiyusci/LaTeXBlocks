using System;
using System.Collections.Generic;
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
        private const double EquationNumberReservePt = 36.0;
        private const double EquationNumberGapPt = 6.0;
        private const double EmusPerPoint = 12700.0;
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
            double fontSizePt = 10, bool displayMathStyle = false)
        {
            return RenderPreviewAsync(source, widthPt, mode, profile, fontSizePt, displayMathStyle)
                .GetAwaiter().GetResult();
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double fontSizePt = 10, bool displayMathStyle = false)
        {
            var normalizedSource = NormalizeSourceText(source);
            var renderSource = displayMathStyle ? PrepareDisplayMathSource(normalizedSource) : normalizedSource;
            var result = await renderers.RenderLatestAsync(profile, renderSource, widthPt,
                mode == LaTeXBlockLayoutMode.Auto, fontSizePt);
            return new LaTeXBlockRender(WriteSvg(result.Bytes), result.Bytes, result.DepthPt, fontSizePt);
        }

        internal WordInterop.InlineShape InsertBlock(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile)
        {
            EnsureDocument();
            var fontSizePt = ResolveFontSize(application.Selection, mode, 10);
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
            var naturalSpaces = MeasureNaturalAdjacentSpaces(target, metadata);
            var document = target.Document;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Insert LaTeX Block");
                undoStarted = true;
                return InsertRenderedAt(target, source, mode, render, metadata, naturalSpaces, true,
                    () => documentMutated = true);
            }
            catch
            {
                if (undoStarted)
                {
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
                    undoStarted = false;
                    if (documentMutated)
                        try { document.Undo(); } catch { }
                }
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        internal WordInterop.InlineShape InsertNumberedBlock(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile)
        {
            EnsureDocument();
            if (mode != LaTeXBlockLayoutMode.Auto)
                throw new InvalidOperationException("Numbered equations use natural-width display math.");
            var fontSizePt = ResolveFontSize(application.Selection, mode, 10);
            var render = RenderPreview(source, widthPt, mode, profile, fontSizePt, true);
            return InsertNumberedRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertNumberedRendered(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (mode != LaTeXBlockLayoutMode.Auto)
                throw new InvalidOperationException("Numbered equations use natural-width display math.");

            var document = application.ActiveDocument;
            var target = application.Selection.Range.Duplicate;
            ValidateNumberedEquationTarget(target);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var layout = GetNumberedEquationLayout(target, render.FontSizePt);
            ValidateNumberedEquationWidth(render.SvgBytes, layout);
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Insert Numbered Equation");
                undoStarted = true;
                documentMutated = true;
                ConfigureNumberedEquationTabs(target.Paragraphs[1], layout);

                var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt, mode, render.FontSizePt,
                    LaTeXBlockRole.NumberedEquation);
                var leadingBreak = NeedsManualBreakBefore(target) ? "\v" : string.Empty;
                var trailingBreak = NeedsManualBreakAfter(target) ? "\v" : string.Empty;
                var scaffoldStart = target.Start;
                document.Range(scaffoldStart, scaffoldStart).Text = leadingBreak + "\t\t()" + trailingBreak;
                var formulaPosition = scaffoldStart + leadingBreak.Length + 1;
                var formulaTarget = document.Range(formulaPosition, formulaPosition);
                var shape = InsertRenderedAt(formulaTarget, source, mode, render, metadata,
                    default(InlineSpaceAdvances), false);
                var fieldPosition = shape.Range.End + 2; // after the second tab and literal '('
                var field = document.Fields.Add(document.Range(fieldPosition, fieldPosition),
                    WordInterop.WdFieldType.wdFieldSequence,
                    EquationSequenceIdentifier + " \\* ARABIC", false);
                if (!field.Update())
                    throw new InvalidOperationException("Word could not create the equation number field.");
                document.Bookmarks.Add(EquationBookmarkName(metadata.Id), field.Result);
                UpdateEquationNumbers(document);
                ValidateNumberedEquationPlacement(shape, render.SvgBytes, render.FontSizePt);
                MoveCaretAfterNumberedEquation(field);
                return shape;
            }
            catch
            {
                if (undoStarted)
                {
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
                    undoStarted = false;
                    if (documentMutated)
                        try { document.Undo(); } catch { }
                }
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        private WordInterop.InlineShape InsertRenderedAt(WordInterop.Range requestedTarget, string source,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, LaTeXBlockMetadata metadata,
            InlineSpaceAdvances naturalSpaces, bool select,
            Action markDocumentMutated = null)
        {
            var target = requestedTarget.Duplicate;
            var hostPosition = ResolveRangePosition(target, 0);
            var replacesText = target.Start != target.End;
            target.Text = string.Empty;
            if (replacesText) markDocumentMutated?.Invoke();
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var insertionPath = PrepareInsertionSvg(render, mode);
            var svgSize = ReadSvgPhysicalSize(render.SvgBytes);
            var shape = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
            markDocumentMutated?.Invoke();
            ApplyContract(shape, source, metadata, hostPosition);
            var sideEffects = MeasureInlineSpaceEffectExtents(shape, metadata, naturalSpaces);
            shape = NormalizeWordInlineDrawing(shape, sideEffects, svgSize);
            ApplyHostRunFormat(shape, metadata, hostPosition);
            shape = ReconcileInsertedInlineSpaceEffectExtents(shape, metadata, sideEffects,
                svgSize, hostPosition);
            if (select) MoveCaretAfterRange(shape.Range);
            return shape;
        }

        internal WordInterop.InlineShape UpdateBlock(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double? fontSizePt = null, bool selectReplacement = true)
        {
            var size = fontSizePt ?? ResolveFontSize(oldShape.Range, mode, 10);
            var displayMathStyle = TryReadContract(oldShape, out var metadata, out _) &&
                                   metadata.Role == LaTeXBlockRole.NumberedEquation;
            var render = RenderPreview(source, widthPt, mode, profile, size, displayMathStyle);
            return UpdateRendered(oldShape, source, widthPt, mode, render, selectReplacement);
        }

        internal static string PrepareDisplayMathSource(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
                throw new ArgumentException("LaTeX source cannot be empty.", nameof(source));

            var body = NormalizeSourceText(source).Trim();
            if (TryStripOuter(body, "\\[", "\\]", out var stripped) ||
                TryStripOuter(body, "\\(", "\\)", out stripped) ||
                TryStripOuter(body, "$$", "$$", out stripped) ||
                TryStripOuter(body, "$", "$", out stripped) ||
                TryStripEnvironment(body, "displaymath", out stripped) ||
                TryStripEnvironment(body, "equation", out stripped) ||
                TryStripEnvironment(body, "equation*", out stripped))
                body = stripped;
            else if (TryStripEnvironment(body, "align", out stripped) ||
                     TryStripEnvironment(body, "align*", out stripped))
                body = "\\begin{aligned}\n" + stripped + "\n\\end{aligned}";
            else if (TryStripEnvironment(body, "gather", out stripped) ||
                     TryStripEnvironment(body, "gather*", out stripped))
                body = "\\begin{gathered}\n" + stripped + "\n\\end{gathered}";

            var uncommented = StemTeXRenderer.RemoveTeXCommentsForDetection(body);
            if (Regex.IsMatch(uncommented, "\\\\(?:tag\\*?\\s*\\{|notag\\b|nonumber\\b)",
                    RegexOptions.CultureInvariant))
                throw new ArgumentException(
                    "TeX-side equation tags and number suppression cannot be combined with Word-owned numbering.",
                    nameof(source));

            if (TryStripEnvironment(body, "split", out stripped))
                body = "\\begin{aligned}\n" + stripped + "\n\\end{aligned}";
            else if (TryStripEnvironment(body, "alignat", out stripped) ||
                     TryStripEnvironment(body, "alignat*", out stripped))
                body = "\\begin{alignedat}\n" + stripped + "\n\\end{alignedat}";
            else if (Regex.IsMatch(uncommented,
                         "^\\s*\\\\begin\\s*\\{\\s*(?:multline\\*?|flalign\\*?|minipage|document)\\s*\\}",
                         RegexOptions.CultureInvariant) ||
                     uncommented.IndexOf("\\[", StringComparison.Ordinal) >= 0 ||
                     uncommented.IndexOf("$$", StringComparison.Ordinal) >= 0 ||
                     Regex.IsMatch(uncommented,
                         "\\\\begin\\s*\\{\\s*(?:displaymath|equation\\*?|align\\*?|gather\\*?)\\s*\\}",
                         RegexOptions.CultureInvariant))
                throw new ArgumentException(
                    "The outer display environment must enclose the entire numbered-equation source, and page-width environments are not supported.",
                    nameof(source));

            if (string.IsNullOrWhiteSpace(body))
                throw new ArgumentException("Display math cannot be empty.", nameof(source));

            // This wrapper exists only in the render request. Alternative Text keeps the
            // canonical source supplied by the user. \displaystyle changes TeX math style;
            // Word, not TeX, owns the display line and its horizontal placement.
            return "\\(\n\\displaystyle\n" + body + "\n\\)";
        }

        private static bool TryStripOuter(string source, string open, string close, out string body)
        {
            body = null;
            if (!source.StartsWith(open, StringComparison.Ordinal) ||
                !source.EndsWith(close, StringComparison.Ordinal) ||
                source.Length < open.Length + close.Length) return false;
            body = source.Substring(open.Length, source.Length - open.Length - close.Length).Trim();
            return true;
        }

        internal static string NormalizeSourceText(string source)
        {
            if (source == null) return null;
            return source.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        private static bool TryStripEnvironment(string source, string environment, out string body)
        {
            body = null;
            var match = Regex.Match(source,
                "^\\s*\\\\begin\\s*\\{\\s*" + Regex.Escape(environment) +
                "\\s*\\}(?<body>[\\s\\S]*)\\\\end\\s*\\{\\s*" + Regex.Escape(environment) +
                "\\s*\\}\\s*$", RegexOptions.CultureInvariant);
            if (!match.Success) return false;
            body = match.Groups["body"].Value.Trim();
            return true;
        }

        internal WordInterop.InlineShape UpdateRendered(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, bool selectReplacement = true)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!TryReadContract(oldShape, out var previous, out _))
                throw new InvalidOperationException("The selected image is not a LaTeX Block.");
            var numbered = previous.Role == LaTeXBlockRole.NumberedEquation;
            var numberedLayout = default(NumberedEquationLayout);
            WordInterop.Field numberedField = null;
            if (numbered)
            {
                if (mode != LaTeXBlockLayoutMode.Auto)
                    throw new InvalidOperationException("Numbered equations use natural-width display math.");
                ValidateNumberedEquationPlacement(oldShape, render.SvgBytes, render.FontSizePt);
                numberedLayout = GetNumberedEquationLayout(oldShape.Range, render.FontSizePt);
                numberedField = FindEquationNumberField(oldShape, previous);
            }

            var target = oldShape.Range.Duplicate;
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt, mode, render.FontSizePt,
                previous.Role);
            var hostPosition = ResolveHostBaselinePosition(oldShape, previous);
            var naturalSpaces = ResolveUpdatedNaturalAdjacentSpaces(oldShape, previous);
            var sideEffects = MeasureInlineSpaceEffectExtents(oldShape, previous, naturalSpaces);
            var svgSize = ReadSvgPhysicalSize(render.SvgBytes);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            WordInterop.InlineShape replacement = null;
            var document = oldShape.Range.Document;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Update LaTeX Block");
                undoStarted = true;
                if (numbered)
                {
                    documentMutated = true;
                    ConfigureNumberedEquationTabs(oldShape.Range.Paragraphs[1], numberedLayout);
                }
                var insertionPath = PrepareInsertionSvg(render, mode);
                replacement = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
                documentMutated = true;
                ApplyContract(replacement, source, metadata, hostPosition);
                replacement = NormalizeWordInlineDrawing(replacement, sideEffects, svgSize);
                ApplyHostRunFormat(replacement, metadata, hostPosition);
                oldShape.Delete();
                if (selectReplacement)
                {
                    if (numbered) MoveCaretAfterNumberedEquation(numberedField);
                    else MoveCaretAfterRange(replacement.Range);
                }
                return replacement;
            }
            catch
            {
                if (undoStarted)
                {
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
                    undoStarted = false;
                    if (documentMutated)
                        try { document.Undo(); } catch { try { replacement?.Delete(); } catch { } }
                }
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
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

            var mainStory = document.StoryRanges[WordInterop.WdStoryType.wdMainTextStory];
            var reflowedParagraphs = new HashSet<int>();
            foreach (WordInterop.InlineShape shape in mainStory.InlineShapes)
            {
                if (!TryReadContract(shape, out var metadata, out _) ||
                    metadata.Role != LaTeXBlockRole.NumberedEquation || shape.Range.Tables.Count > 0)
                    continue;
                var paragraph = shape.Range.Paragraphs[1];
                if (!reflowedParagraphs.Add(paragraph.Range.Start)) continue;
                ConfigureNumberedEquationTabs(paragraph,
                    GetNumberedEquationLayout(shape.Range, metadata.FontSizePt));
            }

            var fields = mainStory.Fields;
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

        private static WordInterop.Field FindEquationNumberField(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata)
        {
            var document = shape.Range.Document;
            var bookmarkName = EquationBookmarkName(metadata.Id);
            if (!document.Bookmarks.Exists(bookmarkName))
                throw new InvalidOperationException("The numbered equation has lost its Word bookmark.");
            var bookmark = document.Bookmarks[bookmarkName].Range;
            foreach (WordInterop.Field field in shape.Range.Paragraphs[1].Range.Fields)
                if (IsEquationSequenceField(field) && field.Result.Start == bookmark.Start &&
                    field.Result.End == bookmark.End) return field;
            throw new InvalidOperationException("The equation bookmark no longer identifies its SEQ field.");
        }

        internal static double ReadSvgWidthPt(byte[] svgBytes)
        {
            return ReadSvgLengthPt(svgBytes, "width");
        }

        internal static double ReadSvgHeightPt(byte[] svgBytes)
        {
            return ReadSvgLengthPt(svgBytes, "height");
        }

        private static SvgPhysicalSize ReadSvgPhysicalSize(byte[] svgBytes)
        {
            return new SvgPhysicalSize(ReadSvgWidthPt(svgBytes), ReadSvgHeightPt(svgBytes));
        }

        private static double ReadSvgLengthPt(byte[] svgBytes, string attribute)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new InvalidDataException("StemTeX returned an empty SVG.");
            var svg = Encoding.UTF8.GetString(svgBytes);
            var match = Regex.Match(svg,
                "<svg\\b[^>]*\\b" + Regex.Escape(attribute) +
                "=(?<q>['\"])(?<value>[-+0-9.eE]+)\\s*(?<unit>[A-Za-z]*)\\k<q>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success || !double.TryParse(match.Groups["value"].Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var value) || !(value > 0) ||
                double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidDataException("StemTeX SVG has no positive physical " + attribute + ".");
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
                default: throw new InvalidDataException("StemTeX SVG " + attribute + " uses an unsupported unit: " +
                    match.Groups["unit"].Value);
            }
        }

        internal static void ValidateNumberedEquationTarget(WordInterop.Range target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (target.Start != target.End)
                throw new InvalidOperationException("Place a collapsed insertion point where the numbered equation belongs.");
            if (target.StoryType != WordInterop.WdStoryType.wdMainTextStory)
                throw new InvalidOperationException("Numbered equations are currently supported only in the main document body.");
            if (Convert.ToBoolean(target.Information[WordInterop.WdInformation.wdWithInTable]))
                throw new InvalidOperationException("Numbered equations inside Word tables are not supported in this version.");
            if (target.Paragraphs[1].Range.ParagraphFormat.LineSpacingRule ==
                WordInterop.WdLineSpacing.wdLineSpaceExactly)
                throw new InvalidOperationException(
                    "A same-paragraph display equation cannot expand a line with Exact line spacing. Use Single, At least, or Multiple line spacing first.");
            ValidateParagraphTabOwnership(target.Paragraphs[1]);
            ValidateParagraphTabStops(target.Paragraphs[1]);
            ValidateEquationInsertionPoint(target);
        }

        private static void ValidateParagraphTabOwnership(WordInterop.Paragraph paragraph)
        {
            var range = paragraph.Range;
            var paragraphEnd = range.End;
            var search = range.Duplicate;
            search.Find.ClearFormatting();
            while (search.Find.Execute(FindText: "\t", MatchCase: false, MatchWholeWord: false,
                       MatchWildcards: false, Forward: true, Wrap: WordInterop.WdFindWrap.wdFindStop,
                       Format: false))
            {
                var tabStart = search.Start;
                var belongsToEquation = IsNumberedShapeAt(range.Document, tabStart + 1) ||
                                        IsNumberedShapeAt(range.Document, tabStart - 1);
                if (!belongsToEquation)
                    throw new InvalidOperationException(
                        "This paragraph already uses tabs for ordinary content. A numbered equation owns its paragraph's center and right tab stops.");
                search.SetRange(search.End, paragraphEnd);
            }
        }

        private static bool IsNumberedShapeAt(WordInterop.Document document, int position)
        {
            if (position < 0 || position >= document.Content.End) return false;
            var candidate = document.Range(position, position + 1);
            if (candidate.InlineShapes.Count != 1) return false;
            return TryReadContract(candidate.InlineShapes[1], out var metadata, out _) &&
                   metadata.Role == LaTeXBlockRole.NumberedEquation;
        }

        private static void ValidateParagraphTabStops(WordInterop.Paragraph paragraph)
        {
            var numberedShapes = 0;
            foreach (WordInterop.InlineShape shape in paragraph.Range.InlineShapes)
                if (TryReadContract(shape, out var metadata, out _) &&
                    metadata.Role == LaTeXBlockRole.NumberedEquation) numberedShapes++;

            var tabs = paragraph.Range.ParagraphFormat.TabStops;
            var customTabs = 0;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                if (!tabs[index].CustomTab) continue;
                customTabs++;
                hasCenter |= tabs[index].Alignment == WordInterop.WdTabAlignment.wdAlignTabCenter;
                hasRight |= tabs[index].Alignment == WordInterop.WdTabAlignment.wdAlignTabRight;
            }

            if (customTabs == 0)
            {
                if (numberedShapes > 0)
                    throw new InvalidOperationException("The existing numbered-equation paragraph has lost its tab stops.");
                return;
            }
            if (numberedShapes == 0)
                throw new InvalidOperationException(
                    "This paragraph already has custom tab stops. A numbered equation must own its center and right tab layout.");
            if (customTabs != 2)
                throw new InvalidOperationException("The existing numbered-equation paragraph has conflicting tab stops.");
            if (!hasCenter || !hasRight)
                throw new InvalidOperationException("The existing numbered-equation paragraph has conflicting tab stops.");
        }

        private static void ValidateEquationInsertionPoint(WordInterop.Range target)
        {
            foreach (WordInterop.InlineShape shape in target.Paragraphs[1].Range.InlineShapes)
            {
                if (!TryReadContract(shape, out var metadata, out _) ||
                    metadata.Role != LaTeXBlockRole.NumberedEquation) continue;
                var line = NumberedEquationLineRange(shape);
                if (target.Start > line.Start && target.Start < line.End)
                    throw new InvalidOperationException(
                        "Place the insertion point before or after an existing numbered-equation line, not inside its scaffold.");
            }
        }

        private static void ValidateNumberedEquationPlacement(WordInterop.InlineShape shape, byte[] svgBytes,
            double fontSizePt)
        {
            var range = shape.Range;
            if (Convert.ToBoolean(range.Information[WordInterop.WdInformation.wdWithInTable]))
                throw new InvalidOperationException(
                    "This numbered equation uses the retired table host. Reinsert it with the current command.");
            if (range.Paragraphs.Count != 1)
                throw new InvalidOperationException("The numbered equation no longer belongs to one Word paragraph.");

            var document = range.Document;
            if (range.Start <= range.Paragraphs[1].Range.Start ||
                document.Range(range.Start - 1, range.Start).Text != "\t" ||
                document.Range(range.End, range.End + 1).Text != "\t")
                throw new InvalidOperationException("The numbered equation's tab scaffold is no longer valid.");

            if (!TryReadContract(shape, out var metadata, out _) ||
                metadata.Role != LaTeXBlockRole.NumberedEquation)
                throw new InvalidOperationException("The image is not a numbered LaTeX equation.");
            var bookmarkName = EquationBookmarkName(metadata.Id);
            if (!document.Bookmarks.Exists(bookmarkName))
                throw new InvalidOperationException("The numbered equation has lost its Word bookmark.");
            var bookmark = document.Bookmarks[bookmarkName].Range;
            if (bookmark.Paragraphs.Count != 1 ||
                bookmark.Paragraphs[1].Range.Start != range.Paragraphs[1].Range.Start)
                throw new InvalidOperationException("The equation number no longer belongs to the formula's paragraph.");

            var matchingField = false;
            foreach (WordInterop.Field field in range.Paragraphs[1].Range.Fields)
                if (IsEquationSequenceField(field) && field.Result.Start == bookmark.Start &&
                    field.Result.End == bookmark.End) { matchingField = true; break; }
            if (!matchingField)
                throw new InvalidOperationException("The equation bookmark no longer identifies its SEQ field.");

            ValidateNumberedEquationWidth(svgBytes, GetNumberedEquationLayout(range, fontSizePt));
        }

        private static void ValidateNumberedEquationWidth(byte[] svgBytes, NumberedEquationLayout layout)
        {
            var renderedWidth = ReadSvgWidthPt(svgBytes);
            if (renderedWidth <= layout.MaximumFormulaWidthPt + 0.5) return;
            throw new InvalidOperationException("The natural formula width (" +
                renderedWidth.ToString("0.#", CultureInfo.InvariantCulture) +
                " pt) leaves no safe space for the equation number in this text column (maximum " +
                layout.MaximumFormulaWidthPt.ToString("0.#", CultureInfo.InvariantCulture) + " pt).");
        }

        private static NumberedEquationLayout GetNumberedEquationLayout(WordInterop.Range target,
            double fontSizePt)
        {
            var page = target.Sections[1].PageSetup;
            var columnWidth = (double)page.PageWidth - page.LeftMargin - page.RightMargin;
            var columns = page.TextColumns;
            if (columns.Count > 1)
                columnWidth = (columnWidth - columns.Spacing * (columns.Count - 1)) / columns.Count;

            if (columnWidth < 72)
                throw new InvalidOperationException("The current paragraph is too narrow for a numbered equation.");
            var numberReserve = Math.Max(EquationNumberReservePt, fontSizePt * 3);
            var maximumFormulaWidth = columnWidth - 2 * (numberReserve + EquationNumberGapPt);
            if (maximumFormulaWidth < 36)
                throw new InvalidOperationException("The current paragraph is too narrow for a centered formula and equation number.");
            // Word stores ordinary custom tab positions as absolute offsets from the
            // text column's left edge. Baking paragraph indents into those positions
            // leaves stale tabs when the user later changes indentation. Display
            // equations therefore own the full column: formula at its center, number
            // at its right edge, independent of running-text paragraph indents.
            return new NumberedEquationLayout(columnWidth / 2, columnWidth, maximumFormulaWidth);
        }

        private static void ConfigureNumberedEquationTabs(WordInterop.Paragraph paragraph,
            NumberedEquationLayout layout)
        {
            var tabs = paragraph.Range.ParagraphFormat.TabStops;
            tabs.ClearAll();
            tabs.Add((float)layout.CenterTabPt, WordInterop.WdTabAlignment.wdAlignTabCenter,
                WordInterop.WdTabLeader.wdTabLeaderSpaces);
            tabs.Add((float)layout.RightTabPt, WordInterop.WdTabAlignment.wdAlignTabRight,
                WordInterop.WdTabLeader.wdTabLeaderSpaces);
        }

        private static bool NeedsManualBreakBefore(WordInterop.Range target)
        {
            var paragraphStart = target.Paragraphs[1].Range.Start;
            if (target.Start <= paragraphStart) return false;
            var previous = target.Document.Range(target.Start - 1, target.Start).Text;
            return previous != "\v" && previous != "\r";
        }

        private static bool NeedsManualBreakAfter(WordInterop.Range target)
        {
            var paragraphEnd = target.Paragraphs[1].Range.End - 1;
            if (target.Start >= paragraphEnd) return false;
            var next = target.Document.Range(target.Start, target.Start + 1).Text;
            return next != "\v" && next != "\r";
        }

        internal static WordInterop.Range NumberedEquationLineRange(WordInterop.InlineShape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var paragraph = shape.Range.Paragraphs[1].Range;
            var document = shape.Range.Document;
            var start = shape.Range.Start;
            while (start > paragraph.Start)
            {
                var previous = document.Range(start - 1, start).Text;
                start--;
                if (previous == "\v") break;
            }
            var end = shape.Range.End;
            while (end < paragraph.End - 1)
            {
                var next = document.Range(end, end + 1).Text;
                end++;
                if (next == "\v") break;
            }

            var hasLeadingBreak = start < paragraph.End && document.Range(start, start + 1).Text == "\v";
            var hasTrailingBreak = end > paragraph.Start && document.Range(end - 1, end).Text == "\v";
            var nextLineIsEquation = hasTrailingBreak && end + 1 < paragraph.End &&
                                     document.Range(end, end + 1).Text == "\t" &&
                                     IsNumberedShapeAt(document, end + 1);
            if (nextLineIsEquation)
            {
                // Keep the shared line boundary so the following display does not join
                // the preceding running-text line.
                end--;
            }
            else if (hasLeadingBreak && PreviousVisualLineIsEquation(document, paragraph, start))
            {
                // Keep the shared boundary before this line when the previous visual
                // line is another numbered equation.
                start++;
            }
            return document.Range(start, end);
        }

        private static bool PreviousVisualLineIsEquation(WordInterop.Document document,
            WordInterop.Range paragraph, int lineBreakPosition)
        {
            var visualStart = paragraph.Start;
            for (var position = lineBreakPosition - 1; position >= paragraph.Start; position--)
                if (document.Range(position, position + 1).Text == "\v")
                {
                    visualStart = position + 1;
                    break;
                }
            return visualStart + 1 < lineBreakPosition &&
                   document.Range(visualStart, visualStart + 1).Text == "\t" &&
                   IsNumberedShapeAt(document, visualStart + 1);
        }

        private struct NumberedEquationLayout
        {
            internal NumberedEquationLayout(double centerTabPt, double rightTabPt,
                double maximumFormulaWidthPt)
            {
                CenterTabPt = centerTabPt;
                RightTabPt = rightTabPt;
                MaximumFormulaWidthPt = maximumFormulaWidthPt;
            }

            internal double CenterTabPt { get; }
            internal double RightTabPt { get; }
            internal double MaximumFormulaWidthPt { get; }
        }

        private static void ApplyContract(WordInterop.InlineShape shape, string source, LaTeXBlockMetadata metadata,
            int hostPosition)
        {
            shape.AlternativeText = NormalizeSourceText(source);
            shape.Title = metadata.ToString();
            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            ApplyHostRunFormat(shape, metadata, hostPosition);
        }

        private static void ApplyHostRunFormat(WordInterop.InlineShape shape, LaTeXBlockMetadata metadata,
            int hostPosition)
        {
            // The image's physical dimensions already come from the SVG. This run size is
            // Word's semantic host size, used by the Font Size UI and by format-change
            // detection. InsertXML drops the drawing run's w:sz along with w:position, so
            // both values must be restored on the final normalized InlineShape.
            if (metadata.Mode == LaTeXBlockLayoutMode.Auto)
                shape.Range.Font.Size = (float)metadata.FontSizePt;
            ApplyBaselinePosition(shape, metadata, hostPosition);
        }

        private static void ApplyBaselinePosition(WordInterop.InlineShape shape, LaTeXBlockMetadata metadata,
            int hostPosition)
        {
            // Word aligns the bottom of an InlineShape to the text baseline. Move the image
            // character down by the TeX box depth. This is always the TeX/Western baseline:
            // CJK glyph extents inside the SVG do not define a second alignment reference.
            // Word persists this API value as whole points. Keep any deliberate baseline
            // position of the surrounding run, then apply the TeX depth relative to it.
            shape.Range.Font.Position = hostPosition -
                (int)Math.Round(metadata.DepthPt, MidpointRounding.AwayFromZero);
        }

        internal static int ResolveHostBaselinePosition(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            var shapePosition = ResolveRangePosition(shape.Range, 0);
            return shapePosition + (int)Math.Round(metadata.DepthPt, MidpointRounding.AwayFromZero);
        }

        internal static int ResolveRangePosition(WordInterop.Range range, int fallback)
        {
            if (range == null) return fallback;
            var value = (double)range.Font.Position;
            return value >= -1584 && value <= 1584 ? (int)Math.Round(value) : fallback;
        }

        private static void MoveCaretAfterRange(WordInterop.Range range)
        {
            var caret = range.Duplicate;
            caret.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
            caret.Select();
        }

        private static void MoveCaretAfterNumberedEquation(WordInterop.Field field)
        {
            // A Word field occupies code/result delimiter characters that are not in
            // paragraph.Text. Result.End + 2 is immediately after the literal closing
            // parenthesis. Cross one existing manual break as well so consecutive
            // numbered equations never leave the caret on the preceding display line.
            var document = field.Result.Document;
            var position = field.Result.End + 2;
            if (position < document.Content.End && document.Range(position, position + 1).Text == "\v")
                position++;
            var caret = document.Range(position, position);
            MoveCaretAfterRange(caret);
        }

        private static InlineSpaceAdvances MeasureNaturalAdjacentSpaces(WordInterop.Range target,
            LaTeXBlockMetadata metadata)
        {
            if (metadata.Mode != LaTeXBlockLayoutMode.Auto || metadata.Role != LaTeXBlockRole.Content)
                return default(InlineSpaceAdvances);

            var measured = new InlineSpaceAdvances(
                MeasureNaturalSpaceBeforeInsertion(AdjacentCharacter(target, true)),
                MeasureNaturalSpaceBeforeInsertion(AdjacentCharacter(target, false)));
            return measured;
        }

        private static double MeasureNaturalSpaceBeforeInsertion(WordInterop.Range space)
        {
            if (!IsUnmodifiedOrdinarySpace(space)) return 0;
            // A lone U+0020 has the natural advance we want to preserve. Two touching
            // spaces are different: Word expands their Range.Information geometry before
            // the formula even exists (for example, 5.40 pt instead of a 2.70 pt Times
            // New Roman space at 11 pt). Use an equivalent isolated space or the same
            // formatting in a scratch document for that collapsed-insertion case.
            if (!HasInlineShapeNeighbor(space) && !HasOrdinarySpaceNeighbor(space) &&
                TryMeasureHorizontalAdvance(space, out var directWidth))
                return directWidth;
            var referenceWidth = FindMatchingNaturalSpaceAdvance(space);
            return referenceWidth > 0 ? referenceWidth : MeasureNaturalSpaceInScratchDocument(space);
        }

        private static InlineEffectExtents MeasureInlineSpaceEffectExtents(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata, InlineSpaceAdvances naturalSpaces)
        {
            if (metadata.Mode != LaTeXBlockLayoutMode.Auto || metadata.Role != LaTeXBlockRole.Content)
                return default(InlineEffectExtents);

            var left = AdjacentCharacter(shape.Range, true);
            var right = AdjacentCharacter(shape.Range, false);
            var measured = new InlineEffectExtents(
                MeasureInlineSpaceExcess(left, naturalSpaces.LeftPt),
                MeasureInlineSpaceExcess(right, naturalSpaces.RightPt));
            return measured;
        }

        private static WordInterop.InlineShape ReconcileInsertedInlineSpaceEffectExtents(
            WordInterop.InlineShape shape, LaTeXBlockMetadata metadata, InlineEffectExtents written,
            SvgPhysicalSize svgSize, int hostPosition)
        {
            if (metadata.Mode != LaTeXBlockLayoutMode.Auto || metadata.Role != LaTeXBlockRole.Content)
                return shape;

            // The first pass preserves the exact pre-insertion space whenever Word exposes
            // its geometry. Validate that result against the final DrawingML object as a
            // postcondition. In a real interactive insertion Word can occasionally make the
            // pre-insertion geometry unavailable even though the completed inline shape is
            // measurable; in that case the first pass legitimately writes zero. The update
            // path below can recover the natural space from an equivalent same-line space,
            // the persisted extent, or the formatted scratch sample. Rewrite only when that
            // final answer differs, so the ordinary insertion path still performs one XML
            // normalization and never waits, polls, or schedules document work.
            var finalNaturalSpaces = ResolveUpdatedNaturalAdjacentSpaces(shape, metadata);
            var finalEffects = MeasureInlineSpaceEffectExtents(shape, metadata, finalNaturalSpaces);
            if (finalEffects.LeftEmu == written.LeftEmu && finalEffects.RightEmu == written.RightEmu)
                return shape;

            shape = NormalizeWordInlineDrawing(shape, finalEffects, svgSize);
            ApplyHostRunFormat(shape, metadata, hostPosition);
            return shape;
        }

        private static InlineSpaceAdvances ResolveUpdatedNaturalAdjacentSpaces(
            WordInterop.InlineShape shape, LaTeXBlockMetadata metadata)
        {
            if (metadata.Mode != LaTeXBlockLayoutMode.Auto || metadata.Role != LaTeXBlockRole.Content)
                return default(InlineSpaceAdvances);

            var existing = ReadInlineEffectExtents(shape.Range.WordOpenXML);
            return new InlineSpaceAdvances(
                ResolveUpdatedNaturalSpaceAdvance(AdjacentCharacter(shape.Range, true), existing.LeftEmu),
                ResolveUpdatedNaturalSpaceAdvance(AdjacentCharacter(shape.Range, false), existing.RightEmu));
        }

        private static double ResolveUpdatedNaturalSpaceAdvance(WordInterop.Range space, long existingEmu)
        {
            if (!IsUnmodifiedOrdinarySpace(space)) return 0;
            if (!TryMeasureHorizontalAdvance(space, out var inlineWidthPt))
            {
                var fallback = FindMatchingNaturalSpaceAdvance(space);
                return fallback > 0 ? fallback : MeasureNaturalSpaceInScratchDocument(space);
            }

            var persistedNaturalWidthPt = existingEmu < 0
                ? inlineWidthPt + existingEmu / EmusPerPoint
                : 0;
            var referenceWidthPt = FindMatchingNaturalSpaceAdvance(space);
            // A real same-line space with identical formatting is authoritative. In
            // particular, it repairs documents whose older extent persisted an incorrect
            // proxy width even when the font size itself has not changed.
            if (referenceWidthPt > 0) return referenceWidthPt;

            // Persisted extents describe the old formatted space. They are a useful
            // guard against Word's sub-point layout quantization for an unchanged
            // edit, but they cannot reveal that the surrounding run changed font or
            // size. With no isolated reference on this line, remeasure the current
            // formatting in the scratch document and retain the persisted value only
            // when both measurements still agree within one layout quantum.
            referenceWidthPt = MeasureNaturalSpaceInScratchDocument(space);
            var naturalWidthPt = persistedNaturalWidthPt;
            // Range.Information is quantized by Word's layout surface. Preserve the
            // side-specific persisted value across an ordinary edit when a nearby
            // reference differs only by one layout quantum; use the reference when a
            // real font/size change makes the difference material.
            if (referenceWidthPt > 0 && (naturalWidthPt <= 0 ||
                Math.Abs(referenceWidthPt - naturalWidthPt) > 0.35))
                naturalWidthPt = referenceWidthPt;
            return naturalWidthPt;
        }

        private static WordInterop.Range AdjacentCharacter(WordInterop.Range shapeRange, bool before)
        {
            var adjacent = shapeRange.Duplicate;
            adjacent.Collapse(before
                ? WordInterop.WdCollapseDirection.wdCollapseStart
                : WordInterop.WdCollapseDirection.wdCollapseEnd);
            var moved = before
                ? adjacent.MoveStart(WordInterop.WdUnits.wdCharacter, -1)
                : adjacent.MoveEnd(WordInterop.WdUnits.wdCharacter, 1);
            return Math.Abs(moved) == 1 ? adjacent : null;
        }

        private static bool IsUnmodifiedOrdinarySpace(WordInterop.Range space)
        {
            if (space == null || space.Text != " ") return false;
            var scaling = (double)space.Font.Scaling;
            var spacing = (double)space.Font.Spacing;
            return Math.Abs(scaling - 100) < 0.001 && Math.Abs(spacing) < 0.001;
        }

        private static long MeasureInlineSpaceExcess(WordInterop.Range space, double naturalWidthPt)
        {
            if (!(naturalWidthPt > 0) || !IsUnmodifiedOrdinarySpace(space) ||
                !TryMeasureHorizontalAdvance(space, out var inlineWidthPt)) return 0;

            var excessPt = inlineWidthPt - naturalWidthPt;
            if (excessPt <= 0.05) return 0;
            var fontSizePt = (double)space.Font.Size;
            if (fontSizePt >= 1 && fontSizePt <= 200 && excessPt > fontSizePt) return 0;

            // Word usually expands a U+0020 next to an InlineShape toward a half-em,
            // but the ordinary width is font-dependent. Absorb only the measured delta
            // from this exact formatted space; never assume a fixed ratio or point value.
            return -(long)Math.Round(excessPt * EmusPerPoint, MidpointRounding.AwayFromZero);
        }

        private static double FindMatchingNaturalSpaceAdvance(WordInterop.Range space)
        {
            if (!IsUnmodifiedOrdinarySpace(space)) return 0;
            try
            {
                var paragraph = space.Paragraphs[1].Range.Duplicate;
                var first = Math.Max(paragraph.Start, space.Start - 512);
                var last = Math.Min(paragraph.End, space.End + 512);
                var page = Convert.ToInt32(space.Information[
                    WordInterop.WdInformation.wdActiveEndPageNumber]);
                var line = Convert.ToInt32(space.Information[
                    WordInterop.WdInformation.wdFirstCharacterLineNumber]);
                var candidate = paragraph.Duplicate;
                var bestDistance = int.MaxValue;
                var bestWidthPt = 0.0;
                for (var position = first; position < last; position++)
                {
                    candidate.SetRange(position, position + 1);
                    if (candidate.Start == space.Start || candidate.Text != " " ||
                        HasInlineShapeNeighbor(candidate) || HasOrdinarySpaceNeighbor(candidate) ||
                        !HasEquivalentSpaceMetrics(space, candidate)) continue;
                    if (Convert.ToInt32(candidate.Information[
                            WordInterop.WdInformation.wdActiveEndPageNumber]) != page ||
                        Convert.ToInt32(candidate.Information[
                            WordInterop.WdInformation.wdFirstCharacterLineNumber]) != line) continue;
                    var distance = Math.Abs(position - space.Start);
                    if (distance < bestDistance && TryMeasureHorizontalAdvance(candidate, out var widthPt) &&
                        widthPt > 0)
                    {
                        bestDistance = distance;
                        bestWidthPt = widthPt;
                    }
                }
                return bestWidthPt;
            }
            catch (COMException) { }
            return 0;
        }

        private static double MeasureNaturalSpaceInScratchDocument(WordInterop.Range formattedSpace)
        {
            WordInterop.Document scratch = null;
            WordInterop.Document host = null;
            try
            {
                host = formattedSpace.Document;
                scratch = formattedSpace.Application.Documents.Add(Visible: false);
                var sample = scratch.Range(0, 0);
                sample.Text = "a b";
                sample.Font = formattedSpace.Font.Duplicate;
                sample.Font.Position = 0;
                sample.ParagraphFormat.Alignment = WordInterop.WdParagraphAlignment.wdAlignParagraphLeft;
                var sampleSpace = scratch.Range(1, 2);
                return TryMeasureHorizontalAdvance(sampleSpace, out var widthPt) ? widthPt : 0;
            }
            catch (COMException)
            {
                return 0;
            }
            finally
            {
                if (scratch != null)
                    try { scratch.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                if (host != null)
                    try { host.Activate(); } catch { }
            }
        }

        private static bool HasInlineShapeNeighbor(WordInterop.Range range)
        {
            var left = AdjacentCharacter(range, true);
            var right = AdjacentCharacter(range, false);
            return (left != null && left.InlineShapes.Count > 0) ||
                   (right != null && right.InlineShapes.Count > 0);
        }

        private static bool HasOrdinarySpaceNeighbor(WordInterop.Range range)
        {
            var left = AdjacentCharacter(range, true);
            var right = AdjacentCharacter(range, false);
            return (left != null && left.Text == " ") || (right != null && right.Text == " ");
        }

        private static bool HasEquivalentSpaceMetrics(WordInterop.Range first, WordInterop.Range second)
        {
            try
            {
                return Math.Abs((double)first.Font.Size - (double)second.Font.Size) < 0.001 &&
                       (int)first.Font.Bold == (int)second.Font.Bold &&
                       (int)first.Font.Italic == (int)second.Font.Italic &&
                       (int)first.Font.Scaling == (int)second.Font.Scaling &&
                       Math.Abs((double)first.Font.Spacing - (double)second.Font.Spacing) < 0.001 &&
                       first.LanguageID == second.LanguageID &&
                       first.LanguageIDFarEast == second.LanguageIDFarEast &&
                       string.Equals(first.Font.Name, second.Font.Name, StringComparison.Ordinal) &&
                       string.Equals(first.Font.NameAscii, second.Font.NameAscii, StringComparison.Ordinal) &&
                       string.Equals(first.Font.NameFarEast, second.Font.NameFarEast, StringComparison.Ordinal);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static InlineEffectExtents ReadInlineEffectExtents(string flatOpc)
        {
            var effect = Regex.Match(flatOpc ?? string.Empty,
                "<wp:effectExtent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!effect.Success) return default(InlineEffectExtents);
            return new InlineEffectExtents(
                ReadLongXmlAttribute(effect.Value, "l"),
                ReadLongXmlAttribute(effect.Value, "r"));
        }

        private static long ReadLongXmlAttribute(string element, string attribute)
        {
            var match = Regex.Match(element, "\\b" + Regex.Escape(attribute) + "=\"(?<value>-?[0-9]+)\"",
                RegexOptions.CultureInvariant);
            return match.Success && long.TryParse(match.Groups["value"].Value,
                NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;
        }

        private static bool TryMeasureHorizontalAdvance(WordInterop.Range range, out double widthPt)
        {
            widthPt = 0;
            try
            {
                var start = range.Duplicate;
                start.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                var end = range.Duplicate;
                end.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
                var startPage = Convert.ToInt32(start.Information[
                    WordInterop.WdInformation.wdActiveEndPageNumber]);
                var endPage = Convert.ToInt32(end.Information[
                    WordInterop.WdInformation.wdActiveEndPageNumber]);
                var startY = Convert.ToDouble(start.Information[
                    WordInterop.WdInformation.wdVerticalPositionRelativeToPage]);
                var endY = Convert.ToDouble(end.Information[
                    WordInterop.WdInformation.wdVerticalPositionRelativeToPage]);
                var startX = Convert.ToDouble(start.Information[
                    WordInterop.WdInformation.wdHorizontalPositionRelativeToPage]);
                var endX = Convert.ToDouble(end.Information[
                    WordInterop.WdInformation.wdHorizontalPositionRelativeToPage]);
                if (startPage != endPage || startX < 0 || endX < 0 || startY < 0 || endY < 0 ||
                    Math.Abs(startY - endY) > 0.5 || endX <= startX) return false;
                widthPt = endX - startX;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static WordInterop.InlineShape NormalizeWordInlineDrawing(WordInterop.InlineShape shape,
            InlineEffectExtents sides, SvgPhysicalSize size)
        {
            // Word imports the SVG through a CSS-pixel-sized drawing canvas. Its COM Width
            // and Height therefore lose the sub-pixel part of dvisvgm's physical point
            // dimensions. Restore the exact SVG size in both DrawingML coordinate systems:
            // wp:extent controls inline layout, while pic:spPr/a:xfrm/a:ext controls the
            // picture transform. This is vector geometry expressed in EMUs; no DPI is
            // involved after this correction.
            //
            // Word also expands an ordinary U+0020 immediately next to an InlineShape.
            // Signed left/right effect extents absorb only the measured excess. Host-only
            // effect extent on the bottom is cleared because it is not part of the SVG or
            // the TeX box. Reinsert the otherwise unchanged Flat OPC package, preserving
            // the SVG relationship, PNG fallback, metadata, and TeX depth.
            var flatOpc = shape.Range.WordOpenXML;
            var effect = Regex.Match(flatOpc,
                "<wp:effectExtent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!effect.Success)
                throw new InvalidDataException("Word inline SVG has no wp:effectExtent element.");

            var normalizedEffect = SetXmlAttribute(effect.Value, "l", sides.LeftEmu);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "r", sides.RightEmu);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "b", 0);
            var patched = flatOpc.Remove(effect.Index, effect.Length)
                .Insert(effect.Index, normalizedEffect);

            var inlineExtent = Regex.Match(patched,
                "<wp:extent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!inlineExtent.Success)
                throw new InvalidDataException("Word inline SVG has no wp:extent element.");
            var normalizedInlineExtent = SetXmlAttribute(inlineExtent.Value, "cx", size.WidthEmu);
            normalizedInlineExtent = SetXmlAttribute(normalizedInlineExtent, "cy", size.HeightEmu);
            patched = patched.Remove(inlineExtent.Index, inlineExtent.Length)
                .Insert(inlineExtent.Index, normalizedInlineExtent);

            // Scope the transform extent to pic:spPr/a:xfrm. The same package also has an
            // a:ext element in the blip extension list whose uri must never be touched.
            var pictureProperties = Regex.Match(patched,
                "<pic:spPr\\b[^>]*>.*?</pic:spPr>",
                RegexOptions.CultureInvariant | RegexOptions.Singleline);
            if (!pictureProperties.Success)
                throw new InvalidDataException("Word inline SVG has no pic:spPr element.");
            var transform = Regex.Match(pictureProperties.Value,
                "<a:xfrm\\b[^>]*>.*?</a:xfrm>",
                RegexOptions.CultureInvariant | RegexOptions.Singleline);
            if (!transform.Success)
                throw new InvalidDataException("Word inline SVG has no picture transform.");
            var transformExtent = Regex.Match(transform.Value,
                "<a:ext\\b(?=[^>]*\\bcx=\"[-+0-9]+\")(?=[^>]*\\bcy=\"[-+0-9]+\")[^>]*/>",
                RegexOptions.CultureInvariant);
            if (!transformExtent.Success)
                throw new InvalidDataException("Word inline SVG has no picture transform extent.");
            var normalizedTransformExtent = SetXmlAttribute(transformExtent.Value, "cx", size.WidthEmu);
            normalizedTransformExtent = SetXmlAttribute(normalizedTransformExtent, "cy", size.HeightEmu);
            var normalizedTransform = transform.Value.Remove(transformExtent.Index, transformExtent.Length)
                .Insert(transformExtent.Index, normalizedTransformExtent);
            var normalizedPictureProperties = pictureProperties.Value.Remove(transform.Index, transform.Length)
                .Insert(transform.Index, normalizedTransform);
            patched = patched.Remove(pictureProperties.Index, pictureProperties.Length)
                .Insert(pictureProperties.Index, normalizedPictureProperties);

            if (patched == flatOpc) return shape;
            var originalStart = shape.Range.Start;
            // InsertXML reconstructs the containing paragraph while importing Flat OPC.
            // Preserve its direct paragraph formatting (notably the equation tab stops)
            // so normalizing one SVG cannot silently rewrite the host paragraph.
            var paragraphFormat = shape.Range.ParagraphFormat.Duplicate;
            // Keep the insertion range in the drawing's own story. Document.Range always
            // addresses the main story and would target unrelated text for a formula in a
            // header, footnote, or text box.
            var insertion = shape.Range.Duplicate;
            insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
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

            var replacementRange = insertion.Duplicate;
            replacementRange.SetRange(originalStart, originalStart + 1);
            if (replacementRange.InlineShapes.Count != 1)
                throw new InvalidDataException("Word did not reinsert the normalized inline SVG.");
            var replacement = replacementRange.InlineShapes[1];

            var separator = replacement.Range.Duplicate;
            separator.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
            if (separator.MoveEnd(WordInterop.WdUnits.wdCharacter, 1) == 1 && separator.Text == "\r")
                separator.Delete();
            replacement.Range.ParagraphFormat = paragraphFormat;
            return replacement;
        }

        private static string SetXmlAttribute(string element, string attribute, long value)
        {
            var pattern = "(\\b" + Regex.Escape(attribute) + "=\")[^\"]*(\")";
            return Regex.Replace(element, pattern,
                match => match.Groups[1].Value + value.ToString(CultureInfo.InvariantCulture) +
                         match.Groups[2].Value,
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        }

        private struct InlineEffectExtents
        {
            internal InlineEffectExtents(long leftEmu, long rightEmu)
            {
                LeftEmu = leftEmu;
                RightEmu = rightEmu;
            }

            internal long LeftEmu { get; }
            internal long RightEmu { get; }
        }

        private struct InlineSpaceAdvances
        {
            internal InlineSpaceAdvances(double leftPt, double rightPt)
            {
                LeftPt = leftPt;
                RightPt = rightPt;
            }

            internal double LeftPt { get; }
            internal double RightPt { get; }
        }

        private struct SvgPhysicalSize
        {
            internal SvgPhysicalSize(double widthPt, double heightPt)
            {
                if (!(widthPt > 0) || !(heightPt > 0) || double.IsNaN(widthPt) ||
                    double.IsNaN(heightPt) || double.IsInfinity(widthPt) || double.IsInfinity(heightPt))
                    throw new ArgumentOutOfRangeException(nameof(widthPt), "SVG dimensions must be finite and positive.");
                WidthEmu = checked((long)Math.Round(widthPt * EmusPerPoint,
                    MidpointRounding.AwayFromZero));
                HeightEmu = checked((long)Math.Round(heightPt * EmusPerPoint,
                    MidpointRounding.AwayFromZero));
            }

            internal long WidthEmu { get; }
            internal long HeightEmu { get; }
        }

        internal static double ResolveFontSize(WordInterop.Range target, LaTeXBlockLayoutMode mode, double fallback)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return fallback;
            var fontSize = (double)target.Font.Size;
            if (fontSize >= 1 && fontSize <= 200) return fontSize;

            // Word reports wdUndefined (normally 9999999) for a mixed-size selection.
            // Insertion replaces that selection at its start, so use the insertion format
            // of the first replaced character instead of silently rendering at 10 pt.
            if (target.Start != target.End)
            {
                var insertion = target.Duplicate;
                insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                fontSize = (double)insertion.Font.Size;
                if (fontSize >= 1 && fontSize <= 200) return fontSize;
            }
            return fallback;
        }

        internal static double ResolveFontSize(WordInterop.Selection selection, LaTeXBlockLayoutMode mode,
            double fallback)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return fallback;

            // At a collapsed run boundary, Selection.Range.Font.Size describes the
            // character to the right, while Selection.Font.Size is Word's actual typing
            // format (and therefore the size TypeText would use). Prefer that value.
            var fontSize = (double)selection.Font.Size;
            if (fontSize >= 1 && fontSize <= 200) return fontSize;
            return ResolveFontSize(selection.Range, mode, fallback);
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
