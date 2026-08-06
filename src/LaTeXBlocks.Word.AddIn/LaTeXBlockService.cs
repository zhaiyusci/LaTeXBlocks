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
        // Word uses the SEQ identifier as the native category of a caption-like
        // item.  Keep it deliberately separate from the SVG metadata role: the
        // former is public Word document semantics and powers Cross-reference;
        // the latter tells this add-in how to lay out the object.
        internal const string EquationCategoryIdentifier = "LaTeXBlockEq";
        internal const string EquationSequenceIdentifier = EquationCategoryIdentifier;
        private const string LegacyEquationSequenceIdentifier = "LaTeXEquation";
        internal const string EquationBookmarkPrefix = "LTXEQ_";
        private const double EquationNumberReservePt = 36.0;
        private const double EquationNumberGapPt = 6.0;
        private const double EmusPerPoint = 12700.0;
        private const string WordJoiner = "\u2060";
        // Current Word exposes a floating SVG as an Office Graphic (28), but the
        // Office 15 PIA bundled with this project predates that enum member.
        private const int WordSvgFloatingShapeType = 28;
        // Word exposes direct font colours as WdColor's BGR integer. Automatic
        // colour is the only non-RGB value we deliberately retain; all other
        // undefined/theme sentinel values fall back to it.
        internal const int AutomaticTextColor = unchecked((int)0xff000000);
        private const int UndefinedTextColor = 9999999;
        private readonly WordInterop.Application application;
        private readonly IStemTeXBackend renderers;
        private readonly string cacheDirectory;
        private bool equationCategoryRegistered;

        internal LaTeXBlockService(WordInterop.Application application, IStemTeXBackend renderers)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.renderers = renderers ?? throw new ArgumentNullException(nameof(renderers));
            cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LaTeXBlocks", "cache");
            Directory.CreateDirectory(cacheDirectory);
        }

        internal string[] Profiles => renderers.Profiles;

        /// <summary>
        /// Makes the public Word category behind numbered LaTeX Blocks visible to
        /// Word's native Caption and Cross-reference UI. Caption labels live in
        /// Word's application settings, while the actual document target remains
        /// the per-equation bookmark created around the SEQ result.
        /// </summary>
        internal void EnsureEquationCategory()
        {
            if (equationCategoryRegistered) return;
            if (HasCaptionLabel(EquationCategoryIdentifier))
            {
                equationCategoryRegistered = true;
                return;
            }

            WordInterop.CaptionLabel added = null;
            try
            {
                added = application.CaptionLabels.Add(EquationCategoryIdentifier);
            }
            finally
            {
                if (added != null) Marshal.FinalReleaseComObject(added);
            }

            if (!HasCaptionLabel(EquationCategoryIdentifier))
                throw new InvalidOperationException(
                    "Word could not register the LaTeXBlockEq equation category.");
            equationCategoryRegistered = true;
        }

        private bool HasCaptionLabel(string name)
        {
            var labels = application.CaptionLabels;
            for (var index = 1; index <= labels.Count; index++)
            {
                object item = index;
                WordInterop.CaptionLabel label = null;
                try
                {
                    label = labels.get_Item(ref item);
                    if (string.Equals(label.Name, name, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                finally
                {
                    if (label != null) Marshal.FinalReleaseComObject(label);
                }
            }
            return false;
        }

        internal double ResolveTextAreaWidth(WordInterop.Range range, double fallbackPt = 360)
        {
            if (range == null) return LaTeXBlockWidthPolicy.NormalizeTextAreaWidth(fallbackPt);
            try
            {
                if (TryResolveTextFrameWidth(range, out var textFrameWidth))
                    return textFrameWidth;

                if (Convert.ToBoolean(range.Information[WordInterop.WdInformation.wdWithInTable]) &&
                    range.Cells.Count > 0)
                {
                    var cell = range.Cells[1];
                    var width = (double)cell.Width;
                    try { width -= cell.LeftPadding + cell.RightPadding; }
                    catch (COMException) { }
                    if (width >= LaTeXBlockWidthPolicy.MinimumWidthPt) return width;
                }

                var page = range.Sections[1].PageSetup;
                var widthPt = (double)page.PageWidth - page.LeftMargin - page.RightMargin;
                var columns = page.TextColumns;
                if (columns.Count > 1)
                {
                    var horizontalPosition = Convert.ToDouble(range.Information[
                        WordInterop.WdInformation.wdHorizontalPositionRelativeToPage]);
                    var columnLeft = (double)page.LeftMargin;
                    var bestWidth = 0.0;
                    for (var index = 1; index <= columns.Count; index++)
                    {
                        var column = columns[index];
                        var columnWidth = (double)column.Width;
                        var spaceAfter = index < columns.Count ? (double)column.SpaceAfter : 0;
                        if (bestWidth <= 0) bestWidth = columnWidth;
                        if (horizontalPosition >= columnLeft - 0.5 &&
                            horizontalPosition <= columnLeft + columnWidth + spaceAfter / 2.0)
                        {
                            bestWidth = columnWidth;
                            break;
                        }
                        columnLeft += columnWidth + spaceAfter;
                    }
                    widthPt = bestWidth;
                }
                if (widthPt >= LaTeXBlockWidthPolicy.MinimumWidthPt) return widthPt;
            }
            catch (COMException) { }
            return LaTeXBlockWidthPolicy.NormalizeTextAreaWidth(fallbackPt);
        }

        private static bool TryResolveTextFrameWidth(WordInterop.Range range,
            out double widthPt)
        {
            widthPt = 0;
            try
            {
                if (range.StoryType != WordInterop.WdStoryType.wdTextFrameStory)
                    return false;
                foreach (WordInterop.Shape shape in range.Document.Shapes)
                {
                    WordInterop.Range textRange = null;
                    try
                    {
                        if (shape.TextFrame.HasText == 0) continue;
                        textRange = shape.TextFrame.TextRange;
                        if (!range.InStory(textRange) || range.Start < textRange.Start ||
                            range.Start > textRange.End) continue;
                        widthPt = shape.Width - shape.TextFrame.MarginLeft -
                                  shape.TextFrame.MarginRight;
                        return widthPt >= LaTeXBlockWidthPolicy.MinimumWidthPt;
                    }
                    catch (COMException) { }
                    finally
                    {
                        if (textRange != null) Marshal.FinalReleaseComObject(textRange);
                    }
                }
            }
            catch (COMException) { }
            return false;
        }

        internal LaTeXBlockRender RenderPreview(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile,
            double fontSizePt = 10, bool displayMathStyle = false,
            int textColor = AutomaticTextColor, LaTeXBlockStyle style = null,
            double? outerHeightPt = null, double? outerWidthPt = null)
        {
            return RenderPreviewAsync(source, widthPt, mode, profile, fontSizePt, displayMathStyle,
                    textColor, style, outerHeightPt, outerWidthPt)
                .GetAwaiter().GetResult();
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double fontSizePt = 10,
            bool displayMathStyle = false, int textColor = AutomaticTextColor,
            LaTeXBlockStyle style = null, double? outerHeightPt = null,
            double? outerWidthPt = null)
        {
            return await RenderAsync(source, widthPt, mode, profile, fontSizePt,
                displayMathStyle, false, textColor, style, outerHeightPt, outerWidthPt);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double fontSizePt = 10,
            bool displayMathStyle = false, int textColor = AutomaticTextColor,
            LaTeXBlockStyle style = null, double? outerHeightPt = null,
            double? outerWidthPt = null)
        {
            return await RenderAsync(source, widthPt, mode, profile, fontSizePt,
                displayMathStyle, true, textColor, style, outerHeightPt, outerWidthPt);
        }

        internal void CancelPreview()
        {
            // Preview renders are intentionally disposable. Committed document work
            // stays in the backend's FIFO queue and is never canceled from the editor.
            renderers.CancelLatestPreview();
        }

        private async Task<LaTeXBlockRender> RenderAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double fontSizePt,
            bool displayMathStyle, bool committed, int textColor, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt)
        {
            var normalizedSource = NormalizeSourceText(source);
            var renderSource = displayMathStyle ? PrepareDisplayMathSource(normalizedSource) : normalizedSource;
            // Fixed Content Blocks use the shared PowerPoint/Word style model. TeX
            // owns leading and foreground paint; the SVG owns the physical shell.
            // Auto formulas and numbered equations deliberately remain on Word's
            // native font-colour path and never acquire a block-style wrapper.
            var styledFixedContent = mode == LaTeXBlockLayoutMode.Fixed && style != null;
            var rendererWidthPt = widthPt;
            if (styledFixedContent)
            {
                textColor = ToWordColor(style.TextColor);
                // A non-null style is an explicit acceptance of the Word Block
                // editor, even when all visible controls happen to show their
                // defaults. Keep that promise literal: 1.20× leading and the
                // foreground colour must still be authored in TeX, while the SVG
                // shell makes the requested outer viewport exact. Older blocks
                // without a style payload continue through the legacy bare path.
                var authoredFrameWidthPt = outerWidthPt ?? widthPt;
                var contentWidthPt = Math.Max(0.1,
                    authoredFrameWidthPt - 2 * style.PaddingPt);
                var contentHeightPt = outerHeightPt.HasValue
                    ? Math.Max(0.1, outerHeightPt.Value - 2 * style.PaddingPt)
                    : (double?)null;
                rendererWidthPt = LaTeXBlockStyle.ToTeXLengthPt(contentWidthPt);
                renderSource = style.WrapSource(renderSource, fontSizePt, true,
                    rendererWidthPt, contentHeightPt.HasValue
                        ? LaTeXBlockStyle.ToTeXLengthPt(contentHeightPt.Value)
                        : (double?)null);
            }
            else
            {
                textColor = NormalizeTextColor(textColor);
                // Styled block requests set PreviewBorder globally at TeX shipout.
                // Restore the legacy border before every unstyled Auto, numbered,
                // or pre-style Fixed render so one Block cannot leak viewport state
                // into the next request in the warm StemTeX worker.
                renderSource = "\\global\\PreviewBorder=1pt\n" + ApplyTextColor(renderSource, textColor,
                    mode == LaTeXBlockLayoutMode.Auto);
            }
            var result = committed
                ? await renderers.RenderQueuedAsync(profile, renderSource, rendererWidthPt,
                    mode == LaTeXBlockLayoutMode.Auto, fontSizePt)
                : await renderers.RenderLatestAsync(profile, renderSource, rendererWidthPt,
                    mode == LaTeXBlockLayoutMode.Auto, fontSizePt);
            var finalSvg = styledFixedContent
                ? LaTeXBlockSvgFrame.Decorate(result.Bytes, style,
                    outerWidthPt ?? widthPt, outerHeightPt)
                : result.Bytes;
            return new LaTeXBlockRender(WriteSvg(finalSvg), finalSvg, result.DepthPt,
                fontSizePt, textColor, result.Bytes, styledFixedContent ? style : null);
        }

        internal WordInterop.InlineShape InsertBlock(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile)
        {
            EnsureDocument();
            var fontSizePt = ResolveFontSize(application.Selection, mode, 10);
            var textColor = ResolveTextColor(application.Selection);
            var render = RenderPreview(source, widthPt, mode, profile, fontSizePt, false,
                textColor);
            return InsertRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertRendered(string source, double widthPt, LaTeXBlockLayoutMode mode,
            LaTeXBlockRender render)
        {
            return InsertRendered(source, widthPt, mode, render, null);
        }

        internal WordInterop.InlineShape InsertRendered(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, LaTeXBlockStyle style)
        {
            return InsertRenderedCore(source, widthPt, mode, render, style, null);
        }

        internal WordInterop.InlineShape InsertRenderedAtHostPosition(string source,
            double widthPt, LaTeXBlockLayoutMode mode, LaTeXBlockRender render,
            int hostPosition)
        {
            return InsertRenderedCore(source, widthPt, mode, render, null, hostPosition);
        }

        private WordInterop.InlineShape InsertRenderedCore(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, LaTeXBlockStyle style,
            int? hostPositionOverride)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));
            // The style editor belongs to Fixed Content Blocks. A caller can reuse
            // the same editor while switching to Auto, but that must never leave
            // latent Block-only metadata on the inline formula.
            if (mode != LaTeXBlockLayoutMode.Fixed) style = null;
            var target = application.Selection.Range.Duplicate;
            var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt, mode, render.FontSizePt,
                LaTeXBlockRole.Content, style);
            var document = target.Document;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Insert LaTeX Block");
                undoStarted = true;
                return InsertRenderedAt(target, source, mode, render, metadata, true,
                    hostPositionOverride,
                    () => documentMutated = true);
            }
            catch (Exception exception)
            {
                var rollbackFailure = TryRollbackCustomRecord(document, ref undoStarted, documentMutated);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete the LaTeX Block insertion and could not remove the partial insertion. " +
                        "Inspect the document before saving.",
                        new AggregateException(exception, rollbackFailure));
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
            var textColor = ResolveTextColor(application.Selection);
            var render = RenderPreview(source, widthPt, mode, profile, fontSizePt, true,
                textColor);
            return InsertNumberedRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertNumberedRendered(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (mode != LaTeXBlockLayoutMode.Auto)
                throw new InvalidOperationException("Numbered equations use natural-width display math.");

            // Register the matching CaptionLabel before mutating the document, so
            // this SEQ identifier is immediately exposed as one Word-native
            // Cross-reference category rather than a private add-in convention.
            EnsureEquationCategory();
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
                var shape = InsertRenderedAt(formulaTarget, source, mode, render, metadata, false);
                var fieldPosition = shape.Range.End + 2; // after the second tab and literal '('
                var field = document.Fields.Add(document.Range(fieldPosition, fieldPosition),
                    WordInterop.WdFieldType.wdFieldSequence,
                    EquationSequenceIdentifier + " \\* ARABIC", false);
                if (!field.Update())
                    throw new InvalidOperationException("Word could not create the equation number field.");
                document.Bookmarks.Add(EquationBookmarkName(metadata.Id), field.Result);
                // The insertion already owns one custom Undo record. Keep the
                // renumbering work in that same transaction rather than nesting a
                // second Word Undo record.
                UpdateEquationNumbers(document, false);
                ValidateNumberedEquationPlacement(shape, render.SvgBytes, render.FontSizePt);
                MoveCaretAfterNumberedEquation(field);
                return shape;
            }
            catch (Exception exception)
            {
                var rollbackFailure = TryRollbackCustomRecord(document, ref undoStarted, documentMutated);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete the numbered-equation insertion and could not remove the partial equation. " +
                        "Inspect the document before saving.",
                        new AggregateException(exception, rollbackFailure));
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
            bool select,
            int? hostPositionOverride = null,
            Action markDocumentMutated = null)
        {
            var target = requestedTarget.Duplicate;
            var hostPosition = hostPositionOverride ?? ResolveRangePosition(target, 0);
            var replacesText = target.Start != target.End;
            target.Text = string.Empty;
            if (replacesText) markDocumentMutated?.Invoke();
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var insertionPath = PrepareInsertionSvg(render, mode);
            var svgSize = ReadSvgPhysicalSize(render.SvgBytes);
            // Title is the only durable per-shape metadata Word exposes on an SVG.
            // Keep the root SVG's physical frame separate from the TeX layout width:
            // once a fixed block becomes floating, its frame is what native resize
            // gestures alter, while WidthPt remains the TeX measure.
            metadata = metadata.WithFrameSize(svgSize.WidthPt, svgSize.HeightPt);
            var shape = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
            markDocumentMutated?.Invoke();
            ApplyContract(shape, source, metadata, render.TextColor, hostPosition);
            shape = NormalizeWordInlineDrawing(shape, svgSize);
            EnsureInlineWordJoinerBoundaries(shape, metadata);
            ApplyHostRunFormat(shape, metadata, render.TextColor, hostPosition);
            if (select) MoveCaretAfterInlineFormula(shape, metadata);
            return shape;
        }

        internal WordInterop.InlineShape UpdateBlock(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile, double? fontSizePt = null, bool selectReplacement = true)
        {
            var size = fontSizePt ?? ResolveFontSize(oldShape.Range, mode, 10);
            var displayMathStyle = TryReadContract(oldShape, out var metadata, out _) &&
                                   metadata.Role == LaTeXBlockRole.NumberedEquation;
            var style = metadata != null && metadata.HasExplicitStyle ? metadata.Style : null;
            var textColor = style != null ? ToWordColor(style.TextColor) : ResolveTextColor(oldShape.Range);
            var render = RenderPreview(source, widthPt, mode, profile, size, displayMathStyle,
                textColor, style);
            // UpdateBlock is also used by callers outside the editor. Preserve a
            // Fixed Content Block's exact Word-owned outer viewport there too;
            // otherwise replacing the SVG would silently restore its natural
            // content size and discard a user resize.
            if (metadata != null && metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                metadata.Role == LaTeXBlockRole.Content && mode == LaTeXBlockLayoutMode.Fixed)
                render = FrameFloatingRender(render, oldShape.Width, oldShape.Height, style);
            return UpdateRendered(oldShape, source, widthPt, mode, render, selectReplacement,
                style);
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
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, bool selectReplacement = true,
            LaTeXBlockStyle style = null)
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
            var svgSize = ReadSvgPhysicalSize(render.SvgBytes);
            var styleData = mode == LaTeXBlockLayoutMode.Fixed &&
                previous.Role == LaTeXBlockRole.Content
                ? (style != null ? style.ToMetadataValue() : previous.StyleData)
                : null;
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt, mode,
                render.FontSizePt, previous.Role, svgSize.WidthPt, svgSize.HeightPt, styleData);
            var previousUsesInlineWordJoinerBoundaries = UsesInlineWordJoinerBoundaries(previous);
            // Baseline placement is derived layout, not state owned by the old
            // drawing run. In particular, a damaged/missing w:position must be
            // repaired by Update instead of being interpreted as an intentional
            // raised host baseline. Resolve the surrounding text baseline before
            // replacing the shape, then apply the new TeX depth below it.
            var hostPosition = ResolveSurroundingTextPosition(oldShape);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            WordInterop.InlineShape replacement = null;
            var document = oldShape.Range.Document;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Update LaTeX Block");
                undoStarted = true;

                // Establish or remove the old formula's joiner boundaries while the
                // old shape is still present. Once Word has deleted it, every remaining
                // document mutation must already be complete so a late caret/COM error
                // can never turn a failed update into a missing formula.
                if (previousUsesInlineWordJoinerBoundaries)
                {
                    documentMutated = true;
                    if (UsesInlineWordJoinerBoundaries(metadata))
                        EnsureInlineWordJoinerBoundaries(oldShape, previous);
                    else
                        RemoveInlineWordJoinerBoundaries(oldShape);
                }
                // Inserting/removing a boundary immediately before the old drawing can
                // shift a live Word Range. Reacquire the insertion point from the old
                // shape after that preparation instead of trusting the stale collapse.
                target = oldShape.Range.Duplicate;
                target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                if (numbered)
                {
                    documentMutated = true;
                    ConfigureNumberedEquationTabs(oldShape.Range.Paragraphs[1], numberedLayout);
                }
                var insertionPath = PrepareInsertionSvg(render, mode);
                replacement = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
                documentMutated = true;
                ApplyContract(replacement, source, metadata, render.TextColor, hostPosition);
                replacement = NormalizeWordInlineDrawing(replacement, svgSize);

                // A fixed block has no owned boundary characters. If it becomes an
                // auto-width formula, install the new boundaries before deleting the
                // old shape. A joiner already after the old shape will naturally become
                // the replacement's trailing boundary after that deletion.
                if (!previousUsesInlineWordJoinerBoundaries && UsesInlineWordJoinerBoundaries(metadata))
                {
                    documentMutated = true;
                    EnsureInlineWordJoiner(replacement, true);
                    if (!IsWordJoiner(AdjacentCharacter(oldShape.Range, false)))
                        EnsureInlineWordJoiner(replacement, false);
                }
                ApplyHostRunFormat(replacement, metadata, render.TextColor, hostPosition);
                oldShape.Delete();
                if (selectReplacement)
                {
                    try
                    {
                        if (numbered) MoveCaretAfterNumberedEquation(numberedField);
                        else MoveCaretAfterInlineFormula(replacement, metadata);
                    }
                    catch
                    {
                        // Selection movement is convenience only. The replacement is
                        // already complete, so a host selection failure must not invoke
                        // document rollback after the old formula has been removed.
                    }
                }
                return replacement;
            }
            catch (Exception exception)
            {
                var rollbackFailure = TryRollbackCustomRecord(document, ref undoStarted, documentMutated);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete the LaTeX Block update and could not restore its previous state. " +
                        "The replacement was left in the document for recovery; inspect the formula before saving.",
                        new AggregateException(exception, rollbackFailure));
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        /// <summary>
        /// Updates a user-floated content block without changing its Word layout
        /// contract. Word has no in-place SVG source replacement for a Shape, so
        /// the object is temporarily converted to its lossless InlineShape form,
        /// updated through the normal SVG path, then converted back and given its
        /// original wrapping and position. This is deliberately opt-in through a
        /// selected, contract-verified Shape; ordinary pictures are never touched.
        /// </summary>
        internal WordInterop.Shape UpdateFloatingRendered(WordInterop.Shape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, bool selectReplacement = true,
            LaTeXBlockStyle style = null)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!TryReadContract(oldShape, out var previous, out _))
                throw new InvalidOperationException("The selected image is not a LaTeX Block.");
            if (previous.Role != LaTeXBlockRole.Content || previous.Mode != LaTeXBlockLayoutMode.Fixed)
                throw new InvalidOperationException(
                    "Only fixed-width LaTeX Blocks can remain floating. Inline formulas and numbered equations must remain inline.");
            if (mode != LaTeXBlockLayoutMode.Fixed)
                throw new InvalidOperationException(
                    "A floating LaTeX Block must remain fixed-width. Convert it back to In Line with Text before changing it to an auto-width formula.");

            var layout = FloatingShapeLayout.Capture(oldShape);
            WordInterop.InlineShape inline = null;
            WordInterop.Shape replacement = null;
            try
            {
                inline = oldShape.ConvertToInlineShape();
                var updated = UpdateRendered(inline, source, widthPt, mode, render, false,
                    style);
                replacement = updated.ConvertToShape();
                layout.Apply(replacement);
                // A LaTeX Block's outer frame has text-box semantics: its width and
                // height are independent user instructions, not an image aspect-ratio
                // constraint.  The SVG inside is reframed before this conversion, so
                // leaving Word free to resize either axis does not become persisted
                // visual scaling after the subsequent reflow.
                replacement.LockAspectRatio = Office.MsoTriState.msoFalse;
                if (selectReplacement)
                    try { replacement.Select(); } catch { }
                return replacement;
            }
            catch
            {
                // Before a successful replacement, Word still has the original
                // drawing as an InlineShape. Best-effort conversion restores the
                // user's chosen floating layout; the original source remains in
                // Alternative Text even if the host rejects the conversion.
                if (replacement == null && inline != null)
                {
                    try
                    {
                        var restored = inline.ConvertToShape();
                        layout.Apply(restored);
                    }
                    catch { }
                }
                throw;
            }
        }

        /// <summary>
        /// Rebuilds a fixed Block at its exact Word frame. Styled blocks start from
        /// the original TeX layout box, so its TeX-owned alignment and the SVG
        /// padding/fill/border shell move together instead of becoming two frames.
        /// </summary>
        internal LaTeXBlockRender FrameFloatingRender(LaTeXBlockRender render,
            double frameWidthPt, double frameHeightPt, LaTeXBlockStyle style = null)
        {
            if (render == null) throw new ArgumentNullException(nameof(render));
            style = style ?? render.Style;
            var framedBytes = style != null
                ? LaTeXBlockSvgFrame.Decorate(render.ContentSvgBytes, style,
                    frameWidthPt, frameHeightPt)
                : FrameSvg(render.SvgBytes, frameWidthPt, frameHeightPt);
            return new LaTeXBlockRender(WriteSvg(framedBytes), framedBytes, render.DepthPt,
                render.FontSizePt, render.TextColor, render.ContentSvgBytes, style);
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

        internal bool TryGetSelectedFloatingBlock(out WordInterop.Shape shape, out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            if (application.Documents.Count == 0 || application.Selection == null) return false;
            try
            {
                var selection = application.Selection;
                if (selection.ShapeRange.Count != 1) return false;
                var candidate = selection.ShapeRange[1];
                if (!TryReadContract(candidate, out metadata, out _)) return false;
                shape = candidate;
                return true;
            }
            catch (COMException)
            {
                metadata = null;
                return false;
            }
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

        internal static bool TryReadContract(WordInterop.Shape shape, out LaTeXBlockMetadata metadata,
            out string source)
        {
            metadata = null;
            source = null;
            if (shape == null) return false;
            try
            {
                // Never treat an arbitrary text box, chart, or OLE object as a
                // block just because it exposes an AlternativeText property. SVGs
                // imported by current Word use the otherwise unnamed type 28.
                if (!IsSupportedFloatingShapeType(shape.Type)) return false;
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

        private static bool IsSupportedFloatingShapeType(Office.MsoShapeType type)
        {
            return type == Office.MsoShapeType.msoPicture ||
                   type == Office.MsoShapeType.msoLinkedPicture ||
                   (int)type == WordSvgFloatingShapeType;
        }

        internal int UpdateEquationNumbers(WordInterop.Document document = null)
        {
            return UpdateEquationNumbers(document, true);
        }

        internal IReadOnlyList<EquationReferenceTarget> GetEquationReferenceTargets(
            WordInterop.Document document = null)
        {
            document = document ?? application.ActiveDocument;
            if (document == null) return new EquationReferenceTarget[0];

            var targets = new List<EquationReferenceTarget>();
            var seenIds = new HashSet<Guid>();
            var mainStory = document.StoryRanges[WordInterop.WdStoryType.wdMainTextStory];
            foreach (WordInterop.InlineShape shape in mainStory.InlineShapes)
            {
                if (!TryReadContract(shape, out var metadata, out var source) ||
                    metadata.Role != LaTeXBlockRole.NumberedEquation ||
                    shape.Range.Tables.Count > 0 || seenIds.Contains(metadata.Id))
                    continue;

                try
                {
                    var field = FindEquationNumberField(shape, metadata);
                    var number = (field.Result.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(number)) continue;
                    targets.Add(new EquationReferenceTarget(metadata.Id,
                        EquationBookmarkName(metadata.Id), number, source, shape.Range.Start));
                    seenIds.Add(metadata.Id);
                }
                catch (COMException)
                {
                    // A copied or manually damaged equation can have metadata but no
                    // matching bookmark/field. It must not silently become a target
                    // for a reference to a different equation.
                }
                catch (InvalidOperationException)
                {
                    // See the COM case above. A future repair command will assign a
                    // fresh identity to copied or otherwise incomplete equations.
                }
            }
            targets.Sort((left, right) => left.Position.CompareTo(right.Position));
            return targets;
        }

        internal WordInterop.Field InsertEquationReference(EquationReferenceTarget reference)
        {
            EnsureDocument();
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            var document = application.ActiveDocument;
            ValidateEquationReferenceTarget(document, reference);
            var target = application.Selection.Range.Duplicate;
            if (target.StoryType != WordInterop.WdStoryType.wdMainTextStory)
                throw new InvalidOperationException(
                    "Equation references can currently be inserted only in the main document text.");
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Insert Equation Reference");
                undoStarted = true;

                // Parentheses are ordinary author-visible text. The number itself is
                // a native REF field so it follows its target on renumbering and is a
                // hyperlink through \h, rather than a copied number managed by us.
                var insertionStart = target.Start;
                target.Text = "(";
                documentMutated = true;
                var fieldRange = document.Range(insertionStart + 1, insertionStart + 1);
                var field = document.Fields.Add(fieldRange, WordInterop.WdFieldType.wdFieldRef,
                    reference.BookmarkName + " \\h", true);
                documentMutated = true;
                if (!field.Update())
                    throw new InvalidOperationException("Word could not create the equation reference field.");

                // Insert the closing parenthesis only after Word has materialized
                // the field. Pre-creating both characters lets Fields.Add consume
                // the closing character on some Word builds.
                var closingPosition = field.Result.End + 1;
                var closingParenthesis = document.Range(closingPosition, closingPosition);
                closingParenthesis.Text = ")";
                var afterReference = document.Range(closingPosition + 1, closingPosition + 1);
                afterReference.Select();
                return field;
            }
            catch (Exception exception)
            {
                var rollbackFailure = TryRollbackCustomRecord(document, ref undoStarted, documentMutated);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete the equation-reference insertion and could not remove the partial reference. " +
                        "Inspect the document before saving.",
                        new AggregateException(exception, rollbackFailure));
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        private int UpdateEquationNumbers(WordInterop.Document document, bool createUndoRecord)
        {
            document = document ?? application.ActiveDocument;
            if (document == null) return 0;

            var mainStory = document.StoryRanges[WordInterop.WdStoryType.wdMainTextStory];
            var reflowedParagraphs = new HashSet<int>();
            var paragraphs = new List<EquationParagraphUpdate>();
            foreach (WordInterop.InlineShape shape in mainStory.InlineShapes)
            {
                if (!TryReadContract(shape, out var metadata, out _) ||
                    metadata.Role != LaTeXBlockRole.NumberedEquation || shape.Range.Tables.Count > 0)
                    continue;
                var paragraph = shape.Range.Paragraphs[1];
                if (!reflowedParagraphs.Add(paragraph.Range.Start)) continue;
                paragraphs.Add(new EquationParagraphUpdate(paragraph,
                    GetNumberedEquationLayout(shape.Range, metadata.FontSizePt)));
            }

            var fields = mainStory.Fields;
            var equationFields = new List<WordInterop.Field>();
            var referenceFields = new List<WordInterop.Field>();
            for (var index = 1; index <= fields.Count; index++)
            {
                var field = fields[index];
                if (IsEquationSequenceField(field)) equationFields.Add(field);
                else if (IsEquationReferenceField(field)) referenceFields.Add(field);
            }

            if (paragraphs.Count == 0 && equationFields.Count == 0 && referenceFields.Count == 0) return 0;
            if (equationFields.Count > 0) EnsureEquationCategory();

            var undoStarted = false;
            var documentMutated = false;
            try
            {
                if (createUndoRecord)
                {
                    application.UndoRecord.StartCustomRecord("Update LaTeX Equation Numbers");
                    undoStarted = true;
                }

                foreach (var paragraph in paragraphs)
                {
                    documentMutated = true;
                    ConfigureNumberedEquationTabs(paragraph.Paragraph, paragraph.Layout);
                }
                for (var index = 0; index < equationFields.Count; index++)
                {
                    documentMutated = true;
                    MigrateLegacyEquationSequenceField(equationFields[index]);
                    if (!equationFields[index].Update())
                        throw new InvalidOperationException("Word could not update equation number " + (index + 1) + ".");
                }
                // REF results cache the previous bookmark text. Update them only
                // after all SEQ fields have settled, otherwise a reference can keep
                // the number from before a preceding equation was moved or deleted.
                for (var index = 0; index < referenceFields.Count; index++)
                {
                    documentMutated = true;
                    referenceFields[index].Update();
                }
                return equationFields.Count;
            }
            catch (Exception exception)
            {
                var rollbackFailure = createUndoRecord
                    ? TryRollbackCustomRecord(document, ref undoStarted, documentMutated)
                    : null;
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not update the equation numbers and could not restore the previous numbering.",
                        new AggregateException(exception, rollbackFailure));
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        private Exception TryRollbackCustomRecord(WordInterop.Document document, ref bool undoStarted,
            bool documentMutated)
        {
            if (!undoStarted) return null;

            Exception endFailure = null;
            try { application.UndoRecord.EndCustomRecord(); }
            catch (Exception exception) { endFailure = exception; }
            undoStarted = false;
            // Word does not guarantee that Undo still targets this transaction after
            // EndCustomRecord failed. Calling document.Undo() in that state could undo
            // an unrelated user edit, so leave the document recoverable and surface
            // the failed transaction instead.
            if (endFailure != null) return endFailure;
            if (!documentMutated) return endFailure;

            try
            {
                document.Undo();
                return null;
            }
            catch (Exception undoFailure)
            {
                return endFailure == null
                    ? undoFailure
                    : new AggregateException(endFailure, undoFailure);
            }
        }

        private sealed class EquationParagraphUpdate
        {
            internal EquationParagraphUpdate(WordInterop.Paragraph paragraph,
                NumberedEquationLayout layout)
            {
                Paragraph = paragraph;
                Layout = layout;
            }

            internal WordInterop.Paragraph Paragraph { get; }
            internal NumberedEquationLayout Layout { get; }
        }

        internal sealed class EquationReferenceTarget
        {
            internal EquationReferenceTarget(Guid id, string bookmarkName, string number,
                string source, int position)
            {
                Id = id;
                BookmarkName = bookmarkName;
                Number = number;
                Source = source;
                Position = position;
            }

            internal Guid Id { get; }
            internal string BookmarkName { get; }
            internal string Number { get; }
            internal string Source { get; }
            internal int Position { get; }
        }

        internal static bool IsEquationSequenceField(WordInterop.Field field)
        {
            if (field == null || field.Type != WordInterop.WdFieldType.wdFieldSequence) return false;
            var code = field.Code.Text ?? string.Empty;
            return Regex.IsMatch(code, "^\\s*SEQ\\s+(?:" + EquationSequenceIdentifier + "|" +
                LegacyEquationSequenceIdentifier + ")(?:\\s|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        internal static bool IsEquationReferenceField(WordInterop.Field field)
        {
            if (field == null || field.Type != WordInterop.WdFieldType.wdFieldRef) return false;
            var code = field.Code.Text ?? string.Empty;
            return Regex.IsMatch(code, "^\\s*REF\\s+" + EquationBookmarkPrefix + "[0-9a-f]{32}(?:\\s|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static void MigrateLegacyEquationSequenceField(WordInterop.Field field)
        {
            if (field == null) return;
            var code = field.Code.Text ?? string.Empty;
            if (!Regex.IsMatch(code, "^\\s*SEQ\\s+" + LegacyEquationSequenceIdentifier + "(?:\\s|$)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return;
            var migrated = Regex.Replace(code,
                "^(\\s*SEQ\\s+)" + LegacyEquationSequenceIdentifier + "(?=\\s|$)",
                "$1" + EquationSequenceIdentifier,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            field.Code.Text = migrated;
        }

        private static void ValidateEquationReferenceTarget(WordInterop.Document document,
            EquationReferenceTarget reference)
        {
            if (!string.Equals(reference.BookmarkName, EquationBookmarkName(reference.Id),
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected equation reference is not a LaTeXBlockEq target.");
            if (!document.Bookmarks.Exists(reference.BookmarkName))
                throw new InvalidOperationException("The selected equation no longer has its Word bookmark.");

            var bookmark = document.Bookmarks[reference.BookmarkName].Range;
            foreach (WordInterop.Field field in bookmark.Paragraphs[1].Range.Fields)
            {
                if (IsEquationSequenceField(field) && field.Result.Start == bookmark.Start &&
                    field.Result.End == bookmark.End)
                    return;
            }
            throw new InvalidOperationException(
                "The selected equation bookmark no longer identifies its Word number field.");
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

        // A floating Word Shape is an outer frame, not a request to scale the
        // mathematics.  Re-author the root viewport at the requested physical
        // dimensions while preserving the original TeX coordinate scale.  The
        // frame therefore adds transparent space when enlarged and clips when
        // reduced; it never stretches glyphs or rules.
        internal static byte[] FrameSvg(byte[] svgBytes, double requestedFrameWidthPt,
            double requestedFrameHeightPt)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new ArgumentException("StemTeX returned an empty SVG.", nameof(svgBytes));

            var naturalWidthPt = ReadSvgWidthPt(svgBytes);
            var naturalHeightPt = ReadSvgHeightPt(svgBytes);
            var frameWidthPt = ClampFloatingFrameExtent(requestedFrameWidthPt);
            var frameHeightPt = ClampFloatingFrameExtent(requestedFrameHeightPt);
            if (Math.Abs(frameWidthPt - naturalWidthPt) < 0.001 &&
                Math.Abs(frameHeightPt - naturalHeightPt) < 0.001)
                return svgBytes;

            var svg = Encoding.UTF8.GetString(svgBytes);
            var root = Regex.Match(svg, "<svg\\b[^>]*>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!root.Success)
                throw new InvalidDataException("StemTeX SVG has no root svg element.");

            var rootTag = root.Value;
            var viewBox = Regex.Match(rootTag,
                "\\bviewBox=(?<q>['\"])(?<x>[-+0-9.eE]+)\\s+(?<y>[-+0-9.eE]+)\\s+" +
                "(?<w>[-+0-9.eE]+)\\s+(?<h>[-+0-9.eE]+)\\k<q>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!viewBox.Success ||
                !TryReadFiniteSvgNumber(viewBox.Groups["w"].Value, out var viewBoxWidth) ||
                !TryReadFiniteSvgNumber(viewBox.Groups["h"].Value, out var viewBoxHeight) ||
                !TryReadFiniteSvgNumber(viewBox.Groups["x"].Value, out var viewBoxX) ||
                !TryReadFiniteSvgNumber(viewBox.Groups["y"].Value, out var viewBoxY) ||
                !(viewBoxWidth > 0) || !(viewBoxHeight > 0))
                throw new InvalidDataException("StemTeX SVG has no numeric root viewBox.");

            // The ratio applied to physical dimensions is also applied to the
            // corresponding viewBox axis.  Consequently one SVG coordinate maps
            // to exactly the same physical length before and after framing.
            var frameViewBoxWidth = viewBoxWidth * frameWidthPt / naturalWidthPt;
            // A fixed Word Block has the same left-anchored TeX layout contract as
            // PowerPoint. Enlarging adds room on the right; shrinking clips there.
            var frameViewBoxX = viewBoxX;
            var frameViewBoxHeight = viewBoxHeight * frameHeightPt / naturalHeightPt;
            // Word has no block vertical-alignment UI.  Match the default
            // PowerPoint block policy: the TeX viewport begins at the top edge,
            // so reducing a frame height preserves the first rendered line.
            var frameViewBoxY = viewBoxY;
            var number = CultureInfo.InvariantCulture;
            var newViewBox = frameViewBoxX.ToString("0.######", number) + " " +
                             frameViewBoxY.ToString("0.######", number) + " " +
                             frameViewBoxWidth.ToString("0.######", number) + " " +
                             frameViewBoxHeight.ToString("0.######", number);
            rootTag = ReplaceSvgRootAttribute(rootTag, "width",
                frameWidthPt.ToString("0.######", number) + "pt");
            rootTag = ReplaceSvgRootAttribute(rootTag, "height",
                frameHeightPt.ToString("0.######", number) + "pt");
            rootTag = ReplaceSvgRootAttribute(rootTag, "viewBox", newViewBox);
            rootTag = ReplaceSvgRootAttribute(rootTag, "overflow", "hidden");
            svg = svg.Substring(0, root.Index) + rootTag +
                  svg.Substring(root.Index + root.Length);
            return Encoding.UTF8.GetBytes(svg);
        }

        internal static double ClampFloatingFrameExtent(double extentPt)
        {
            // This is a validity guard, not a layout policy. A Word Block owns its
            // native outer frame, including widths beyond the editor's historical
            // 2000 pt typesetting range. Silently capping that frame changes the
            // user-visible SVG viewport after a drag. Word cannot retain a zero-size
            // picture, so only preserve a tiny positive floor for malformed/zero
            // values; otherwise retain the supplied finite physical extent exactly.
            if (double.IsNaN(extentPt) || double.IsInfinity(extentPt) || !(extentPt > 0))
                return 0.01;
            return extentPt;
        }

        // The native gesture path deliberately observes only dimensions. This is
        // shared with smoke tests so a move/rotation cannot silently regress into a
        // render-triggering geometry change.
        internal static bool HasNativeFrameGeometryChanged(double previousWidthPt,
            double previousHeightPt, double currentWidthPt, double currentHeightPt)
        {
            const double tolerancePt = 0.05;
            return Math.Abs(previousWidthPt - currentWidthPt) > tolerancePt ||
                   Math.Abs(previousHeightPt - currentHeightPt) > tolerancePt;
        }

        internal static double ComposeNativeFrameLayoutWidth(double previousLayoutWidthPt,
            double previousFrameWidthPt, double currentFrameWidthPt)
        {
            var widthPt = previousLayoutWidthPt + currentFrameWidthPt - previousFrameWidthPt;
            if (double.IsNaN(widthPt) || double.IsInfinity(widthPt))
                return LaTeXBlockWidthPolicy.MinimumWidthPt;
            // The TeX layout policy remains independently bounded, while FrameSvg
            // above preserves the outer Word frame exactly even beyond this range.
            return Math.Max(LaTeXBlockWidthPolicy.MinimumWidthPt,
                Math.Min(LaTeXBlockWidthPolicy.MaximumWidthPt, widthPt));
        }

        private static bool TryReadFiniteSvgNumber(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                       out value) && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string ReplaceSvgRootAttribute(string rootTag, string name, string value)
        {
            var attribute = Regex.Match(rootTag,
                "\\b" + Regex.Escape(name) + "=(?<q>['\"])[^'\"]*\\k<q>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var replacement = name + "='" + value + "'";
            if (attribute.Success)
                return rootTag.Substring(0, attribute.Index) + replacement +
                       rootTag.Substring(attribute.Index + attribute.Length);
            var insertion = rootTag.EndsWith("/>", StringComparison.Ordinal)
                ? rootTag.Length - 2
                : rootTag.Length - 1;
            return rootTag.Insert(insertion, " " + replacement);
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

        private static void ApplyContract(WordInterop.InlineShape shape, string source,
            LaTeXBlockMetadata metadata, int textColor, int hostPosition)
        {
            shape.AlternativeText = NormalizeSourceText(source);
            shape.Title = metadata.ToString();
            shape.LockAspectRatio = HasIndependentFrameResize(metadata)
                ? Office.MsoTriState.msoFalse
                : Office.MsoTriState.msoTrue;
            ApplyHostRunFormat(shape, metadata, textColor, hostPosition);
        }

        private static void ApplyHostRunFormat(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata, int textColor, int hostPosition)
        {
            // The image's physical dimensions already come from the SVG. This run size is
            // Word's semantic host size, used by the Font Size UI and by format-change
            // detection. InsertXML drops the drawing run's w:sz along with w:position, so
            // both values must be restored on the final normalized InlineShape.
            if (metadata.Mode == LaTeXBlockLayoutMode.Auto)
                shape.Range.Font.Size = (float)metadata.FontSizePt;
            // The native Word Font Color is the authoritative user-facing colour.
            // Its RGB value is also used when generating the SVG, while Alternative
            // Text remains exactly the author-written TeX source.
            shape.Range.Font.Color = (WordInterop.WdColor)NormalizeTextColor(textColor);
            ApplyBaselinePosition(shape, metadata, hostPosition);
            // InsertXML used by NormalizeWordInlineDrawing rebuilds the DrawingML
            // object and may restore Word's default aspect lock. Fixed Content Blocks
            // intentionally own a two-dimensional frame, so restore that setting
            // after every normalize/update. Do not rewrite an Auto/numbered drawing
            // here: toggling its aspect lock after normalization can make Word add
            // effect extents to an otherwise exact inline SVG.
            if (HasIndependentFrameResize(metadata))
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
        }

        private static bool HasIndependentFrameResize(LaTeXBlockMetadata metadata)
        {
            return metadata != null && metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                   metadata.Role == LaTeXBlockRole.Content;
        }

        internal static int ResolveTextColor(WordInterop.Range target,
            int fallback = AutomaticTextColor)
        {
            fallback = NormalizeTextColor(fallback);
            if (target == null) return fallback;
            if (TryReadTextColor(target.Font, out var color)) return color;

            // Word returns wdUndefined for a mixed selection. The formula replaces
            // at its start, so use the insertion character's colour just as we do
            // for a mixed font-size selection.
            if (target.Start != target.End)
            {
                var insertion = target.Duplicate;
                insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                if (TryReadTextColor(insertion.Font, out color)) return color;
            }
            return fallback;
        }

        internal static int ResolveTextColor(WordInterop.Selection selection,
            int fallback = AutomaticTextColor)
        {
            fallback = NormalizeTextColor(fallback);
            if (selection == null) return fallback;
            // For a collapsed run boundary Selection.Font is Word's actual typing
            // formatting. Selection.Range.Font may instead describe the character
            // on the right of the caret.
            if (TryReadTextColor(selection.Font, out var color)) return color;
            return ResolveTextColor(selection.Range, fallback);
        }

        internal static bool TextColorsEqual(int left, int right)
        {
            return NormalizeTextColor(left) == NormalizeTextColor(right);
        }

        internal static int NormalizeTextColor(int color)
        {
            return color >= 0 && color <= 0x00ffffff
                ? color
                : AutomaticTextColor;
        }

        internal static int ToWordColor(System.Drawing.Color color)
        {
            // WdColor is BGR in the low 24 bits, while System.Drawing.Color is RGB.
            return color.R | (color.G << 8) | (color.B << 16);
        }

        internal static string ApplyTextColor(string source, int textColor,
            bool trimTerminalLineBreaks = false)
        {
            textColor = NormalizeTextColor(textColor);
            if (textColor == AutomaticTextColor) return source;
            // StemTeX's auto-width measurement wrapper trims terminal line breaks
            // before it builds its hbox. Do the same before adding this color wrapper,
            // otherwise a user-facing terminal newline becomes an interior hbox space.
            // Fixed blocks retain their original source verbatim.
            if (trimTerminalLineBreaks && source != null)
                source = source.TrimEnd('\r', '\n');
            var red = textColor & 0xff;
            var green = (textColor >> 8) & 0xff;
            var blue = (textColor >> 16) & 0xff;
            return "\\begingroup\\color[HTML]{" + red.ToString("X2", CultureInfo.InvariantCulture) +
                   green.ToString("X2", CultureInfo.InvariantCulture) +
                   blue.ToString("X2", CultureInfo.InvariantCulture) + "}%\n" +
                   // This percent is part of the wrapper, not the author's source: an
                   // un-commented newline after an inline fragment becomes a TeX
                   // interword glue node inside StemTeXRenderer's measurement hbox.
                   // That would make changing Word Font.Color change the formula's
                   // width, even though colour must affect paint only.
                   source + "%\n\\endgroup";
        }

        private static bool TryReadTextColor(WordInterop.Font font, out int color)
        {
            color = AutomaticTextColor;
            if (font == null) return false;
            var raw = (int)font.Color;
            if (raw == UndefinedTextColor) return false;
            color = NormalizeTextColor(raw);
            return true;
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

        internal static int ResolveSurroundingTextPosition(WordInterop.InlineShape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var anchor = shape.Range;

            // Prefer preceding prose because it is the typing context in which an
            // inline formula was inserted. Skip our U+2060 boundaries and drawing
            // characters; neither owns a text baseline. Falling back to following
            // prose also repairs a formula at the start of a paragraph.
            if (TryResolveAdjacentTextPosition(anchor, true, out var position))
                return position;
            if (TryResolveAdjacentTextPosition(anchor, false, out position))
                return position;
            return 0;
        }

        private static bool TryResolveAdjacentTextPosition(WordInterop.Range anchor,
            bool before, out int position)
        {
            position = 0;
            const int maximumScan = 64;
            for (var distance = 1; distance <= maximumScan; distance++)
            {
                var start = before ? anchor.Start - distance : anchor.End + distance - 1;
                if (start < 0) return false;
                var probe = anchor.Duplicate;
                try
                {
                    probe.SetRange(start, start + 1);
                }
                catch (COMException)
                {
                    return false;
                }

                var text = probe.Text;
                if (string.IsNullOrEmpty(text)) continue;
                var character = text[0];
                if (character == '\u2060' || character == '\u0001') continue;
                if (character == '\r' || character == '\a') return false;
                position = ResolveRangePosition(probe, 0);
                return true;
            }
            return false;
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

        private static bool UsesInlineWordJoinerBoundaries(LaTeXBlockMetadata metadata)
        {
            return metadata != null && metadata.Mode == LaTeXBlockLayoutMode.Auto &&
                   metadata.Role == LaTeXBlockRole.Content;
        }

        private static void EnsureInlineWordJoinerBoundaries(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata)
        {
            if (!UsesInlineWordJoinerBoundaries(metadata)) return;

            // The order matters during an update: when the old image is deleted its
            // trailing joiner can become the new image's trailing joiner. Check each
            // immediate neighbor before inserting anything, and never coalesce joiners
            // that were already present in the user's document.
            EnsureInlineWordJoiner(shape, false);
            EnsureInlineWordJoiner(shape, true);
        }

        private static void EnsureInlineWordJoiner(WordInterop.InlineShape shape, bool before)
        {
            if (IsWordJoiner(AdjacentCharacter(shape.Range, before))) return;
            var insertion = shape.Range.Duplicate;
            insertion.Collapse(before
                ? WordInterop.WdCollapseDirection.wdCollapseStart
                : WordInterop.WdCollapseDirection.wdCollapseEnd);
            insertion.Text = WordJoiner;
        }

        private static void RemoveInlineWordJoinerBoundaries(WordInterop.InlineShape shape)
        {
            // A former auto formula may be edited into a fixed-width block. Remove the
            // boundary characters that belonged only to it, but preserve a joiner that
            // is simultaneously the boundary of an immediately adjacent auto formula.
            RemoveInlineWordJoiner(shape, false);
            RemoveInlineWordJoiner(shape, true);
        }

        private static void RemoveInlineWordJoiner(WordInterop.InlineShape shape, bool before)
        {
            var joiner = AdjacentCharacter(shape.Range, before);
            if (!IsWordJoiner(joiner)) return;

            var adjacentShapePosition = before ? joiner.Start - 1 : joiner.End;
            if (IsAutoContentShapeAt(joiner.Document, adjacentShapePosition)) return;
            joiner.Delete();
        }

        private static bool IsAutoContentShapeAt(WordInterop.Document document, int position)
        {
            if (document == null || position < document.Content.Start || position >= document.Content.End)
                return false;
            var candidate = document.Range(position, position + 1);
            return candidate.InlineShapes.Count == 1 &&
                   TryReadContract(candidate.InlineShapes[1], out var metadata, out _) &&
                   UsesInlineWordJoinerBoundaries(metadata);
        }

        private static bool IsWordJoiner(WordInterop.Range character)
        {
            return character != null && string.Equals(character.Text, WordJoiner,
                StringComparison.Ordinal);
        }

        private static void MoveCaretAfterInlineFormula(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata)
        {
            if (UsesInlineWordJoinerBoundaries(metadata))
            {
                var trailingJoiner = AdjacentCharacter(shape.Range, false);
                if (IsWordJoiner(trailingJoiner))
                {
                    MoveCaretAfterRange(trailingJoiner);
                    return;
                }
            }
            MoveCaretAfterRange(shape.Range);
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

        private static WordInterop.InlineShape NormalizeWordInlineDrawing(WordInterop.InlineShape shape,
            SvgPhysicalSize size)
        {
            // Word imports the SVG through a CSS-pixel-sized drawing canvas. Its COM Width
            // and Height therefore lose the sub-pixel part of dvisvgm's physical point
            // dimensions. Restore the exact SVG size in both DrawingML coordinate systems:
            // wp:extent controls inline layout, while pic:spPr/a:xfrm/a:ext controls the
            // picture transform. This is vector geometry expressed in EMUs; no DPI is
            // involved after this correction.
            //
            // The SVG is a faithful TeX box. All effect extents are host-side drawing
            // margins, so they must be zero; U+2060 boundary characters handle Word's
            // adjacent-space behavior for ordinary inline formulas. Reinsert the
            // otherwise unchanged Flat OPC package, preserving the SVG relationship,
            // PNG fallback, metadata, and TeX depth.
            var flatOpc = shape.Range.WordOpenXML;
            var effect = Regex.Match(flatOpc,
                "<wp:effectExtent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!effect.Success)
                throw new InvalidDataException("Word inline SVG has no wp:effectExtent element.");

            var normalizedEffect = SetXmlAttribute(effect.Value, "l", 0);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "t", 0);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "r", 0);
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

        private sealed class FloatingShapeLayout
        {
            private readonly int wrapType;
            private readonly int wrapSide;
            private readonly float distanceLeft;
            private readonly float distanceRight;
            private readonly float distanceTop;
            private readonly float distanceBottom;
            private readonly int allowOverlap;
            private readonly int relativeHorizontalPosition;
            private readonly int relativeVerticalPosition;
            private readonly float left;
            private readonly float top;
            private readonly float rotation;
            private readonly int layoutInCell;
            private readonly int lockAnchor;

            private FloatingShapeLayout(WordInterop.Shape shape)
            {
                var wrap = shape.WrapFormat;
                wrapType = (int)wrap.Type;
                wrapSide = (int)wrap.Side;
                distanceLeft = wrap.DistanceLeft;
                distanceRight = wrap.DistanceRight;
                distanceTop = wrap.DistanceTop;
                distanceBottom = wrap.DistanceBottom;
                allowOverlap = wrap.AllowOverlap;
                relativeHorizontalPosition = (int)shape.RelativeHorizontalPosition;
                relativeVerticalPosition = (int)shape.RelativeVerticalPosition;
                left = shape.Left;
                top = shape.Top;
                rotation = shape.Rotation;
                layoutInCell = shape.LayoutInCell;
                lockAnchor = shape.LockAnchor;
            }

            internal static FloatingShapeLayout Capture(WordInterop.Shape shape)
            {
                if (shape == null) throw new ArgumentNullException(nameof(shape));
                return new FloatingShapeLayout(shape);
            }

            internal void Apply(WordInterop.Shape shape)
            {
                if (shape == null) throw new ArgumentNullException(nameof(shape));
                var wrap = shape.WrapFormat;
                // Word derives several values from each other during conversion.
                // Restore the reference frame before the wrap, then absolute
                // coordinates, then margins; the reverse order shifts a correctly
                // positioned floating block back to Word's default margin anchor.
                shape.RelativeHorizontalPosition =
                    (WordInterop.WdRelativeHorizontalPosition)relativeHorizontalPosition;
                shape.RelativeVerticalPosition =
                    (WordInterop.WdRelativeVerticalPosition)relativeVerticalPosition;
                wrap.Type = (WordInterop.WdWrapType)wrapType;
                shape.Left = left;
                shape.Top = top;
                wrap.Side = (WordInterop.WdWrapSideType)wrapSide;
                wrap.DistanceLeft = distanceLeft;
                wrap.DistanceRight = distanceRight;
                wrap.DistanceTop = distanceTop;
                wrap.DistanceBottom = distanceBottom;
                wrap.AllowOverlap = allowOverlap;
                shape.Rotation = rotation;
                shape.LayoutInCell = layoutInCell;
                shape.LockAnchor = lockAnchor;
            }
        }

        private struct SvgPhysicalSize
        {
            internal SvgPhysicalSize(double widthPt, double heightPt)
            {
                if (!(widthPt > 0) || !(heightPt > 0) || double.IsNaN(widthPt) ||
                    double.IsNaN(heightPt) || double.IsInfinity(widthPt) || double.IsInfinity(heightPt))
                    throw new ArgumentOutOfRangeException(nameof(widthPt), "SVG dimensions must be finite and positive.");
                WidthPt = widthPt;
                HeightPt = heightPt;
                WidthEmu = checked((long)Math.Round(widthPt * EmusPerPoint,
                    MidpointRounding.AwayFromZero));
                HeightEmu = checked((long)Math.Round(heightPt * EmusPerPoint,
                    MidpointRounding.AwayFromZero));
            }

            internal double WidthPt { get; }
            internal double HeightPt { get; }
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
        internal LaTeXBlockRender(string svgPath, byte[] svgBytes, double depthPt,
            double fontSizePt, int textColor = LaTeXBlockService.AutomaticTextColor,
            byte[] contentSvgBytes = null, LaTeXBlockStyle style = null)
        {
            SvgPath = svgPath;
            SvgBytes = svgBytes ?? throw new ArgumentNullException(nameof(svgBytes));
            DepthPt = depthPt;
            FontSizePt = fontSizePt;
            TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
            ContentSvgBytes = contentSvgBytes ?? SvgBytes;
            Style = style;
        }
        internal string SvgPath { get; }
        internal byte[] SvgBytes { get; }
        internal double DepthPt { get; }
        internal double FontSizePt { get; }
        internal int TextColor { get; }
        // The decorated SVG is what Office inserts and previews.  Keep the raw TeX
        // content alongside it so a subsequent native resize can repaint the shell
        // at the new frame edges rather than nesting transparent SVG frames.
        internal byte[] ContentSvgBytes { get; }
        internal LaTeXBlockStyle Style { get; }
    }
}
