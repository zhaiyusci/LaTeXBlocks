using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
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
        private const string BatchMetadataTitlePrefix = "LaTeXBlocksBatch/";
        private const string WordprocessingNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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
            // owns layout and leading; the SVG owns foreground paint and the shell.
            // Auto formulas and numbered equations deliberately remain on Word's
            // native font-colour path and never acquire a block-style wrapper.
            var styledFixedContent = mode == LaTeXBlockLayoutMode.Fixed && style != null;
            var rendererWidthPt = widthPt;
            if (styledFixedContent)
            {
                textColor = ToWordColor(style.TextColor);
                // A non-null style is an explicit acceptance of the Word Block
                // editor, even when all visible controls happen to show their
                // defaults. Keep that promise literal: 1.20× leading is still
                // authored in TeX, while the SVG supplies inherited foreground
                // paint and the exact outer viewport. Older blocks
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
                // A natural-width hbox now owns a standard zero-width TeX strut, so
                // Auto content and natural-width displaystyle math use no unrelated
                // preview border. Legacy unstyled Fixed renders retain the historical
                // 1pt border. Set either value on every request so one Block cannot
                // leak viewport state into the next request in the warm worker.
                renderSource = "\\global\\PreviewBorder=" +
                    (mode == LaTeXBlockLayoutMode.Auto ? "0pt\n" : "1pt\n") +
                    (mode == LaTeXBlockLayoutMode.Auto
                        ? renderSource
                        : ApplyTextColor(renderSource, textColor));
            }
            var result = committed
                ? await renderers.RenderQueuedAsync(profile, renderSource, rendererWidthPt,
                    mode == LaTeXBlockLayoutMode.Auto, fontSizePt)
                : await renderers.RenderLatestAsync(profile, renderSource, rendererWidthPt,
                    mode == LaTeXBlockLayoutMode.Auto, fontSizePt);
            var finalSvg = styledFixedContent
                ? LaTeXBlockSvgFrame.Decorate(result.Bytes, style,
                    outerWidthPt ?? widthPt, outerHeightPt)
                : mode == LaTeXBlockLayoutMode.Auto
                    ? MarkHostTextColorUnset(result.Bytes)
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
            return InsertRenderedCore(source, widthPt, mode, render, style);
        }

        private WordInterop.InlineShape InsertRenderedCore(string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render, LaTeXBlockStyle style)
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
            Action markDocumentMutated = null)
        {
            var target = requestedTarget.Duplicate;
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
            ApplyContractMetadata(shape, source, metadata);
            shape = NormalizeWordInlineDrawing(shape, svgSize);
            EnsureInlineWordJoinerBoundaries(shape, metadata);
            ApplyHostRunFormat(shape, metadata, render.TextColor);
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
            var preserveFixedFrame = metadata != null &&
                metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                metadata.Role == LaTeXBlockRole.Content &&
                mode == LaTeXBlockLayoutMode.Fixed;
            var render = RenderPreview(source, widthPt, mode, profile, size, displayMathStyle,
                textColor, style,
                preserveFixedFrame && style != null ? oldShape.Height : (double?)null,
                preserveFixedFrame && style != null ? oldShape.Width : (double?)null);
            // UpdateBlock is also used by callers outside the editor. Preserve a
            // Fixed Content Block's exact Word-owned outer viewport there too;
            // otherwise replacing the SVG would silently restore its natural
            // content size and discard a user resize.
            // Styled Blocks already gave that exact frame to TeX, which owns the
            // vertical alignment. Only legacy unstyled Blocks need an SVG-only frame.
            if (preserveFixedFrame && style == null)
                render = FrameFloatingRender(render, oldShape.Width, oldShape.Height);
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
                ? (style != null
                    ? style.ToMetadataValue()
                    : previous.HasExplicitStyle
                        ? previous.Style.ToMetadataValue()
                        : null)
                : null;
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt, mode,
                render.FontSizePt, previous.Role, svgSize.WidthPt, svgSize.HeightPt, styleData);
            var previousUsesInlineWordJoinerBoundaries = UsesInlineWordJoinerBoundaries(previous);
            var nativeTextColor = NativeTextColorDescriptor.Automatic;
            var preserveNativeTextColor = !previous.HasExplicitStyle &&
                NativeTextColorDescriptor.TryCapture(oldShape.Range,
                    out nativeTextColor);
            // Baseline placement is derived layout, not state owned by either the
            // old drawing run or neighboring text. Word supplies the current line
            // baseline; the new TeX depth is the only required offset.
            WordInterop.InlineShape replacement = null;
            WordInlineRunFormatSnapshot hostRunFormat = null;
            var document = oldShape.Range.Document;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                // Replacing an InlineShape creates a new Word drawing character.
                // Preserve the old character's ordinary text-run formatting, then
                // deliberately overwrite only the properties owned by the freshly
                // rendered formula (size, colour and derived baseline) below.
                hostRunFormat = WordInlineRunFormatSnapshot.Capture(oldShape.Range);
                application.UndoRecord.StartCustomRecord("Update LaTeX Block");
                undoStarted = true;

                // Establish the final boundary state while the old formula still
                // identifies the exact character being replaced.
                if (previousUsesInlineWordJoinerBoundaries ||
                    UsesInlineWordJoinerBoundaries(metadata))
                {
                    documentMutated = true;
                    if (UsesInlineWordJoinerBoundaries(metadata))
                        EnsureInlineWordJoinerBoundaries(oldShape, metadata);
                    else
                        RemoveInlineWordJoinerBoundaries(oldShape);
                }
                // Word does not actually replace an existing InlineShape when its
                // one-character Range is supplied to AddPicture; it inserts beside
                // the drawing. Delete the old character first, retain its exact start,
                // and let the custom undo record restore it if any later step fails.
                // This avoids both a temporary duplicate and a surviving old copy.
                var replacementStart = oldShape.Range.Start;
                if (numbered)
                {
                    documentMutated = true;
                    ConfigureNumberedEquationTabs(oldShape.Range.Paragraphs[1], numberedLayout);
                }
                var insertionPath = PrepareInsertionSvg(render, mode);
                documentMutated = true;
                oldShape.Delete();
                target = document.Range(replacementStart, replacementStart);
                replacement = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
                ApplyContractMetadata(replacement, source, metadata);
                replacement = NormalizeWordInlineDrawing(replacement, svgSize);
                hostRunFormat.Apply(replacement.Range);
                ApplyHostRunFormat(replacement, metadata, render.TextColor);
                // ApplyContract and NormalizeWordInlineDrawing rebuild the drawing run,
                // and ApplyHostRunFormat deliberately writes the resolved render RGB.
                // Restore Word's richer native value last so a theme colour remains a
                // live theme slot+tint rather than being silently downgraded to RGB.
                if (preserveNativeTextColor)
                    nativeTextColor.ApplyTo(replacement.Range);
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
                hostRunFormat?.Dispose();
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        internal void UpdateRenderedBatch(IList<LaTeXBlockBatchUpdate> updates)
        {
            if (updates == null) throw new ArgumentNullException(nameof(updates));
            if (updates.Count < 2)
                throw new ArgumentException("A batch update requires at least two formulas.",
                    nameof(updates));

            var states = new List<BatchInlineUpdateState>(updates.Count);
            WordInterop.Document document = null;
            var undoStarted = false;
            var documentMutated = false;
            try
            {
                foreach (var update in updates)
                {
                    if (update == null || update.Shape == null || update.Render == null)
                        throw new ArgumentException("A batch update contains an empty item.",
                            nameof(updates));
                    var previous = update.Metadata;
                    if (previous == null ||
                        previous.Mode != LaTeXBlockLayoutMode.Auto ||
                        previous.Role != LaTeXBlockRole.Content)
                        throw new InvalidOperationException(
                            "Only ordinary Auto inline formulas can share a batch replacement.");

                    var range = update.Range;
                    if (document == null) document = range.Document;
                    var svgSize = ReadSvgPhysicalSize(update.Render.SvgBytes);
                    var metadata = new LaTeXBlockMetadata(previous.Id, update.WidthPt,
                        update.Render.DepthPt, LaTeXBlockLayoutMode.Auto,
                        update.Render.FontSizePt, previous.Role, svgSize.WidthPt,
                        svgSize.HeightPt, null);
                    states.Add(new BatchInlineUpdateState(update, range.Start,
                        range.StoryType, update.ParagraphStart, update.ParagraphEnd,
                        metadata, svgSize,
                        PrepareInsertionSvg(update.Render, LaTeXBlockLayoutMode.Auto)));
                }

                states.Sort((left, right) => left.Start.CompareTo(right.Start));
                var storyType = states[0].StoryType;
                var paragraphStart = states[0].ParagraphStart;
                var paragraphEnd = states[0].ParagraphEnd;
                for (var index = 0; index < states.Count; index++)
                {
                    if (states[index].StoryType != storyType ||
                        states[index].ParagraphStart != paragraphStart ||
                        states[index].ParagraphEnd != paragraphEnd ||
                        (index > 0 && states[index - 1].Start >= states[index].Start))
                        throw new InvalidOperationException(
                            "A formula batch must contain distinct shapes in one Word paragraph.");
                }

                // Read the original drawing runs once. Their rPr elements contain all
                // host-owned character formatting (highlight, language, emphasis,
                // spacing, scaling, theme colour, and so on). Moving those elements
                // in OpenXML is both more complete and substantially cheaper than
                // round-tripping a Word Font object for every formula.
                var originalFirst = states[0].Update.Range;
                var originalLast = states[states.Count - 1].Update.Range;
                var originalEnvelope = originalFirst.Duplicate;
                originalEnvelope.SetRange(originalFirst.Start, originalLast.End);
                var envelopeStart = originalEnvelope.Start;
                var finalEnvelopeEnd = originalEnvelope.End;
                var expectedInlineShapeCount = originalEnvelope.InlineShapes.Count;
                var formats = CaptureWordInlineFormatsXml(originalEnvelope.WordOpenXML,
                    states);

                var following = states[states.Count - 1].Update.Range.Duplicate;
                following.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
                var hadParagraphMarkAfter =
                    following.MoveEnd(WordInterop.WdUnits.wdCharacter, 1) == 1 &&
                    following.Text == "\r";

                application.UndoRecord.StartCustomRecord("Update LaTeX Blocks");
                undoStarted = true;
                for (var index = states.Count - 1; index >= 0; index--)
                {
                    var state = states[index];
                    var insertion = state.Update.Range.Duplicate;
                    insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                    documentMutated = true;
                    state.ImportedShape = insertion.InlineShapes.AddPicture(
                        state.InsertionPath, LinkToFile: false,
                        SaveWithDocument: true, Range: insertion);
                    state.ImportedShape.Title = BatchMetadataTitlePrefix +
                        state.Metadata.Id.ToString("D");
                }

                var first = states[0].ImportedShape.Range;
                var last = states[states.Count - 1].Update.Shape.Range;
                var envelope = first.Duplicate;
                envelope.SetRange(first.Start, last.End);
                var importedXml = envelope.WordOpenXML;
                var normalized = NormalizeWordInlineDrawingsXml(importedXml,
                    formats, out var normalizedCount, out var removedCount);
                if (normalizedCount != states.Count)
                    throw new InvalidDataException("Word batch XML did not contain every target formula.");
                if (removedCount != states.Count)
                    throw new InvalidDataException("Word batch XML did not contain every old formula.");

                var insertionRange = envelope.Duplicate;
                insertionRange.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                envelope.Delete();
                insertionRange.InsertXML(normalized);

                if (!hadParagraphMarkAfter)
                {
                    var separator = insertionRange.Duplicate;
                    separator.SetRange(finalEnvelopeEnd, finalEnvelopeEnd + 1);
                    if (separator.Text == "\r") separator.Delete();
                }
                var restoredEnvelope = insertionRange.Duplicate;
                restoredEnvelope.SetRange(envelopeStart, finalEnvelopeEnd);
                if (restoredEnvelope.InlineShapes.Count != expectedInlineShapeCount)
                    throw new InvalidDataException(
                        "Word did not preserve every drawing during batch replacement.");
            }
            catch (Exception exception)
            {
                var rollbackFailure = document == null
                    ? null
                    : TryRollbackCustomRecord(document, ref undoStarted,
                        documentMutated);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete or roll back the batched LaTeX Block update.",
                        new AggregateException(exception, rollbackFailure));
                throw;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        internal bool TryUpdateTextColorsBatch(
            IList<LaTeXBlockColorUpdate> updates, out bool changedDocument)
        {
            changedDocument = false;
            if (updates == null) throw new ArgumentNullException(nameof(updates));
            if (updates.Count == 0) return false;

            var ordered = new List<LaTeXBlockColorUpdate>(updates);
            ordered.Sort((left, right) => left.Range.Start.CompareTo(right.Range.Start));
            var story = ordered[0].Range.StoryType;
            for (var index = 0; index < ordered.Count; index++)
            {
                var update = ordered[index];
                if (update.Shape == null || update.Range == null ||
                    update.Metadata == null ||
                    update.Metadata.Mode != LaTeXBlockLayoutMode.Auto ||
                    update.Metadata.Role != LaTeXBlockRole.Content ||
                    update.Range.StoryType != story ||
                    (index > 0 && ordered[index - 1].Range.Start >= update.Range.Start))
                    return false;
            }
            if (story != WordInterop.WdStoryType.wdMainTextStory) return false;

            var groups = new List<List<LaTeXBlockColorUpdate>>();
            foreach (var update in ordered)
            {
                if (groups.Count == 0 ||
                    groups[groups.Count - 1][0].ParagraphStart !=
                        update.ParagraphStart ||
                    groups[groups.Count - 1][0].ParagraphEnd != update.ParagraphEnd)
                    groups.Add(new List<LaTeXBlockColorUpdate>());
                groups[groups.Count - 1].Add(update);
            }
            var preparedGroups = new List<PreparedColorBatchGroup>(groups.Count);
            foreach (var group in groups)
            {
                var first = group[0].Range;
                var last = group[group.Count - 1].Range;
                var envelope = first.Duplicate;
                envelope.SetRange(first.Start, last.End);
                if (!TryDetachHostColorWordMediaXml(envelope.WordOpenXML, group,
                        out var recoloredXml, out var requiresWrite))
                    return false;
                if (!requiresWrite) continue;
                var following = last.Duplicate;
                following.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
                var hadParagraphMarkAfter =
                    following.MoveEnd(WordInterop.WdUnits.wdCharacter, 1) == 1 &&
                    following.Text == "\r";
                preparedGroups.Add(new PreparedColorBatchGroup(envelope,
                    envelope.Start, envelope.End, envelope.InlineShapes.Count,
                    hadParagraphMarkAfter, recoloredXml));
            }

            // New SVGs deliberately leave their default paint unset. Word's native
            // run colour has already completed the operation, so there is no media
            // or drawing character to replace and therefore nothing that can flash.
            if (preparedGroups.Count == 0) return true;

            var document = ordered[0].Range.Document;
            var undoStarted = false;
            try
            {
                application.UndoRecord.StartCustomRecord("Update LaTeX Block Colors");
                undoStarted = true;
                for (var index = preparedGroups.Count - 1; index >= 0; index--)
                {
                    var group = preparedGroups[index];
                    var insertion = group.Envelope.Duplicate;
                    insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                    changedDocument = true;
                    group.Envelope.Delete();
                    insertion.InsertXML(group.RecoloredXml);
                    if (!group.HadParagraphMarkAfter)
                    {
                        var separator = insertion.Duplicate;
                        separator.SetRange(group.End, group.End + 1);
                        if (separator.Text == "\r") separator.Delete();
                    }
                    var restored = insertion.Duplicate;
                    restored.SetRange(group.Start, group.End);
                    if (restored.InlineShapes.Count !=
                        group.ExpectedInlineShapeCount)
                        throw new InvalidDataException(
                            "Word did not preserve every formula during colour refresh.");
                }
                return true;
            }
            catch (Exception exception)
            {
                var rollbackFailure = TryRollbackCustomRecord(document,
                    ref undoStarted, changedDocument);
                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not complete or roll back the formula colour refresh.",
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
        /// Gives a legacy unstyled fixed Block its exact Word-owned transparent
        /// viewport. Styled Blocks receive the frame before TeX layout and never
        /// pass through this SVG-only fallback.
        /// </summary>
        internal LaTeXBlockRender FrameFloatingRender(LaTeXBlockRender render,
            double frameWidthPt, double frameHeightPt)
        {
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (render.Style != null)
                throw new InvalidOperationException(
                    "Styled Blocks must be rendered against their exact outer frame before SVG composition.");
            var framedBytes = FrameSvg(render.SvgBytes, frameWidthPt, frameHeightPt);
            return new LaTeXBlockRender(WriteSvg(framedBytes), framedBytes, render.DepthPt,
                render.FontSizePt, render.TextColor, render.ContentSvgBytes, null);
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

        /// <summary>
        /// Returns an auto-width formula only while Word is exposing its ordinary
        /// exact InlineShape selection. For ordinary inline content the two U+2060
        /// placement boundaries are not part of the user's selection: including them
        /// changes copy, navigation and picture-handle semantics even though they have
        /// no visible width. Numbered auto-width formulas use the same exact selection
        /// contract without those boundaries.
        /// </summary>
        internal bool TryGetExactlySelectedInlineFormula(out WordInterop.InlineShape shape)
        {
            shape = null;
            if (application.Documents.Count == 0 || application.Selection == null)
                return false;
            var selection = application.Selection;
            if (selection.Type != WordInterop.WdSelectionType.wdSelectionInlineShape ||
                selection.InlineShapes.Count != 1)
                return false;
            var candidate = selection.InlineShapes[1];
            if (!TryReadContract(candidate, out var metadata, out _) ||
                metadata.Mode != LaTeXBlockLayoutMode.Auto ||
                selection.Start != candidate.Range.Start ||
                selection.End != candidate.Range.End)
                return false;
            shape = candidate;
            return true;
        }

        private static WordInterop.Range TryGetInlineFormulaNativeTextRange(
            WordInterop.InlineShape shape)
        {
            if (shape == null || !TryReadContract(shape, out var metadata, out _) ||
                !UsesInlineWordJoinerBoundaries(metadata))
                return null;
            var leadingJoiner = AdjacentCharacter(shape.Range, true);
            var trailingJoiner = AdjacentCharacter(shape.Range, false);
            if (!IsWordJoiner(leadingJoiner) || !IsWordJoiner(trailingJoiner)) return null;
            return shape.Range.Document.Range(leadingJoiner.Start, trailingJoiner.End);
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
            // The shared block contract is top-left: reducing a frame height
            // preserves the first rendered line.
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

        private static void ApplyContractMetadata(WordInterop.InlineShape shape,
            string source, LaTeXBlockMetadata metadata)
        {
            shape.AlternativeText = NormalizeSourceText(source);
            shape.Title = metadata.ToString();
            shape.LockAspectRatio = HasIndependentFrameResize(metadata)
                ? Office.MsoTriState.msoFalse
                : Office.MsoTriState.msoTrue;
        }

        private static void ApplyHostRunFormat(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata, int textColor)
        {
            // Script placement belongs to the TeX source. Word's Subscript and
            // Superscript flags are a second size/baseline transform, so they must
            // never survive onto either a newly inserted or a replacement formula.
            shape.Range.Font.Subscript = 0;
            shape.Range.Font.Superscript = 0;
            // The image's physical dimensions already come from the SVG. This run size is
            // Word's semantic host size, used by the Font Size UI and by format-change
            // detection. InsertXML drops the drawing run's w:sz along with w:position, so
            // both values must be restored on the final normalized InlineShape.
            if (metadata.Mode == LaTeXBlockLayoutMode.Auto)
            {
                shape.Range.Font.Size = (float)metadata.FontSizePt;
                shape.Range.Font.SizeBi = (float)metadata.FontSizePt;
            }
            // This is the colour actually used by the completed render: native Word
            // Font Color for Auto/legacy objects, or durable style colour for a styled
            // Fixed Block. Alternative Text remains exactly the author-written source.
            shape.Range.Font.Color = (WordInterop.WdColor)NormalizeTextColor(textColor);
            if (metadata.Mode == LaTeXBlockLayoutMode.Auto)
                ApplyGraphicFill(shape, textColor);
            ApplyBaselinePosition(shape, metadata);
            // InsertXML used by NormalizeWordInlineDrawing rebuilds the DrawingML
            // object and may restore Word's default aspect lock. Fixed Content Blocks
            // intentionally own a two-dimensional frame, so restore that setting
            // after every normalize/update. Do not rewrite an Auto/numbered drawing
            // here: toggling its aspect lock after normalization can make Word add
            // effect extents to an otherwise exact inline SVG.
            if (HasIndependentFrameResize(metadata))
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
        }

        internal static void ApplyGraphicFill(WordInterop.InlineShape shape,
            int textColor)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            WordInterop.FillFormat fill = null;
            WordInterop.ColorFormat foreground = null;
            try
            {
                fill = shape.Fill;
                fill.Visible = Office.MsoTriState.msoTrue;
                fill.Solid();
                foreground = fill.ForeColor;
                var color = NormalizeTextColor(textColor);
                foreground.RGB = color == AutomaticTextColor ? 0 : color;
            }
            finally
            {
                if (foreground != null) Marshal.ReleaseComObject(foreground);
                if (fill != null) Marshal.ReleaseComObject(fill);
            }
        }

        internal bool TryApplyGraphicFillsBatch(
            IList<LaTeXBlockColorUpdate> updates)
        {
            if (updates == null) throw new ArgumentNullException(nameof(updates));
            if (updates.Count == 0) return false;
            foreach (var update in updates)
                if (update?.Shape == null || update.Metadata == null ||
                    update.Metadata.Mode != LaTeXBlockLayoutMode.Auto ||
                    update.Metadata.Role != LaTeXBlockRole.Content)
                    return false;

            var undoStarted = false;
            try
            {
                application.UndoRecord.StartCustomRecord(
                    "Update LaTeX Block Graphic Fills");
                undoStarted = true;
                foreach (var update in updates)
                    ApplyGraphicFill(update.Shape, update.TargetTextColor);
                return true;
            }
            finally
            {
                if (undoStarted)
                    try { application.UndoRecord.EndCustomRecord(); } catch { }
            }
        }

        /// <summary>
        /// Captures direct Word character formatting that is independent of formula
        /// rendering. Font.Duplicate keeps the host font-family, emphasis, underline,
        /// spacing/scaling, hidden state, and related values without copying the
        /// drawing itself. Range-level proofing, highlight, and language values are
        /// captured separately. Size and colour are restored here only as a base
        /// state and are then overwritten from the merged formula request; vertical
        /// placement is always cleared and recomputed by ApplyHostRunFormat.
        /// </summary>
        private sealed class WordInlineRunFormatSnapshot : IDisposable
        {
            private WordInterop.Font font;
            private readonly int noProofing;
            private readonly WordInterop.WdColorIndex highlightColorIndex;
            private readonly WordInterop.WdLanguageID languageId;
            private readonly WordInterop.WdLanguageID languageIdFarEast;
            private readonly WordInterop.WdLanguageID languageIdOther;

            private WordInlineRunFormatSnapshot(WordInterop.Font font, int noProofing,
                WordInterop.WdColorIndex highlightColorIndex,
                WordInterop.WdLanguageID languageId,
                WordInterop.WdLanguageID languageIdFarEast,
                WordInterop.WdLanguageID languageIdOther)
            {
                this.font = font;
                this.noProofing = noProofing;
                this.highlightColorIndex = highlightColorIndex;
                this.languageId = languageId;
                this.languageIdFarEast = languageIdFarEast;
                this.languageIdOther = languageIdOther;
            }

            internal static WordInlineRunFormatSnapshot Capture(WordInterop.Range range)
            {
                if (range == null) throw new ArgumentNullException(nameof(range));
                WordInterop.Font sourceFont = null;
                WordInterop.Font duplicate = null;
                try
                {
                    sourceFont = range.Font;
                    duplicate = sourceFont.Duplicate;
                    return new WordInlineRunFormatSnapshot(duplicate, range.NoProofing,
                        range.HighlightColorIndex, range.LanguageID,
                        range.LanguageIDFarEast, range.LanguageIDOther);
                }
                catch
                {
                    if (duplicate != null) Marshal.FinalReleaseComObject(duplicate);
                    throw;
                }
                finally
                {
                    if (sourceFont != null) Marshal.ReleaseComObject(sourceFont);
                }
            }

            internal void Apply(WordInterop.Range range)
            {
                if (range == null) throw new ArgumentNullException(nameof(range));
                if (font == null) throw new ObjectDisposedException(nameof(WordInlineRunFormatSnapshot));
                range.Font = font;
                range.NoProofing = noProofing;
                range.HighlightColorIndex = highlightColorIndex;
                range.LanguageID = languageId;
                range.LanguageIDFarEast = languageIdFarEast;
                range.LanguageIDOther = languageIdOther;
            }

            public void Dispose()
            {
                if (font == null) return;
                Marshal.FinalReleaseComObject(font);
                font = null;
            }
        }

        private static bool HasIndependentFrameResize(LaTeXBlockMetadata metadata)
        {
            return metadata != null && metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                   metadata.Role == LaTeXBlockRole.Content;
        }

        internal enum NativeTextColorKind
        {
            Automatic,
            Direct,
            Theme
        }

        /// <summary>
        /// Word's native font colour is not always an RGB value. Theme colours retain
        /// a theme slot and tint/shade, while Automatic is a distinct sentinel. Keep
        /// that semantic value long enough to restore it after Word replaces the SVG;
        /// the separately resolved BGR integer remains the paint input for StemTeX.
        /// </summary>
        internal readonly struct NativeTextColorDescriptor : IEquatable<NativeTextColorDescriptor>
        {
            private NativeTextColorDescriptor(NativeTextColorKind kind, int directWordColor,
                WordInterop.WdThemeColorIndex themeColor, float tintAndShade)
            {
                Kind = kind;
                DirectWordColor = directWordColor;
                ThemeColor = themeColor;
                TintAndShade = tintAndShade;
            }

            internal NativeTextColorKind Kind { get; }
            internal int DirectWordColor { get; }
            internal WordInterop.WdThemeColorIndex ThemeColor { get; }
            internal float TintAndShade { get; }

            internal static NativeTextColorDescriptor Automatic =>
                new NativeTextColorDescriptor(NativeTextColorKind.Automatic,
                    AutomaticTextColor, WordInterop.WdThemeColorIndex.wdNotThemeColor, 0);

            internal static NativeTextColorDescriptor Direct(int wordColor)
            {
                if (wordColor < 0 || wordColor > 0x00ffffff)
                    throw new ArgumentOutOfRangeException(nameof(wordColor));
                return new NativeTextColorDescriptor(NativeTextColorKind.Direct, wordColor,
                    WordInterop.WdThemeColorIndex.wdNotThemeColor, 0);
            }

            internal static NativeTextColorDescriptor Theme(
                WordInterop.WdThemeColorIndex themeColor, float tintAndShade)
            {
                if (themeColor == WordInterop.WdThemeColorIndex.wdNotThemeColor)
                    throw new ArgumentOutOfRangeException(nameof(themeColor));
                if (float.IsNaN(tintAndShade) || float.IsInfinity(tintAndShade) ||
                    tintAndShade < -1f || tintAndShade > 1f)
                    throw new ArgumentOutOfRangeException(nameof(tintAndShade));
                return new NativeTextColorDescriptor(NativeTextColorKind.Theme, 0,
                    themeColor, tintAndShade);
            }

            internal static bool TryCaptureCollapsedSelection(WordInterop.Selection selection,
                out NativeTextColorDescriptor descriptor)
            {
                descriptor = Automatic;
                if (selection == null || selection.Start != selection.End) return false;
                WordInterop.Font font = null;
                try
                {
                    // A collapsed Range.Font describes the character on the right in
                    // several boundary cases. Selection.Font is Word's actual typing
                    // format, which is exactly what ExecuteMso updates for the probe.
                    font = selection.Font;
                    return TryCapture(font, out descriptor);
                }
                catch (COMException)
                {
                    return false;
                }
                finally
                {
                    if (font != null) Marshal.ReleaseComObject(font);
                }
            }

            internal static bool TryCapture(WordInterop.Range range,
                out NativeTextColorDescriptor descriptor)
            {
                descriptor = Automatic;
                if (range == null) return false;
                WordInterop.Font font = null;
                try
                {
                    font = range.Font;
                    return TryCapture(font, out descriptor);
                }
                catch (COMException)
                {
                    return false;
                }
                finally
                {
                    if (font != null) Marshal.ReleaseComObject(font);
                }
            }

            private static bool TryCapture(WordInterop.Font font,
                out NativeTextColorDescriptor descriptor)
            {
                descriptor = Automatic;
                if (font == null) return false;
                var raw = (int)font.Color;
                if (raw == UndefinedTextColor) return false;
                if (raw == AutomaticTextColor)
                {
                    descriptor = Automatic;
                    return true;
                }
                if (raw >= 0 && raw <= 0x00ffffff)
                {
                    descriptor = Direct(raw);
                    return true;
                }

                WordInterop.ColorFormat colorFormat = null;
                try
                {
                    colorFormat = font.TextColor;
                    if (colorFormat.Type != Office.MsoColorType.msoColorTypeScheme ||
                        colorFormat.ObjectThemeColor ==
                            WordInterop.WdThemeColorIndex.wdNotThemeColor)
                        return false;
                    var tint = colorFormat.TintAndShade;
                    if (float.IsNaN(tint) || float.IsInfinity(tint) ||
                        tint < -1f || tint > 1f)
                        return false;
                    descriptor = Theme(colorFormat.ObjectThemeColor, tint);
                    return true;
                }
                catch (COMException)
                {
                    return false;
                }
                finally
                {
                    if (colorFormat != null) Marshal.ReleaseComObject(colorFormat);
                }
            }

            internal void ApplyTo(WordInterop.Range range)
            {
                if (range == null) throw new ArgumentNullException(nameof(range));
                WordInterop.Font font = null;
                WordInterop.ColorFormat colorFormat = null;
                try
                {
                    font = range.Font;
                    if (Kind == NativeTextColorKind.Automatic)
                    {
                        font.Color = (WordInterop.WdColor)AutomaticTextColor;
                    }
                    else if (Kind == NativeTextColorKind.Direct)
                    {
                        font.Color = (WordInterop.WdColor)DirectWordColor;
                    }
                    else
                    {
                        colorFormat = font.TextColor;
                        colorFormat.ObjectThemeColor = ThemeColor;
                        colorFormat.TintAndShade = TintAndShade;
                    }
                }
                finally
                {
                    if (colorFormat != null) Marshal.ReleaseComObject(colorFormat);
                    if (font != null) Marshal.ReleaseComObject(font);
                }
            }

            public bool Equals(NativeTextColorDescriptor other)
            {
                return Kind == other.Kind && DirectWordColor == other.DirectWordColor &&
                       ThemeColor == other.ThemeColor &&
                       Math.Abs(TintAndShade - other.TintAndShade) < 0.000001f;
            }

            public override bool Equals(object value)
            {
                return value is NativeTextColorDescriptor other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hash = (int)Kind;
                    hash = (hash * 397) ^ DirectWordColor;
                    hash = (hash * 397) ^ (int)ThemeColor;
                    hash = (hash * 397) ^ TintAndShade.GetHashCode();
                    return hash;
                }
            }
        }

        internal static int ResolveTextColor(WordInterop.Range target,
            int fallback = AutomaticTextColor)
        {
            fallback = NormalizeTextColor(fallback);
            if (target == null) return fallback;
            if (TryReadTextColor(target.Font, out var color, out var isThemeColor))
                return color;
            if (isThemeColor && TryResolveTextColorFromWordOpenXml(target, out color))
                return color;

            // Word returns wdUndefined for a mixed selection. The formula replaces
            // at its start, so use the insertion character's colour just as we do
            // for a mixed font-size selection.
            if (target.Start != target.End)
            {
                var insertion = target.Duplicate;
                insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                if (TryReadTextColor(insertion.Font, out color, out isThemeColor))
                    return color;
                if (isThemeColor &&
                    TryResolveTextColorFromWordOpenXml(insertion, out color))
                    return color;
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
            if (TryReadTextColor(selection.Font, out var color, out var isThemeColor))
                return color;
            if (isThemeColor &&
                TryResolveTextColorFromWordOpenXml(selection.Range, out color))
                return color;
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

        private static byte[] MarkHostTextColorUnset(byte[] svgBytes)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new InvalidDataException("StemTeX returned an empty SVG.");
            var svg = Encoding.UTF8.GetString(svgBytes);
            var root = Regex.Match(svg, "<svg\\b[^>]*>",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (!root.Success)
                throw new InvalidDataException("StemTeX returned a malformed SVG root.");
            var markedRoot = SetSvgStringAttribute(root.Value,
                "data-latexblocks-host-color", "unset");
            svg = svg.Remove(root.Index, root.Length).Insert(root.Index, markedRoot);
            return Encoding.UTF8.GetBytes(svg);
        }

        private static bool TryReadTextColor(WordInterop.Font font, out int color,
            out bool isThemeColor)
        {
            color = AutomaticTextColor;
            isThemeColor = false;
            if (font == null) return false;
            var raw = (int)font.Color;
            if (raw == UndefinedTextColor) return false;
            if (raw == AutomaticTextColor)
            {
                color = AutomaticTextColor;
                return true;
            }
            if (raw >= 0 && raw <= 0x00ffffff)
            {
                color = raw;
                return true;
            }
            // Word exposes theme colours as a negative encoded WdColor through
            // Font.Color and TextColor.RGB.  That value is not RGB and must not be
            // collapsed to Automatic.  Word's Flat OPC serialisation of the complete
            // formula scaffold carries the currently resolved RGB in w:color/@w:val.
            try
            {
                isThemeColor = font.TextColor.Type == Office.MsoColorType.msoColorTypeScheme;
            }
            catch (COMException)
            {
                isThemeColor = false;
            }
            return false;
        }

        private static bool TryResolveTextColorFromWordOpenXml(WordInterop.Range target,
            out int color)
        {
            color = AutomaticTextColor;
            if (target == null) return false;
            try
            {
                var serializationRange = target;
                if (target.InlineShapes.Count == 1)
                {
                    var formulaRange = TryGetInlineFormulaNativeTextRange(
                        target.InlineShapes[1]);
                    if (formulaRange != null) serializationRange = formulaRange;
                }
                return TryParseResolvedTextColorFromWordOpenXml(
                    serializationRange.WordOpenXML, out color);
            }
            catch (COMException)
            {
                return false;
            }
        }

        internal static bool TryParseResolvedTextColorFromWordOpenXml(string wordOpenXml,
            out int color)
        {
            color = AutomaticTextColor;
            if (string.IsNullOrWhiteSpace(wordOpenXml)) return false;
            try
            {
                var document = new XmlDocument { XmlResolver = null };
                document.LoadXml(wordOpenXml);
                var namespaces = new XmlNamespaceManager(document.NameTable);
                namespaces.AddNamespace("pkg",
                    "http://schemas.microsoft.com/office/2006/xmlPackage");
                namespaces.AddNamespace("w",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                var mainDocument = document.SelectSingleNode(
                    "//pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData",
                    namespaces) ?? (XmlNode)document;
                var run = mainDocument.SelectSingleNode(".//w:r[w:drawing]", namespaces);
                if (run == null) return false;
                var colorNode = run.SelectSingleNode("w:rPr/w:color", namespaces);
                // Automatic is identified from Font.Color before this parser runs.
                // A missing node here therefore means the negative COM value was not
                // a serialised theme colour; retain the caller's normal fallback.
                if (colorNode == null) return false;
                var value = colorNode.Attributes?["val",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"]?.Value;
                if (string.IsNullOrWhiteSpace(value) ||
                    string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    color = AutomaticTextColor;
                    return true;
                }
                if (value.Length != 6 ||
                    !int.TryParse(value, NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var rgb))
                    return false;
                var red = (rgb >> 16) & 0xff;
                var green = (rgb >> 8) & 0xff;
                var blue = rgb & 0xff;
                color = red | (green << 8) | (blue << 16);
                return true;
            }
            catch (XmlException)
            {
                return false;
            }
        }

        private static void ApplyBaselinePosition(WordInterop.InlineShape shape,
            LaTeXBlockMetadata metadata)
        {
            if (metadata.Mode != LaTeXBlockLayoutMode.Auto)
            {
                // A fixed-width Content Block is a page-like frame, not one
                // baseline-bearing text line. Its vertical placement is already
                // resolved inside TeX, so no surrounding-run offset applies here.
                shape.Range.Font.Position = 0;
                return;
            }
            // Word aligns the bottom of an InlineShape to the text baseline. Move the image
            // character down by the TeX box depth. This is always the TeX/Western baseline:
            // CJK glyph extents inside the SVG do not define a second alignment reference.
            // Font.Position is relative to Word's current line baseline. Moving the
            // image down by its TeX depth makes the baseline inside the SVG coincide
            // with that line baseline. Neighboring run positions are unrelated.
            shape.Range.Font.Position =
                -(int)Math.Round(metadata.DepthPt, MidpointRounding.AwayFromZero);
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

            // Check each immediate neighbor before inserting anything, and never
            // coalesce joiners that were already present in the user's document.
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
            // InsertXML normally appends a temporary paragraph boundary to the
            // imported Flat OPC run. At the physical end of an existing paragraph,
            // however, Word reuses the paragraph mark that was already after the
            // drawing. Remember that distinction before deleting the original
            // shape: an existing paragraph mark belongs to the document and must
            // never be removed as normalization scaffolding.
            var hadParagraphMarkAfter = false;
            var originalFollowingCharacter = shape.Range.Duplicate;
            originalFollowingCharacter.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
            if (originalFollowingCharacter.MoveEnd(WordInterop.WdUnits.wdCharacter, 1) == 1)
                hadParagraphMarkAfter = originalFollowingCharacter.Text == "\r";
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
            if (!hadParagraphMarkAfter &&
                separator.MoveEnd(WordInterop.WdUnits.wdCharacter, 1) == 1 &&
                separator.Text == "\r")
                separator.Delete();
            replacement.Range.ParagraphFormat = paragraphFormat;
            return replacement;
        }

        private static Dictionary<Guid, BatchInlineXmlFormat> CaptureWordInlineFormatsXml(
            string flatOpc, IList<BatchInlineUpdateState> states)
        {
            if (flatOpc == null) throw new ArgumentNullException(nameof(flatOpc));
            if (states == null) throw new ArgumentNullException(nameof(states));
            var targets = new Dictionary<Guid, BatchInlineUpdateState>();
            foreach (var state in states) targets[state.Metadata.Id] = state;
            var formats = new Dictionary<Guid, BatchInlineXmlFormat>();
            foreach (Match run in Regex.Matches(flatOpc,
                "<w:r\\b[^>]*>[\\s\\S]*?</w:r>", RegexOptions.CultureInvariant,
                TimeSpan.FromSeconds(2)))
            {
                if (!TryReadWordInlineMetadata(run.Value, out var metadata) ||
                    !targets.TryGetValue(metadata.Id, out var state))
                    continue;
                var properties = Regex.Match(run.Value,
                    "<w:rPr\\b(?:[^>]*/>|[^>]*>[\\s\\S]*?</w:rPr>)",
                    RegexOptions.CultureInvariant);
                formats[metadata.Id] = new BatchInlineXmlFormat(state.SvgSize,
                    state.Metadata.FontSizePt, state.Metadata.DepthPt,
                    properties.Success ? properties.Value : "<w:rPr></w:rPr>",
                    state.Metadata, state.Update.Source);
            }
            if (formats.Count != states.Count)
                throw new InvalidDataException(
                    "Word batch XML did not contain every original formula run.");
            return formats;
        }

        private static bool ContainsExplicitColourSource(string source)
        {
            return Regex.IsMatch(source ?? string.Empty,
                "\\\\(?:color|textcolor|pagecolor|colorbox|fcolorbox|includegraphics)\\b|\\\\begin\\s*\\{(?:tikzpicture|pgfpicture)\\}",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                TimeSpan.FromSeconds(1));
        }

        private static bool TryDetachHostColorWordMediaXml(string flatOpc,
            IList<LaTeXBlockColorUpdate> updates, out string normalizedXml,
            out bool requiresWrite)
        {
            normalizedXml = null;
            requiresWrite = false;
            try
            {
                var document = new XmlDocument
                {
                    XmlResolver = null,
                    PreserveWhitespace = true
                };
                document.LoadXml(flatOpc);
                const string packageNamespace =
                    "http://schemas.microsoft.com/office/2006/xmlPackage";
                const string wordNamespace =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                const string drawingWordNamespace =
                    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                const string drawingNamespace =
                    "http://schemas.openxmlformats.org/drawingml/2006/main";
                const string officeRelationshipNamespace =
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                const string svgNamespace =
                    "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
                const string packageRelationshipNamespace =
                    "http://schemas.openxmlformats.org/package/2006/relationships";
                var namespaces = new XmlNamespaceManager(document.NameTable);
                namespaces.AddNamespace("pkg", packageNamespace);
                namespaces.AddNamespace("w", wordNamespace);
                namespaces.AddNamespace("wp", drawingWordNamespace);
                namespaces.AddNamespace("a", drawingNamespace);
                namespaces.AddNamespace("r", officeRelationshipNamespace);
                namespaces.AddNamespace("asvg", svgNamespace);
                namespaces.AddNamespace("pr", packageRelationshipNamespace);

                var mainPart = document.SelectSingleNode(
                    "//pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData",
                    namespaces);
                var relationshipPart = document.SelectSingleNode(
                    "//pkg:part[@pkg:name='/word/_rels/document.xml.rels']/pkg:xmlData",
                    namespaces);
                if (mainPart == null || relationshipPart == null)
                {
                    return false;
                }

                var relationshipTargets = new Dictionary<string, string>(
                    StringComparer.Ordinal);
                foreach (XmlNode relationship in relationshipPart.SelectNodes(
                    ".//pr:Relationship", namespaces))
                {
                    var id = relationship.Attributes?["Id"]?.Value;
                    var target = relationship.Attributes?["Target"]?.Value;
                    if (!string.IsNullOrWhiteSpace(id) &&
                        !string.IsNullOrWhiteSpace(target) &&
                        !target.Contains(":"))
                        relationshipTargets[id] = "/word/" +
                            target.Replace('\\', '/').TrimStart('/');
                }
                var packageParts = new Dictionary<string, XmlNode>(
                    StringComparer.OrdinalIgnoreCase);
                foreach (XmlNode part in document.SelectNodes("//pkg:part", namespaces))
                {
                    var name = part.Attributes?["name", packageNamespace]?.Value;
                    if (!string.IsNullOrWhiteSpace(name)) packageParts[name] = part;
                }

                var targets = new Dictionary<Guid, LaTeXBlockColorUpdate>();
                foreach (var update in updates) targets[update.Metadata.Id] = update;
                var found = new HashSet<Guid>();
                var changedParts = new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
                foreach (XmlNode run in mainPart.SelectNodes(".//w:r[w:drawing]",
                    namespaces))
                {
                    var documentProperties = run.SelectSingleNode(".//wp:docPr",
                        namespaces);
                    var title = documentProperties?.Attributes?["title"]?.Value;
                    if (!LaTeXBlockMetadata.TryParse(title, out var metadata) ||
                        !targets.TryGetValue(metadata.Id, out var update))
                        continue;
                    var svgBlip = run.SelectSingleNode(".//asvg:svgBlip[@r:embed]",
                        namespaces);
                    var svgId = svgBlip?.Attributes?["embed",
                        officeRelationshipNamespace]?.Value;
                    if (string.IsNullOrWhiteSpace(svgId))
                    {
                        return false;
                    }
                    if (!relationshipTargets.TryGetValue(svgId, out var svgPartName))
                    {
                        return false;
                    }
                    if (!packageParts.TryGetValue(svgPartName, out var svgPart))
                    {
                        return false;
                    }
                    var svgBinary = svgPart.SelectSingleNode("pkg:binaryData",
                        namespaces);
                    var svgXml = svgPart.SelectSingleNode("pkg:xmlData", namespaces);
                    if (svgBinary == null && svgXml == null)
                    {
                        return false;
                    }
                    if (!changedParts.Contains(svgPartName))
                    {
                        var svgBytes = svgBinary != null
                            ? Convert.FromBase64String(svgBinary.InnerText)
                            : Encoding.UTF8.GetBytes(svgXml.InnerXml);
                        if (!TryDetachHostColorSvg(svgBytes, update.Source,
                                update.PreviousTextColor, out var normalizedSvg,
                                out var svgChanged))
                        {
                            return false;
                        }
                        if (svgChanged)
                        {
                            if (svgBinary != null)
                                svgBinary.InnerText = Convert.ToBase64String(normalizedSvg);
                            else
                                svgXml.InnerXml = Encoding.UTF8.GetString(normalizedSvg);
                            // Migration replaces the drawing run once. Reassert its
                            // derived baseline in that same transaction; already-unset
                            // formulas never enter this path or touch w:position.
                            NormalizeWordFormulaBaseline(run, namespaces,
                                update.Metadata.DepthPt);
                            requiresWrite = true;
                        }
                        changedParts.Add(svgPartName);
                    }
                    found.Add(metadata.Id);
                }
                if (found.Count != updates.Count)
                {
                    return false;
                }
                normalizedXml = document.OuterXml;
                return true;
            }
            catch (XmlException) { return false; }
            catch (FormatException) { return false; }
            catch (ArgumentException) { return false; }
            catch (System.Runtime.InteropServices.ExternalException) { return false; }
        }

        private static bool TryDetachHostColorSvg(byte[] svgBytes, string source,
            int previousTextColor, out byte[] normalizedBytes, out bool changed)
        {
            normalizedBytes = svgBytes;
            changed = false;
            var svg = Encoding.UTF8.GetString(svgBytes);
            var hostElement = Regex.Match(svg,
                "<[^!?/][^>]*\\bdata-latexblocks-host-color\\s*=\\s*['\"][^'\"]*['\"][^>]*>",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (hostElement.Success)
            {
                var marker = ReadSvgStringAttribute(hostElement.Value,
                    "data-latexblocks-host-color");
                if (string.Equals(marker, "unset", StringComparison.OrdinalIgnoreCase))
                    return true;
                var normalizedElement = SetSvgStringAttribute(hostElement.Value,
                    "data-latexblocks-host-color", "unset");
                normalizedElement = RemoveSvgStringAttribute(normalizedElement, "color");
                normalizedElement = RemoveSvgStringAttribute(normalizedElement, "fill");
                svg = svg.Remove(hostElement.Index, hostElement.Length)
                    .Insert(hostElement.Index, normalizedElement);
                normalizedBytes = Encoding.UTF8.GetBytes(svg);
                changed = true;
                return true;
            }
            // Word may strip unknown data-* attributes while importing SVG. An
            // authored image or explicit TeX colour already owns its local paint;
            // host Font.Color applies only to the remaining unset SVG content and
            // therefore requires no media rewrite.
            if (Regex.IsMatch(svg, "<image\\b", RegexOptions.CultureInvariant |
                    RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1)) ||
                ContainsExplicitColourSource(source))
                return true;
            var oldHex = WordColorToRgbHex(previousTextColor);
            var paintPattern =
                "(?<space>\\s+)(?<name>fill|stroke)\\s*=\\s*(?<q>['\"])(?<color>#[0-9a-fA-F]{3}(?:[0-9a-fA-F]{3})?)\\k<q>";
            var paintMatches = Regex.Matches(svg, paintPattern,
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                TimeSpan.FromSeconds(1));
            var explicitPaints = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (Match match in paintMatches)
                explicitPaints.Add(ExpandRgbHex(match.Groups["color"].Value));

            if (explicitPaints.Count == 0)
                return true;
            if (explicitPaints.Count > 0)
            {
                if (explicitPaints.Count != 1 ||
                    !explicitPaints.Contains(oldHex)) return false;
                svg = Regex.Replace(svg, paintPattern, match =>
                    string.Equals(match.Groups["name"].Value, "fill",
                        StringComparison.OrdinalIgnoreCase)
                        ? string.Empty
                        : match.Groups["space"].Value + "stroke=\"currentColor\"",
                    RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                    TimeSpan.FromSeconds(1));
            }
            var root = Regex.Match(svg, "<svg\\b[^>]*>",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (!root.Success) return false;
            var markedRoot = SetSvgStringAttribute(root.Value,
                "data-latexblocks-host-color", "unset");
            svg = svg.Remove(root.Index, root.Length).Insert(root.Index, markedRoot);
            normalizedBytes = Encoding.UTF8.GetBytes(svg);
            changed = true;
            return true;
        }

        private static void NormalizeWordFormulaBaseline(XmlNode run,
            XmlNamespaceManager namespaces, double depthPt)
        {
            if (run == null) throw new ArgumentNullException(nameof(run));
            var owner = run.OwnerDocument ??
                throw new InvalidDataException("Word formula run has no XML document.");
            var runProperties = run.SelectSingleNode("w:rPr", namespaces);
            if (runProperties == null)
            {
                runProperties = owner.CreateElement("w", "rPr",
                    WordprocessingNamespace);
                run.PrependChild(runProperties);
            }
            var position = runProperties.SelectSingleNode("w:position", namespaces)
                as XmlElement;
            if (position == null)
            {
                position = owner.CreateElement("w", "position",
                    WordprocessingNamespace);
                runProperties.AppendChild(position);
            }
            var baselineHalfPoints = checked((long)(-2 * Math.Round(depthPt,
                MidpointRounding.AwayFromZero)));
            position.SetAttribute("val", WordprocessingNamespace,
                baselineHalfPoints.ToString(CultureInfo.InvariantCulture));
        }

        private static string SetSvgStringAttribute(string element,
            string attribute, string value)
        {
            var pattern = "(\\b" + Regex.Escape(attribute) +
                "\\s*=\\s*['\"])[^'\"]*(['\"])";
            var escaped = SecurityElement.Escape(value ?? string.Empty);
            if (Regex.IsMatch(element, pattern, RegexOptions.CultureInvariant |
                RegexOptions.IgnoreCase))
                return Regex.Replace(element, pattern,
                    "$1" + escaped.Replace("$", "$$") + "$2",
                    RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                    TimeSpan.FromSeconds(1));
            var closing = element.LastIndexOf('>');
            if (closing < 0)
                throw new InvalidDataException("SVG element is malformed.");
            return element.Insert(closing, " " + attribute + "=\"" +
                escaped + "\"");
        }

        private static string ReadSvgStringAttribute(string element,
            string attribute)
        {
            var match = Regex.Match(element, "\\b" + Regex.Escape(attribute) +
                "\\s*=\\s*(['\"])(?<value>[^'\"]*)\\1",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["value"].Value : null;
        }

        private static string RemoveSvgStringAttribute(string element,
            string attribute)
        {
            return Regex.Replace(element, "\\s+" + Regex.Escape(attribute) +
                "\\s*=\\s*(['\"])[^'\"]*\\1", string.Empty,
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                TimeSpan.FromSeconds(1));
        }

        private static string ExpandRgbHex(string value)
        {
            var hex = value.TrimStart('#');
            if (hex.Length == 3)
                hex = new string(new[] { hex[0], hex[0], hex[1], hex[1],
                    hex[2], hex[2] });
            return hex.ToUpperInvariant();
        }

        private static string WordColorToRgbHex(int wordColor)
        {
            var color = WordColorToDrawingColor(wordColor);
            return color.R.ToString("X2", CultureInfo.InvariantCulture) +
                   color.G.ToString("X2", CultureInfo.InvariantCulture) +
                   color.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static System.Drawing.Color WordColorToDrawingColor(int wordColor)
        {
            wordColor = NormalizeTextColor(wordColor);
            if (wordColor == AutomaticTextColor) return System.Drawing.Color.Black;
            return System.Drawing.Color.FromArgb(wordColor & 0xff,
                (wordColor >> 8) & 0xff, (wordColor >> 16) & 0xff);
        }

        private static bool TryReadWordInlineMetadata(string xml,
            out LaTeXBlockMetadata metadata)
        {
            metadata = null;
            var title = Regex.Match(xml,
                "<wp:docPr\\b[^>]*\\btitle=\"(?<title>[^\"]*)\"[^>]*/>",
                RegexOptions.CultureInvariant);
            return title.Success && LaTeXBlockMetadata.TryParse(
                WebUtility.HtmlDecode(title.Groups["title"].Value), out metadata);
        }

        private static bool TryReadWordInlineBatchId(string xml, out Guid id)
        {
            id = Guid.Empty;
            var title = Regex.Match(xml,
                "<wp:docPr\\b[^>]*\\btitle=\"(?<title>[^\"]*)\"[^>]*/>",
                RegexOptions.CultureInvariant);
            if (!title.Success) return false;
            var value = WebUtility.HtmlDecode(title.Groups["title"].Value);
            return value.StartsWith(BatchMetadataTitlePrefix,
                       StringComparison.Ordinal) &&
                   Guid.TryParse(value.Substring(BatchMetadataTitlePrefix.Length),
                       out id);
        }

        private static string NormalizeWordInlineDrawingsXml(string flatOpc,
            IDictionary<Guid, BatchInlineXmlFormat> formats,
            out int normalizedCount, out int removedCount)
        {
            if (flatOpc == null) throw new ArgumentNullException(nameof(flatOpc));
            if (formats == null) throw new ArgumentNullException(nameof(formats));
            var count = 0;
            var removed = 0;
            var normalized = Regex.Replace(flatOpc,
                "<w:r\\b[^>]*>[\\s\\S]*?</w:r>", match =>
                {
                    if (!TryReadWordInlineBatchId(match.Value, out var id) ||
                        !formats.TryGetValue(id, out var format))
                    {
                        if (TryReadWordInlineMetadata(match.Value, out var oldMetadata) &&
                            formats.ContainsKey(oldMetadata.Id))
                        {
                            removed++;
                            return string.Empty;
                        }
                        return match.Value;
                    }
                    var inline = Regex.Match(match.Value,
                        "<wp:inline\\b[\\s\\S]*?</wp:inline>",
                        RegexOptions.CultureInvariant);
                    if (!inline.Success)
                        throw new InvalidDataException(
                            "Word formula run has no inline drawing.");
                    var normalizedInline = NormalizeWordInlineXml(inline.Value,
                        format);
                    var patched = match.Value.Remove(inline.Index, inline.Length)
                        .Insert(inline.Index, normalizedInline);
                    var runProperties = NormalizeWordFormulaRunProperties(
                        format.RunPropertiesXml, format.FontSizePt, format.DepthPt);
                    var currentProperties = Regex.Match(patched,
                        "<w:rPr\\b(?:[^>]*/>|[^>]*>[\\s\\S]*?</w:rPr>)",
                        RegexOptions.CultureInvariant);
                    if (currentProperties.Success)
                        patched = patched.Remove(currentProperties.Index,
                                currentProperties.Length)
                            .Insert(currentProperties.Index, runProperties);
                    else
                    {
                        var openingEnd = patched.IndexOf('>') + 1;
                        if (openingEnd <= 0)
                            throw new InvalidDataException("Word formula run is malformed.");
                        patched = patched.Insert(openingEnd, runProperties);
                    }
                    count++;
                    return patched;
                }, RegexOptions.CultureInvariant,
                TimeSpan.FromSeconds(2));
            normalizedCount = count;
            removedCount = removed;
            return normalized;
        }

        private static string NormalizeWordFormulaRunProperties(string runProperties,
            double fontSizePt, double depthPt)
        {
            var normalized = Regex.Replace(runProperties,
                "^<w:rPr\\b(?<attributes>[^>]*)/>$",
                "<w:rPr${attributes}></w:rPr>",
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            normalized = Regex.Replace(normalized,
                "<w:vertAlign\\b[^>]*/>", string.Empty,
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            var halfPoints = checked((long)Math.Round(fontSizePt * 2,
                MidpointRounding.AwayFromZero));
            var baselineHalfPoints = checked((long)(-2 * Math.Round(depthPt,
                MidpointRounding.AwayFromZero)));
            normalized = SetWordRunProperty(normalized, "sz", halfPoints);
            normalized = SetWordRunProperty(normalized, "szCs", halfPoints);
            return SetWordRunProperty(normalized, "position", baselineHalfPoints);
        }

        private static string SetWordRunProperty(string runProperties,
            string property, long value)
        {
            var pattern = "<w:" + Regex.Escape(property) + "\\b[^>]*/>";
            var replacement = "<w:" + property + " w:val=\"" +
                value.ToString(CultureInfo.InvariantCulture) + "\"/>";
            var existing = Regex.Match(runProperties, pattern,
                RegexOptions.CultureInvariant);
            if (existing.Success)
                return runProperties.Remove(existing.Index, existing.Length)
                    .Insert(existing.Index, replacement);
            var closing = runProperties.LastIndexOf("</w:rPr>",
                StringComparison.Ordinal);
            if (closing < 0)
                throw new InvalidDataException("Word formula run properties are malformed.");
            return runProperties.Insert(closing, replacement);
        }

        private static string NormalizeWordInlineXml(string inline,
            BatchInlineXmlFormat format)
        {
            var size = format.SvgSize;
            var effect = Regex.Match(inline,
                "<wp:effectExtent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!effect.Success)
                throw new InvalidDataException("Word inline SVG has no wp:effectExtent element.");
            var normalizedEffect = SetXmlAttribute(effect.Value, "l", 0);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "t", 0);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "r", 0);
            normalizedEffect = SetXmlAttribute(normalizedEffect, "b", 0);
            var patched = inline.Remove(effect.Index, effect.Length)
                .Insert(effect.Index, normalizedEffect);

            var inlineExtent = Regex.Match(patched,
                "<wp:extent\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!inlineExtent.Success)
                throw new InvalidDataException("Word inline SVG has no wp:extent element.");
            var normalizedInlineExtent = SetXmlAttribute(inlineExtent.Value,
                "cx", size.WidthEmu);
            normalizedInlineExtent = SetXmlAttribute(normalizedInlineExtent,
                "cy", size.HeightEmu);
            patched = patched.Remove(inlineExtent.Index, inlineExtent.Length)
                .Insert(inlineExtent.Index, normalizedInlineExtent);

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
                throw new InvalidDataException(
                    "Word inline SVG has no picture transform extent.");
            var normalizedTransformExtent = SetXmlAttribute(transformExtent.Value,
                "cx", size.WidthEmu);
            normalizedTransformExtent = SetXmlAttribute(normalizedTransformExtent,
                "cy", size.HeightEmu);
            var normalizedTransform = transform.Value.Remove(transformExtent.Index,
                    transformExtent.Length)
                .Insert(transformExtent.Index, normalizedTransformExtent);
            var normalizedPictureProperties = pictureProperties.Value.Remove(
                    transform.Index, transform.Length)
                .Insert(transform.Index, normalizedTransform);
            patched = patched.Remove(pictureProperties.Index, pictureProperties.Length)
                .Insert(pictureProperties.Index, normalizedPictureProperties);

            var documentProperties = Regex.Match(patched,
                "<wp:docPr\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (!documentProperties.Success)
                throw new InvalidDataException("Word inline SVG has no wp:docPr element.");
            var normalizedDocumentProperties = SetXmlStringAttribute(
                documentProperties.Value, "title", format.Metadata.ToString());
            normalizedDocumentProperties = SetXmlStringAttribute(
                normalizedDocumentProperties, "descr", format.Source);
            patched = patched.Remove(documentProperties.Index, documentProperties.Length)
                .Insert(documentProperties.Index, normalizedDocumentProperties);

            var picturePropertiesMetadata = Regex.Match(patched,
                "<pic:cNvPr\\b[^>]*/>", RegexOptions.CultureInvariant);
            if (picturePropertiesMetadata.Success)
            {
                var normalizedPictureMetadata = SetXmlStringAttribute(
                    picturePropertiesMetadata.Value, "title",
                    format.Metadata.ToString());
                normalizedPictureMetadata = SetXmlStringAttribute(
                    normalizedPictureMetadata, "descr", format.Source);
                patched = patched.Remove(picturePropertiesMetadata.Index,
                        picturePropertiesMetadata.Length)
                    .Insert(picturePropertiesMetadata.Index,
                        normalizedPictureMetadata);
            }
            return patched;
        }

        private static string SetXmlStringAttribute(string element,
            string attribute, string value)
        {
            var escaped = SecurityElement.Escape(value ?? string.Empty)
                .Replace("\r", "&#xD;").Replace("\n", "&#xA;")
                .Replace("\t", "&#x9;");
            var pattern = "(\\b" + Regex.Escape(attribute) + "=\")[^\"]*(\")";
            var replacement = "$1" + escaped.Replace("$", "$$") + "$2";
            if (Regex.IsMatch(element, pattern, RegexOptions.CultureInvariant))
                return Regex.Replace(element, pattern, replacement,
                    RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            var closing = element.LastIndexOf("/>", StringComparison.Ordinal);
            if (closing < 0)
                throw new InvalidDataException("Word XML element is malformed.");
            return element.Insert(closing, " " + attribute + "=\"" + escaped + "\"");
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

        internal static bool TryClassifyHostFormatChange(LaTeXBlockLayoutMode mode,
            double selectedSizePt, int selectedTextColor, double currentSizePt,
            int currentTextColor, double renderedSizePt, out bool fontSizeChanged,
            out bool textColorChanged)
        {
            fontSizeChanged = mode == LaTeXBlockLayoutMode.Auto &&
                ShouldRefreshForHostFontSizeChange(selectedSizePt, currentSizePt,
                    renderedSizePt);
            textColorChanged = !TextColorsEqual(selectedTextColor, currentTextColor);
            return fontSizeChanged || textColorChanged;
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

        private sealed class BatchInlineXmlFormat
        {
            internal BatchInlineXmlFormat(SvgPhysicalSize svgSize,
                double fontSizePt, double depthPt, string runPropertiesXml,
                LaTeXBlockMetadata metadata, string source)
            {
                SvgSize = svgSize;
                FontSizePt = fontSizePt;
                DepthPt = depthPt;
                RunPropertiesXml = runPropertiesXml;
                Metadata = metadata;
                Source = source;
            }

            internal SvgPhysicalSize SvgSize { get; }
            internal double FontSizePt { get; }
            internal double DepthPt { get; }
            internal string RunPropertiesXml { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal string Source { get; }
        }

        private sealed class PreparedColorBatchGroup
        {
            internal PreparedColorBatchGroup(WordInterop.Range envelope, int start,
                int end, int expectedInlineShapeCount,
                bool hadParagraphMarkAfter, string recoloredXml)
            {
                Envelope = envelope;
                Start = start;
                End = end;
                ExpectedInlineShapeCount = expectedInlineShapeCount;
                HadParagraphMarkAfter = hadParagraphMarkAfter;
                RecoloredXml = recoloredXml;
            }

            internal WordInterop.Range Envelope { get; }
            internal int Start { get; }
            internal int End { get; }
            internal int ExpectedInlineShapeCount { get; }
            internal bool HadParagraphMarkAfter { get; }
            internal string RecoloredXml { get; }
        }

        private sealed class BatchInlineUpdateState
        {
            internal BatchInlineUpdateState(LaTeXBlockBatchUpdate update, int start,
                WordInterop.WdStoryType storyType, int paragraphStart,
                int paragraphEnd, LaTeXBlockMetadata metadata,
                SvgPhysicalSize svgSize, string insertionPath)
            {
                Update = update;
                Start = start;
                StoryType = storyType;
                ParagraphStart = paragraphStart;
                ParagraphEnd = paragraphEnd;
                Metadata = metadata;
                SvgSize = svgSize;
                InsertionPath = insertionPath;
            }

            internal LaTeXBlockBatchUpdate Update { get; }
            internal int Start { get; }
            internal WordInterop.WdStoryType StoryType { get; }
            internal int ParagraphStart { get; }
            internal int ParagraphEnd { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal SvgPhysicalSize SvgSize { get; }
            internal string InsertionPath { get; }
            internal WordInterop.InlineShape ImportedShape { get; set; }
        }

        private void EnsureDocument()
        {
            if (application.Documents.Count == 0)
                throw new InvalidOperationException("Open a Word document before inserting a LaTeX Block.");
        }
    }

    internal sealed class LaTeXBlockBatchUpdate
    {
        internal LaTeXBlockBatchUpdate(WordInterop.InlineShape shape, string source,
            double widthPt, LaTeXBlockRender render, LaTeXBlockMetadata metadata,
            WordInterop.Range range, int paragraphStart, int paragraphEnd)
        {
            Shape = shape ?? throw new ArgumentNullException(nameof(shape));
            Source = LaTeXBlockService.NormalizeSourceText(source) ??
                throw new ArgumentNullException(nameof(source));
            WidthPt = widthPt;
            Render = render ?? throw new ArgumentNullException(nameof(render));
            Metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
            Range = range ?? throw new ArgumentNullException(nameof(range));
            ParagraphStart = paragraphStart;
            ParagraphEnd = paragraphEnd;
        }

        internal WordInterop.InlineShape Shape { get; }
        internal string Source { get; }
        internal double WidthPt { get; }
        internal LaTeXBlockRender Render { get; }
        internal LaTeXBlockMetadata Metadata { get; }
        internal WordInterop.Range Range { get; }
        internal int ParagraphStart { get; }
        internal int ParagraphEnd { get; }
    }

    internal sealed class LaTeXBlockColorUpdate
    {
        internal LaTeXBlockColorUpdate(WordInterop.InlineShape shape,
            WordInterop.Range range, LaTeXBlockMetadata metadata, string source,
            int previousTextColor, int targetTextColor, int paragraphStart,
            int paragraphEnd)
        {
            Shape = shape ?? throw new ArgumentNullException(nameof(shape));
            Range = range ?? throw new ArgumentNullException(nameof(range));
            Metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
            Source = LaTeXBlockService.NormalizeSourceText(source) ?? string.Empty;
            PreviousTextColor = LaTeXBlockService.NormalizeTextColor(previousTextColor);
            TargetTextColor = LaTeXBlockService.NormalizeTextColor(targetTextColor);
            ParagraphStart = paragraphStart;
            ParagraphEnd = paragraphEnd;
        }

        internal WordInterop.InlineShape Shape { get; }
        internal WordInterop.Range Range { get; }
        internal LaTeXBlockMetadata Metadata { get; }
        internal string Source { get; }
        internal int PreviousTextColor { get; }
        internal int TargetTextColor { get; }
        internal int ParagraphStart { get; }
        internal int ParagraphEnd { get; }
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
