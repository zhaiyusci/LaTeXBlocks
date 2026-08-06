using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace LaTeXBlocks.PowerPoint
{
    internal sealed class PowerPointBlockService
    {
        internal const string KindTag = "LATEXBLOCKS_KIND";
        internal const string KindValue = "LATEX_BLOCK";
        internal const string SvgWidthTag = "LATEXBLOCKS_SVG_WIDTH_PT";
        internal const string SvgHeightTag = "LATEXBLOCKS_SVG_HEIGHT_PT";
        // Older releases wrote a serialized default style tag for every Block,
        // even though a default-valued editor state still used the bare legacy
        // SVG route. This separate marker records the newer, literal meaning:
        // an editor-accepted default is real 1.20× TeX leading plus an SVG shell.
        internal const string StyleAppliedTag = "LATEXBLOCKS_TEX_STYLE_APPLIED";
        private readonly PowerPointInterop.Application application;
        private readonly IStemTeXBackend renderers;
        private readonly string cacheDirectory;
        private const int CacheRetentionDays = 7;

        internal PowerPointBlockService(PowerPointInterop.Application application, IStemTeXBackend renderers)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.renderers = renderers ?? throw new ArgumentNullException(nameof(renderers));
            cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LaTeXBlocks", "PowerPointCache");
            Directory.CreateDirectory(cacheDirectory);
            SweepExpiredCacheFiles();
        }

        internal string[] Profiles => renderers.Profiles;

        internal void CancelPreview()
        {
            // Closing an editor must not leave an obsolete latest-only XeTeX request
            // occupying the shared worker. This never cancels committed document work.
            renderers.CancelLatestPreview();
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            string profile, double fontSizePt)
        {
            return await RenderPreviewAsync(source, widthPt, profile, fontSizePt,
                LaTeXBlockStyle.Default, null);
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt)
        {
            return await RenderPreviewAsync(source, widthPt, profile, fontSizePt,
                style, outerHeightPt, null);
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, style,
                outerHeightPt, outerWidthPt, false, false);
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt, bool styleWasExplicit)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, style,
                outerHeightPt, outerWidthPt, false, styleWasExplicit);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            string profile, double fontSizePt)
        {
            return await RenderCommittedAsync(source, widthPt, profile, fontSizePt,
                LaTeXBlockStyle.Default, null);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt)
        {
            return await RenderCommittedAsync(source, widthPt, profile, fontSizePt,
                style, outerHeightPt, null);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, style,
                outerHeightPt, outerWidthPt, true, false);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt, bool styleWasExplicit)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, style,
                outerHeightPt, outerWidthPt, true, styleWasExplicit);
        }

        private async Task<LaTeXBlockRender> RenderAsync(string source, double widthPt,
            string profile, double fontSizePt, LaTeXBlockStyle style,
            double? outerHeightPt, double? outerWidthPt, bool committed,
            bool styleWasExplicit)
        {
            source = NormalizeSourceText(source);
            if (string.IsNullOrWhiteSpace(source))
                throw new ArgumentException("LaTeX source cannot be empty.", nameof(source));
            // The 30–450 pt range belongs to the editor/Ribbon controls. Existing
            // documents and direct shape gestures keep the broader host contract,
            // but native frame fitting must at least be able to use the same lower
            // endpoint that the user can enter in the Ribbon.
            if (!(widthPt >= BlockLayoutWidthPolicy.MinimumPt) || widthPt > 2000)
                throw new ArgumentOutOfRangeException(nameof(widthPt));
            if (!(fontSizePt >= 1) || fontSizePt > 200)
                throw new ArgumentOutOfRangeException(nameof(fontSizePt));

            style = style ?? LaTeXBlockStyle.Default;
            // TeX owns the contents (glyphs, paragraph layout, leading, vertical
            // placement and text colour). The final SVG owns the outer shell. Do not
            // ask TeX \fbox / \colorbox to paint a full PowerPoint frame: TeX's
            // preview coordinates can lie outside dvisvgm's root viewport.
            var styleIsApplied = styleWasExplicit || !style.IsDefault;
            var authoredFrameWidthPt = outerWidthPt ?? widthPt;
            var contentWidthPt = styleIsApplied
                ? Math.Max(0.1, authoredFrameWidthPt - 2 * style.PaddingPt)
                : widthPt;
            var contentHeightPt = styleIsApplied && outerHeightPt.HasValue
                ? Math.Max(0.1, outerHeightPt.Value - 2 * style.PaddingPt)
                : (double?)null;
            var renderSource = !styleIsApplied
                // Styled requests reset PreviewBorder globally at TeX shipout. The
                // next unstyled request must restore the legacy profile border.
                ? "\\global\\PreviewBorder=1pt\n" + source
                : style.WrapSource(source, fontSizePt, true,
                    LaTeXBlockStyle.ToTeXLengthPt(contentWidthPt),
                    contentHeightPt.HasValue
                        ? LaTeXBlockStyle.ToTeXLengthPt(contentHeightPt.Value)
                        : (double?)null);
            var rendererWidthPt = !styleIsApplied ? widthPt :
                LaTeXBlockStyle.ToTeXLengthPt(contentWidthPt);

            // A PowerPoint block is always a real block of the requested width. Font
            // size remains an independent TeX renderer input; the SVG is never scaled
            // to imitate a different TeX design size.
            var result = committed
                ? await renderers.RenderQueuedAsync(profile, renderSource, rendererWidthPt,
                    false, fontSizePt)
                : await renderers.RenderLatestAsync(profile, renderSource, rendererWidthPt,
                    false, fontSizePt);
            var finalSvg = !styleIsApplied
                ? result.Bytes
                : LaTeXBlockSvgFrame.Decorate(result.Bytes, style,
                    authoredFrameWidthPt, outerHeightPt);
            return new LaTeXBlockRender(WriteSvg(finalSvg), finalSvg, result.DepthPt,
                fontSizePt, styleIsApplied);
        }

        internal PowerPointInterop.Shape InsertRendered(string source, double widthPt,
            LaTeXBlockRender render)
        {
            return InsertRendered(source, widthPt, render, LaTeXBlockStyle.Default, false);
        }

        internal PowerPointInterop.Shape InsertRendered(string source, double widthPt,
            LaTeXBlockRender render, LaTeXBlockStyle style, bool styleWasExplicit = false)
        {
            if (render == null) throw new ArgumentNullException(nameof(render));
            style = style ?? LaTeXBlockStyle.Default;
            var styleIsApplied = styleWasExplicit || !style.IsDefault ||
                render.StyleWasApplied;
            if (styleIsApplied && !render.StyleWasApplied)
                throw new InvalidOperationException(
                    "The styled LaTeX Block render does not match the requested style state.");
            var slide = GetActiveSlide();
            // A decorated block already owns its complete shell in its SVG. In
            // particular, background and border must end at the SVG frame edge,
            // rather than being followed by a PowerPoint-created transparent
            // extension.
            var framedSvg = styleIsApplied
                ? render.SvgBytes
                : FrameSvg(render.SvgBytes, ReadSvgWidthPt(render.SvgBytes),
                    ReadSvgHeightPt(render.SvgBytes));
            var size = ReadSvgSize(framedSvg);
            var framedPath = WriteSvg(framedSvg);
            ResolveInsertionPoint(slide, size.WidthPt, size.HeightPt, out var left, out var top);
            var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt,
                LaTeXBlockLayoutMode.Fixed, render.FontSizePt, LaTeXBlockRole.Content);

            try { application.StartNewUndoEntry(); } catch { }
            PowerPointInterop.Shape shape = null;
            try
            {
                shape = slide.Shapes.AddPicture(framedPath, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue, left, top, (float)size.WidthPt, (float)size.HeightPt);
                // A LaTeX Block has an independently editable host frame. Do not
                // inherit PowerPoint's picture-aspect lock, which would make a
                // nominally horizontal resize mutate height as well.
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                ApplyContract(shape, source, metadata, style, styleIsApplied,
                    TemporaryShapeName(metadata.Id), size.WidthPt, size.HeightPt);
                shape.Name = StableShapeName(metadata.Id);
                shape.Select(Office.MsoTriState.msoTrue);
                DeleteCachedSvg(framedPath);
                DeleteCachedSvg(render.SvgPath);
                return shape;
            }
            catch
            {
                try { shape?.Delete(); } catch { }
                DeleteCachedSvg(framedPath);
                throw;
            }
        }

        internal PowerPointInterop.Shape UpdateRendered(PowerPointInterop.Shape oldShape,
            string source, double widthPt, LaTeXBlockRender render,
            bool selectReplacement = true, double? frameHeightPt = null,
            double? frameWidthPt = null, LaTeXBlockStyle style = null,
            bool styleWasExplicit = false)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!TryReadContract(oldShape, out var previous, out _))
                throw new InvalidOperationException("The selected shape is not a LaTeX Block.");
            style = style ?? ReadStyle(oldShape);
            var styleIsApplied = styleWasExplicit || !style.IsDefault ||
                HasExplicitStyleMarker(oldShape) || render.StyleWasApplied;
            if (styleIsApplied && !render.StyleWasApplied)
                throw new InvalidOperationException(
                    "The styled LaTeX Block render does not match the requested style state.");

            var slide = GetOwningSlide(oldShape);
            // A block has one native PowerPoint frame and one genuine StemTeX
            // layout width. Native size gestures have already queued a fresh TeX
            // layout request before they reach this replacement step; this method
            // only embeds that unscaled result in the requested host frame.
            var requestedFrameHeightPt = frameHeightPt ?? ReadFrameHeightPt(oldShape);
            var requestedFrameWidthPt = frameWidthPt ?? ReadFrameWidthPt(oldShape);
            // A non-default render already contains its complete SVG shell:
            // padding, background, and border. Its vertical placement was resolved
            // inside the TeX content box. Never add a
            // second host-side transparent frame; it would make the PowerPoint
            // extent disagree with the authored SVG frame. The requested frame is
            // authoritative: if its viewport is smaller than the TeX result, SVG
            // clips the overflow instead of silently enlarging the PowerPoint box.
            var framedSvg = styleIsApplied
                ? render.SvgBytes
                : FrameSvg(render.SvgBytes, requestedFrameWidthPt,
                    requestedFrameHeightPt);
            var sourceSize = ReadSvgSize(framedSvg);
            var framedPath = WriteSvg(framedSvg);
            var newWidth = (float)sourceSize.WidthPt;
            var newHeight = (float)sourceSize.HeightPt;
            var left = oldShape.Left;
            var top = oldShape.Top;
            var rotation = oldShape.Rotation;
            var oldName = oldShape.Name;
            var oldZ = TryGetZOrder(oldShape);
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt,
                LaTeXBlockLayoutMode.Fixed, render.FontSizePt, LaTeXBlockRole.Content);

            try { application.StartNewUndoEntry(); } catch { }
            PowerPointInterop.Shape replacement = null;
            try
            {
                replacement = slide.Shapes.AddPicture(framedPath, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue, left, top, newWidth, newHeight);
                replacement.LockAspectRatio = Office.MsoTriState.msoFalse;
                replacement.Rotation = rotation;
                ApplyContract(replacement, source, metadata, style, styleIsApplied,
                    TemporaryShapeName(metadata.Id), sourceSize.WidthPt, sourceSize.HeightPt);

                // While the old shape still exists, place the newly inserted topmost
                // shape immediately above it. Deleting the old shape then leaves the
                // replacement at precisely the former z-order position.
                if (oldZ > 0)
                {
                    var desiredBeforeDelete = oldZ + 1;
                    var guard = slide.Shapes.Count + 2;
                    while (replacement.ZOrderPosition > desiredBeforeDelete && guard-- > 0)
                        replacement.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                }

                oldShape.Delete();
                try { RestoreZOrder(replacement, oldZ, slide.Shapes.Count + 2); } catch { }
                try { replacement.Name = oldName; }
                catch { try { replacement.Name = StableShapeName(metadata.Id); } catch { } }
                if (selectReplacement)
                    try { replacement.Select(Office.MsoTriState.msoTrue); } catch { }
                DeleteCachedSvg(framedPath);
                DeleteCachedSvg(render.SvgPath);
                return replacement;
            }
            catch
            {
                try { replacement?.Delete(); } catch { }
                DeleteCachedSvg(framedPath);
                throw;
            }
        }

        internal bool TryGetSelectedBlock(out PowerPointInterop.Shape shape,
            out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            PowerPointInterop.Selection selection;
            try { selection = application.ActiveWindow?.Selection; }
            catch { return false; }
            if (selection == null || selection.Type != PowerPointInterop.PpSelectionType.ppSelectionShapes)
                return false;
            try
            {
                if (selection.ShapeRange.Count != 1) return false;
                var candidate = selection.ShapeRange[1];
                if (!TryReadContract(candidate, out metadata, out _)) return false;
                shape = candidate;
                return true;
            }
            catch (COMException) { return false; }
        }

        internal static bool TryReadContract(PowerPointInterop.Shape shape,
            out LaTeXBlockMetadata metadata, out string source)
        {
            metadata = null;
            source = null;
            if (shape == null) return false;
            try
            {
                if (!string.Equals(shape.Tags[KindTag], KindValue,
                        StringComparison.OrdinalIgnoreCase)) return false;
                if (!LaTeXBlockMetadata.TryParse(shape.Title, out metadata) ||
                    metadata.Role != LaTeXBlockRole.Content ||
                    metadata.Mode != LaTeXBlockLayoutMode.Fixed)
                {
                    metadata = null;
                    return false;
                }
                source = NormalizeSourceText(shape.AlternativeText);
                if (string.IsNullOrWhiteSpace(source))
                {
                    metadata = null;
                    source = null;
                    return false;
                }
                return true;
            }
            catch (COMException)
            {
                metadata = null;
                source = null;
                return false;
            }
        }

        internal static LaTeXBlockStyle ReadStyle(PowerPointInterop.Shape shape)
        {
            if (shape == null) return LaTeXBlockStyle.Default;
            try
            {
                return LaTeXBlockStyle.ReadFromTag(shape.Tags[LaTeXBlockStyle.TagName]);
            }
            catch (COMException)
            {
                return LaTeXBlockStyle.Default;
            }
        }

        internal static bool IsStyleApplied(PowerPointInterop.Shape shape)
        {
            var style = ReadStyle(shape);
            return !style.IsDefault || HasExplicitStyleMarker(shape);
        }

        private static bool HasExplicitStyleMarker(PowerPointInterop.Shape shape)
        {
            if (shape == null) return false;
            try
            {
                return string.Equals(shape.Tags[StyleAppliedTag], "1",
                    StringComparison.Ordinal);
            }
            catch (COMException)
            {
                return false;
            }
        }

        internal static double ResolveFontSize(PowerPointInterop.Application application,
            double fallbackPt = 18)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            try
            {
                var selection = application.ActiveWindow?.Selection;
                if (selection == null) return fallbackPt;
                if (selection.Type == PowerPointInterop.PpSelectionType.ppSelectionText)
                {
                    var size = TryReadTextRange2Size(selection);
                    if (IsFontSize(size)) return size;
                    size = TryReadTextRangeSize(selection);
                    if (IsFontSize(size)) return size;
                    size = TryReadFirstCharacterSize(selection);
                    if (IsFontSize(size)) return size;
                }
                if (selection.Type == PowerPointInterop.PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count == 1)
                {
                    var shape = selection.ShapeRange[1];
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                        shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        var size = Convert.ToDouble(shape.TextFrame.TextRange.Font.Size);
                        if (IsFontSize(size)) return size;
                        if (shape.TextFrame.TextRange.Length > 0)
                        {
                            size = Convert.ToDouble(shape.TextFrame.TextRange.Characters(1, 1).Font.Size);
                            if (IsFontSize(size)) return size;
                        }
                    }
                }
            }
            catch (COMException) { }
            return fallbackPt;
        }

        internal double ResolveInitialWidth(double fallbackPt = 360)
        {
            return BlockLayoutWidthPolicy.Clamp(fallbackPt);
        }

        internal double GetActiveSlideWidthPt()
        {
            if (application.ActivePresentation == null)
                throw new InvalidOperationException("Open a PowerPoint presentation first.");
            var width = Convert.ToDouble(application.ActivePresentation.PageSetup.SlideWidth);
            if (!(width > 0) || double.IsNaN(width) || double.IsInfinity(width))
                throw new InvalidOperationException("PowerPoint reported an invalid slide width.");
            return width;
        }

        internal static PowerPointShapeKey GetShapeKey(PowerPointInterop.Shape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var slide = GetOwningSlide(shape);
            var presentation = slide.Parent as PowerPointInterop.Presentation;
            if (presentation == null)
                throw new InvalidOperationException(
                    "The LaTeX Block is no longer attached to a presentation.");
            var unknown = Marshal.GetIUnknownForObject(presentation);
            try
            {
                return new PowerPointShapeKey(unknown.ToInt64(), slide.SlideID,
                    shape.Id);
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }

        internal static PowerPointFrameUpdate CaptureFrameResize(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            var expectedWidth = intrinsicWidth;
            var expectedHeight = intrinsicHeight;
            if (!(expectedWidth > 0) || !(expectedHeight > 0))
                return PowerPointFrameUpdate.None;

            const double tolerancePt = 0.01;
            var widthChanged = Math.Abs(shape.Width - expectedWidth) > tolerancePt;
            var heightChanged = Math.Abs(shape.Height - expectedHeight) > tolerancePt;
            if (!widthChanged && !heightChanged) return PowerPointFrameUpdate.None;

            // Every native handle means the same thing: the user changed the host
            // frame. The flags retain only which axes the user explicitly changed;
            // they do not select different side/corner/vertical modes. Every true
            // size change enters the async TeX reflow path. It never introduces
            // visual scale.
            return PowerPointFrameUpdate.Create(widthChanged, heightChanged,
                ClampHostFrameWidth(shape.Width),
                ClampHostFrameHeight(shape.Height));
        }

        private static double ClampHostFrameWidth(double widthPt)
        {
            if (double.IsNaN(widthPt) || double.IsInfinity(widthPt)) return 1;
            return Math.Max(1, Math.Min(2000, widthPt));
        }

        private static double ClampHostFrameHeight(double heightPt)
        {
            if (double.IsNaN(heightPt) || double.IsInfinity(heightPt)) return 1;
            return Math.Max(1, Math.Min(2000, heightPt));
        }

        internal static double ReadFrameHeightPt(PowerPointInterop.Shape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var height = Convert.ToDouble(shape.Height);
            return ClampHostFrameHeight(height);
        }

        internal static double ReadFrameWidthPt(PowerPointInterop.Shape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var width = Convert.ToDouble(shape.Width);
            return ClampHostFrameWidth(width);
        }

        internal static void RestoreStoredGeometry(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            var snapshot = ShapeGeometrySnapshot.Capture(shape);
            try
            {
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                shape.Width = (float)intrinsicWidth;
                shape.Height = (float)intrinsicHeight;
            }
            catch (Exception exception)
            {
                var recoveryFailure = RestoreGeometry(shape, snapshot);
                if (recoveryFailure != null)
                    throw new InvalidOperationException(
                        "PowerPoint could not restore the LaTeX Block geometry and could not fully restore its prior state: " +
                        recoveryFailure.Message, exception);
                throw;
            }
        }

        private static Exception RestoreGeometry(PowerPointInterop.Shape shape,
            ShapeGeometrySnapshot snapshot)
        {
            Exception firstFailure = null;
            RestoreStep(() => shape.LockAspectRatio = Office.MsoTriState.msoFalse,
                ref firstFailure);
            RestoreStep(() => shape.Width = snapshot.Width, ref firstFailure);
            RestoreStep(() => shape.Height = snapshot.Height, ref firstFailure);
            RestoreStep(() => shape.Left = snapshot.Left, ref firstFailure);
            RestoreStep(() => shape.Top = snapshot.Top, ref firstFailure);
            RestoreStep(() => shape.LockAspectRatio = snapshot.LockAspectRatio,
                ref firstFailure);
            return firstFailure;
        }

        private static void RestoreStep(Action action, ref Exception firstFailure)
        {
            try { action(); }
            catch (Exception exception)
            {
                if (firstFailure == null) firstFailure = exception;
            }
        }

        private sealed class ShapeGeometrySnapshot
        {
            private ShapeGeometrySnapshot(float left, float top, float width, float height,
                Office.MsoTriState lockAspectRatio)
            {
                Left = left;
                Top = top;
                Width = width;
                Height = height;
                LockAspectRatio = lockAspectRatio;
            }

            internal float Left { get; }
            internal float Top { get; }
            internal float Width { get; }
            internal float Height { get; }
            internal Office.MsoTriState LockAspectRatio { get; }

            internal static ShapeGeometrySnapshot Capture(PowerPointInterop.Shape shape)
            {
                return new ShapeGeometrySnapshot(shape.Left, shape.Top, shape.Width, shape.Height,
                    shape.LockAspectRatio);
            }
        }

        internal static string NormalizeSourceText(string source)
        {
            return source?.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        internal static double ReadSvgWidthPt(byte[] svgBytes) => ReadSvgLengthPt(svgBytes, "width");
        internal static double ReadSvgHeightPt(byte[] svgBytes) => ReadSvgLengthPt(svgBytes, "height");

        // Reframe the root SVG without changing the TeX coordinate scale. A larger
        // target adds transparent viewport space; a smaller target selects a
        // sub-viewport and clips overflow. Placement is top-left aligned. The
        // physical SVG dimensions therefore always equal the user-specified
        // PowerPoint frame.
        internal static byte[] FrameSvg(byte[] svgBytes, double requestedFrameWidthPt,
            double requestedFrameHeightPt)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new ArgumentException("StemTeX returned an empty SVG.", nameof(svgBytes));

            var naturalSize = ReadSvgSize(svgBytes);
            var frameWidthPt = ClampHostFrameWidth(requestedFrameWidthPt);
            var frameHeightPt = ClampHostFrameHeight(requestedFrameHeightPt);
            if (Math.Abs(frameWidthPt - naturalSize.WidthPt) < 0.001 &&
                Math.Abs(frameHeightPt - naturalSize.HeightPt) < 0.001) return svgBytes;

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
                !TryReadFinitePositive(viewBox.Groups["w"].Value, out var viewBoxWidth) ||
                !TryReadFinitePositive(viewBox.Groups["h"].Value, out var viewBoxHeight) ||
                !TryReadFinite(viewBox.Groups["x"].Value, out var viewBoxX) ||
                !TryReadFinite(viewBox.Groups["y"].Value, out var viewBoxY))
                throw new InvalidDataException("StemTeX SVG has no numeric root viewBox.");

            // The root's physical dimensions and viewBox can use different units.
            // Change each axis by the same ratio in physical and viewBox space.
            // This preserves the original TeX coordinate scale exactly; a larger
            // frame adds transparent space and a smaller one crops. The bare SVG
            // path uses the same top-left origin as decorated blocks.
            var frameViewBoxWidth = viewBoxWidth * frameWidthPt / naturalSize.WidthPt;
            var frameViewBoxX = viewBoxX;
            var frameViewBoxHeight = viewBoxHeight * frameHeightPt / naturalSize.HeightPt;
            var frameViewBoxY = viewBoxY;
            var number = CultureInfo.InvariantCulture;
            var newViewBox = frameViewBoxX.ToString("0.######", number) + " " +
                             frameViewBoxY.ToString("0.######", number) + " " +
                             frameViewBoxWidth.ToString("0.######", number) + " " +
                             frameViewBoxHeight.ToString("0.######", number);
            rootTag = ReplaceSvgAttribute(rootTag, "width",
                frameWidthPt.ToString("0.######", number) + "pt");
            rootTag = ReplaceSvgAttribute(rootTag, "height",
                frameHeightPt.ToString("0.######", number) + "pt");
            rootTag = ReplaceSvgAttribute(rootTag, "viewBox", newViewBox);
            rootTag = ReplaceSvgAttribute(rootTag, "overflow", "hidden");
            svg = svg.Substring(0, root.Index) + rootTag +
                  svg.Substring(root.Index + root.Length);
            return Encoding.UTF8.GetBytes(svg);
        }

        private static bool TryReadFinitePositive(string text, out double value)
        {
            return TryReadFinite(text, out value) && value > 0;
        }

        private static bool TryReadFinite(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                       out value) && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string ReplaceSvgAttribute(string rootTag, string name, string value)
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

        private static double TryReadTextRange2Size(PowerPointInterop.Selection selection)
        {
            try { return Convert.ToDouble(selection.TextRange2.Font.Size); }
            catch { return double.NaN; }
        }

        private static double TryReadTextRangeSize(PowerPointInterop.Selection selection)
        {
            try { return Convert.ToDouble(selection.TextRange.Font.Size); }
            catch { return double.NaN; }
        }

        private static double TryReadFirstCharacterSize(PowerPointInterop.Selection selection)
        {
            try
            {
                if (selection.TextRange2.Length > 0)
                    return Convert.ToDouble(selection.TextRange2.get_Characters(1, 1).Font.Size);
            }
            catch { }
            try
            {
                if (selection.TextRange.Length > 0)
                    return Convert.ToDouble(selection.TextRange.Characters(1, 1).Font.Size);
            }
            catch { }
            return double.NaN;
        }

        private static bool IsFontSize(double size)
        {
            return size >= 1 && size <= 200 && !double.IsNaN(size) && !double.IsInfinity(size);
        }

        private PowerPointInterop.Slide GetActiveSlide()
        {
            if (application.Presentations.Count == 0 || application.ActiveWindow == null)
                throw new InvalidOperationException("Open a PowerPoint presentation before inserting a LaTeX Block.");
            try
            {
                var slide = application.ActiveWindow.View.Slide as PowerPointInterop.Slide;
                if (slide != null) return slide;
            }
            catch (COMException) { }
            throw new InvalidOperationException("Select an editable slide before inserting a LaTeX Block.");
        }

        private static PowerPointInterop.Slide GetOwningSlide(PowerPointInterop.Shape shape)
        {
            try
            {
                var slide = shape.Parent as PowerPointInterop.Slide;
                if (slide != null) return slide;
            }
            catch (COMException) { }
            throw new InvalidOperationException(
                "The LaTeX Block is no longer attached to an editable slide.");
        }

        private void ResolveInsertionPoint(PowerPointInterop.Slide slide, double widthPt,
            double heightPt, out float left, out float top)
        {
            try
            {
                var selection = application.ActiveWindow.Selection;
                if ((selection.Type == PowerPointInterop.PpSelectionType.ppSelectionText ||
                     selection.Type == PowerPointInterop.PpSelectionType.ppSelectionShapes) &&
                    selection.ShapeRange.Count == 1)
                {
                    left = selection.ShapeRange[1].Left;
                    top = selection.ShapeRange[1].Top;
                    return;
                }
            }
            catch (COMException) { }

            var page = application.ActivePresentation.PageSetup;
            left = (float)Math.Max(0, (page.SlideWidth - widthPt) / 2.0);
            top = (float)Math.Max(0, (page.SlideHeight - heightPt) / 2.0);
        }

        private static void ApplyContract(PowerPointInterop.Shape shape, string source,
            LaTeXBlockMetadata metadata, LaTeXBlockStyle style, bool styleIsApplied,
            string name, double svgWidthPt, double svgHeightPt)
        {
            // The picture geometry is the persisted host frame. All native handles
            // feed the same frame-update path; committed SVGs are framed to exactly
            // this geometry instead of retaining a visual zoom state.
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.AlternativeText = NormalizeSourceText(source);
            shape.Title = metadata.ToString();
            shape.Tags.Add(KindTag, KindValue);
            shape.Tags.Add(LaTeXBlockStyle.TagName,
                (style ?? LaTeXBlockStyle.Default).ToString());
            if (styleIsApplied)
                shape.Tags.Add(StyleAppliedTag, "1");
            shape.Tags.Add(SvgWidthTag,
                svgWidthPt.ToString("R", CultureInfo.InvariantCulture));
            shape.Tags.Add(SvgHeightTag,
                svgHeightPt.ToString("R", CultureInfo.InvariantCulture));
            shape.Name = name;
        }

        internal static double ReadPositiveTag(PowerPointInterop.Shape shape, string name,
            double fallback)
        {
            try
            {
                if (double.TryParse(shape.Tags[name], NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var value) && value > 0 &&
                    !double.IsNaN(value) && !double.IsInfinity(value)) return value;
            }
            catch (COMException) { }
            return fallback;
        }

        private static int TryGetZOrder(PowerPointInterop.Shape shape)
        {
            try { return shape.ZOrderPosition; }
            catch { return 0; }
        }

        private static void RestoreZOrder(PowerPointInterop.Shape shape, int target, int guard)
        {
            if (target <= 0) return;
            while (shape.ZOrderPosition > target && guard-- > 0)
                shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            while (shape.ZOrderPosition < target && guard-- > 0)
                shape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
        }

        private static string StableShapeName(Guid id) => "LaTeXBlock_" + id.ToString("N");
        private static string TemporaryShapeName(Guid id) => StableShapeName(id) + "_new";

        private string WriteSvg(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                throw new InvalidDataException("StemTeX returned an empty SVG.");
            var path = Path.Combine(cacheDirectory, Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        private void SweepExpiredCacheFiles()
        {
            try
            {
                var cutoff = DateTime.UtcNow.AddDays(-CacheRetentionDays);
                foreach (var path in Directory.GetFiles(cacheDirectory, "*.svg"))
                    if (File.GetLastWriteTimeUtc(path) < cutoff)
                        try { File.Delete(path); } catch { }
            }
            catch { }
        }

        private void DeleteCachedSvg(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                var cacheRoot = Path.GetFullPath(cacheDirectory).TrimEnd(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                var candidate = Path.GetFullPath(path);
                if (!candidate.StartsWith(cacheRoot, StringComparison.OrdinalIgnoreCase)) return;
                File.Delete(candidate);
            }
            catch { }
        }

        private static SvgSize ReadSvgSize(byte[] svgBytes)
        {
            return new SvgSize(ReadSvgWidthPt(svgBytes), ReadSvgHeightPt(svgBytes));
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
            if (!match.Success || !double.TryParse(match.Groups["value"].Value,
                    NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ||
                !(value > 0) || double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidDataException("StemTeX SVG has no positive physical " + attribute + ".");
            switch (match.Groups["unit"].Value.ToLowerInvariant())
            {
                case "pt":
                case "bp": return value;
                case "px":
                case "": return value * 72.0 / 96.0;
                case "in": return value * 72.0;
                case "cm": return value * 72.0 / 2.54;
                case "mm": return value * 72.0 / 25.4;
                case "pc": return value * 12.0;
                default: throw new InvalidDataException("StemTeX SVG " + attribute +
                    " uses an unsupported unit: " + match.Groups["unit"].Value);
            }
        }

        private struct SvgSize
        {
            internal SvgSize(double widthPt, double heightPt)
            {
                WidthPt = widthPt;
                HeightPt = heightPt;
            }
            internal double WidthPt { get; }
            internal double HeightPt { get; }
        }
    }

    internal struct PowerPointShapeKey : IEquatable<PowerPointShapeKey>
    {
        internal PowerPointShapeKey(long presentationIdentity, int slideId, int shapeId)
        {
            PresentationIdentity = presentationIdentity;
            SlideId = slideId;
            ShapeId = shapeId;
        }

        internal long PresentationIdentity { get; }
        internal int SlideId { get; }
        internal int ShapeId { get; }

        public bool Equals(PowerPointShapeKey other)
        {
            return PresentationIdentity == other.PresentationIdentity &&
                   SlideId == other.SlideId && ShapeId == other.ShapeId;
        }

        public override bool Equals(object obj)
        {
            return obj is PowerPointShapeKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = PresentationIdentity.GetHashCode();
                hash = hash * 397 ^ SlideId;
                return hash * 397 ^ ShapeId;
            }
        }
    }

    internal sealed class PowerPointFrameUpdate
    {
        private PowerPointFrameUpdate(bool widthChanged, bool heightChanged,
            double frameWidthPt, double frameHeightPt)
        {
            WidthChanged = widthChanged;
            HeightChanged = heightChanged;
            FrameWidthPt = frameWidthPt;
            FrameHeightPt = frameHeightPt;
        }

        internal static readonly PowerPointFrameUpdate None =
            new PowerPointFrameUpdate(false, false, 0, 0);
        internal static PowerPointFrameUpdate Create(bool widthChanged, bool heightChanged,
            double widthPt, double heightPt) =>
            new PowerPointFrameUpdate(widthChanged, heightChanged, widthPt, heightPt);

        internal bool HasChange => WidthChanged || HeightChanged;
        // These flags only let two sequential Office events be merged into one frame
        // update. They do not give the handles different user-facing semantics.
        internal bool WidthChanged { get; }
        internal bool HeightChanged { get; }
        internal double FrameWidthPt { get; }
        internal double FrameHeightPt { get; }
    }

    internal sealed class LaTeXBlockRender
    {
        internal LaTeXBlockRender(string svgPath, byte[] svgBytes, double depthPt,
            double fontSizePt, bool styleWasApplied = false)
        {
            SvgPath = svgPath;
            SvgBytes = svgBytes ?? throw new ArgumentNullException(nameof(svgBytes));
            DepthPt = depthPt;
            FontSizePt = fontSizePt;
            StyleWasApplied = styleWasApplied;
        }

        internal string SvgPath { get; }
        internal byte[] SvgBytes { get; }
        internal double DepthPt { get; }
        internal double FontSizePt { get; }
        internal bool StyleWasApplied { get; }
    }
}
