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
        internal const string VisualScaleTag = "LATEXBLOCKS_VISUAL_SCALE";
        private readonly PowerPointInterop.Application application;
        private readonly StemTeXBackend renderers;
        private readonly string cacheDirectory;

        internal PowerPointBlockService(PowerPointInterop.Application application, StemTeXBackend renderers)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.renderers = renderers ?? throw new ArgumentNullException(nameof(renderers));
            cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LaTeXBlocks", "PowerPointCache");
            Directory.CreateDirectory(cacheDirectory);
        }

        internal string[] Profiles => renderers.Profiles;

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            string profile, double fontSizePt)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, false);
        }

        internal async Task<LaTeXBlockRender> RenderCommittedAsync(string source, double widthPt,
            string profile, double fontSizePt)
        {
            return await RenderAsync(source, widthPt, profile, fontSizePt, true);
        }

        private async Task<LaTeXBlockRender> RenderAsync(string source, double widthPt,
            string profile, double fontSizePt, bool committed)
        {
            source = NormalizeSourceText(source);
            if (string.IsNullOrWhiteSpace(source))
                throw new ArgumentException("LaTeX source cannot be empty.", nameof(source));
            // The 30–450 pt range belongs to the editor/Ribbon controls. Existing
            // documents and direct shape gestures keep the broader host contract.
            if (!(widthPt >= 36) || widthPt > 2000)
                throw new ArgumentOutOfRangeException(nameof(widthPt));
            if (!(fontSizePt >= 1) || fontSizePt > 200)
                throw new ArgumentOutOfRangeException(nameof(fontSizePt));

            // A PowerPoint block is always a real block of the requested width. Font
            // size remains an independent TeX renderer input; the SVG is never scaled
            // to imitate a different TeX design size.
            var result = committed
                ? await renderers.RenderQueuedAsync(profile, source, widthPt, false, fontSizePt)
                : await renderers.RenderLatestAsync(profile, source, widthPt, false, fontSizePt);
            return new LaTeXBlockRender(WriteSvg(result.Bytes), result.Bytes, result.DepthPt, fontSizePt);
        }

        internal PowerPointInterop.Shape InsertRendered(string source, double widthPt,
            LaTeXBlockRender render)
        {
            if (render == null) throw new ArgumentNullException(nameof(render));
            var slide = GetActiveSlide();
            var size = ReadSvgSize(render.SvgBytes);
            ResolveInsertionPoint(slide, size.WidthPt, size.HeightPt, out var left, out var top);
            var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt,
                LaTeXBlockLayoutMode.Fixed, render.FontSizePt, LaTeXBlockRole.Content);

            try { application.StartNewUndoEntry(); } catch { }
            PowerPointInterop.Shape shape = null;
            try
            {
                shape = slide.Shapes.AddPicture(render.SvgPath, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue, left, top, (float)size.WidthPt, (float)size.HeightPt);
                ApplyContract(shape, source, metadata, TemporaryShapeName(metadata.Id),
                    size.WidthPt, size.HeightPt, 1.0);
                shape.Name = StableShapeName(metadata.Id);
                shape.Select(Office.MsoTriState.msoTrue);
                return shape;
            }
            catch
            {
                try { shape?.Delete(); } catch { }
                throw;
            }
        }

        internal PowerPointInterop.Shape UpdateRendered(PowerPointInterop.Shape oldShape,
            string source, double widthPt, LaTeXBlockRender render,
            bool selectReplacement = true)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!TryReadContract(oldShape, out var previous, out _))
                throw new InvalidOperationException("The selected shape is not a LaTeX Block.");

            var slide = GetOwningSlide(oldShape);
            var sourceSize = ReadSvgSize(render.SvgBytes);
            var scale = ReadVisualScale(oldShape, previous);
            var newWidth = (float)(sourceSize.WidthPt * scale);
            var newHeight = (float)(sourceSize.HeightPt * scale);
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
                replacement = slide.Shapes.AddPicture(render.SvgPath, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue, left, top, newWidth, newHeight);
                replacement.Rotation = rotation;
                ApplyContract(replacement, source, metadata, TemporaryShapeName(metadata.Id),
                    sourceSize.WidthPt, sourceSize.HeightPt, scale);

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
                return replacement;
            }
            catch
            {
                try { replacement?.Delete(); } catch { }
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

        internal static PowerPointResizeResult ClassifyResize(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            var visualScale = ReadPositiveTag(shape, VisualScaleTag,
                InferVisualScale(shape, intrinsicWidth, intrinsicHeight));
            if (!(visualScale > 0)) visualScale = 1.0;
            var expectedWidth = intrinsicWidth * visualScale;
            var expectedHeight = intrinsicHeight * visualScale;
            if (!(expectedWidth > 0) || !(expectedHeight > 0))
                return PowerPointResizeResult.None;

            var widthFactor = shape.Width / expectedWidth;
            var heightFactor = shape.Height / expectedHeight;
            const double tolerancePt = 0.01;
            var widthChanged = Math.Abs(shape.Width - expectedWidth) > tolerancePt;
            var heightChanged = Math.Abs(shape.Height - expectedHeight) > tolerancePt;
            if (!widthChanged && !heightChanged) return PowerPointResizeResult.None;

            // This mirrors Scholia's handles: a horizontal-only resize changes the
            // TeX paragraph width; any vertical component is a visual-scale gesture.
            if (!heightChanged)
            {
                var layoutWidth = Math.Max(36, Math.Min(2000,
                    LayoutWidthForVisibleWidth(shape, metadata, shape.Width)));
                return PowerPointResizeResult.Layout(layoutWidth, visualScale);
            }

            var scaleFactor = widthChanged && Math.Abs(widthFactor - 1.0) >
                Math.Abs(heightFactor - 1.0) ? widthFactor : heightFactor;
            var newScale = visualScale * scaleFactor;
            return newScale > 0 && !double.IsNaN(newScale) && !double.IsInfinity(newScale)
                ? PowerPointResizeResult.Scale(newScale)
                : PowerPointResizeResult.None;
        }

        internal static double LayoutWidthForVisibleWidth(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata, double visibleWidthPt)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            if (!(visibleWidthPt > 0)) throw new ArgumentOutOfRangeException(nameof(visibleWidthPt));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            var visualScale = ReadPositiveTag(shape, VisualScaleTag,
                InferVisualScale(shape, intrinsicWidth, intrinsicHeight));
            var expectedVisibleWidth = intrinsicWidth * visualScale;
            if (!(expectedVisibleWidth > 0)) expectedVisibleWidth = metadata.WidthPt;
            return metadata.WidthPt * visibleWidthPt / expectedVisibleWidth;
        }

        internal static void NormalizeVisualScale(PowerPointInterop.Shape shape,
            double visualScale)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!(visualScale > 0) || double.IsNaN(visualScale) ||
                double.IsInfinity(visualScale))
                throw new ArgumentOutOfRangeException(nameof(visualScale));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, shape.Width);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            SetGeometryAroundCurrentCenter(shape, intrinsicWidth * visualScale,
                intrinsicHeight * visualScale);
            shape.Tags.Add(VisualScaleTag,
                visualScale.ToString("R", CultureInfo.InvariantCulture));
        }

        internal static void RestoreStoredGeometry(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            var scale = ReadPositiveTag(shape, VisualScaleTag,
                InferVisualScale(shape, intrinsicWidth, intrinsicHeight));
            if (!(scale > 0)) scale = 1;
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.Width = (float)(intrinsicWidth * scale);
            shape.Height = (float)(intrinsicHeight * scale);
        }

        private static void SetGeometryAroundCurrentCenter(PowerPointInterop.Shape shape,
            double widthPt, double heightPt)
        {
            var centerX = shape.Left + shape.Width / 2.0;
            var centerY = shape.Top + shape.Height / 2.0;
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.Width = (float)widthPt;
            shape.Height = (float)heightPt;
            shape.Left = (float)(centerX - widthPt / 2.0);
            shape.Top = (float)(centerY - heightPt / 2.0);
        }

        internal static string NormalizeSourceText(string source)
        {
            return source?.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        internal static double ReadSvgWidthPt(byte[] svgBytes) => ReadSvgLengthPt(svgBytes, "width");
        internal static double ReadSvgHeightPt(byte[] svgBytes) => ReadSvgLengthPt(svgBytes, "height");

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
            LaTeXBlockMetadata metadata, string name, double svgWidthPt, double svgHeightPt,
            double visualScale)
        {
            // Unlocking the picture exposes PowerPoint's horizontal side handles as
            // layout-width handles. AfterShapeSizeChange restores uniform visual
            // scaling for vertical/corner gestures.
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.AlternativeText = NormalizeSourceText(source);
            shape.Title = metadata.ToString();
            shape.Tags.Add(KindTag, KindValue);
            shape.Tags.Add(SvgWidthTag,
                svgWidthPt.ToString("R", CultureInfo.InvariantCulture));
            shape.Tags.Add(SvgHeightTag,
                svgHeightPt.ToString("R", CultureInfo.InvariantCulture));
            shape.Tags.Add(VisualScaleTag,
                visualScale.ToString("R", CultureInfo.InvariantCulture));
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

        private static double ReadVisualScale(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            var intrinsicWidth = ReadPositiveTag(shape, SvgWidthTag, metadata.WidthPt);
            var intrinsicHeight = ReadPositiveTag(shape, SvgHeightTag, shape.Height);
            return ReadPositiveTag(shape, VisualScaleTag,
                InferVisualScale(shape, intrinsicWidth, intrinsicHeight));
        }

        private static double InferVisualScale(PowerPointInterop.Shape shape,
            double intrinsicWidth, double intrinsicHeight)
        {
            var widthScale = intrinsicWidth > 0 ? shape.Width / intrinsicWidth : 0;
            var heightScale = intrinsicHeight > 0 ? shape.Height / intrinsicHeight : 0;
            if (widthScale > 0 && heightScale > 0)
            {
                var relativeDifference = Math.Abs(widthScale - heightScale) /
                    Math.Max(widthScale, heightScale);
                return relativeDifference < 0.01 ? (widthScale + heightScale) / 2.0 : heightScale;
            }
            return widthScale > 0 ? widthScale : heightScale > 0 ? heightScale : 1.0;
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

    internal enum PowerPointResizeKind { None, LayoutWidth, VisualScale }

    internal sealed class PowerPointResizeResult
    {
        private PowerPointResizeResult(PowerPointResizeKind kind, double layoutWidthPt,
            double visualScale)
        {
            Kind = kind;
            LayoutWidthPt = layoutWidthPt;
            VisualScale = visualScale;
        }

        internal static readonly PowerPointResizeResult None =
            new PowerPointResizeResult(PowerPointResizeKind.None, 0, 0);
        internal static PowerPointResizeResult Layout(double widthPt, double scale) =>
            new PowerPointResizeResult(PowerPointResizeKind.LayoutWidth, widthPt, scale);
        internal static PowerPointResizeResult Scale(double scale) =>
            new PowerPointResizeResult(PowerPointResizeKind.VisualScale, 0, scale);

        internal PowerPointResizeKind Kind { get; }
        internal double LayoutWidthPt { get; }
        internal double VisualScale { get; }
    }

    internal sealed class LaTeXBlockRender
    {
        internal LaTeXBlockRender(string svgPath, byte[] svgBytes, double depthPt,
            double fontSizePt)
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
