using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    internal sealed class LaTeXBlockService
    {
        private const double StemTeXBaseFontSizePt = 10.0;
        // Word writes wp:effectExtent b="9525" for an inline SVG. That bottom
        // layout extent is 0.75 pt and participates in the host baseline anchor.
        private const double WordInlineSvgEffectBottomPt = 9525.0 / 12700.0;
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

        internal LaTeXBlockRender RenderPreview(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile)
        {
            return RenderPreviewAsync(source, widthPt, mode, profile).GetAwaiter().GetResult();
        }

        internal async Task<LaTeXBlockRender> RenderPreviewAsync(string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile)
        {
            var result = await renderers.RenderLatestAsync(profile, source, widthPt, mode == LaTeXBlockLayoutMode.Auto);
            return new LaTeXBlockRender(WriteSvg(result.Bytes), result.Bytes, result.DepthPt);
        }

        internal WordInterop.InlineShape InsertBlock(string source, double widthPt, LaTeXBlockLayoutMode mode, string profile)
        {
            var render = RenderPreview(source, widthPt, mode, profile);
            return InsertRendered(source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape InsertRendered(string source, double widthPt, LaTeXBlockLayoutMode mode,
            LaTeXBlockRender render)
        {
            EnsureDocument();
            if (render == null) throw new ArgumentNullException(nameof(render));
            var target = application.Selection.Range.Duplicate;
            var inlineScale = GetInlineScale(target, mode);
            var metadata = LaTeXBlockMetadata.Create(widthPt, render.DepthPt * inlineScale, mode);
            target.Text = string.Empty;
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var insertionPath = PrepareInsertionSvg(render, inlineScale, mode);
            var shape = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
            ScaleInlineShape(shape, inlineScale);
            ApplyContract(shape, source, metadata);
            shape.Range.Select();
            return shape;
        }

        internal WordInterop.InlineShape UpdateBlock(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile)
        {
            var render = RenderPreview(source, widthPt, mode, profile);
            return UpdateRendered(oldShape, source, widthPt, mode, render);
        }

        internal WordInterop.InlineShape UpdateRendered(WordInterop.InlineShape oldShape, string source, double widthPt,
            LaTeXBlockLayoutMode mode, LaTeXBlockRender render)
        {
            if (oldShape == null) throw new ArgumentNullException(nameof(oldShape));
            if (render == null) throw new ArgumentNullException(nameof(render));
            if (!LaTeXBlockMetadata.TryParse(oldShape.Title, out var previous))
                throw new InvalidOperationException("The selected image is not a LaTeX Block.");

            var target = oldShape.Range.Duplicate;
            var inlineScale = GetInlineScale(target, mode);
            var metadata = new LaTeXBlockMetadata(previous.Id, widthPt, render.DepthPt * inlineScale, mode);
            target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            WordInterop.InlineShape replacement = null;
            try
            {
                var insertionPath = PrepareInsertionSvg(render, inlineScale, mode);
                replacement = target.InlineShapes.AddPicture(insertionPath, LinkToFile: false, SaveWithDocument: true, Range: target);
                ScaleInlineShape(replacement, inlineScale);
                ApplyContract(replacement, source, metadata);
                oldShape.Delete();
                replacement.Range.Select();
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
                if (LaTeXBlockMetadata.TryParse(candidate.Title, out metadata))
                {
                    shape = candidate;
                    return true;
                }
            }
            return false;
        }

        private static void ApplyContract(WordInterop.InlineShape shape, string source, LaTeXBlockMetadata metadata)
        {
            shape.AlternativeText = source;
            shape.Title = metadata.ToString();
            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            // Word aligns the bottom of an InlineShape to the text baseline. Move the image
            // character down by the TeX box depth. This is always the TeX/Western baseline:
            // CJK glyph extents inside the SVG do not define a second alignment reference.
            // Word persists this API value as whole points.
            var hostDepth = metadata.DepthPt +
                (metadata.Mode == LaTeXBlockLayoutMode.Auto ? WordInlineSvgEffectBottomPt : 0.0);
            shape.Range.Font.Position = -(int)Math.Round(hostDepth, MidpointRounding.AwayFromZero);
        }

        private static double GetInlineScale(WordInterop.Range target, LaTeXBlockLayoutMode mode)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return 1.0;
            var fontSize = (double)target.Font.Size;
            if (!(fontSize > 0) || fontSize > 1000) return 1.0;
            return fontSize / StemTeXBaseFontSizePt;
        }

        private static void ScaleInlineShape(WordInterop.InlineShape shape, double scale)
        {
            if (Math.Abs(scale - 1.0) < 0.0001) return;
            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            shape.Width = shape.Width * (float)scale;
        }

        private string PrepareInsertionSvg(LaTeXBlockRender render, double scale, LaTeXBlockLayoutMode mode)
        {
            if (mode != LaTeXBlockLayoutMode.Auto) return render.SvgPath;
            return WriteSvg(ApplyFractionalBaselineCompensation(render.SvgBytes, render.DepthPt, scale));
        }

        internal static byte[] ApplyFractionalBaselineCompensation(byte[] svgBytes, double depthPt, double scale)
        {
            if (svgBytes == null) throw new ArgumentNullException(nameof(svgBytes));
            if (!(scale > 0)) throw new ArgumentOutOfRangeException(nameof(scale));

            var scaledDepth = depthPt * scale;
            var hostDepth = scaledDepth + WordInlineSvgEffectBottomPt;
            var wordDepth = Math.Round(hostDepth, MidpointRounding.AwayFromZero);
            // Word supplies the whole-point component. Shift the SVG viewport so the
            // remaining fraction, including Word's SVG effect extent, moves the TeX
            // baseline by exactly the residual amount.
            var residualInSvgUnits = (hostDepth - wordDepth) / scale;
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
        internal LaTeXBlockRender(string svgPath, byte[] svgBytes, double depthPt)
        {
            SvgPath = svgPath;
            SvgBytes = svgBytes ?? throw new ArgumentNullException(nameof(svgBytes));
            DepthPt = depthPt;
        }
        internal string SvgPath { get; }
        internal byte[] SvgBytes { get; }
        internal double DepthPt { get; }
    }
}
