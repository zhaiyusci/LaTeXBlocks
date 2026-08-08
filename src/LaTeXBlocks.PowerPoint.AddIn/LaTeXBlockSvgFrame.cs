using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    // Host-neutral SVG shell composition for styled fixed blocks.  TeX produces
    // the complete content layout box; this class gives it a precise outer
    // viewport without scaling glyphs. Both Office hosts therefore share padding,
    // background, border, and crop semantics exactly; Office Graphics Fill owns
    // the inherited default foreground. Paragraph layout and vertical placement
    // remain in TeX.
    internal static class LaTeXBlockSvgFrame
    {
        internal static byte[] Decorate(byte[] svgBytes, LaTeXBlockStyle style,
            double requestedFrameWidthPt, double? requestedFrameHeightPt)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new ArgumentException("StemTeX returned an empty SVG.", nameof(svgBytes));
            if (style == null) throw new ArgumentNullException(nameof(style));

            var naturalSize = ReadSvgSize(svgBytes);
            // A native Office frame is an explicit author instruction. Keep its
            // physical size exact, even when unchanged TeX content must be clipped.
            var frameWidthPt = ClampFrameExtent(requestedFrameWidthPt);
            var requestedHeight = requestedFrameHeightPt.HasValue &&
                requestedFrameHeightPt.Value > 0 &&
                !double.IsNaN(requestedFrameHeightPt.Value) &&
                !double.IsInfinity(requestedFrameHeightPt.Value)
                ? ClampFrameExtent(requestedFrameHeightPt.Value)
                : 0;
            var frameHeightPt = requestedHeight > 0
                ? requestedHeight
                : naturalSize.HeightPt + 2 * style.PaddingPt;

            // TeX has already made the exact inner layout box and aligned its
            // contents. SVG only places that box at the top-left content origin,
            // paints the shell, and clips any overflow at the authored frame.
            var leftPt = style.PaddingPt;
            var rightPt = frameWidthPt - naturalSize.WidthPt - leftPt;
            var topPt = style.PaddingPt;
            var bottomPt = frameHeightPt - naturalSize.HeightPt - topPt;

            var svg = Encoding.UTF8.GetString(svgBytes);
            var root = Regex.Match(svg, "<svg\\b[^>]*>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!root.Success)
                throw new InvalidDataException("StemTeX SVG has no root svg element.");
            var closingIndex = svg.LastIndexOf("</svg>", StringComparison.OrdinalIgnoreCase);
            if (closingIndex < root.Index + root.Length)
                throw new InvalidDataException("StemTeX SVG has no closing root svg element.");

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

            var xUnitsPerPt = viewBoxWidth / naturalSize.WidthPt;
            var yUnitsPerPt = viewBoxHeight / naturalSize.HeightPt;
            var frameViewBoxX = viewBoxX - leftPt * xUnitsPerPt;
            var frameViewBoxY = viewBoxY - topPt * yUnitsPerPt;
            var frameViewBoxWidth = viewBoxWidth + (leftPt + rightPt) * xUnitsPerPt;
            var frameViewBoxHeight = viewBoxHeight + (topPt + bottomPt) * yUnitsPerPt;
            var newViewBox = FormatNumber(frameViewBoxX) + " " +
                FormatNumber(frameViewBoxY) + " " + FormatNumber(frameViewBoxWidth) +
                " " + FormatNumber(frameViewBoxHeight);
            rootTag = ReplaceAttribute(rootTag, "width", FormatNumber(frameWidthPt) + "pt");
            rootTag = ReplaceAttribute(rootTag, "height", FormatNumber(frameHeightPt) + "pt");
            rootTag = ReplaceAttribute(rootTag, "viewBox", newViewBox);
            rootTag = ReplaceAttribute(rootTag, "overflow", "hidden");
            var frame = new StringBuilder();
            frame.Append("<g data-latexblocks-frame='1'>");
            if (style.HasBackgroundFill)
                AppendRect(frame, frameViewBoxX, frameViewBoxY, frameViewBoxWidth,
                    frameViewBoxHeight, SvgColor(style.BackgroundColor));
            frame.Append("</g>\n");

            var border = new StringBuilder();
            if (style.HasBorder)
            {
                // Four filled strips stay entirely inside the viewport. A stroked
                // rectangle would extend half its line width outside and clip.
                var borderX = style.BorderThicknessPt * xUnitsPerPt;
                var borderY = style.BorderThicknessPt * yUnitsPerPt;
                var rightX = frameViewBoxX + frameViewBoxWidth - borderX;
                var bottomY = frameViewBoxY + frameViewBoxHeight - borderY;
                var color = SvgColor(style.BorderColor);
                border.Append("<g data-latexblocks-border='1'>");
                AppendRect(border, frameViewBoxX, frameViewBoxY, frameViewBoxWidth,
                    borderY, color);
                AppendRect(border, frameViewBoxX, bottomY, frameViewBoxWidth,
                    borderY, color);
                AppendRect(border, frameViewBoxX, frameViewBoxY, borderX,
                    frameViewBoxHeight, color);
                AppendRect(border, rightX, frameViewBoxY, borderX,
                    frameViewBoxHeight, color);
                border.Append("</g>\n");
            }

            var result = new StringBuilder(svg.Length + frame.Length + border.Length + 512);
            result.Append(svg, 0, root.Index);
            result.Append(rootTag);
            result.Append('\n');
            result.Append(frame);
            result.Append(svg, root.Index + root.Length,
                closingIndex - (root.Index + root.Length));
            result.Append(border);
            result.Append(svg, closingIndex, svg.Length - closingIndex);
            return Encoding.UTF8.GetBytes(result.ToString());
        }

        private static void AppendRect(StringBuilder svg, double x, double y,
            double width, double height, string fill)
        {
            if (!(width > 0) || !(height > 0)) return;
            svg.Append("<rect x='").Append(FormatNumber(x)).Append("' y='")
                .Append(FormatNumber(y)).Append("' width='").Append(FormatNumber(width))
                .Append("' height='").Append(FormatNumber(height)).Append("' fill='")
                .Append(fill).Append("'/>");
        }

        private static string SvgColor(Color color)
        {
            return "#" + color.R.ToString("X2", CultureInfo.InvariantCulture) +
                color.G.ToString("X2", CultureInfo.InvariantCulture) +
                color.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("0.#########", CultureInfo.InvariantCulture);
        }

        private static double ClampFrameExtent(double extentPt)
        {
            if (double.IsNaN(extentPt) || double.IsInfinity(extentPt) || !(extentPt > 0))
                return 0.01;
            return extentPt;
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

        private static string ReplaceAttribute(string rootTag, string name, string value)
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

        private static SvgSize ReadSvgSize(byte[] svgBytes)
        {
            return new SvgSize(ReadSvgLengthPt(svgBytes, "width"),
                ReadSvgLengthPt(svgBytes, "height"));
        }

        private static double ReadSvgLengthPt(byte[] svgBytes, string attribute)
        {
            var svg = Encoding.UTF8.GetString(svgBytes);
            var match = Regex.Match(svg,
                "<svg\\b[^>]*\\b" + Regex.Escape(attribute) +
                "=(?<q>['\"])(?<value>[-+0-9.eE]+)\\s*(?<unit>[A-Za-z]*)\\k<q>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success || !TryReadFinitePositive(match.Groups["value"].Value,
                    out var value))
                throw new InvalidDataException("StemTeX SVG has no positive physical " +
                    attribute + ".");
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
}
