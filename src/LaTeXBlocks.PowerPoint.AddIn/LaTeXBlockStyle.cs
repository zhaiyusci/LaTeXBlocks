using System;
using System.Drawing;
using System.Globalization;
using System.Text;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    // The layout-facing part of a block style belongs to TeX: exact inner size,
    // paragraph indentation, leading, and vertical placement. The final SVG owns
    // padding, background, border, and clipping; Office owns default foreground paint.
    internal enum LaTeXBlockVerticalAlignment
    {
        Top,
        Middle,
        Bottom
    }

    internal sealed class LaTeXBlockStyle : IEquatable<LaTeXBlockStyle>
    {
        internal const string TagName = "LATEXBLOCKS_TEX_STYLE";
        private const string Prefix = "LaTeXBlocksStyle/1;";
        internal const double DefaultLineSpacing = 1.2;
        internal const double MinimumLineSpacing = 0.5;
        internal const double MaximumLineSpacing = 4.0;
        internal const double MaximumPaddingPt = 144.0;
        internal const double MaximumBorderThicknessPt = 24.0;
        // PowerPoint's point is 1/72 in while TeX's pt is 1/72.27 in. Outer-frame
        // geometry stays in PowerPoint/SVG points; only the inner typesetting width
        // passed to TeX needs this conversion.
        private const double TeXPointsPerOfficePoint = 72.27 / 72.0;

        internal LaTeXBlockStyle(double lineSpacing = DefaultLineSpacing,
            double paddingPt = 0, LaTeXBlockVerticalAlignment verticalAlignment =
                LaTeXBlockVerticalAlignment.Top, Color? textColor = null,
            bool hasBackgroundFill = false, Color? backgroundColor = null,
            double borderThicknessPt = 0, Color? borderColor = null)
        {
            LineSpacing = Clamp(lineSpacing, MinimumLineSpacing, MaximumLineSpacing,
                DefaultLineSpacing);
            PaddingPt = Clamp(paddingPt, 0, MaximumPaddingPt, 0);
            VerticalAlignment = verticalAlignment;
            TextColor = NormalizeColor(textColor ?? Color.Black, Color.Black);
            HasBackgroundFill = hasBackgroundFill;
            BackgroundColor = NormalizeColor(backgroundColor ?? Color.White, Color.White);
            BorderThicknessPt = Clamp(borderThicknessPt, 0, MaximumBorderThicknessPt, 0);
            BorderColor = NormalizeColor(borderColor ?? Color.Black, Color.Black);
        }

        internal static LaTeXBlockStyle Default { get; } = new LaTeXBlockStyle();

        internal double LineSpacing { get; }
        internal double PaddingPt { get; }
        internal LaTeXBlockVerticalAlignment VerticalAlignment { get; }
        internal Color TextColor { get; }
        internal bool HasBackgroundFill { get; }
        internal Color BackgroundColor { get; }
        internal double BorderThicknessPt { get; }
        internal Color BorderColor { get; }
        internal bool HasBorder => BorderThicknessPt > 0.0001;

        // The border stroke is drawn inside the SVG viewport. Its full thickness,
        // together with the requested padding, separates the content SVG from the
        // outer edge just as \fboxsep + \fboxrule would have done in TeX.
        // Hosts can use this to retain their legacy bare-snippet route for blocks
        // written before the style editor existed. An explicitly accepted style
        // may still apply the default typography (see WrapSource).
        internal bool IsDefault =>
            NearlyEqual(LineSpacing, DefaultLineSpacing) &&
            NearlyEqual(PaddingPt, 0) &&
            VerticalAlignment == LaTeXBlockVerticalAlignment.Top &&
            TextColor.ToArgb() == Color.Black.ToArgb() &&
            !HasBackgroundFill &&
            !HasBorder;

        internal static LaTeXBlockStyle ReadFromTag(string serialized)
        {
            if (!TryParse(serialized, out var style)) return Default;
            return style;
        }

        internal static bool TryParse(string serialized, out LaTeXBlockStyle style)
        {
            style = null;
            if (string.IsNullOrWhiteSpace(serialized)) return false;
            if (!serialized.StartsWith(Prefix, StringComparison.Ordinal)) return false;

            var lineSpacing = DefaultLineSpacing;
            var paddingPt = 0.0;
            var verticalAlignment = LaTeXBlockVerticalAlignment.Top;
            var textColor = Color.Black;
            var hasBackgroundFill = false;
            var backgroundColor = Color.White;
            var borderThicknessPt = 0.0;
            var borderColor = Color.Black;

            foreach (var part in serialized.Substring(Prefix.Length).Split(';'))
            {
                var separator = part.IndexOf('=');
                if (separator <= 0) continue;
                var key = part.Substring(0, separator);
                var value = part.Substring(separator + 1);
                if (string.Equals(key, "leading", StringComparison.OrdinalIgnoreCase))
                {
                    if (double.TryParse(value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var parsedLineSpacing))
                        lineSpacing = parsedLineSpacing;
                }
                else if (string.Equals(key, "padding", StringComparison.OrdinalIgnoreCase))
                {
                    if (double.TryParse(value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var parsedPaddingPt))
                        paddingPt = parsedPaddingPt;
                }
                else if (string.Equals(key, "valign", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(value, "middle", StringComparison.OrdinalIgnoreCase))
                        verticalAlignment = LaTeXBlockVerticalAlignment.Middle;
                    else if (string.Equals(value, "bottom", StringComparison.OrdinalIgnoreCase))
                        verticalAlignment = LaTeXBlockVerticalAlignment.Bottom;
                }
                else if (string.Equals(key, "text", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryParseHexColor(value, out var color)) textColor = color;
                }
                else if (string.Equals(key, "fill", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        hasBackgroundFill = false;
                    }
                    else if (TryParseHexColor(value, out var color))
                    {
                        hasBackgroundFill = true;
                        backgroundColor = color;
                    }
                }
                else if (string.Equals(key, "border", StringComparison.OrdinalIgnoreCase))
                {
                    if (double.TryParse(value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var parsedBorderThicknessPt))
                        borderThicknessPt = parsedBorderThicknessPt;
                }
                else if (string.Equals(key, "bordercolor", StringComparison.OrdinalIgnoreCase) &&
                         TryParseHexColor(value, out var color))
                {
                    borderColor = color;
                }
            }

            style = new LaTeXBlockStyle(lineSpacing, paddingPt, verticalAlignment, textColor,
                hasBackgroundFill, backgroundColor, borderThicknessPt,
                borderColor);
            return true;
        }

        // Keep this compact representation as the in-memory bridge used by
        // LaTeXBlockMetadata.StyleData. The persisted magic-header contract
        // in both hosts and exposes the style as named fields.
        internal string ToMetadataValue()
        {
            return "1," + FormatDecimal(LineSpacing) + "," +
                FormatDecimal(PaddingPt) + "," + MetadataVerticalAlignment(VerticalAlignment) +
                "," + ToHex(TextColor) + "," +
                (HasBackgroundFill ? ToHex(BackgroundColor) : "-") + "," +
                FormatDecimal(BorderThicknessPt) + "," + ToHex(BorderColor);
        }

        internal static LaTeXBlockStyle ReadFromMetadataValue(string value)
        {
            if (TryParseMetadataValue(value, out var style)) return style;
            return Default;
        }

        internal static bool TryParseMetadataValue(string value, out LaTeXBlockStyle style)
        {
            style = null;
            if (string.IsNullOrWhiteSpace(value)) return false;
            var parts = value.Split(',');
            if (parts.Length != 8 ||
                !string.Equals(parts[0], "1", StringComparison.Ordinal)) return false;
            if (!double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var lineSpacing) ||
                !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var paddingPt) ||
                !double.TryParse(parts[6], NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var borderThicknessPt) ||
                !TryParseHexColor(parts[4], out var textColor) ||
                !TryParseHexColor(parts[7], out var borderColor))
                return false;

            var verticalAlignment = LaTeXBlockVerticalAlignment.Top;
            if (string.Equals(parts[3], "m", StringComparison.OrdinalIgnoreCase))
                verticalAlignment = LaTeXBlockVerticalAlignment.Middle;
            else if (string.Equals(parts[3], "b", StringComparison.OrdinalIgnoreCase))
                verticalAlignment = LaTeXBlockVerticalAlignment.Bottom;
            else if (!string.Equals(parts[3], "t", StringComparison.OrdinalIgnoreCase))
                return false;

            var hasBackgroundFill = !string.Equals(parts[5], "-", StringComparison.Ordinal);
            var backgroundColor = Color.White;
            if (hasBackgroundFill && !TryParseHexColor(parts[5], out backgroundColor))
                return false;

            style = new LaTeXBlockStyle(lineSpacing, paddingPt, verticalAlignment, textColor,
                hasBackgroundFill, backgroundColor, borderThicknessPt,
                borderColor);
            return true;
        }

        internal string WrapSource(string source, double fontSizePt,
            bool applyDefaultTypography = false, double? contentWidthPt = null,
            double? contentHeightPt = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!(fontSizePt > 0) || double.IsNaN(fontSizePt) || double.IsInfinity(fontSizePt))
                throw new ArgumentOutOfRangeException(nameof(fontSizePt));
            // A host can distinguish a legacy bare block from a block that the
            // user explicitly accepted in the style editor.  In the latter case,
            // 1.20× is a real requested leading, not merely a display value in
            // the editor, so allow the caller to force the typography wrapper
            // even when every visual style field has its default value.
            if (IsDefault && !applyDefaultTypography) return source;

            var tex = new StringBuilder();
            // The host wrapper may already be in horizontal mode. Remove only a
            // pending delimiter space; do not end the paragraph or shift the TeX
            // box vertically. Every setup line remains space-neutral until the
            // exact content box is inserted at the current origin.
            tex.AppendLine("\\ifhmode\\unskip\\fi%");
            tex.AppendLine("\\begingroup%");
            // preview reads this at shipout, after the local group ends. It must
            // therefore be global; RenderAsync restores 1pt for each plain block.
            tex.AppendLine("\\global\\PreviewBorder=0pt%");
            tex.AppendLine("\\renewcommand{\\baselinestretch}{" +
                FormatDecimal(LineSpacing) + "}%");
            tex.AppendLine("\\selectfont%");
            // StemTeX has already selected the requested design size before the
            // request file is read. Set the concrete baseline distance as well as
            // baselinestretch: otherwise the existing \fontsize baseline remains
            // cached and a leading-only style change has no visible effect.
            tex.AppendLine("\\setlength{\\baselineskip}{" + FormatDecimal(
                fontSizePt * LineSpacing) + "pt}%");
            // A styled Fixed Block is a genuine TeX layout box. Office supplies
            // its outer dimensions; the host has already subtracted padding before
            // passing these content dimensions. Keep paragraph and vertical
            // placement semantics here, leaving SVG composition to paint only the
            // shell and its inside border around the resulting box.
            var hasFixedWidth = contentWidthPt.HasValue && contentWidthPt.Value > 0;
            var hasFixedHeight = contentHeightPt.HasValue && contentHeightPt.Value > 0;
            if (hasFixedWidth)
            {
                tex.AppendLine("\\setlength{\\hsize}{" +
                    FormatDecimal(contentWidthPt.Value) + "pt}%");
                tex.AppendLine("\\setlength{\\linewidth}{\\hsize}%");
            }
            tex.AppendLine("\\setlength{\\parindent}{0pt}%");
            tex.AppendLine("\\setlength{\\leftskip}{0pt}%");
            tex.AppendLine("\\setlength{\\rightskip}{0pt}%");
            tex.AppendLine("\\setlength{\\parfillskip}{0pt plus 1fil}%");
            tex.AppendLine("\\hangindent=0pt\\hangafter=1\\parshape=0%");
            tex.AppendLine("\\everypar{}%");
            tex.AppendLine("\\setbox0=\\vbox{%");
            // Fixed LaTeX Blocks own their paint in TeX. This wrapper is not used
            // by the three formula kinds, whose host-colour behaviour is unchanged.
            tex.AppendLine("\\color[HTML]{" + ToHex(TextColor) + "}%");
            // A standalone display already owns a vertical TeX list. Do not insert
            // *anything* which opens a paragraph before or after it: \noindent,
            // \color and \par can all change a tight preview's display geometry.
            var standaloneDisplay =
                StemTeXRenderer.StartsWithFullDisplayOrPageWidthEnvironment(source);
            if (standaloneDisplay)
            {
                // Explicit colours authored in TeX remain local overrides.
                tex.Append(source);
                tex.AppendLine("%");
            }
            else
            {
                // TeX derives a paragraph line's height/depth from the glyphs on
                // that line. A lone lowercase run would therefore expose only its
                // x-height at the top of a fixed Block. Define the Block's stable
                // typographic line box from the selected baseline distance and put
                // it on the first and final text lines. Tall content can still grow
                // beyond it naturally. This remains entirely inside TeX; the SVG
                // shell does not infer or add any vertical offset.
                tex.AppendLine("\\setbox\\strutbox=\\hbox{\\vrule height .7\\baselineskip depth .3\\baselineskip width 0pt}%");
                tex.AppendLine("\\noindent\\strut%");
                // Preserve the author source byte-for-byte, including terminal and
                // repeated line endings. The wrapper's percent is outside the
                // source: when the source has no terminal newline it prevents our
                // following newline becoming TeX space; when it does, it simply
                // occupies the next wrapper line without deleting the authored
                // newline or paragraph break.
                tex.Append(source);
                tex.AppendLine("%");
                tex.AppendLine("\\ifhmode\\strut\\fi");
                tex.AppendLine("\\par");
            }
            tex.AppendLine("}%");
            if (hasFixedHeight)
            {
                tex.AppendLine("\\setbox2=\\vbox to " +
                    FormatDecimal(contentHeightPt.Value) + "pt{%");
                switch (VerticalAlignment)
                {
                    case LaTeXBlockVerticalAlignment.Middle:
                        tex.AppendLine("\\vss\\box0\\vss%");
                        break;
                    case LaTeXBlockVerticalAlignment.Bottom:
                        tex.AppendLine("\\vss\\box0%");
                        break;
                    default:
                        tex.AppendLine("\\box0\\vss%");
                        break;
                }
                tex.AppendLine("}%");
            }
            else
            {
                tex.AppendLine("\\setbox2=\\box0%");
            }
            if (hasFixedWidth)
            {
                var contentBox = "\\hbox to " + FormatDecimal(contentWidthPt.Value) +
                    "pt{\\box2\\hss}";
                if (HasBackgroundFill)
                {
                    tex.AppendLine("\\setlength{\\fboxsep}{" +
                        FormatDecimal(ToTeXLengthPt(PaddingPt)) + "pt}%");
                    tex.AppendLine("\\noindent\\colorbox[HTML]{" +
                        ToHex(BackgroundColor) + "}{" + contentBox + "}\\par");
                }
                else
                {
                    tex.AppendLine("\\noindent" + contentBox + "\\par");
                }
            }
            else
                tex.AppendLine("\\noindent\\box2\\par");
            tex.AppendLine("\\endgroup");
            return tex.ToString();
        }

        public override string ToString()
        {
            return Prefix + "leading=" + FormatDecimal(LineSpacing) + ";padding=" +
                   FormatDecimal(PaddingPt) + ";valign=" + VerticalAlignmentToTag(
                       VerticalAlignment) + ";text=" + ToHex(TextColor) + ";fill=" +
                   (HasBackgroundFill ? ToHex(BackgroundColor) : "none") + ";border=" +
                   FormatDecimal(BorderThicknessPt) + ";bordercolor=" + ToHex(BorderColor);
        }

        public bool Equals(LaTeXBlockStyle other)
        {
            if (ReferenceEquals(this, other)) return true;
            if (ReferenceEquals(other, null)) return false;
            return NearlyEqual(LineSpacing, other.LineSpacing) &&
                   NearlyEqual(PaddingPt, other.PaddingPt) &&
                   VerticalAlignment == other.VerticalAlignment &&
                   TextColor.ToArgb() == other.TextColor.ToArgb() &&
                   HasBackgroundFill == other.HasBackgroundFill &&
                   BackgroundColor.ToArgb() == other.BackgroundColor.ToArgb() &&
                   NearlyEqual(BorderThicknessPt, other.BorderThicknessPt) &&
                   BorderColor.ToArgb() == other.BorderColor.ToArgb();
        }

        public override bool Equals(object obj) => Equals(obj as LaTeXBlockStyle);

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = LineSpacing.GetHashCode();
                hash = hash * 31 + PaddingPt.GetHashCode();
                hash = hash * 31 + (int)VerticalAlignment;
                hash = hash * 31 + TextColor.ToArgb();
                hash = hash * 31 + HasBackgroundFill.GetHashCode();
                hash = hash * 31 + BackgroundColor.ToArgb();
                hash = hash * 31 + BorderThicknessPt.GetHashCode();
                return hash * 31 + BorderColor.ToArgb();
            }
        }

        private static string VerticalAlignmentToTag(LaTeXBlockVerticalAlignment alignment)
        {
            switch (alignment)
            {
                case LaTeXBlockVerticalAlignment.Middle: return "middle";
                case LaTeXBlockVerticalAlignment.Bottom: return "bottom";
                default: return "top";
            }
        }

        private static string MetadataVerticalAlignment(LaTeXBlockVerticalAlignment alignment)
        {
            switch (alignment)
            {
                case LaTeXBlockVerticalAlignment.Middle: return "m";
                case LaTeXBlockVerticalAlignment.Bottom: return "b";
                default: return "t";
            }
        }

        private static string FormatDecimal(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static bool NearlyEqual(double left, double right)
        {
            return Math.Abs(left - right) < 0.0005;
        }

        private static double Clamp(double value, double minimum, double maximum,
            double fallback)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return fallback;
            return Math.Max(minimum, Math.Min(maximum, value));
        }

        internal static double ToTeXLengthPt(double officePt)
        {
            return officePt * TeXPointsPerOfficePoint;
        }

        private static Color NormalizeColor(Color color, Color fallback)
        {
            return color.IsEmpty ? fallback : Color.FromArgb(color.R, color.G, color.B);
        }

        private static string ToHex(Color color)
        {
            return color.R.ToString("X2", CultureInfo.InvariantCulture) +
                   color.G.ToString("X2", CultureInfo.InvariantCulture) +
                   color.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static bool TryParseHexColor(string text, out Color color)
        {
            color = Color.Empty;
            if (string.IsNullOrWhiteSpace(text)) return false;
            text = text.Trim();
            if (text.Length == 7 && text[0] == '#') text = text.Substring(1);
            if (text.Length != 6) return false;
            if (!int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                    out var value)) return false;
            color = Color.FromArgb((value >> 16) & 0xff, (value >> 8) & 0xff,
                value & 0xff);
            return true;
        }
    }
}
