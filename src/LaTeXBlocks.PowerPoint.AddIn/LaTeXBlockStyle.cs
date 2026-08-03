using System;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace LaTeXBlocks.PowerPoint
{
    // The source-facing part of a PowerPoint block style belongs to TeX: leading
    // and foreground colour affect the actual mathematics and text. The outer
    // shell (padding, fill, border and vertical placement) is composed into the
    // final SVG by PowerPointBlockService. Keeping those coordinate systems
    // separate prevents a TeX box from painting past dvisvgm's SVG viewport.
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
        internal double OuterInsetPt => PaddingPt + (HasBorder ? BorderThicknessPt : 0);

        // A default block keeps the historical bare-snippet route byte-for-byte at
        // the TeX level. This avoids silently changing existing slides merely by
        // opening them with a newer add-in.
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
            if (string.IsNullOrWhiteSpace(serialized) ||
                !serialized.StartsWith(Prefix, StringComparison.Ordinal)) return false;

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

            style = new LaTeXBlockStyle(lineSpacing, paddingPt, verticalAlignment,
                textColor, hasBackgroundFill, backgroundColor, borderThicknessPt,
                borderColor);
            return true;
        }

        internal string WrapSource(string source, double fontSizePt)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!(fontSizePt > 0) || double.IsNaN(fontSizePt) || double.IsInfinity(fontSizePt))
                throw new ArgumentOutOfRangeException(nameof(fontSizePt));
            if (IsDefault) return source;

            var tex = new StringBuilder();
            tex.AppendLine("\\begingroup");
            // preview reads this at shipout, after the local group ends. It must
            // therefore be global; RenderAsync restores 1pt for each plain block.
            tex.AppendLine("\\global\\PreviewBorder=0pt");
            tex.AppendLine("\\definecolor{latexblocksforeground}{HTML}{" +
                ToHex(TextColor) + "}");
            tex.AppendLine("\\renewcommand{\\baselinestretch}{" +
                FormatDecimal(LineSpacing) + "}");
            tex.AppendLine("\\selectfont");
            // StemTeX has already selected the requested design size before the
            // request file is read. Set the concrete baseline distance as well as
            // baselinestretch: otherwise the existing \fontsize baseline remains
            // cached and a leading-only style change has no visible effect.
            tex.AppendLine("\\setlength{\\baselineskip}{" + FormatDecimal(
                fontSizePt * LineSpacing) + "pt}");
            // \color may enter horizontal mode. Establish the paragraph with no
            // indent first, so the content SVG uses the requested typesetting
            // width rather than a profile-defined paragraph indent.
            tex.AppendLine("\\noindent");
            tex.AppendLine("\\color{latexblocksforeground}");

            tex.AppendLine(source);
            // Finish the author paragraph while the leading/color scope is still
            // active. Otherwise the worker's outer \par runs only after
            // \endgroup and silently restores the profile's leading.
            tex.AppendLine("\\par");
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
