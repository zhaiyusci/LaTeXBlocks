using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    internal enum LaTeXBlockLayoutMode { Fixed, Auto }
    internal enum LaTeXBlockRole { Content, NumberedEquation }
    internal enum LaTeXBlockKind
    {
        Unspecified,
        InlineMath,
        DisplayMath,
        NumberedMath,
        LaTeXBlock
    }

    internal sealed class LaTeXBlockMetadata
    {
        internal const int ContractVersion = 1;
        internal const string HeaderLine = "% !latexblocks 1";
        internal const string EndLine = "% !end-latexblocks";

        internal LaTeXBlockMetadata(Guid id, double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed, double fontSizePt = 10,
            LaTeXBlockRole role = LaTeXBlockRole.Content, double frameWidthPt = 0,
            double frameHeightPt = 0, string styleData = null,
            LaTeXBlockKind kind = LaTeXBlockKind.Unspecified)
        {
            Id = id;
            WidthPt = widthPt;
            DepthPt = depthPt;
            Mode = mode;
            FontSizePt = fontSizePt;
            Role = role;
            FrameWidthPt = NormalizeOptionalExtent(frameWidthPt);
            FrameHeightPt = NormalizeOptionalExtent(frameHeightPt);
            StyleData = string.IsNullOrWhiteSpace(styleData) ? null : styleData;
            Kind = kind;
        }

        internal Guid Id { get; }
        internal double WidthPt { get; }
        internal double DepthPt { get; }
        internal LaTeXBlockLayoutMode Mode { get; }
        internal double FontSizePt { get; }
        internal LaTeXBlockRole Role { get; }
        internal double FrameWidthPt { get; }
        internal double FrameHeightPt { get; }
        internal string StyleData { get; }
        internal LaTeXBlockKind Kind { get; }
        internal bool HasExplicitStyle => !string.IsNullOrWhiteSpace(StyleData);
        internal LaTeXBlockStyle Style => LaTeXBlockStyle.ReadFromMetadataValue(StyleData);

        internal static LaTeXBlockMetadata Create(double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed, double fontSizePt = 10,
            LaTeXBlockRole role = LaTeXBlockRole.Content, LaTeXBlockStyle style = null,
            LaTeXBlockKind kind = LaTeXBlockKind.Unspecified)
        {
            if (mode != LaTeXBlockLayoutMode.Fixed || role != LaTeXBlockRole.Content)
                style = null;
            return new LaTeXBlockMetadata(Guid.NewGuid(), widthPt, depthPt, mode, fontSizePt,
                role, 0, 0, style?.ToMetadataValue(), NormalizeKind(kind, mode, role));
        }

        internal LaTeXBlockMetadata WithFrameSize(double frameWidthPt, double frameHeightPt)
        {
            return new LaTeXBlockMetadata(Id, WidthPt, DepthPt, Mode, FontSizePt, Role,
                frameWidthPt, frameHeightPt, StyleData, Kind);
        }

        internal LaTeXBlockMetadata WithStyle(LaTeXBlockStyle style)
        {
            if (Mode != LaTeXBlockLayoutMode.Fixed || Role != LaTeXBlockRole.Content)
                style = null;
            return new LaTeXBlockMetadata(Id, WidthPt, DepthPt, Mode, FontSizePt, Role,
                FrameWidthPt, FrameHeightPt, style?.ToMetadataValue(), Kind);
        }

        internal LaTeXBlockMetadata WithKind(LaTeXBlockKind kind)
        {
            return new LaTeXBlockMetadata(Id, WidthPt, DepthPt, Mode, FontSizePt, Role,
                FrameWidthPt, FrameHeightPt, StyleData, kind);
        }

        internal LaTeXBlockMetadata WithObservedState(double widthPt, double frameWidthPt,
            double frameHeightPt, double depthPt = 0)
        {
            return new LaTeXBlockMetadata(Id,
                Mode == LaTeXBlockLayoutMode.Fixed ? WidthPt : widthPt, depthPt, Mode,
                FontSizePt, Role, frameWidthPt, frameHeightPt, StyleData, Kind);
        }

        internal string Serialize(string source)
        {
            var kind = NormalizeKind(Kind, Mode, Role);
            if (kind == LaTeXBlockKind.Unspecified)
                throw new InvalidOperationException("A persisted LaTeX Blocks object needs a kind.");
            var builder = new StringBuilder();
            builder.AppendLine(HeaderLine);
            Append(builder, "kind", FormatKind(kind));
            Append(builder, "id", Id.ToString("D"));
            Append(builder, "mode", Mode == LaTeXBlockLayoutMode.Auto ? "auto" : "fixed");
            if (Mode == LaTeXBlockLayoutMode.Fixed)
                Append(builder, "width-pt", FormatNumber(WidthPt));
            Append(builder, "font-size-pt", FormatNumber(FontSizePt));
            if (kind == LaTeXBlockKind.LaTeXBlock && HasExplicitStyle)
                AppendStyle(builder, Style);
            builder.AppendLine(EndLine);
            builder.Append(source ?? string.Empty);
            return builder.ToString();
        }

        internal static bool TryParse(string value, out LaTeXBlockMetadata metadata,
            out string source)
        {
            metadata = null;
            source = null;
            if (value == null || !value.StartsWith(HeaderLine, StringComparison.Ordinal))
                return false;
            var firstEnd = FindLineEnd(value, 0, out var next);
            if (!string.Equals(value.Substring(0, firstEnd), HeaderLine,
                    StringComparison.Ordinal)) return false;
            var fields = new Dictionary<string, string>(StringComparer.Ordinal);
            var position = next;
            while (position <= value.Length)
            {
                var lineEnd = FindLineEnd(value, position, out next);
                var line = value.Substring(position, lineEnd - position);
                if (line == EndLine)
                {
                    source = next <= value.Length ? value.Substring(next) : string.Empty;
                    break;
                }
                if (!line.StartsWith("% ", StringComparison.Ordinal)) return false;
                var separator = line.IndexOf(": ", 2, StringComparison.Ordinal);
                if (separator <= 2) return false;
                var key = line.Substring(2, separator - 2);
                if (!IsKey(key) || fields.ContainsKey(key)) return false;
                fields.Add(key, line.Substring(separator + 2));
                if (next <= position) return false;
                position = next;
            }
            if (source == null || !TryRequired(fields, "kind", out var kindText) ||
                !TryRequired(fields, "id", out var idText) ||
                !TryRequired(fields, "mode", out var modeText) ||
                !TryRequired(fields, "font-size-pt", out var fontText) ||
                !Guid.TryParseExact(idText, "D", out var id) ||
                !TryParseMode(modeText, out var mode) ||
                !TryNumber(fontText, 1, 200, out var fontSizePt)) return false;
            var kind = ParseKind(kindText);
            if (kind == LaTeXBlockKind.Unspecified) return false;
            var widthPt = 0.0;
            if (mode == LaTeXBlockLayoutMode.Fixed &&
                (!TryRequired(fields, "width-pt", out var widthText) ||
                 !TryNumber(widthText, double.Epsilon, double.MaxValue, out widthPt)))
                return false;
            if (mode == LaTeXBlockLayoutMode.Auto && fields.ContainsKey("width-pt"))
                return false;
            var role = kind == LaTeXBlockKind.NumberedMath
                ? LaTeXBlockRole.NumberedEquation : LaTeXBlockRole.Content;
            if (!TryReadStyle(fields, kind, out var style, out var hasStyle)) return false;
            metadata = new LaTeXBlockMetadata(id, widthPt, 0, mode, fontSizePt, role,
                0, 0, hasStyle ? style.ToMetadataValue() : null, kind);
            return true;
        }

        internal static bool TryParse(string value, out LaTeXBlockMetadata metadata)
        {
            return TryParse(value, out metadata, out _);
        }

        internal static string ReadSource(string value)
        {
            return TryParse(value, out _, out var source) ? source : null;
        }

        public override string ToString() => Serialize(string.Empty);

        private static void AppendStyle(StringBuilder builder, LaTeXBlockStyle style)
        {
            Append(builder, "line-spacing", FormatNumber(style.LineSpacing));
            Append(builder, "padding-pt", FormatNumber(style.PaddingPt));
            Append(builder, "vertical-alignment", style.VerticalAlignment ==
                LaTeXBlockVerticalAlignment.Middle ? "center" : style.VerticalAlignment ==
                LaTeXBlockVerticalAlignment.Bottom ? "bottom" : "top");
            Append(builder, "text-color", ToHex(style.TextColor));
            if (style.HasBackgroundFill)
                Append(builder, "background-color", ToHex(style.BackgroundColor));
            Append(builder, "border-width-pt", FormatNumber(style.BorderThicknessPt));
            Append(builder, "border-color", ToHex(style.BorderColor));
        }

        private static bool TryReadStyle(IDictionary<string, string> fields,
            LaTeXBlockKind kind, out LaTeXBlockStyle style, out bool hasStyle)
        {
            style = LaTeXBlockStyle.Default;
            var keys = new[] { "line-spacing", "padding-pt", "vertical-alignment",
                "text-color", "background-color", "border-width-pt", "border-color" };
            hasStyle = false;
            foreach (var key in keys) hasStyle |= fields.ContainsKey(key);
            if (!hasStyle) return true;
            if (kind != LaTeXBlockKind.LaTeXBlock) return false;
            var d = LaTeXBlockStyle.Default;
            var leading = d.LineSpacing;
            var padding = d.PaddingPt;
            var vertical = d.VerticalAlignment;
            var text = d.TextColor;
            var hasBackground = fields.TryGetValue("background-color", out var backgroundText);
            var background = d.BackgroundColor;
            var borderWidth = d.BorderThicknessPt;
            var border = d.BorderColor;
            if (fields.TryGetValue("line-spacing", out var leadingText) &&
                !TryNumber(leadingText, LaTeXBlockStyle.MinimumLineSpacing,
                    LaTeXBlockStyle.MaximumLineSpacing, out leading)) return false;
            if (fields.TryGetValue("padding-pt", out var paddingText) &&
                !TryNumber(paddingText, 0, LaTeXBlockStyle.MaximumPaddingPt,
                    out padding)) return false;
            if (fields.TryGetValue("vertical-alignment", out var verticalText))
            {
                if (verticalText == "center") vertical = LaTeXBlockVerticalAlignment.Middle;
                else if (verticalText == "bottom") vertical = LaTeXBlockVerticalAlignment.Bottom;
                else if (verticalText != "top") return false;
            }
            if (fields.TryGetValue("text-color", out var textValue) &&
                !TryColor(textValue, out text)) return false;
            if (hasBackground && !TryColor(backgroundText, out background)) return false;
            if (fields.TryGetValue("border-width-pt", out var borderWidthText) &&
                !TryNumber(borderWidthText, 0,
                    LaTeXBlockStyle.MaximumBorderThicknessPt, out borderWidth)) return false;
            if (fields.TryGetValue("border-color", out var borderText) &&
                !TryColor(borderText, out border)) return false;
            style = new LaTeXBlockStyle(leading, padding, vertical, text,
                hasBackground, background, borderWidth, border);
            return true;
        }

        private static int FindLineEnd(string value, int start, out int next)
        {
            var lf = value.IndexOf('\n', start);
            var cr = value.IndexOf('\r', start);
            var end = lf < 0 ? cr : cr < 0 ? lf : Math.Min(lf, cr);
            if (end < 0) { next = value.Length + 1; return value.Length; }
            next = end + 1;
            if (value[end] == '\r' && next < value.Length && value[next] == '\n') next++;
            return end;
        }

        private static bool IsKey(string key)
        {
            if (key.Length == 0) return false;
            foreach (var c in key)
                if ((c < 'a' || c > 'z') && c != '-') return false;
            return true;
        }

        private static bool TryRequired(IDictionary<string, string> fields,
            string key, out string value) => fields.TryGetValue(key, out value) && value.Length > 0;

        private static bool TryNumber(string text, double minimum, double maximum,
            out double value) => double.TryParse(text, NumberStyles.Float,
                CultureInfo.InvariantCulture, out value) && !double.IsNaN(value) &&
                !double.IsInfinity(value) && value >= minimum && value <= maximum;

        private static bool TryColor(string text, out System.Drawing.Color color)
        {
            color = System.Drawing.Color.Empty;
            if (text == null || text.Length != 7 || text[0] != '#' ||
                !int.TryParse(text.Substring(1), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture, out var rgb)) return false;
            color = System.Drawing.Color.FromArgb((rgb >> 16) & 255,
                (rgb >> 8) & 255, rgb & 255);
            return true;
        }

        private static string ToHex(System.Drawing.Color color) => "#" +
            color.R.ToString("X2", CultureInfo.InvariantCulture) +
            color.G.ToString("X2", CultureInfo.InvariantCulture) +
            color.B.ToString("X2", CultureInfo.InvariantCulture);

        private static string FormatNumber(double value) =>
            value.ToString("0.######", CultureInfo.InvariantCulture);

        private static void Append(StringBuilder builder, string key, string value) =>
            builder.Append("% ").Append(key).Append(": ").AppendLine(value);

        private static double NormalizeOptionalExtent(double value) =>
            value > 0 && !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0;

        private static LaTeXBlockKind NormalizeKind(LaTeXBlockKind kind,
            LaTeXBlockLayoutMode mode, LaTeXBlockRole role)
        {
            if (kind != LaTeXBlockKind.Unspecified) return kind;
            if (role == LaTeXBlockRole.NumberedEquation) return LaTeXBlockKind.NumberedMath;
            if (mode == LaTeXBlockLayoutMode.Fixed) return LaTeXBlockKind.LaTeXBlock;
            return LaTeXBlockKind.Unspecified;
        }

        private static LaTeXBlockKind ParseKind(string value)
        {
            switch (value)
            {
                case "inline-math": return LaTeXBlockKind.InlineMath;
                case "display-math": return LaTeXBlockKind.DisplayMath;
                case "numbered-math": return LaTeXBlockKind.NumberedMath;
                case "latex-block": return LaTeXBlockKind.LaTeXBlock;
                default: return LaTeXBlockKind.Unspecified;
            }
        }

        private static string FormatKind(LaTeXBlockKind kind)
        {
            switch (kind)
            {
                case LaTeXBlockKind.InlineMath: return "inline-math";
                case LaTeXBlockKind.DisplayMath: return "display-math";
                case LaTeXBlockKind.NumberedMath: return "numbered-math";
                case LaTeXBlockKind.LaTeXBlock: return "latex-block";
                default: return "unspecified";
            }
        }

        private static bool TryParseMode(string value, out LaTeXBlockLayoutMode mode)
        {
            mode = LaTeXBlockLayoutMode.Fixed;
            if (value == "fixed") return true;
            if (value != "auto") return false;
            mode = LaTeXBlockLayoutMode.Auto;
            return true;
        }
    }
}
