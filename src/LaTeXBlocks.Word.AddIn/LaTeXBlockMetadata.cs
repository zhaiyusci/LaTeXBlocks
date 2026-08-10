using System;
using System.Globalization;

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
        internal const string Prefix = "LaTeXBlocks/1;";

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
        // Fixed floating Word blocks distinguish the TeX measure from the exact
        // authored SVG frame.  The latter is persisted here because Word Shapes
        // do not offer PowerPoint's per-shape Tags collection.  Zero means a
        // document written before frame persistence was introduced.
        internal double FrameWidthPt { get; }
        internal double FrameHeightPt { get; }
        // Word persists the style in the same Title contract as the block identity
        // because both InlineShape and floating Shape support it.  AlternativeText
        // intentionally remains only the exact author-written TeX source.
        internal string StyleData { get; }
        internal LaTeXBlockKind Kind { get; }
        internal bool HasExplicitStyle => !string.IsNullOrWhiteSpace(StyleData);
        internal LaTeXBlockStyle Style => LaTeXBlockStyle.ReadFromMetadataValue(StyleData);

        internal static LaTeXBlockMetadata Create(double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed, double fontSizePt = 10,
            LaTeXBlockRole role = LaTeXBlockRole.Content, LaTeXBlockStyle style = null,
            LaTeXBlockKind kind = LaTeXBlockKind.Unspecified)
        {
            // Styling is meaningful only for a fixed content viewport.  Do not
            // let a caller accidentally smuggle fixed-Block state into an Auto
            // formula or a Word-owned numbered equation merely because the same
            // editor instance was used to change layout mode.
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

        internal static bool TryParse(string title, out LaTeXBlockMetadata metadata)
        {
            metadata = null;
            if (string.IsNullOrEmpty(title) || !title.StartsWith(Prefix, StringComparison.Ordinal)) return false;

            Guid id = Guid.Empty;
            var width = 360.0;
            var depth = 0.0;
            var mode = LaTeXBlockLayoutMode.Fixed;
            var fontSize = 10.0;
            var role = LaTeXBlockRole.Content;
            var frameWidth = 0.0;
            var frameHeight = 0.0;
            string styleData = null;
            var kind = LaTeXBlockKind.Unspecified;
            foreach (var part in title.Substring(Prefix.Length).Split(';'))
            {
                var separator = part.IndexOf('=');
                if (separator <= 0) continue;
                var key = part.Substring(0, separator);
                var value = part.Substring(separator + 1);
                if (key == "id") Guid.TryParse(value, out id);
                else if (key == "width") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out width);
                else if (key == "depth") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out depth);
                else if (key == "size") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out fontSize);
                else if (key == "framewidth") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out frameWidth);
                else if (key == "frameheight") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out frameHeight);
                else if (key == "style") styleData = value;
                else if (key == "kind") kind = ParseKind(value);
                else if (key == "mode" && string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) mode = LaTeXBlockLayoutMode.Auto;
                else if (key == "role" && string.Equals(value, "numbered-equation", StringComparison.OrdinalIgnoreCase))
                    role = LaTeXBlockRole.NumberedEquation;
            }
            if (id == Guid.Empty || width <= 0) return false;
            metadata = new LaTeXBlockMetadata(id, width, Math.Max(0, depth), mode,
                fontSize >= 1 && fontSize <= 200 ? fontSize : 10, role, frameWidth, frameHeight,
                styleData, kind);
            return true;
        }

        public override string ToString()
        {
            return Prefix + "id=" + Id.ToString("D") + ";width=" +
                   WidthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";depth=" +
                   DepthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";mode=" +
                   (Mode == LaTeXBlockLayoutMode.Auto ? "auto" : "fixed") + ";size=" +
                   FontSizePt.ToString("0.###", CultureInfo.InvariantCulture) + ";role=" +
                   (Role == LaTeXBlockRole.NumberedEquation ? "numbered-equation" : "content") +
                   ";kind=" + FormatKind(Kind) +
                   FormatOptionalExtent("framewidth", FrameWidthPt) +
                   FormatOptionalExtent("frameheight", FrameHeightPt) +
                   (HasExplicitStyle ? ";style=" + StyleData : string.Empty);
        }

        private static string FormatOptionalExtent(string name, double value)
        {
            return value > 0
                ? ";" + name + "=" + value.ToString("0.######", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private static double NormalizeOptionalExtent(double value)
        {
            return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0;
        }

        private static LaTeXBlockKind NormalizeKind(LaTeXBlockKind kind,
            LaTeXBlockLayoutMode mode, LaTeXBlockRole role)
        {
            if (kind != LaTeXBlockKind.Unspecified) return kind;
            if (role == LaTeXBlockRole.NumberedEquation)
                return LaTeXBlockKind.NumberedMath;
            if (mode == LaTeXBlockLayoutMode.Fixed)
                return LaTeXBlockKind.LaTeXBlock;
            // Legacy Auto Content did not persist inline-vs-display identity.
            return LaTeXBlockKind.Unspecified;
        }

        private static LaTeXBlockKind ParseKind(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
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
    }
}
