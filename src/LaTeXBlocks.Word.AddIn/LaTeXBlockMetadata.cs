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

    internal sealed class LaTeXBlockMetadata
    {
        internal const string Prefix = "LaTeXBlocks/1;";

        internal LaTeXBlockMetadata(Guid id, double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed, double fontSizePt = 10,
            LaTeXBlockRole role = LaTeXBlockRole.Content)
        {
            Id = id;
            WidthPt = widthPt;
            DepthPt = depthPt;
            Mode = mode;
            FontSizePt = fontSizePt;
            Role = role;
        }

        internal Guid Id { get; }
        internal double WidthPt { get; }
        internal double DepthPt { get; }
        internal LaTeXBlockLayoutMode Mode { get; }
        internal double FontSizePt { get; }
        internal LaTeXBlockRole Role { get; }

        internal static LaTeXBlockMetadata Create(double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed, double fontSizePt = 10,
            LaTeXBlockRole role = LaTeXBlockRole.Content)
        {
            return new LaTeXBlockMetadata(Guid.NewGuid(), widthPt, depthPt, mode, fontSizePt, role);
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
                else if (key == "mode" && string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) mode = LaTeXBlockLayoutMode.Auto;
                else if (key == "role" && string.Equals(value, "numbered-equation", StringComparison.OrdinalIgnoreCase))
                    role = LaTeXBlockRole.NumberedEquation;
            }
            if (id == Guid.Empty || width <= 0) return false;
            metadata = new LaTeXBlockMetadata(id, width, Math.Max(0, depth), mode,
                fontSize >= 1 && fontSize <= 200 ? fontSize : 10, role);
            return true;
        }

        public override string ToString()
        {
            return Prefix + "id=" + Id.ToString("D") + ";width=" +
                   WidthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";depth=" +
                   DepthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";mode=" +
                   (Mode == LaTeXBlockLayoutMode.Auto ? "auto" : "fixed") + ";size=" +
                   FontSizePt.ToString("0.###", CultureInfo.InvariantCulture) + ";role=" +
                   (Role == LaTeXBlockRole.NumberedEquation ? "numbered-equation" : "content");
        }
    }
}
