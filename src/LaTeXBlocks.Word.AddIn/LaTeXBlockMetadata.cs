using System;
using System.Globalization;

namespace LaTeXBlocks.Word
{
    internal enum LaTeXBlockLayoutMode { Fixed, Auto }

    internal sealed class LaTeXBlockMetadata
    {
        internal const string Prefix = "LaTeXBlocks/1;";

        internal LaTeXBlockMetadata(Guid id, double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed)
        {
            Id = id;
            WidthPt = widthPt;
            DepthPt = depthPt;
            Mode = mode;
        }

        internal Guid Id { get; }
        internal double WidthPt { get; }
        internal double DepthPt { get; }
        internal LaTeXBlockLayoutMode Mode { get; }

        internal static LaTeXBlockMetadata Create(double widthPt, double depthPt = 0,
            LaTeXBlockLayoutMode mode = LaTeXBlockLayoutMode.Fixed)
        {
            return new LaTeXBlockMetadata(Guid.NewGuid(), widthPt, depthPt, mode);
        }

        internal static bool TryParse(string title, out LaTeXBlockMetadata metadata)
        {
            metadata = null;
            if (string.IsNullOrEmpty(title) || !title.StartsWith(Prefix, StringComparison.Ordinal)) return false;

            Guid id = Guid.Empty;
            var width = 360.0;
            var depth = 0.0;
            var mode = LaTeXBlockLayoutMode.Fixed;
            foreach (var part in title.Substring(Prefix.Length).Split(';'))
            {
                var separator = part.IndexOf('=');
                if (separator <= 0) continue;
                var key = part.Substring(0, separator);
                var value = part.Substring(separator + 1);
                if (key == "id") Guid.TryParse(value, out id);
                else if (key == "width") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out width);
                else if (key == "depth") double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out depth);
                else if (key == "mode" && string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) mode = LaTeXBlockLayoutMode.Auto;
            }
            if (id == Guid.Empty || width <= 0) return false;
            metadata = new LaTeXBlockMetadata(id, width, Math.Max(0, depth), mode);
            return true;
        }

        public override string ToString()
        {
            return Prefix + "id=" + Id.ToString("D") + ";width=" +
                   WidthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";depth=" +
                   DepthPt.ToString("0.###", CultureInfo.InvariantCulture) + ";mode=" +
                   (Mode == LaTeXBlockLayoutMode.Auto ? "auto" : "fixed");
        }
    }
}
