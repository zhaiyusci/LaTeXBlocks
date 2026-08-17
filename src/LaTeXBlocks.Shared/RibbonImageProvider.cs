using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace LaTeXBlocks.Branding
{
    internal static class RibbonImageProvider
    {
        private const string ResourcePrefix = "LaTeXBlocks.Branding.Ribbon.";
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, Bitmap> Bitmaps =
            new Dictionary<string, Bitmap>(StringComparer.Ordinal);
        private static readonly Dictionary<string, object> Pictures =
            new Dictionary<string, object>(StringComparer.Ordinal);

        internal static object GetImage(string controlId)
        {
            var resourceName = GetResourceName(controlId);
            if (resourceName == null)
                throw new ArgumentException("No branded Ribbon image is registered for " + controlId,
                    nameof(controlId));

            lock (Sync)
            {
                if (Pictures.TryGetValue(resourceName, out var picture))
                    return picture;

                using (var stream = typeof(RibbonImageProvider).Assembly
                    .GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                        throw new InvalidOperationException(
                            "Missing embedded Ribbon image: " + resourceName);
                    using (var decoded = new Bitmap(stream))
                        Bitmaps[resourceName] = new Bitmap(decoded);
                }

                picture = PictureDispHost.Convert(Bitmaps[resourceName]);
                Pictures[resourceName] = picture;
                return picture;
            }
        }

        internal static string GetResourceName(string controlId)
        {
            switch (controlId)
            {
                case "LaTeXBlocks.InsertFormula":
                    return ResourcePrefix + "InlineMath.png";
                case "LaTeXBlocks.InsertDisplayMath":
                    return ResourcePrefix + "DisplayMath.png";
                case "LaTeXBlocks.InsertNumberedEquation":
                    return ResourcePrefix + "NumberedMath.png";
                case "LaTeXBlocks.InsertBlock":
                case "LaTeXBlocks.PowerPoint.Insert":
                    return ResourcePrefix + "LaTeXBlock.png";
                case "LaTeXBlocks.InsertEquationReference":
                    return ResourcePrefix + "EquationReference.png";
                default:
                    return null;
            }
        }

        private sealed class PictureDispHost : AxHost
        {
            private PictureDispHost() : base(string.Empty) { }

            internal static object Convert(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
