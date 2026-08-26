using System;
using System.Runtime.InteropServices;
using LaTeXBlocks.OfficeColor;
using Office = Microsoft.Office.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace LaTeXBlocks.PowerPoint
{
    /// <summary>Maps PowerPoint's ColorFormat/FillFormat to the shared model.</summary>
    internal static class PowerPointGraphicColorAdapter
    {
        internal static bool TryReadFill(PowerPointInterop.Shape shape,
            out OfficeGraphicColor color)
        {
            color = OfficeGraphicColor.Automatic;
            if (shape == null) return false;
            PowerPointInterop.FillFormat fill = null;
            PowerPointInterop.ColorFormat foreground = null;
            try
            {
                fill = shape.Fill;
                foreground = fill.ForeColor;
                if (fill.Visible != Office.MsoTriState.msoTrue ||
                    fill.Type != Office.MsoFillType.msoFillSolid) return false;
                if (foreground.Type == Office.MsoColorType.msoColorTypeScheme &&
                    TryFromTheme(foreground.ObjectThemeColor, out var theme))
                {
                    color = OfficeGraphicColor.FromTheme(theme,
                        OfficeColorTransform.FromSignedAdjustment(
                            foreground.TintAndShade));
                }
                else
                    color = OfficeGraphicColor.FromRgb(
                        OfficeRgb.FromColorRef(foreground.RGB));
                return true;
            }
            catch (COMException) { return false; }
            finally
            {
                if (foreground != null) Marshal.ReleaseComObject(foreground);
                if (fill != null) Marshal.ReleaseComObject(fill);
            }
        }

        internal static void ApplyToFill(PowerPointInterop.Shape shape,
            OfficeGraphicColor color)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (TryReadFill(shape, out var current) && current.Equals(color)) return;
            PowerPointInterop.FillFormat fill = null;
            PowerPointInterop.ColorFormat foreground = null;
            try
            {
                fill = shape.Fill;
                if (fill.Visible != Office.MsoTriState.msoTrue)
                    fill.Visible = Office.MsoTriState.msoTrue;
                if (fill.Type != Office.MsoFillType.msoFillSolid) fill.Solid();
                foreground = fill.ForeColor;
                if (color.Kind == OfficeGraphicColorKind.Theme)
                {
                    foreground.ObjectThemeColor = ToTheme(color.Theme);
                    Marshal.ReleaseComObject(foreground);
                    foreground = fill.ForeColor;
                    // PowerPoint shape fills own TintAndShade. Brightness is a
                    // different picture/OLE projection and is not Word's SVG path.
                    foreground.TintAndShade =
                        color.Transform.ToSignedAdjustment();
                }
                else
                    foreground.RGB = color.Kind == OfficeGraphicColorKind.Rgb
                        ? OfficeRgb.ToColorRef(color.Rgb) : 0;
            }
            finally
            {
                if (foreground != null) Marshal.ReleaseComObject(foreground);
                if (fill != null) Marshal.ReleaseComObject(fill);
            }
        }

        private static Office.MsoThemeColorIndex ToTheme(OfficeThemeColor theme) =>
            (Office.MsoThemeColorIndex)((int)theme + 1);

        private static bool TryFromTheme(Office.MsoThemeColorIndex value,
            out OfficeThemeColor theme)
        {
            var raw = (int)value;
            if (raw >= 1 && raw <= 16)
            {
                theme = (OfficeThemeColor)(raw - 1);
                return true;
            }
            theme = default(OfficeThemeColor);
            return false;
        }
    }
}
