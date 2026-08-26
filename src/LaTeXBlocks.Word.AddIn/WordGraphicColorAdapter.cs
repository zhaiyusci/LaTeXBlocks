using System;
using System.Drawing;
using System.Runtime.InteropServices;
using LaTeXBlocks.OfficeColor;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    /// <summary>Maps Word's two distinct ColorFormat contracts to our model.</summary>
    internal static class WordGraphicColorAdapter
    {
        internal static bool TryCaptureFont(WordInterop.Range range,
            out OfficeGraphicColor color)
        {
            color = OfficeGraphicColor.Automatic;
            if (range == null) return false;
            WordInterop.Font font = null;
            WordInterop.ColorFormat format = null;
            try
            {
                font = range.Font;
                var raw = (int)font.Color;
                if (raw == 9999999) return false;
                if (raw == LaTeXBlockService.AutomaticTextColor)
                {
                    color = OfficeGraphicColor.Automatic;
                    return true;
                }
                if (raw >= 0 && raw <= 0x00ffffff)
                {
                    color = OfficeGraphicColor.FromRgb(OfficeRgb.FromColorRef(raw));
                    return true;
                }
                format = font.TextColor;
                if (format.Type != Office.MsoColorType.msoColorTypeScheme ||
                    !TryFromWordTheme(format.ObjectThemeColor, out var theme))
                    return false;
                if (!IsVariant(format.TintAndShade)) return false;
                color = OfficeGraphicColor.FromTheme(theme,
                    OfficeColorTransform.FromSignedAdjustment(
                        format.TintAndShade));
                return true;
            }
            catch (COMException) { return false; }
            finally
            {
                if (format != null) Marshal.ReleaseComObject(format);
                if (font != null) Marshal.ReleaseComObject(font);
            }
        }

        internal static string DescribeFontForDiagnostics(WordInterop.Range range)
        {
            if (range == null) return "range=null";
            WordInterop.Font font = null;
            WordInterop.ColorFormat format = null;
            try
            {
                font = range.Font;
                format = font.TextColor;
                return "fontColor=" + (int)font.Color + " type=" +
                    (int)format.Type + " theme=" +
                    (int)format.ObjectThemeColor + " tintAndShade=" +
                    format.TintAndShade.ToString("R",
                        System.Globalization.CultureInfo.InvariantCulture) +
                    " brightness=" + format.Brightness.ToString("R",
                        System.Globalization.CultureInfo.InvariantCulture) +
                    " rgb=" + format.RGB;
            }
            catch (Exception exception)
            {
                return "unavailable=" + exception.GetBaseException().Message;
            }
            finally
            {
                if (format != null) Marshal.ReleaseComObject(format);
                if (font != null) Marshal.ReleaseComObject(font);
            }
        }

        internal static void ApplyToRange(WordInterop.Range range,
            OfficeGraphicColor color)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));
            WordInterop.Font font = null;
            WordInterop.ColorFormat format = null;
            try
            {
                font = range.Font;
                if (color.Kind == OfficeGraphicColorKind.Automatic)
                    font.Color = (WordInterop.WdColor)LaTeXBlockService.AutomaticTextColor;
                else if (color.Kind == OfficeGraphicColorKind.Rgb)
                    font.Color = (WordInterop.WdColor)OfficeRgb.ToColorRef(color.Rgb);
                else
                {
                    format = font.TextColor;
                    format.ObjectThemeColor = ToWordTheme(color.Theme);
                    format.TintAndShade =
                        color.Transform.ToSignedAdjustment();
                }
            }
            finally
            {
                if (format != null) Marshal.ReleaseComObject(format);
                if (font != null) Marshal.ReleaseComObject(font);
            }
        }

        internal static bool TryReadFill(WordInterop.InlineShape shape,
            out OfficeGraphicColor color)
        {
            color = OfficeGraphicColor.Automatic;
            if (shape == null) return false;
            WordInterop.FillFormat fill = null;
            WordInterop.ColorFormat foreground = null;
            try
            {
                fill = shape.Fill;
                foreground = fill.ForeColor;
                if (fill.Visible != Office.MsoTriState.msoTrue ||
                    fill.Type != Office.MsoFillType.msoFillSolid) return false;
                if (foreground.Type == Office.MsoColorType.msoColorTypeScheme &&
                    TryFromGraphicsTheme(foreground.ObjectThemeColor, out var theme))
                {
                    if (!IsVariant(foreground.Brightness)) return false;
                    color = OfficeGraphicColor.FromTheme(theme,
                        OfficeColorTransform.FromSignedAdjustment(
                            foreground.Brightness));
                    return true;
                }
                color = OfficeGraphicColor.FromRgb(OfficeRgb.FromColorRef(foreground.RGB));
                return true;
            }
            catch (COMException) { return false; }
            finally
            {
                if (foreground != null) Marshal.ReleaseComObject(foreground);
                if (fill != null) Marshal.ReleaseComObject(fill);
            }
        }

        internal static void ApplyToFill(WordInterop.InlineShape shape,
            OfficeGraphicColor color)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (color.Kind != OfficeGraphicColorKind.Theme &&
                TryReadFill(shape, out var current) && current.Equals(color)) return;
            WordInterop.FillFormat fill = null;
            WordInterop.ColorFormat foreground = null;
            try
            {
                OfficeGraphicColor before;
                var hasBefore = TryReadFill(shape, out before);
                ColorSyncDiagnostics.Write("fill-adapter", "before=" +
                    (hasBefore ? Describe(before) : "unavailable") +
                    " target=" + Describe(color));
                fill = shape.Fill;
                if (fill.Visible != Office.MsoTriState.msoTrue)
                    fill.Visible = Office.MsoTriState.msoTrue;
                if (fill.Type != Office.MsoFillType.msoFillSolid) fill.Solid();
                foreground = fill.ForeColor;
                if (color.Kind == OfficeGraphicColorKind.Theme)
                {
                    foreground.ObjectThemeColor = ToGraphicsTheme(color.Theme);
                    Marshal.ReleaseComObject(foreground);
                    foreground = fill.ForeColor;
                    // Word's SVG/Office Graphic fill persists the canonical theme
                    // transform through Brightness. Font.TextColor uses
                    // TintAndShade; writing both properties mixes two distinct
                    // owner-specific projections of the theme transform.
                    foreground.Brightness =
                        color.Transform.ToSignedAdjustment();
                }
                else
                {
                    foreground.RGB = color.Kind == OfficeGraphicColorKind.Rgb
                        ? OfficeRgb.ToColorRef(color.Rgb) : 0;
                }
                OfficeGraphicColor after;
                var hasAfter = TryReadFill(shape, out after);
                ColorSyncDiagnostics.Write("fill-adapter", "after=" +
                    (hasAfter ? Describe(after) : "unavailable"));
            }
            finally
            {
                if (foreground != null) Marshal.ReleaseComObject(foreground);
                if (fill != null) Marshal.ReleaseComObject(fill);
            }
        }

        private static string Describe(OfficeGraphicColor color)
        {
            return "kind=" + color.Kind + " theme=" + color.Theme +
                " transform=" + color.Transform.Kind + ":" +
                color.Transform.Amount.ToString("R",
                    System.Globalization.CultureInfo.InvariantCulture) +
                " rgb=" + color.Rgb.R + "," + color.Rgb.G + "," +
                color.Rgb.B;
        }

        private static bool IsVariant(float value) => !float.IsNaN(value) &&
            !float.IsInfinity(value) && value >= -1f && value <= 1f;

        internal static WordInterop.WdThemeColorIndex ToWordTheme(OfficeThemeColor theme)
        {
            switch (theme)
            {
                case OfficeThemeColor.Background1: return WordInterop.WdThemeColorIndex.wdThemeColorBackground1;
                case OfficeThemeColor.Text1: return WordInterop.WdThemeColorIndex.wdThemeColorText1;
                case OfficeThemeColor.Background2: return WordInterop.WdThemeColorIndex.wdThemeColorBackground2;
                case OfficeThemeColor.Text2: return WordInterop.WdThemeColorIndex.wdThemeColorText2;
                default: return (WordInterop.WdThemeColorIndex)(int)theme;
            }
        }

        internal static WordInterop.WdThemeColorIndex ToGraphicsTheme(OfficeThemeColor theme)
        {
            switch (theme)
            {
                case OfficeThemeColor.Text1: return (WordInterop.WdThemeColorIndex)13;
                case OfficeThemeColor.Background1: return (WordInterop.WdThemeColorIndex)14;
                case OfficeThemeColor.Text2: return (WordInterop.WdThemeColorIndex)15;
                case OfficeThemeColor.Background2: return (WordInterop.WdThemeColorIndex)16;
                default: return (WordInterop.WdThemeColorIndex)((int)theme + 1);
            }
        }

        internal static bool TryFromWordTheme(WordInterop.WdThemeColorIndex value,
            out OfficeThemeColor theme)
        {
            var raw = (int)value;
            if (raw >= 0 && raw <= 11) { theme = (OfficeThemeColor)raw; return true; }
            switch (value)
            {
                case WordInterop.WdThemeColorIndex.wdThemeColorText1: theme = OfficeThemeColor.Text1; return true;
                case WordInterop.WdThemeColorIndex.wdThemeColorBackground1: theme = OfficeThemeColor.Background1; return true;
                case WordInterop.WdThemeColorIndex.wdThemeColorText2: theme = OfficeThemeColor.Text2; return true;
                case WordInterop.WdThemeColorIndex.wdThemeColorBackground2: theme = OfficeThemeColor.Background2; return true;
                default: theme = default(OfficeThemeColor); return false;
            }
        }

        private static bool TryFromGraphicsTheme(WordInterop.WdThemeColorIndex value,
            out OfficeThemeColor theme)
        {
            var raw = (int)value;
            if (raw >= 1 && raw <= 12) { theme = (OfficeThemeColor)(raw - 1); return true; }
            switch (raw)
            {
                case 13: theme = OfficeThemeColor.Text1; return true;
                case 14: theme = OfficeThemeColor.Background1; return true;
                case 15: theme = OfficeThemeColor.Text2; return true;
                case 16: theme = OfficeThemeColor.Background2; return true;
                default: theme = default(OfficeThemeColor); return false;
            }
        }
    }
}
