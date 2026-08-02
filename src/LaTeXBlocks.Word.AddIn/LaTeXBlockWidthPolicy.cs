using System;
using System.Globalization;

namespace LaTeXBlocks.Word
{
    internal static class LaTeXBlockWidthPolicy
    {
        // Matches the visible StemTeX GUI. This is an Office editor policy, not a
        // renderer constraint: the native API still receives the exact fractional
        // point value selected here.
        internal const double MinimumWidthPt = 30.0;
        internal const double MaximumWidthPt = 450.0;
        internal const double DefaultWidthPt = 360.0;
        internal const double WidthStepPt = 0.5;

        internal static double NormalizeTextAreaWidth(double textAreaWidthPt,
            double fallbackPt = 360)
        {
            if (IsFinitePositive(textAreaWidthPt)) return textAreaWidthPt;
            return IsFinitePositive(fallbackPt) ? fallbackPt : DefaultWidthPt;
        }

        internal static double ResolveDefaultFixedWidth()
        {
            return DefaultWidthPt;
        }

        internal static bool IsValidWidth(double widthPt)
        {
            return IsFinitePositive(widthPt) && widthPt >= MinimumWidthPt &&
                   widthPt <= MaximumWidthPt;
        }

        internal static bool TryParseWidth(string text, out double widthPt)
        {
            var styles = NumberStyles.Float;
            if ((!double.TryParse(text, styles, CultureInfo.CurrentCulture, out widthPt) &&
                 !double.TryParse(text, styles, CultureInfo.InvariantCulture, out widthPt)) ||
                !IsValidWidth(widthPt))
            {
                widthPt = 0;
                return false;
            }
            return true;
        }

        internal static double ClampWidth(double widthPt)
        {
            if (!IsFinitePositive(widthPt)) return DefaultWidthPt;
            return Math.Max(MinimumWidthPt, Math.Min(MaximumWidthPt, widthPt));
        }

        private static bool IsFinitePositive(double value)
        {
            return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
