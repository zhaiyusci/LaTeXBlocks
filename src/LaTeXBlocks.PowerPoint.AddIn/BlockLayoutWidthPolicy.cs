using System;

namespace LaTeXBlocks.PowerPoint
{
    // The renderer and PowerPoint host-frame contract both support widths through
    // 2000 pt. Keep the editor and Ribbon on that same range so a value is never
    // silently clamped merely because it passed through a different UI surface.
    internal static class BlockLayoutWidthPolicy
    {
        internal const double MinimumPt = 30.0;
        internal const double MaximumPt = 2000.0;
        internal const double DefaultPt = 360.0;
        internal const double StepPt = 0.5;

        internal static bool IsValid(double widthPt)
        {
            return !double.IsNaN(widthPt) && !double.IsInfinity(widthPt) &&
                   widthPt >= MinimumPt && widthPt <= MaximumPt;
        }

        internal static double Clamp(double widthPt)
        {
            if (double.IsNaN(widthPt) || double.IsInfinity(widthPt))
                return DefaultPt;
            return Math.Max(MinimumPt, Math.Min(MaximumPt, widthPt));
        }
    }
}
