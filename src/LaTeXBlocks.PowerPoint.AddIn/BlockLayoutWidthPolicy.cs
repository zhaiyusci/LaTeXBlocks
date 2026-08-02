using System;

namespace LaTeXBlocks.PowerPoint
{
    // This is deliberately the same front-end policy as StemTeX GUI. The renderer
    // itself accepts any positive fractional width; Office chooses this compact,
    // editable range for a block's persisted typesetting width.
    internal static class BlockLayoutWidthPolicy
    {
        internal const double MinimumPt = 30.0;
        internal const double MaximumPt = 450.0;
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
