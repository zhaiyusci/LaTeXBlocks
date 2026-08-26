using System;
using System.Drawing;

namespace LaTeXBlocks.OfficeColor
{
    internal static class OfficeRgb
    {
        // Office's integer RGB properties use COLORREF byte order (0x00BBGGRR).
        internal static int ToColorRef(Color color) =>
            color.R | color.G << 8 | color.B << 16;

        internal static Color FromColorRef(int value) => Color.FromArgb(255,
            value & 0xff, value >> 8 & 0xff, value >> 16 & 0xff);
    }

    /// <summary>
    /// Host-independent theme slots.  These are semantic Office theme slots, not
    /// the numeric values of Word's or PowerPoint's interop enums.
    /// </summary>
    internal enum OfficeThemeColor
    {
        Dark1,
        Light1,
        Dark2,
        Light2,
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Hyperlink,
        FollowedHyperlink,
        Text1,
        Background1,
        Text2,
        Background2
    }

    internal enum OfficeGraphicColorKind
    {
        Automatic,
        Rgb,
        Theme
    }

    internal enum OfficeColorTransformKind
    {
        None,
        Tint,
        Shade
    }

    /// <summary>
    /// A host-independent theme transform. Amount is always positive: Tint moves
    /// toward white and Shade moves toward black. Adapters decide which COM
    /// property represents it.
    /// </summary>
    internal readonly struct OfficeColorTransform : IEquatable<OfficeColorTransform>
    {
        private OfficeColorTransform(OfficeColorTransformKind kind, float amount)
        {
            if (float.IsNaN(amount) || float.IsInfinity(amount) ||
                amount < 0f || amount > 1f)
                throw new ArgumentOutOfRangeException(nameof(amount));
            if (kind == OfficeColorTransformKind.None && amount != 0f)
                throw new ArgumentException("A neutral transform has no amount.",
                    nameof(amount));
            Kind = kind;
            Amount = amount;
        }

        internal OfficeColorTransformKind Kind { get; }
        internal float Amount { get; }
        internal static OfficeColorTransform None =>
            new OfficeColorTransform(OfficeColorTransformKind.None, 0);
        internal static OfficeColorTransform Tint(float amount) => amount == 0
            ? None : new OfficeColorTransform(OfficeColorTransformKind.Tint, amount);
        internal static OfficeColorTransform Shade(float amount) => amount == 0
            ? None : new OfficeColorTransform(OfficeColorTransformKind.Shade, amount);

        internal static OfficeColorTransform FromSignedAdjustment(float value)
        {
            if (float.IsNaN(value) || float.IsInfinity(value) ||
                value < -1f || value > 1f)
                throw new ArgumentOutOfRangeException(nameof(value));
            return value > 0 ? Tint(value) : value < 0 ? Shade(-value) : None;
        }

        internal float ToSignedAdjustment() =>
            Kind == OfficeColorTransformKind.Tint ? Amount :
            Kind == OfficeColorTransformKind.Shade ? -Amount : 0;

        public bool Equals(OfficeColorTransform other) =>
            Kind == other.Kind && Math.Abs(Amount - other.Amount) < 0.000001f;
        public override bool Equals(object value) =>
            value is OfficeColorTransform other && Equals(other);
        public override int GetHashCode() => ((int)Kind * 397) ^ Amount.GetHashCode();
    }

    /// <summary>
    /// Colour owned by an Office graphic.  It deliberately contains no Word or
    /// PowerPoint interop types, so each host can map its own ColorFormat without
    /// leaking PIA-specific enum values into rendering code.
    /// </summary>
    internal readonly struct OfficeGraphicColor : IEquatable<OfficeGraphicColor>
    {
        private OfficeGraphicColor(OfficeGraphicColorKind kind, Color rgb,
            OfficeThemeColor theme, OfficeColorTransform transform)
        {
            Kind = kind;
            Rgb = Color.FromArgb(255, rgb.R, rgb.G, rgb.B);
            Theme = theme;
            Transform = transform;
        }

        internal OfficeGraphicColorKind Kind { get; }
        internal Color Rgb { get; }
        internal OfficeThemeColor Theme { get; }

        internal OfficeColorTransform Transform { get; }

        internal static OfficeGraphicColor Automatic =>
            new OfficeGraphicColor(OfficeGraphicColorKind.Automatic,
                Color.Empty, default(OfficeThemeColor), OfficeColorTransform.None);

        internal static OfficeGraphicColor FromRgb(Color color) =>
            new OfficeGraphicColor(OfficeGraphicColorKind.Rgb, color,
                default(OfficeThemeColor), OfficeColorTransform.None);

        internal static OfficeGraphicColor FromTheme(OfficeThemeColor theme,
            OfficeColorTransform transform) =>
            new OfficeGraphicColor(OfficeGraphicColorKind.Theme, Color.Empty,
                theme, transform);

        public bool Equals(OfficeGraphicColor other) =>
            Kind == other.Kind && Rgb.ToArgb() == other.Rgb.ToArgb() &&
            Theme == other.Theme && Transform.Equals(other.Transform);

        public override bool Equals(object value) =>
            value is OfficeGraphicColor other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = (int)Kind;
                hash = (hash * 397) ^ Rgb.ToArgb();
                hash = (hash * 397) ^ (int)Theme;
                hash = (hash * 397) ^ Transform.GetHashCode();
                return hash;
            }
        }
    }

    internal enum TeXColorKind
    {
        InheritOffice,
        ExplicitRgb
    }

    /// <summary>
    /// Colour owned by TeX/SVG generation.  InheritOffice is a rendering policy,
    /// not an Office colour value; ExplicitRgb is selected by LaTeXBlocks UI.
    /// </summary>
    internal readonly struct TeXColor : IEquatable<TeXColor>
    {
        private TeXColor(TeXColorKind kind, Color rgb)
        {
            Kind = kind;
            Rgb = Color.FromArgb(255, rgb.R, rgb.G, rgb.B);
        }

        internal TeXColorKind Kind { get; }
        internal Color Rgb { get; }
        internal static TeXColor InheritOffice =>
            new TeXColor(TeXColorKind.InheritOffice, Color.Empty);
        internal static TeXColor Explicit(Color color) =>
            new TeXColor(TeXColorKind.ExplicitRgb, color);

        internal Color Resolve(Color resolvedOfficeColor) =>
            Kind == TeXColorKind.ExplicitRgb ? Rgb : resolvedOfficeColor;

        public bool Equals(TeXColor other) =>
            Kind == other.Kind && Rgb.ToArgb() == other.Rgb.ToArgb();
        public override bool Equals(object value) =>
            value is TeXColor other && Equals(other);
        public override int GetHashCode() => ((int)Kind * 397) ^ Rgb.ToArgb();
    }
}
