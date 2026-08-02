using System;
using System.Globalization;

namespace LaTeXBlocks.PowerPoint
{
    internal static class PowerPointRibbonContract
    {
        internal const string FontSizeControlId = "LaTeXBlocks.PowerPoint.FontSize";
        internal const string LayoutWidthControlId = "LaTeXBlocks.PowerPoint.LayoutWidth";

        internal static string BuildCustomUi()
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">" +
                   "<ribbon><tabs><tab id=\"LaTeXBlocks.PowerPoint.Tab\" label=\"LaTeX Blocks\">" +
                   "<group id=\"LaTeXBlocks.PowerPoint.Blocks\" label=\"LaTeX Blocks\">" +
                   "<button id=\"LaTeXBlocks.PowerPoint.Insert\" label=\"Insert Block\" size=\"large\" imageMso=\"TextBoxInsert\" onAction=\"OnInsertBlock\"/>" +
                   "<button id=\"LaTeXBlocks.PowerPoint.Edit\" label=\"Edit Block\" size=\"large\" imageMso=\"ObjectEdit\" onAction=\"OnEditBlock\"/>" +
                   "<editBox id=\"" + LayoutWidthControlId + "\" label=\"Typesetting width (pt)\" sizeString=\"000.0\" getText=\"GetLayoutWidthText\" getEnabled=\"GetLayoutWidthEnabled\" onChange=\"OnLayoutWidthChanged\"/>" +
                   "<editBox id=\"" + FontSizeControlId + "\" label=\"TeX size (pt)\" sizeString=\"000.0\" getText=\"GetFontSizeText\" getEnabled=\"GetFontSizeEnabled\" onChange=\"OnFontSizeChanged\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        internal static bool TryParseFontSize(string text, out double fontSizePt)
        {
            var styles = NumberStyles.Float;
            if ((!double.TryParse(text, styles, CultureInfo.CurrentCulture, out fontSizePt) &&
                 !double.TryParse(text, styles, CultureInfo.InvariantCulture, out fontSizePt)) ||
                fontSizePt < 1 || fontSizePt > 200 || double.IsNaN(fontSizePt) ||
                double.IsInfinity(fontSizePt))
            {
                fontSizePt = 0;
                return false;
            }
            return true;
        }

        internal static bool TryParseLayoutWidthPt(string text, out double widthPt)
        {
            var styles = NumberStyles.Float;
            if ((!double.TryParse(text, styles, CultureInfo.CurrentCulture, out widthPt) &&
                 !double.TryParse(text, styles, CultureInfo.InvariantCulture, out widthPt)) ||
                !BlockLayoutWidthPolicy.IsValid(widthPt))
            {
                widthPt = 0;
                return false;
            }
            return true;
        }
    }
}
