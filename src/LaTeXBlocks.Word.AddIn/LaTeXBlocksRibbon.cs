using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    [ComVisible(true)]
    public sealed class LaTeXBlocksRibbon : Office.IRibbonExtensibility
    {
        internal const string WidthControlId = "LaTeXBlocks.WidthPt";
        private readonly ThisAddIn addIn;
        private Office.IRibbonUI ribbonUi;

        internal LaTeXBlocksRibbon(ThisAddIn addIn) { this.addIn = addIn ?? throw new ArgumentNullException(nameof(addIn)); }

        public string GetCustomUI(string ribbonId)
        {
            return BuildCustomUi();
        }

        internal static string BuildCustomUi()
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">" +
                   "<ribbon><tabs><tab id=\"LaTeXBlocks.Tab\" label=\"LaTeX Blocks\">" +
                   "<group id=\"LaTeXBlocks.Blocks\" label=\"LaTeX\">" +
                   "<button id=\"LaTeXBlocks.InsertFormula\" label=\"Insert Formula\" size=\"large\" imageMso=\"EquationInsertNew\" onAction=\"OnInsertFormula\"/>" +
                   "<button id=\"LaTeXBlocks.InsertBlock\" label=\"Insert Block\" size=\"large\" imageMso=\"TextBoxInsert\" onAction=\"OnInsertBlock\"/>" +
                   "<button id=\"LaTeXBlocks.InsertNumberedEquation\" label=\"Numbered Equation\" size=\"large\" imageMso=\"CaptionInsert\" onAction=\"OnInsertNumberedEquation\"/>" +
                   "<button id=\"LaTeXBlocks.Edit\" label=\"Edit Block\" size=\"large\" imageMso=\"ObjectEdit\" onAction=\"OnEdit\"/>" +
                   "<button id=\"LaTeXBlocks.UpdateEquationNumbers\" label=\"Update Numbers\" imageMso=\"FieldsUpdate\" onAction=\"OnUpdateEquationNumbers\"/>" +
                   "<editBox id=\"" + WidthControlId +
                   "\" label=\"Typesetting width (pt)\" sizeString=\"000.0\" getText=\"GetWidthText\" getEnabled=\"GetWidthEnabled\" onChange=\"OnWidthChanged\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnLoad(Office.IRibbonUI ui) { ribbonUi = ui; }
        public void OnInsertFormula(Office.IRibbonControl control) { Run(addIn.ShowInsertFormulaEditor); }
        public void OnInsertBlock(Office.IRibbonControl control) { Run(addIn.ShowInsertBlockEditor); }
        public void OnInsertNumberedEquation(Office.IRibbonControl control) { Run(addIn.ShowInsertNumberedEquationEditor); }
        public void OnEdit(Office.IRibbonControl control) { Run(addIn.ShowEditEditor); }
        public void OnUpdateEquationNumbers(Office.IRibbonControl control) { Run(addIn.UpdateEquationNumbers); }
        public string GetWidthText(Office.IRibbonControl control)
        {
            try { return addIn.GetSelectedFixedBlockWidthText(); }
            catch { return string.Empty; }
        }
        public bool GetWidthEnabled(Office.IRibbonControl control)
        {
            try { return addIn.HasSelectedFixedBlockWidth(); }
            catch { return false; }
        }
        public void OnWidthChanged(Office.IRibbonControl control, string text)
        {
            Run(() => addIn.ApplySelectedFixedBlockWidth(text));
        }

        internal void InvalidateWidthControl()
        {
            try { ribbonUi?.InvalidateControl(WidthControlId); }
            catch (COMException) { }
        }

        private void Run(Action action)
        {
            try { action(); }
            catch (Exception exception)
            {
                MessageBox.Show(new WordWindow(addIn.WordApplication), exception.Message, "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { InvalidateWidthControl(); }
        }

        internal sealed class WordWindow : IWin32Window
        {
            internal WordWindow(WordInterop.Application application)
            {
                try { Handle = new IntPtr(application.ActiveWindow.Hwnd); } catch { Handle = IntPtr.Zero; }
            }
            public IntPtr Handle { get; }
        }
    }
}
