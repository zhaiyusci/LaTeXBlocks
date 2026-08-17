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
        internal const string ReflowFrameControlId = "LaTeXBlocks.ReflowFrame";
        internal const string DontExpandShiftEnterControlId =
            "LaTeXBlocks.DontExpandShiftEnter";
        // EditText is Office's standard pencil/text-edit glyph and is available
        // in both Word and PowerPoint Ribbon hosts.
        internal const string EditBlockImageMso = "EditText";
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
                   "<button id=\"LaTeXBlocks.InsertFormula\" label=\"Inline Math\" size=\"large\" getImage=\"GetCommandImage\" onAction=\"OnInsertFormula\"/>" +
                   "<button id=\"LaTeXBlocks.InsertDisplayMath\" label=\"Display Math\" size=\"large\" getImage=\"GetCommandImage\" onAction=\"OnInsertDisplayMath\"/>" +
                   "<button id=\"LaTeXBlocks.InsertNumberedEquation\" label=\"Numbered Math\" size=\"large\" getImage=\"GetCommandImage\" onAction=\"OnInsertNumberedEquation\"/>" +
                   "<button id=\"LaTeXBlocks.InsertBlock\" label=\"LaTeX Block\" size=\"large\" getImage=\"GetCommandImage\" onAction=\"OnInsertBlock\"/>" +
                   "<button id=\"LaTeXBlocks.InsertEquationReference\" label=\"Equation Reference\" getImage=\"GetCommandImage\" onAction=\"OnInsertEquationReference\"/>" +
                   "<button id=\"LaTeXBlocks.Edit\" label=\"Edit Block\" size=\"large\" imageMso=\"" +
                   EditBlockImageMso + "\" onAction=\"OnEdit\"/>" +
                   "<button id=\"" + ReflowFrameControlId +
                   "\" label=\"Reflow Frame\" imageMso=\"RefreshAll\" getEnabled=\"GetReflowFrameEnabled\" onAction=\"OnReflowFrame\"/>" +
                   "<button id=\"LaTeXBlocks.UpdateEquationNumbers\" label=\"Update Numbers\" imageMso=\"FieldsUpdate\" onAction=\"OnUpdateEquationNumbers\"/>" +
                   "<button id=\"LaTeXBlocks.CopyAsLaTeX\" label=\"Copy as LaTeX\" imageMso=\"Copy\" onAction=\"OnCopyAsLaTeX\"/>" +
                   "<button id=\"LaTeXBlocks.PasteFromLaTeX\" label=\"Paste from LaTeX\" imageMso=\"Paste\" onAction=\"OnPasteFromLaTeX\"/>" +
                   "<toggleButton id=\"" + DontExpandShiftEnterControlId +
                   "\" label=\"Don't Expand Shift+Enter Lines\" getEnabled=\"GetDontExpandShiftEnterEnabled\" getPressed=\"GetDontExpandShiftEnterPressed\" onAction=\"OnDontExpandShiftEnter\"/>" +
                   "<editBox id=\"" + WidthControlId +
                   "\" label=\"Typesetting width (pt)\" sizeString=\"000.0\" getText=\"GetWidthText\" getEnabled=\"GetWidthEnabled\" onChange=\"OnWidthChanged\"/>" +
                   "</group><group id=\"LaTeXBlocks.AboutGroup\" label=\"LaTeX Blocks\">" +
                   "<button id=\"LaTeXBlocks.About\" label=\"About\" imageMso=\"Info\" onAction=\"OnAbout\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnLoad(Office.IRibbonUI ui) { ribbonUi = ui; }
        public object GetCommandImage(Office.IRibbonControl control)
        {
            return Branding.RibbonImageProvider.GetImage(control.Id);
        }
        public void OnInsertFormula(Office.IRibbonControl control) { Run(addIn.ShowInsertFormulaEditor); }
        public void OnInsertDisplayMath(Office.IRibbonControl control) { Run(addIn.ShowInsertDisplayMathEditor); }
        public void OnInsertBlock(Office.IRibbonControl control) { Run(addIn.ShowInsertBlockEditor); }
        public void OnInsertNumberedEquation(Office.IRibbonControl control) { Run(addIn.ShowInsertNumberedEquationEditor); }
        public void OnInsertEquationReference(Office.IRibbonControl control) { Run(addIn.ShowInsertEquationReference); }
        public void OnEdit(Office.IRibbonControl control) { Run(addIn.ShowEditEditor); }
        public void OnReflowFrame(Office.IRibbonControl control) { Run(addIn.ReflowSelectedBlockFrame); }
        public void OnUpdateEquationNumbers(Office.IRibbonControl control) { Run(addIn.UpdateEquationNumbers); }
        public void OnCopyAsLaTeX(Office.IRibbonControl control) { Run(addIn.CopySelectionAsLaTeX); }
        public void OnPasteFromLaTeX(Office.IRibbonControl control) { Run(addIn.PasteFromLaTeX); }
        public void OnAbout(Office.IRibbonControl control) { Run(addIn.ShowAbout); }
        public bool GetDontExpandShiftEnterEnabled(Office.IRibbonControl control)
        {
            try { return addIn.HasActiveDocument(); }
            catch { return false; }
        }
        public bool GetDontExpandShiftEnterPressed(Office.IRibbonControl control)
        {
            try { return addIn.GetDontExpandShiftEnter(); }
            catch { return false; }
        }
        public void OnDontExpandShiftEnter(Office.IRibbonControl control, bool pressed)
        {
            Run(() => addIn.SetDontExpandShiftEnter(pressed));
        }
        public bool GetReflowFrameEnabled(Office.IRibbonControl control)
        {
            try { return addIn.HasSelectedBlockFrame(); }
            catch { return false; }
        }
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
            try
            {
                ribbonUi?.InvalidateControl(WidthControlId);
                ribbonUi?.InvalidateControl(ReflowFrameControlId);
                ribbonUi?.InvalidateControl(DontExpandShiftEnterControlId);
            }
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
