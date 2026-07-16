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
        private readonly ThisAddIn addIn;

        internal LaTeXBlocksRibbon(ThisAddIn addIn) { this.addIn = addIn ?? throw new ArgumentNullException(nameof(addIn)); }

        public string GetCustomUI(string ribbonId)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                   "<ribbon><tabs><tab id=\"LaTeXBlocks.Tab\" label=\"LaTeX Blocks\">" +
                   "<group id=\"LaTeXBlocks.Blocks\" label=\"LaTeX\">" +
                   "<button id=\"LaTeXBlocks.InsertFormula\" label=\"Insert Formula\" size=\"large\" imageMso=\"EquationInsertNew\" onAction=\"OnInsertFormula\"/>" +
                   "<button id=\"LaTeXBlocks.InsertBlock\" label=\"Insert Block\" size=\"large\" imageMso=\"TextBoxInsert\" onAction=\"OnInsertBlock\"/>" +
                   "<button id=\"LaTeXBlocks.Edit\" label=\"Edit Block\" size=\"large\" imageMso=\"ObjectEdit\" onAction=\"OnEdit\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnInsertFormula(Office.IRibbonControl control) { Run(addIn.ShowInsertFormulaEditor); }
        public void OnInsertBlock(Office.IRibbonControl control) { Run(addIn.ShowInsertBlockEditor); }
        public void OnEdit(Office.IRibbonControl control) { Run(addIn.ShowEditEditor); }

        private void Run(Action action)
        {
            try { action(); }
            catch (Exception exception)
            {
                MessageBox.Show(new WordWindow(addIn.WordApplication), exception.Message, "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
