using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace LaTeXBlocks.PowerPoint
{
    [ComVisible(true)]
    public sealed class LaTeXBlocksRibbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn addIn;
        private Office.IRibbonUI ribbonUi;

        internal LaTeXBlocksRibbon(ThisAddIn addIn)
        {
            this.addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public string GetCustomUI(string ribbonId)
        {
            return PowerPointRibbonContract.BuildCustomUi();
        }

        public void OnLoad(Office.IRibbonUI ui) { ribbonUi = ui; }
        public object GetCommandImage(Office.IRibbonControl control)
        {
            return Branding.RibbonImageProvider.GetImage(control.Id);
        }
        public void OnInsertBlock(Office.IRibbonControl control) { Run(addIn.ShowInsertBlockEditor); }
        public void OnEditBlock(Office.IRibbonControl control) { Run(addIn.ShowEditBlockEditor); }
        public void OnAbout(Office.IRibbonControl control) { Run(addIn.ShowAbout); }
        public string GetLayoutWidthText(Office.IRibbonControl control)
        {
            try { return addIn.GetSelectedBlockLayoutWidthText(); }
            catch { return string.Empty; }
        }
        public bool GetLayoutWidthEnabled(Office.IRibbonControl control)
        {
            try { return addIn.HasSelectedBlockLayoutWidth(); }
            catch { return false; }
        }
        public void OnLayoutWidthChanged(Office.IRibbonControl control, string text)
        {
            Run(() => addIn.ApplySelectedBlockLayoutWidth(text));
        }
        public string GetFontSizeText(Office.IRibbonControl control)
        {
            try { return addIn.GetSelectedBlockFontSizeText(); }
            catch { return string.Empty; }
        }
        public bool GetFontSizeEnabled(Office.IRibbonControl control)
        {
            try { return addIn.HasSelectedBlockFontSize(); }
            catch { return false; }
        }
        public void OnFontSizeChanged(Office.IRibbonControl control, string text)
        {
            Run(() => addIn.ApplySelectedBlockFontSize(text));
        }

        internal void InvalidateBlockControls()
        {
            try { ribbonUi?.InvalidateControl(PowerPointRibbonContract.FontSizeControlId); }
            catch (COMException) { }
            try { ribbonUi?.InvalidateControl(PowerPointRibbonContract.LayoutWidthControlId); }
            catch (COMException) { }
        }

        private void Run(Action action)
        {
            try { action(); }
            catch (Exception exception)
            {
                MessageBox.Show(new PowerPointWindow(addIn.PowerPointApplication), exception.Message,
                    "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { InvalidateBlockControls(); }
        }

        internal sealed class PowerPointWindow : IWin32Window
        {
            internal PowerPointWindow(PowerPointInterop.Application application)
            {
                try { Handle = new IntPtr(application.HWND); }
                catch { Handle = IntPtr.Zero; }
            }

            public IntPtr Handle { get; }
        }
    }
}
