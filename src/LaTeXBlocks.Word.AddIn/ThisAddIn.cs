using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    public partial class ThisAddIn
    {
        private StemTeXBackend rendererPool;
        private LaTeXBlockService blocks;
        private RuntimeDiagnostics diagnostics;
        private string currentProfile;
        private string backendStartupError;
        private string backendStatus = "not-started";
        private const string SettingsKey = @"Software\LaTeXBlocks";
        internal WordInterop.Application WordApplication => Application;

        private StemTeXBackend Renderers => rendererPool ?? (rendererPool = new StemTeXBackend());
        private LaTeXBlockService Blocks => blocks ?? (blocks = new LaTeXBlockService(Application, Renderers));

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
            try
            {
                var pool = Renderers;
                currentProfile = LoadCurrentProfile(pool);
                var startupProfile = currentProfile;
                backendStatus = "warming:" + startupProfile;
                pool.SwitchProfile(startupProfile);
                backendStatus = pool.Status;
            }
            catch (Exception exception)
            {
                backendStartupError = exception.Message;
                backendStatus = "failed";
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
            rendererPool?.Dispose();
            rendererPool = null;
            blocks = null;
        }

        internal void ShowInsertFormulaEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            using (var editor = new LaTeXBlockEditorForm(Blocks, "$E=mc^2$", 360, LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender);
                }
            }
        }

        internal void ShowInsertBlockEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[E=mc^2\\]", 360, LaTeXBlockLayoutMode.Fixed,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender);
                }
            }
        }

        internal void ShowEditEditor()
        {
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select a LaTeX Block first.");
            var source = shape.AlternativeText;
            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt, metadata.Mode,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, true))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    Blocks.UpdateRendered(shape, editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender);
                }
            }
        }

        private void Application_WindowBeforeDoubleClick(WordInterop.Selection selection, ref bool cancel)
        {
            try
            {
                if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata)) return;
                cancel = true;
                ShowEditEditor();
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message, "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string LoadCurrentProfile(StemTeXBackend pool)
        {
            string saved = null;
            using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey)) saved = key?.GetValue("Profile") as string;
            foreach (var profile in pool.Profiles)
                if (string.Equals(profile, saved, StringComparison.OrdinalIgnoreCase)) return profile;
            return pool.DefaultAvailableProfile;
        }

        private void SetCurrentProfile(string profile)
        {
            var valid = false;
            foreach (var candidate in Renderers.Profiles)
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase)) { profile = candidate; valid = true; break; }
            if (!valid) throw new ArgumentException("Unknown StemTeX profile: " + profile, nameof(profile));
            currentProfile = profile;
            Renderers.SwitchProfile(profile);
            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey)) key.SetValue("Profile", profile, RegistryValueKind.String);
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() { return new LaTeXBlocksRibbon(this); }
        protected override object RequestComAddInAutomationService() { return diagnostics ?? (diagnostics = new RuntimeDiagnostics(this)); }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.AutoDual)]
        public sealed class RuntimeDiagnostics
        {
            private readonly ThisAddIn addIn;
            internal RuntimeDiagnostics(ThisAddIn addIn) { this.addIn = addIn; }
            public string AssemblyVersion => typeof(ThisAddIn).Assembly.GetName().Version.ToString();
            public string AssemblyLocation => typeof(ThisAddIn).Assembly.Location;
            public string ObjectContract => "svg+alt-tex+title-metadata-v1";
            public string StemTeXHome
            {
                get { try { return addIn.Renderers.StemTeXHome; } catch (Exception exception) { return "Unavailable: " + exception.Message; } }
            }
            public string Profiles { get { try { return string.Join(",", addIn.Renderers.Profiles); } catch { return string.Empty; } } }
            public string CurrentProfile => addIn.currentProfile ?? string.Empty;
            public string BackendStartup => addIn.backendStartupError == null ? addIn.Renderers.Status : addIn.backendStatus + ": " + addIn.backendStartupError;
        }
    }
}
