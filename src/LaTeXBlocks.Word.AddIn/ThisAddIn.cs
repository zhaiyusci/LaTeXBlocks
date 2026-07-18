using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
        private Office.CommandBarComboBox nativeFontSizeControl;
        private bool refreshingNativeFontSize;
        private List<SelectionFontSnapshot> previousSelectionFontSnapshots = new List<SelectionFontSnapshot>();
        private Control wordUiDispatcher;
        private long formatRefreshGeneration;
        private int programmaticMutationDepth;
        private bool shuttingDown;
        private const int NativeFontSizeControlId = 1731;
        private const string SettingsKey = @"Software\LaTeXBlocks";
        internal WordInterop.Application WordApplication => Application;

        private StemTeXBackend Renderers => rendererPool ?? (rendererPool = new StemTeXBackend());
        private LaTeXBlockService Blocks => blocks ?? (blocks = new LaTeXBlockService(Application, Renderers));

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            wordUiDispatcher = new Control();
            wordUiDispatcher.CreateControl();
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            AttachNativeFontSizeControl();
            if (Application.Documents.Count > 0) RememberSelection(Application.Selection);
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
            shuttingDown = true;
            Interlocked.Increment(ref formatRefreshGeneration);
            Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
            Application.WindowSelectionChange -= Application_WindowSelectionChange;
            DetachNativeFontSizeControl();
            ClearPreviousSelectionSnapshot();
            rendererPool?.Dispose();
            rendererPool = null;
            blocks = null;
            wordUiDispatcher?.Dispose();
            wordUiDispatcher = null;
        }

        internal void ShowInsertFormulaEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection.Range,
                LaTeXBlockLayoutMode.Auto, 10);
            using (var editor = new LaTeXBlockEditorForm(Blocks, "$E=mc^2$", 360, LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false, fontSizePt))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender));
                }
            }
        }

        internal void ShowInsertBlockEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[E=mc^2\\]", 360, LaTeXBlockLayoutMode.Fixed,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false, 10))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender));
                }
            }
        }

        internal void ShowInsertNumberedEquationEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            LaTeXBlockService.ValidateNumberedEquationTarget(Application.Selection.Range);
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection.Range,
                LaTeXBlockLayoutMode.Auto, 10);
            const double widthPt = 360;
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[E=mc^2\\]", widthPt,
                LaTeXBlockLayoutMode.Auto, currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, false, fontSizePt, "Insert Numbered Equation", true))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.InsertNumberedRendered(editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender));
                }
            }
        }

        internal void UpdateEquationNumbers()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var count = 0;
            RunProgrammaticMutation(() => count = Blocks.UpdateEquationNumbers(Application.ActiveDocument));
            Application.StatusBar = count == 1
                ? "Updated 1 LaTeX equation number."
                : "Updated " + count + " LaTeX equation numbers.";
        }

        internal void ShowEditEditor()
        {
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select a LaTeX Block first.");
            var source = shape.AlternativeText;
            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt, metadata.Mode,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, true, metadata.FontSizePt,
                null, metadata.Role == LaTeXBlockRole.NumberedEquation))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.UpdateRendered(shape, editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender));
                }
            }
        }

        internal void ApplyFontSizeToSelection(double fontSizePt)
        {
            if (fontSizePt < 1 || fontSizePt > 200) throw new ArgumentOutOfRangeException(nameof(fontSizePt));
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");

            var selection = Application.Selection;
            selection.Font.Size = (float)fontSizePt;
            QueueFormatRefresh(CaptureAutoBlocks(selection.Range, fontSizePt, false));
            RememberSelection(selection);
        }

        private void AttachNativeFontSizeControl()
        {
            try
            {
                nativeFontSizeControl = Application.CommandBars.FindControl(
                    Office.MsoControlType.msoControlComboBox, NativeFontSizeControlId,
                    Type.Missing, Type.Missing) as Office.CommandBarComboBox;
                if (nativeFontSizeControl != null)
                    nativeFontSizeControl.Change += NativeFontSizeControl_Change;
            }
            catch
            {
                nativeFontSizeControl = null;
            }
        }

        private void DetachNativeFontSizeControl()
        {
            if (nativeFontSizeControl == null) return;
            try { nativeFontSizeControl.Change -= NativeFontSizeControl_Change; } catch { }
            nativeFontSizeControl = null;
        }

        private void NativeFontSizeControl_Change(Office.CommandBarComboBox control)
        {
            if (refreshingNativeFontSize || Application.Documents.Count == 0) return;
            try
            {
                refreshingNativeFontSize = true;
                var selection = Application.Selection;
                var size = (double)selection.Font.Size;
                if (size < 1 || size > 200) return;
                QueueFormatRefresh(CaptureAutoBlocks(selection.Range, size, true));
                RememberSelection(selection);
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message,
                    "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                refreshingNativeFontSize = false;
            }
        }

        private void Application_WindowSelectionChange(WordInterop.Selection selection)
        {
            if (refreshingNativeFontSize || programmaticMutationDepth > 0)
                return;

            try
            {
                refreshingNativeFontSize = true;
                // Word exposes no general formatting-changed event. If a size was
                // applied through a shortcut or another native command, validate the
                // range when the user next leaves it. This is event driven, not polling.
                QueueFormatRefresh(CaptureAutoBlocksWhoseHostSizeChanged(previousSelectionFontSnapshots));
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message,
                    "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                refreshingNativeFontSize = false;
                RememberSelection(selection);
            }
        }

        private List<FormatRefreshRequest> CaptureAutoBlocks(WordInterop.Range range, double fontSizePt,
            bool onlyChanged)
        {
            var requests = new List<FormatRefreshRequest>();
            foreach (WordInterop.InlineShape shape in range.InlineShapes)
            {
                if (LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) &&
                    metadata.Mode == LaTeXBlockLayoutMode.Auto &&
                    (!onlyChanged || Math.Abs(metadata.FontSizePt - fontSizePt) > 0.001))
                    requests.Add(new FormatRefreshRequest(metadata.Id, source, metadata, fontSizePt));
            }
            return requests;
        }

        private List<FormatRefreshRequest> CaptureAutoBlocksWhoseHostSizeChanged(
            IList<SelectionFontSnapshot> snapshots)
        {
            var requests = new List<FormatRefreshRequest>();
            if (snapshots == null || snapshots.Count == 0) return requests;
            foreach (var snapshot in snapshots)
            {
                try
                {
                    var shape = FindBlock(snapshot.Id);
                    if (shape == null ||
                        !LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) ||
                        metadata.Mode != LaTeXBlockLayoutMode.Auto) continue;
                    var size = (double)shape.Range.Font.Size;
                    if (!LaTeXBlockService.ShouldRefreshForHostFontSizeChange(snapshot.HostFontSizePt, size,
                            metadata.FontSizePt)) continue;
                    requests.Add(new FormatRefreshRequest(metadata.Id, source, metadata, size));
                }
                catch (COMException)
                {
                    // The block may have been deleted or moved across a story while
                    // the selection changed. Size refresh is opportunistic.
                }
            }
            return requests;
        }

        private List<SelectionFontSnapshot> CaptureSelectionFontSnapshots(WordInterop.Selection selection)
        {
            var snapshots = new List<SelectionFontSnapshot>();
            if (selection == null) return snapshots;
            try
            {
                foreach (WordInterop.InlineShape shape in selection.Range.InlineShapes)
                {
                    if (!LaTeXBlockService.TryReadContract(shape, out var metadata, out _) ||
                        metadata.Mode != LaTeXBlockLayoutMode.Auto) continue;
                    var size = (double)shape.Range.Font.Size;
                    if (size >= 1 && size <= 200)
                        snapshots.Add(new SelectionFontSnapshot(metadata.Id, size));
                }
            }
            catch (COMException) { }
            return snapshots;
        }

        private void QueueFormatRefresh(List<FormatRefreshRequest> requests)
        {
            if (shuttingDown || requests == null || requests.Count == 0) return;
            var generation = Interlocked.Increment(ref formatRefreshGeneration);
            var service = Blocks;
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            RefreshBlocksAsync(service, profile, requests, generation);
        }

        private async void RefreshBlocksAsync(LaTeXBlockService service, string profile,
            List<FormatRefreshRequest> requests, long generation)
        {
            try
            {
                // Render serially on the StemTeX background queue. Only the small Word
                // object-model replacement is marshalled back to Office's UI thread.
                for (var index = requests.Count - 1; index >= 0; index--)
                {
                    var request = requests[index];
                    var render = await service.RenderPreviewAsync(request.Source, request.Metadata.WidthPt,
                        request.Metadata.Mode, profile, request.FontSizePt,
                        request.Metadata.Role == LaTeXBlockRole.NumberedEquation).ConfigureAwait(false);
                    if (shuttingDown || generation != Interlocked.Read(ref formatRefreshGeneration)) return;
                    await InvokeOnWordUiAsync(() =>
                    {
                        if (shuttingDown || generation != Interlocked.Read(ref formatRefreshGeneration)) return;
                        var shape = FindBlock(request.Id);
                        if (shape != null)
                            RunProgrammaticMutation(() => service.UpdateRendered(shape, request.Source,
                                request.Metadata.WidthPt, request.Metadata.Mode, render, false), false);
                    }).ConfigureAwait(false);
                }
            }
            catch (TaskCanceledException) { }
            catch (ObjectDisposedException) { }
            catch (Exception exception)
            {
                if (shuttingDown || generation != Interlocked.Read(ref formatRefreshGeneration)) return;
                await InvokeOnWordUiAsync(() => MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                    exception.GetBaseException().Message, "LaTeX Blocks", MessageBoxButtons.OK,
                    MessageBoxIcon.Error)).ConfigureAwait(false);
            }
        }

        private Task InvokeOnWordUiAsync(Action action)
        {
            var completion = new TaskCompletionSource<bool>();
            var dispatcher = wordUiDispatcher;
            if (shuttingDown || dispatcher == null || dispatcher.IsDisposed)
            {
                completion.SetCanceled();
                return completion.Task;
            }
            try
            {
                dispatcher.BeginInvoke(new Action(() =>
                {
                    try { action(); completion.TrySetResult(true); }
                    catch (Exception exception) { completion.TrySetException(exception); }
                }));
            }
            catch (InvalidOperationException) { completion.TrySetCanceled(); }
            return completion.Task;
        }

        private WordInterop.InlineShape FindBlock(Guid id)
        {
            foreach (WordInterop.Document document in Application.Documents)
                foreach (WordInterop.InlineShape shape in document.InlineShapes)
                    if (LaTeXBlockService.TryReadContract(shape, out var metadata, out _) && metadata.Id == id) return shape;
            return null;
        }

        private void RunProgrammaticMutation(Action action, bool recaptureSelectionSnapshot = true)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            var outermost = programmaticMutationDepth == 0;
            programmaticMutationDepth++;
            if (outermost && recaptureSelectionSnapshot) ClearPreviousSelectionSnapshot();
            try
            {
                action();
            }
            finally
            {
                programmaticMutationDepth--;
                if (outermost && recaptureSelectionSnapshot && !shuttingDown && Application.Documents.Count > 0)
                {
                    try { RememberSelection(Application.Selection); }
                    catch { ClearPreviousSelectionSnapshot(); }
                }
            }
        }

        private void RememberSelection(WordInterop.Selection selection)
        {
            previousSelectionFontSnapshots = CaptureSelectionFontSnapshots(selection);
        }

        private void ClearPreviousSelectionSnapshot()
        {
            previousSelectionFontSnapshots = new List<SelectionFontSnapshot>();
        }

        private sealed class SelectionFontSnapshot
        {
            internal SelectionFontSnapshot(Guid id, double hostFontSizePt)
            { Id = id; HostFontSizePt = hostFontSizePt; }
            internal Guid Id { get; }
            internal double HostFontSizePt { get; }
        }

        private sealed class FormatRefreshRequest
        {
            internal FormatRefreshRequest(Guid id, string source, LaTeXBlockMetadata metadata, double fontSizePt)
            { Id = id; Source = source; Metadata = metadata; FontSizePt = fontSizePt; }
            internal Guid Id { get; }
            internal string Source { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal double FontSizePt { get; }
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
