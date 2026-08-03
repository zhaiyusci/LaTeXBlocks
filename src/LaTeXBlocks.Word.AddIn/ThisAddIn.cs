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
        private LaTeXBlocksRibbon ribbon;
        private readonly Dictionary<long, PendingFormatRefresh> pendingFormatRefreshes =
            new Dictionary<long, PendingFormatRefresh>();
        private readonly HashSet<long> formatRefreshesInFlight = new HashSet<long>();
        private long formatRefreshSequence;
        private int programmaticMutationDepth;
        private bool shuttingDown;
        private bool hostEventProcessingEnabled;
        private const int NativeFontSizeControlId = 1731;
        // A profile is a host-level preference: selecting a Word profile must
        // not change the one PowerPoint starts with.
        private const string SettingsKey = @"Software\LaTeXBlocks\Word";
        private const string LegacySettingsKey = @"Software\LaTeXBlocks";
        internal WordInterop.Application WordApplication => Application;

        private StemTeXBackend Renderers => rendererPool ?? (rendererPool = new StemTeXBackend());
        private LaTeXBlockService Blocks => blocks ?? (blocks = new LaTeXBlockService(Application, Renderers));

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                wordUiDispatcher = new Control();
                wordUiDispatcher.CreateControl();
                Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                AttachNativeFontSizeControl();
                if (Application.Documents.Count > 0) RememberSelection(Application.Selection);

                var pool = Renderers;
                currentProfile = LoadCurrentProfile(pool);
                var startupProfile = currentProfile;
                backendStatus = "warming:" + startupProfile;
                pool.SwitchProfile(startupProfile);
                backendStatus = pool.Status;
                hostEventProcessingEnabled = true;
            }
            catch (Exception exception)
            {
                backendStartupError = exception.Message;
                backendStatus = "failed";
                ReleaseHostResources();
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            shuttingDown = true;
            ReleaseHostResources();
        }

        private void ReleaseHostResources()
        {
            hostEventProcessingEnabled = false;
            Interlocked.Increment(ref formatRefreshSequence);
            pendingFormatRefreshes.Clear();
            formatRefreshesInFlight.Clear();
            // Word is already part-way through COM teardown when this event runs. One
            // failed event unsubscription must never prevent the renderer shutdown
            // path from being reached: otherwise the background worker is left alive
            // until the process finally exits.
            RunBestEffortCleanup(() => Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick);
            RunBestEffortCleanup(() => Application.WindowSelectionChange -= Application_WindowSelectionChange);
            RunBestEffortCleanup(DetachNativeFontSizeControl);
            RunBestEffortCleanup(ClearPreviousSelectionSnapshot);

            var pool = rendererPool;
            rendererPool = null;
            blocks = null;
            RunBestEffortCleanup(() => pool?.Dispose());

            var dispatcher = wordUiDispatcher;
            wordUiDispatcher = null;
            RunBestEffortCleanup(() => dispatcher?.Dispose());
        }

        private static void RunBestEffortCleanup(Action cleanup)
        {
            try { cleanup?.Invoke(); }
            catch
            {
                // Startup rollback and VSTO shutdown can both race Word COM teardown.
                // Cleanup is best-effort, but every independently owned step still runs.
            }
        }

        internal void ShowInsertFormulaEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            using (var editor = new LaTeXBlockEditorForm(Blocks, "$E=mc^2$", 360,
                LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt))
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
            var widthPt = LaTeXBlockWidthPolicy.ResolveDefaultFixedWidth();
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[E=mc^2\\]", widthPt,
                LaTeXBlockLayoutMode.Fixed,
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
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            const double widthPt = 360;
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[E=mc^2\\]", widthPt,
                LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt, "Insert Numbered Equation", true))
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
            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt,
                metadata.Mode, currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, true, metadata.FontSizePt, null,
                metadata.Role == LaTeXBlockRole.NumberedEquation))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.UpdateRendered(shape, editor.Source, editor.WidthPt, editor.Mode, editor.CurrentRender));
                }
            }
        }

        internal bool HasSelectedFixedBlockWidth()
        {
            return TryGetSelectedFixedContentBlock(out _, out _);
        }

        internal string GetSelectedFixedBlockWidthText()
        {
            if (!TryGetSelectedFixedContentBlock(out _, out var metadata))
                return string.Empty;
            return metadata.WidthPt.ToString("0.0", System.Globalization.CultureInfo.CurrentCulture);
        }

        internal void ApplySelectedFixedBlockWidth(string text)
        {
            if (!LaTeXBlockWidthPolicy.TryParseWidth(text, out var requestedWidthPt))
                throw new ArgumentException("Enter a typesetting width from 30 to 450 pt.",
                    nameof(text));
            if (!TryGetSelectedFixedContentBlock(out var selectedShape, out var metadata))
                throw new InvalidOperationException("Select one fixed-width LaTeX Block first.");
            var source = selectedShape.AlternativeText;
            QueueFormatRefresh(new List<FormatRefreshRequest>
            {
                new FormatRefreshRequest(selectedShape, source, metadata, requestedWidthPt,
                    metadata.FontSizePt)
            });
        }

        private bool TryGetSelectedFixedContentBlock(out WordInterop.InlineShape shape,
            out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            if (!Blocks.TryGetSelectedBlock(out var candidate, out var candidateMetadata) ||
                candidateMetadata.Mode != LaTeXBlockLayoutMode.Fixed ||
                candidateMetadata.Role != LaTeXBlockRole.Content) return false;
            shape = candidate;
            metadata = candidateMetadata;
            return true;
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
            if (shuttingDown || !hostEventProcessingEnabled || refreshingNativeFontSize ||
                Application.Documents.Count == 0) return;
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
            if (shuttingDown || !hostEventProcessingEnabled || refreshingNativeFontSize ||
                programmaticMutationDepth > 0)
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
                ribbon?.InvalidateWidthControl();
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
                    (!onlyChanged || Math.Abs(metadata.FontSizePt - fontSizePt) > 0.001 ||
                     HasPendingFontTarget(shape, fontSizePt)))
                    requests.Add(new FormatRefreshRequest(shape, source, metadata, fontSizePt));
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
                    var shape = snapshot.Shape;
                    if (shape == null ||
                        !LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) ||
                        metadata.Mode != LaTeXBlockLayoutMode.Auto) continue;
                    var size = (double)shape.Range.Font.Size;
                    if (!LaTeXBlockService.ShouldRefreshForHostFontSizeChange(snapshot.HostFontSizePt,
                            size, metadata.FontSizePt) &&
                        !HasPendingFontTarget(shape, size)) continue;
                    requests.Add(new FormatRefreshRequest(shape, source, metadata, size));
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
                        snapshots.Add(new SelectionFontSnapshot(shape, size));
                }
            }
            catch (COMException) { }
            return snapshots;
        }

        private void QueueFormatRefresh(List<FormatRefreshRequest> requests)
        {
            if (shuttingDown || requests == null || requests.Count == 0) return;
            foreach (var request in requests) QueueFormatRefresh(request);
        }

        private void QueueFormatRefresh(FormatRefreshRequest request)
        {
            if (shuttingDown || request == null) return;
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var targetWidthPt = request.Metadata.WidthPt;
            var targetFontSizePt = request.Metadata.FontSizePt;
            if (pendingFormatRefreshes.TryGetValue(request.ShapeKey, out var existing) &&
                SameBaseState(existing, request.Metadata, request.Source, profile))
            {
                targetWidthPt = existing.TargetWidthPt;
                targetFontSizePt = existing.TargetFontSizePt;
            }
            if (request.ChangesWidth)
                targetWidthPt = request.WidthPt;
            if (request.ChangesFontSize)
                targetFontSizePt = request.FontSizePt;

            if (Math.Abs(targetWidthPt - request.Metadata.WidthPt) < 0.001 &&
                Math.Abs(targetFontSizePt - request.Metadata.FontSizePt) < 0.001)
            {
                pendingFormatRefreshes.Remove(request.ShapeKey);
                ribbon?.InvalidateWidthControl();
                return;
            }

            var sequence = Interlocked.Increment(ref formatRefreshSequence);
            var pending = new PendingFormatRefresh(request.ShapeKey, request.Shape,
                request.Metadata, request.Source, profile, targetWidthPt,
                targetFontSizePt, sequence);
            pendingFormatRefreshes[request.ShapeKey] = pending;
            StartNextFormatRefresh(request.ShapeKey);
        }

        private void StartNextFormatRefresh(long shapeKey)
        {
            if (shuttingDown || formatRefreshesInFlight.Contains(shapeKey) ||
                !pendingFormatRefreshes.TryGetValue(shapeKey, out var pending)) return;
            formatRefreshesInFlight.Add(shapeKey);
            _ = RefreshBlockAsync(Blocks, pending);
        }

        private async Task RefreshBlockAsync(LaTeXBlockService service,
            PendingFormatRefresh pending)
        {
            try
            {
                var render = await service.RenderCommittedAsync(pending.Source,
                    pending.TargetWidthPt, pending.BaseMetadata.Mode, pending.Profile,
                    pending.TargetFontSizePt,
                    pending.BaseMetadata.Role == LaTeXBlockRole.NumberedEquation)
                    .ConfigureAwait(false);
                await InvokeOnWordUiAsync(() => CompleteFormatRefresh(service, pending,
                    render)).ConfigureAwait(false);
            }
            catch (TaskCanceledException)
            {
                await AbandonFormatRefreshAsync(pending, null).ConfigureAwait(false);
            }
            catch (ObjectDisposedException)
            {
                await AbandonFormatRefreshAsync(pending, null).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                await AbandonFormatRefreshAsync(pending, exception).ConfigureAwait(false);
            }
        }

        private void CompleteFormatRefresh(LaTeXBlockService service,
            PendingFormatRefresh pending, LaTeXBlockRender render)
        {
            formatRefreshesInFlight.Remove(pending.ShapeKey);
            if (shuttingDown) return;
            if (!IsCurrentPending(pending))
            {
                StartNextFormatRefresh(pending.ShapeKey);
                return;
            }
            var shape = pending.Shape;
            if (shape != null && LaTeXBlockService.TryReadContract(shape,
                    out var currentMetadata, out var currentSource) &&
                currentSource == pending.Source &&
                SameMetadataState(currentMetadata, pending.BaseMetadata) &&
                string.Equals(pending.Profile, currentProfile,
                    StringComparison.OrdinalIgnoreCase))
                RunProgrammaticMutation(() => service.UpdateRendered(shape, pending.Source,
                    pending.TargetWidthPt, pending.BaseMetadata.Mode, render, false), false);
            pendingFormatRefreshes.Remove(pending.ShapeKey);
            ribbon?.InvalidateWidthControl();
            StartNextFormatRefresh(pending.ShapeKey);
        }

        private Task AbandonFormatRefreshAsync(PendingFormatRefresh pending,
            Exception exception)
        {
            if (shuttingDown) return Task.FromResult(false);
            return InvokeOnWordUiAsync(() =>
            {
                formatRefreshesInFlight.Remove(pending.ShapeKey);
                if (IsCurrentPending(pending))
                {
                    pendingFormatRefreshes.Remove(pending.ShapeKey);
                    ribbon?.InvalidateWidthControl();
                    if (exception != null)
                        MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                            exception.GetBaseException().Message, "LaTeX Blocks",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                StartNextFormatRefresh(pending.ShapeKey);
            });
        }

        private bool IsCurrentPending(PendingFormatRefresh pending)
        {
            return pendingFormatRefreshes.TryGetValue(pending.ShapeKey,
                       out var current) && current.Sequence == pending.Sequence;
        }

        private bool HasPendingFontTarget(WordInterop.InlineShape shape, double fontSizePt)
        {
            return shape != null && pendingFormatRefreshes.TryGetValue(
                       GetComIdentity(shape), out var pending) &&
                   Math.Abs(pending.TargetFontSizePt - fontSizePt) > 0.001;
        }

        private static bool SameBaseState(PendingFormatRefresh pending,
            LaTeXBlockMetadata metadata, string source, string profile)
        {
            return SameMetadataState(pending.BaseMetadata, metadata) &&
                   pending.Source == source && string.Equals(pending.Profile, profile,
                       StringComparison.OrdinalIgnoreCase);
        }

        private static bool SameMetadataState(LaTeXBlockMetadata left,
            LaTeXBlockMetadata right)
        {
            return left != null && right != null && left.Id == right.Id &&
                   left.Mode == right.Mode && left.Role == right.Role &&
                   Math.Abs(left.WidthPt - right.WidthPt) < 0.001 &&
                   Math.Abs(left.FontSizePt - right.FontSizePt) < 0.001;
        }

        private static long GetComIdentity(object value)
        {
            if (value == null) return 0;
            var unknown = Marshal.GetIUnknownForObject(value);
            try { return unknown.ToInt64(); }
            finally { Marshal.Release(unknown); }
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
            internal SelectionFontSnapshot(WordInterop.InlineShape shape,
                double hostFontSizePt)
            { Shape = shape; HostFontSizePt = hostFontSizePt; }
            internal WordInterop.InlineShape Shape { get; }
            internal double HostFontSizePt { get; }
        }

        private sealed class FormatRefreshRequest
        {
            internal FormatRefreshRequest(WordInterop.InlineShape shape, string source,
                LaTeXBlockMetadata metadata, double fontSizePt)
            {
                Shape = shape;
                ShapeKey = GetComIdentity(shape);
                Source = source;
                Metadata = metadata;
                WidthPt = metadata.WidthPt;
                FontSizePt = fontSizePt;
                ChangesFontSize = true;
            }
            internal FormatRefreshRequest(WordInterop.InlineShape shape, string source,
                LaTeXBlockMetadata metadata, double widthPt, double fontSizePt)
            {
                Shape = shape;
                ShapeKey = GetComIdentity(shape);
                Source = source;
                Metadata = metadata;
                WidthPt = widthPt;
                FontSizePt = fontSizePt;
                ChangesWidth = true;
            }
            internal WordInterop.InlineShape Shape { get; }
            internal long ShapeKey { get; }
            internal string Source { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal double WidthPt { get; }
            internal double FontSizePt { get; }
            internal bool ChangesWidth { get; }
            internal bool ChangesFontSize { get; }
        }

        private sealed class PendingFormatRefresh
        {
            internal PendingFormatRefresh(long shapeKey, WordInterop.InlineShape shape,
                LaTeXBlockMetadata baseMetadata, string source, string profile,
                double targetWidthPt, double targetFontSizePt, long sequence)
            {
                ShapeKey = shapeKey;
                Shape = shape;
                BaseMetadata = baseMetadata;
                Source = source;
                Profile = profile;
                TargetWidthPt = targetWidthPt;
                TargetFontSizePt = targetFontSizePt;
                Sequence = sequence;
            }

            internal long ShapeKey { get; }
            internal WordInterop.InlineShape Shape { get; }
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal string Source { get; }
            internal string Profile { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFontSizePt { get; }
            internal long Sequence { get; }
        }

        private void Application_WindowBeforeDoubleClick(WordInterop.Selection selection, ref bool cancel)
        {
            if (shuttingDown || !hostEventProcessingEnabled) return;
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
            // Upgrade from the former shared preference once. Subsequent saves
            // always go to Word's own key and can no longer affect PowerPoint.
            if (string.IsNullOrWhiteSpace(saved))
                using (var key = Registry.CurrentUser.OpenSubKey(LegacySettingsKey)) saved = key?.GetValue("Profile") as string;
            foreach (var profile in pool.Profiles)
                if (string.Equals(profile, saved, StringComparison.OrdinalIgnoreCase)) return profile;
            return pool.DefaultAvailableProfile;
        }

        private void SetCurrentProfile(string profile)
        {
            if (shuttingDown) throw new ObjectDisposedException(nameof(ThisAddIn));
            var valid = false;
            foreach (var candidate in Renderers.Profiles)
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase)) { profile = candidate; valid = true; break; }
            if (!valid) throw new ArgumentException("Unknown StemTeX profile: " + profile, nameof(profile));
            if (string.Equals(currentProfile, profile, StringComparison.OrdinalIgnoreCase)) return;

            // Do not expose the new profile to the editor, queued format updates, or
            // the next Word session until both immediate state transitions succeeded.
            // SwitchProfile can still finish warming asynchronously; its synchronous
            // validation/shutdown failures must leave the old host preference intact.
            var previousProfile = currentProfile;
            var previousSetting = ReadProfileSetting();
            Renderers.SwitchProfile(profile);
            try
            {
                WriteProfileSetting(profile);
            }
            catch (Exception persistenceException)
            {
                Exception rollbackFailure = null;
                try { RestoreProfileSetting(previousSetting); }
                catch (Exception exception) { rollbackFailure = exception; }
                if (!string.IsNullOrWhiteSpace(previousProfile))
                {
                    try { Renderers.SwitchProfile(previousProfile); }
                    catch (Exception exception) { rollbackFailure = rollbackFailure ?? exception; }
                }

                if (rollbackFailure != null)
                    throw new InvalidOperationException(
                        "Word could not save the selected StemTeX profile and could not restore the previous profile.",
                        new AggregateException(persistenceException, rollbackFailure));
                throw;
            }

            currentProfile = profile;
            pendingFormatRefreshes.Clear();
            Interlocked.Increment(ref formatRefreshSequence);
        }

        private static ProfileSetting ReadProfileSetting()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey))
            {
                var value = key?.GetValue("Profile") as string;
                return new ProfileSetting(value != null, value);
            }
        }

        private static void WriteProfileSetting(string profile)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
                key.SetValue("Profile", profile, RegistryValueKind.String);
        }

        private static void RestoreProfileSetting(ProfileSetting setting)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
            {
                if (setting.HasValue)
                    key.SetValue("Profile", setting.Value, RegistryValueKind.String);
                else
                    key.DeleteValue("Profile", false);
            }
        }

        private struct ProfileSetting
        {
            internal ProfileSetting(bool hasValue, string value)
            {
                HasValue = hasValue;
                Value = value;
            }

            internal bool HasValue { get; }
            internal string Value { get; }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return ribbon = new LaTeXBlocksRibbon(this);
        }
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
