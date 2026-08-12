using System;
using System.Collections.Generic;
using System.Globalization;
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
        private IStemTeXBackend rendererPool;
        private LaTeXBlockService blocks;
        private RuntimeDiagnostics diagnostics;
        private string currentProfile;
        private string backendStartupError;
        private string backendStatus = "not-started";
        private Office.CommandBarComboBox nativeFontSizeControl;
        private bool refreshingNativeFontSize;
        private long lastFontSizeCommitUtcTicks;
        private double lastFontSizeCommitPt;
        private SelectionRangeLease lastFontSizeCommitLease;
        private List<SelectionFontSnapshot> previousSelectionFontSnapshots = new List<SelectionFontSnapshot>();
        private SelectionRangeLease previousSelectionRangeLease;
        private BlockFrameSnapshot previousBlockFrameSnapshot;
        private Control wordUiDispatcher;
        private WordMouseCaptureMonitor wordMouseCaptureMonitor;
        private IWordFormatInteractionSource wordFormatInteractionSource;
        private PendingFontColorInteraction pendingFontColorInteraction;
        private WordInterop.ApplicationEvents4_Event applicationEvents;
        private LaTeXBlocksRibbon ribbon;
        private readonly Dictionary<long, PendingFormatRefresh> pendingFormatRefreshes =
            new Dictionary<long, PendingFormatRefresh>();
        private readonly HashSet<long> formatRefreshesInFlight = new HashSet<long>();
        private long formatRefreshSequence;
        private readonly Dictionary<long, PendingFormatBatchTarget>
            pendingFormatBatchTargets =
                new Dictionary<long, PendingFormatBatchTarget>();
        private long formatBatchSequence;
        // A metadata id is intentionally not a scheduling key: Word retains the
        // Alternative Text magic-header metadata
        // metadata when a user copies a Block, while the two COM picture objects
        // must still be allowed to render independently.
        private readonly Dictionary<long, PendingBlockFrameReflow> pendingBlockFrameReflows =
            new Dictionary<long, PendingBlockFrameReflow>();
        private readonly HashSet<long> blockFrameReflowsInFlight = new HashSet<long>();
        private long blockFrameReflowSequence;
        private int programmaticMutationDepth;
        private int hostResourcesReleased;
        private bool shuttingDown;
        private bool hostEventProcessingEnabled;
        private const int NativeFontSizeControlId = 1731;
        // A profile is a host-level preference: selecting a Word profile must
        // not change the one PowerPoint starts with.
        private const string SettingsKey = @"Software\LaTeXBlocks\Word";
        private const string LegacySettingsKey = @"Software\LaTeXBlocks";
        internal WordInterop.Application WordApplication => Application;

        private IStemTeXBackend Renderers => rendererPool ?? (rendererPool = new RenderHostClientBackend());
        private LaTeXBlockService Blocks => blocks ?? (blocks = new LaTeXBlockService(Application, Renderers));

        internal bool HasActiveDocument()
        {
            return Application.Documents.Count > 0;
        }

        internal bool GetDontExpandShiftEnter()
        {
            // The current Word Advanced-options checkbox is the inverse of the
            // legacy compatibility flag exposed by the PIA. Keep that inversion
            // here so the Ribbon's pressed state describes the visible behavior.
            return HasActiveDocument() && !Application.ActiveDocument.Compatibility[
                WordInterop.WdCompatibility.wdExpandShiftReturn];
        }

        internal void SetDontExpandShiftEnter(bool enabled)
        {
            if (!HasActiveDocument())
                throw new InvalidOperationException("Open a Word document first.");
            Application.ActiveDocument.Compatibility[
                WordInterop.WdCompatibility.wdExpandShiftReturn] = !enabled;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                wordUiDispatcher = new Control();
                wordUiDispatcher.CreateControl();
                // Word's application Quit event is raised while the COM server is
                // still accepting event unsubscriptions.  VSTO's AddIn Shutdown can
                // arrive later in the host's teardown sequence, after Word has begun
                // waiting on event sinks.  Clean up at the earlier boundary too.
                applicationEvents = (WordInterop.ApplicationEvents4_Event)Application;
                applicationEvents.Quit += Application_Quit;
                Application.DocumentBeforeClose += Application_DocumentBeforeClose;
                Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                AttachNativeFontSizeControl();
                if (Application.Documents.Count > 0)
                {
                    RememberSelection(Application.Selection);
                }

                var pool = Renderers;
                currentProfile = LoadCurrentProfile(pool);
                var startupProfile = currentProfile;
                backendStatus = "warming:" + startupProfile;
                pool.SwitchProfile(startupProfile);
                backendStatus = pool.Status;
                hostEventProcessingEnabled = true;
                AttachWordMouseCaptureMonitor();
                AttachWordFontColorMonitor();

                // A CaptionLabel is Word application state rather than DOCX data.
                // Register the category when the add-in starts so Word recognizes
                // existing numbered equations under one native category. The actual
                // per-equation target remains its document bookmark; manual-break
                // lines are not independently listed by Word's built-in dialog.
                // Failure here must not disable the renderer or the rest of the
                // add-in; number insertion retries the registration transactionally.
                try
                {
                    Blocks.EnsureEquationCategory();
                }
                catch { }
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

        private void Application_Quit()
        {
            shuttingDown = true;
            ReleaseHostResources();
        }

        private void ReleaseHostResources()
        {
            if (Interlocked.Exchange(ref hostResourcesReleased, 1) != 0) return;
            hostEventProcessingEnabled = false;
            Interlocked.Increment(ref formatRefreshSequence);
            pendingFormatRefreshes.Clear();
            formatRefreshesInFlight.Clear();
            Interlocked.Increment(ref formatBatchSequence);
            pendingFormatBatchTargets.Clear();
            Interlocked.Increment(ref blockFrameReflowSequence);
            pendingBlockFrameReflows.Clear();
            blockFrameReflowsInFlight.Clear();
            // Once Word has raised Quit/Shutdown, its COM connection-point server is
            // already being dismantled. Even an event -= operation is an RPC and can
            // wait for Office's teardown timeout. The process owns those connection
            // points and will destroy them; shutdown must make no Office COM calls.
            // Startup rollback is different: Word remains alive, so detach normally.
            if (shuttingDown)
            {
                applicationEvents = null;
                nativeFontSizeControl = null;
            }
            else
            {
                var events = applicationEvents;
                applicationEvents = null;
                if (events != null)
                    RunBestEffortCleanup(() => events.Quit -= Application_Quit);
                RunBestEffortCleanup(() => Application.DocumentBeforeClose -=
                    Application_DocumentBeforeClose);
                RunBestEffortCleanup(() => Application.WindowBeforeDoubleClick -=
                    Application_WindowBeforeDoubleClick);
                RunBestEffortCleanup(() => Application.WindowSelectionChange -=
                    Application_WindowSelectionChange);
                RunBestEffortCleanup(DetachNativeFontSizeControl);
            }
            RunBestEffortCleanup(DetachWordMouseCaptureMonitor);
            RunBestEffortCleanup(DetachWordFontColorMonitor);
            RunBestEffortCleanup(ClearPreviousSelectionSnapshot);

            var pool = rendererPool;
            rendererPool = null;
            blocks = null;
            RunBestEffortCleanup(() => pool?.Dispose());

            var dispatcher = wordUiDispatcher;
            wordUiDispatcher = null;
            // Destroying a hidden WinForms handle can pump teardown messages. During
            // process exit Windows will reclaim it; only rollback needs Dispose.
            if (!shuttingDown)
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

        private void Application_DocumentBeforeClose(WordInterop.Document document,
            ref bool cancel)
        {
            if (shuttingDown) return;

            // Discard formula Shape snapshots and pending work while the document is
            // still valid. Selection leases themselves contain scalar coordinates
            // only; no close-time COM release or global GC is required.
            pendingFontColorInteraction = null;
            Interlocked.Increment(ref formatRefreshSequence);
            pendingFormatRefreshes.Clear();
            formatRefreshesInFlight.Clear();
            Interlocked.Increment(ref formatBatchSequence);
            pendingFormatBatchTargets.Clear();
            Interlocked.Increment(ref blockFrameReflowSequence);
            pendingBlockFrameReflows.Clear();
            blockFrameReflowsInFlight.Clear();
            ClearPreviousSelectionSnapshot();
        }

        private void AttachWordMouseCaptureMonitor()
        {
            // The monitor is an enhancement over Word's selection-change fallback.
            // A locked-down desktop can refuse a WinEvent hook; formula rendering
            // and normal document editing must still start successfully in that case.
            try
            {
                wordMouseCaptureMonitor = new WordMouseCaptureMonitor(wordUiDispatcher);
                wordMouseCaptureMonitor.CaptureStarted += WordMouseCaptureMonitor_CaptureStarted;
                wordMouseCaptureMonitor.CaptureEnded += WordMouseCaptureMonitor_CaptureEnded;
                wordMouseCaptureMonitor.Start();
            }
            catch
            {
                DetachWordMouseCaptureMonitor();
            }
        }

        private void DetachWordMouseCaptureMonitor()
        {
            var monitor = wordMouseCaptureMonitor;
            wordMouseCaptureMonitor = null;
            if (monitor == null) return;
            try { monitor.CaptureStarted -= WordMouseCaptureMonitor_CaptureStarted; } catch { }
            try { monitor.CaptureEnded -= WordMouseCaptureMonitor_CaptureEnded; } catch { }
            try { monitor.Dispose(); } catch { }
        }

        private void AttachWordFontColorMonitor()
        {
            // Font Color is a Fluent gallery, not a Word object-model command event.
            // UI Automation exposes its stable control ids and commit patterns. If
            // accessibility is unavailable, fail closed: ordinary formula editing and
            // the selection-change format fallback continue to work.
            try
            {
                wordFormatInteractionSource = new WordFontColorMonitor(wordUiDispatcher);
                wordFormatInteractionSource.FormatInteraction +=
                    WordFormatInteractionSource_FormatInteraction;
                wordFormatInteractionSource.Start();
                UpdateNativeFormatMonitorContext(Application.Selection);
            }
            catch
            {
                DetachWordFontColorMonitor();
            }
        }

        private void DetachWordFontColorMonitor()
        {
            var monitor = wordFormatInteractionSource;
            wordFormatInteractionSource = null;
            pendingFontColorInteraction = null;
            if (monitor == null) return;
            try { monitor.FormatInteraction -=
                WordFormatInteractionSource_FormatInteraction; } catch { }
            try
            {
                var nativeMonitor = monitor as WordFontColorMonitor;
                if (shuttingDown && nativeMonitor != null)
                {
                    nativeMonitor.SetInteractionContext(false);
                    nativeMonitor.DisposeForHostShutdown();
                }
                else
                    monitor.Dispose();
            }
            catch { }
        }

        private void WordMouseCaptureMonitor_CaptureStarted(object sender, EventArgs e)
        {
            if (shuttingDown || !hostEventProcessingEnabled || programmaticMutationDepth > 0)
                return;
            // Selection change normally captured the pre-drag geometry already.
            // An out-of-context WinEvent callback may be dispatched only after
            // Word's modal drag loop has returned, so overwriting that snapshot here
            // would turn the final geometry into its own baseline. Capture only when
            // there is no valid prior selection snapshot (for example a keyboard
            // selection path that did not raise WindowSelectionChange).
            if (!IsBlockFrameSnapshotStillSelected(previousBlockFrameSnapshot))
                previousBlockFrameSnapshot = CaptureBlockFrameSnapshot();
        }

        private void WordMouseCaptureMonitor_CaptureEnded(object sender, EventArgs e)
        {
            if (shuttingDown || !hostEventProcessingEnabled || programmaticMutationDepth > 0)
                return;
            try
            {
                // This path runs after one queued UI turn following WM_LBUTTONUP;
                // it is a native gesture completion, not a recurring geometry poll.
                // Preserve the block selection if the user left it selected while
                // StemTeX renders; later selection changes never get stolen back.
                QueueBlockFrameReflow(CreateBlockFrameReflowRequest(
                    previousBlockFrameSnapshot, false, true));
            }
            catch (COMException)
            {
                // The shape can disappear through Undo/Delete before the queued UI
                // turn runs.  In that case there is simply nothing to commit.
            }
            finally
            {
                ribbon?.InvalidateWidthControl();
            }
        }

        private void WordFormatInteractionSource_FormatInteraction(object sender,
            WordFormatInteractionEventArgs e)
        {
            if (e == null || shuttingDown || !hostEventProcessingEnabled ||
                programmaticMutationDepth > 0 || Application.Documents.Count == 0)
                return;
            try
            {
                if (e.Property == WordFormatProperty.FontSize)
                {
                    if (e.Phase == WordFormatInteractionPhase.Committed)
                        _ = CommitFontSizeInteractionAfterHostAsync();
                    return;
                }
                if (e.Property != WordFormatProperty.TextColor) return;
                switch (e.Phase)
                {
                    case WordFormatInteractionPhase.Began:
                        pendingFontColorInteraction =
                            CaptureFontColorInteraction(e.InteractionId);
                        break;
                    case WordFormatInteractionPhase.Canceled:
                        if (pendingFontColorInteraction?.InteractionId == e.InteractionId)
                            pendingFontColorInteraction = null;
                        break;
                    case WordFormatInteractionPhase.Committed:
                        QueueFontColorInteractionCommit(e.InteractionId);
                        break;
                }
            }
            catch (Exception exception)
            {
                pendingFontColorInteraction = null;
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                    exception.GetBaseException().Message, "LaTeX Blocks",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (e?.Property == WordFormatProperty.FontSize)
                    ribbon?.InvalidateWidthControl();
            }
        }

        private async Task CommitFontSizeInteractionAfterHostAsync(
            double? requestedFontSizePt = null)
        {
            try
            {
                await Task.Delay(75).ConfigureAwait(false);
                await InvokeOnWordUiAsync(() =>
                    RefreshSelectedFontSize(requestedFontSizePt)).ConfigureAwait(false);
            }
            catch (TaskCanceledException) { }
            catch (ObjectDisposedException) { }
        }

        private void QueueFontColorInteractionCommit(long interactionId)
        {
            // Palette commits are already deferred by WordFontColorMonitor until
            // after WM_LBUTTONUP has returned to Word. Posting once to the Word UI
            // queue is therefore the completion boundary. A second fixed delay made
            // formula Graphics Fill visibly trail the surrounding text colour.
            _ = CommitFontColorInteractionOnWordUiAsync(interactionId);
        }

        private async Task CommitFontColorInteractionOnWordUiAsync(long interactionId)
        {
            try
            {
                await InvokeOnWordUiAsync(() =>
                {
                    if (shuttingDown || !hostEventProcessingEnabled ||
                        programmaticMutationDepth > 0 || Application.Documents.Count == 0)
                        return;
                    CommitFontColorInteraction(interactionId);
                }).ConfigureAwait(false);
            }
            catch (TaskCanceledException) { }
            catch (ObjectDisposedException) { }
            catch (Exception exception)
            {
                if (shuttingDown) return;
                try
                {
                    await InvokeOnWordUiAsync(() =>
                    {
                        pendingFontColorInteraction = null;
                        MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                            exception.GetBaseException().Message, "LaTeX Blocks",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }

        private PendingFontColorInteraction CaptureFontColorInteraction(long interactionId)
        {
            if (interactionId <= 0 || previousSelectionRangeLease == null ||
                !previousSelectionRangeLease.Matches(Application.Selection))
                return null;
            try
            {
                // An exact formula selection is an Office Graphic. Its colour is
                // owned by the native Graphics Fill command, not Font Color.
                if (Blocks.TryGetExactlySelectedInlineFormula(out _)) return null;
            }
            catch (COMException) { return null; }
            var lease = previousSelectionRangeLease.Clone();
            if (lease == null) return null;
            return new PendingFontColorInteraction(interactionId, lease,
                new List<SelectionFontSnapshot>(previousSelectionFontSnapshots));
        }

        private void CommitFontColorInteraction(long interactionId)
        {
            var interaction = pendingFontColorInteraction;
            pendingFontColorInteraction = null;
            if (interaction == null || interaction.InteractionId != interactionId ||
                !interaction.SelectionLease.Matches(Application.Selection))
                return;

            // For a normal text range Word has already written the chosen native
            // colour to each drawing character. Reconcile only the formulas that were
            // present when the interaction began; ordinary text remains entirely
            // owned by Word and no picker probe is needed.
            var targetTextColor = LaTeXBlockService.ResolveTextColor(
                Application.Selection);
            LaTeXBlockService.NativeTextColorDescriptor targetNativeTextColor;
            var hasTargetNativeTextColor = TryCaptureSelectionNativeTextColor(
                Application.Selection, out targetNativeTextColor);
            var requests = CaptureTextColorRefreshes(interaction.Formulas,
                interaction.SelectionLease, targetTextColor);
            RememberSelection(Application.Selection);
            TryApplyExternalColorRequests(requests, hasTargetNativeTextColor
                ? targetNativeTextColor
                : (LaTeXBlockService.NativeTextColorDescriptor?)null);
        }

        private static bool TryCaptureSelectionNativeTextColor(
            WordInterop.Selection selection,
            out LaTeXBlockService.NativeTextColorDescriptor descriptor)
        {
            descriptor = LaTeXBlockService.NativeTextColorDescriptor.Automatic;
            if (selection == null) return false;
            if (LaTeXBlockService.NativeTextColorDescriptor.TryCapture(
                    selection.Range, out descriptor))
                return true;

            WordInterop.Range insertion = null;
            try
            {
                insertion = selection.Range.Duplicate;
                insertion.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                return LaTeXBlockService.NativeTextColorDescriptor.TryCapture(
                    insertion, out descriptor);
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (insertion != null) Marshal.ReleaseComObject(insertion);
            }
        }

        private bool TryApplyExternalColorRequests(
            IList<FormatRefreshRequest> requests,
            LaTeXBlockService.NativeTextColorDescriptor? targetNativeTextColor = null)
        {
            if (shuttingDown || requests == null || requests.Count == 0)
                return false;
            var updates = new List<LaTeXBlockColorUpdate>(requests.Count);
            foreach (var request in requests)
            {
                var shape = request.Shape;
                if (shape == null || !request.ChangesTextColor ||
                    request.ChangesFontSize || request.ChangesWidth ||
                    !request.PreviousTextColor.HasValue ||
                    !LaTeXBlockService.TryReadContract(shape, out var metadata,
                        out var source) || source != request.Source ||
                    !SameMetadataState(metadata, request.Metadata))
                    return false;
                updates.Add(new LaTeXBlockColorUpdate(shape, request.TextColor,
                    targetNativeTextColor));
            }

            var applied = false;
            RunProgrammaticMutation(() =>
            {
                applied = Blocks.TryApplySvgForegroundFillsBatch(updates);
            });
            if (!applied) return false;
            foreach (var request in requests)
            {
                pendingFormatRefreshes.Remove(request.ShapeKey);
                pendingFormatBatchTargets.Remove(request.ShapeKey);
            }
            return true;
        }

        private void QueueFormatBatch(IList<FormatRefreshRequest> requests,
            SelectionRangeLease selectionLease)
        {
            if (shuttingDown || requests == null || requests.Count == 0) return;
            var sequence = Interlocked.Increment(ref formatBatchSequence);
            pendingFormatBatchTargets.Clear();
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var batch = new PendingFormatBatch(sequence, profile, selectionLease,
                new List<FormatRefreshRequest>(requests));
            foreach (var request in batch.Requests)
            {
                pendingFormatRefreshes.Remove(request.ShapeKey);
                pendingFormatBatchTargets[request.ShapeKey] =
                    new PendingFormatBatchTarget(sequence, request.TextColor,
                        request.FontSizePt);
            }
            _ = RefreshFormatBatchAsync(Blocks, batch);
        }

        private async Task RefreshFormatBatchAsync(LaTeXBlockService service,
            PendingFormatBatch batch)
        {
            try
            {
                // External colour-only requests are removed by QueueFormatRefresh
                // before a render batch is created. Every request here therefore
                // changes a genuine renderer input (normally TeX design size).
                var tasks = new List<Task<LaTeXBlockRender>>(batch.Requests.Count);
                foreach (var request in batch.Requests)
                    tasks.Add(service.RenderCommittedAsync(request.Source,
                        request.Metadata.WidthPt, request.Metadata.Mode, batch.Profile,
                        request.FontSizePt,
                        request.Metadata.Kind == LaTeXBlockKind.DisplayMath ||
                        request.Metadata.Kind == LaTeXBlockKind.NumberedMath,
                        request.TextColor, renderKind: request.Metadata.Kind));
                var renders = await Task.WhenAll(tasks).ConfigureAwait(false);
                await InvokeOnWordUiAsync(() => CompleteFormatBatch(service, batch,
                    renders)).ConfigureAwait(false);
            }
            catch (TaskCanceledException)
            {
                await AbandonFormatBatchAsync(batch, null).ConfigureAwait(false);
            }
            catch (ObjectDisposedException)
            {
                await AbandonFormatBatchAsync(batch, null).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                await AbandonFormatBatchAsync(batch, exception).ConfigureAwait(false);
            }
        }

        private void CompleteFormatBatch(LaTeXBlockService service,
            PendingFormatBatch batch, LaTeXBlockRender[] renders)
        {
            if (shuttingDown || batch.Sequence != formatBatchSequence) return;
            var restoreSelection = batch.SelectionLease != null &&
                batch.SelectionLease.Matches(Application.Selection);
            Exception firstFailure = null;
            var screenUpdating = true;
            try { screenUpdating = Application.ScreenUpdating; }
            catch (COMException) { }
            try
            {
                try { Application.ScreenUpdating = false; }
                catch (COMException) { }
                RunProgrammaticMutation(() =>
                {
                    var liveUpdates = new List<LaTeXBlockBatchUpdate>();
                    var canUseOpenXmlBatch = batch.Requests.Count > 1;
                    var canReplaceSvgMediaDirectly = batch.Requests.Count > 0;
                    int? batchParagraphStart = null;
                    int? batchParagraphEnd = null;
                    for (var index = 0; index < batch.Requests.Count; index++)
                    {
                        var request = batch.Requests[index];
                        try
                        {
                            var shape = request.Shape;
                            if (shape == null ||
                                !LaTeXBlockService.TryReadContract(shape,
                                    out var currentMetadata, out var currentSource) ||
                                currentSource != request.Source ||
                                !SameMetadataState(currentMetadata, request.Metadata) ||
                                !string.Equals(batch.Profile, currentProfile,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            var shapeRange = shape.Range;
                            var liveColor = LaTeXBlockService.ResolveTextColor(shapeRange);
                            var liveSize = (double)shapeRange.Font.Size;
                            if (!LaTeXBlockService.TextColorsEqual(liveColor,
                                    request.TextColor) || liveSize < 1 || liveSize > 200 ||
                                Math.Abs(liveSize - request.FontSizePt) >= 0.001)
                                continue;
                            var paragraph = shapeRange.Paragraphs[1].Range;
                            var autoInline = LaTeXBlockService.CanShareAutoInlineFormatBatch(
                                currentMetadata, request.ChangesWidth);
                            if (!autoInline ||
                                (batchParagraphStart.HasValue &&
                                 (batchParagraphStart.Value != paragraph.Start ||
                                  batchParagraphEnd.Value != paragraph.End)))
                                canUseOpenXmlBatch = false;
                            else if (!batchParagraphStart.HasValue)
                            {
                                batchParagraphStart = paragraph.Start;
                                batchParagraphEnd = paragraph.End;
                            }
                            liveUpdates.Add(new LaTeXBlockBatchUpdate(shape,
                                request.Source, request.Metadata.WidthPt,
                                renders[index], currentMetadata, shapeRange,
                                paragraph.Start, paragraph.End));
                            if (!autoInline || !request.ChangesFontSize ||
                                request.ChangesTextColor)
                                canReplaceSvgMediaDirectly = false;
                        }
                        catch (Exception exception)
                        {
                            if (firstFailure == null) firstFailure = exception;
                        }
                    }
                    if (liveUpdates.Count != batch.Requests.Count)
                    {
                        canUseOpenXmlBatch = false;
                        canReplaceSvgMediaDirectly = false;
                    }
                    if (canReplaceSvgMediaDirectly) canUseOpenXmlBatch = true;
                    try
                    {
                        if (canUseOpenXmlBatch)
                            service.UpdateRenderedBatch(liveUpdates,
                                canReplaceSvgMediaDirectly);
                        else
                            for (var index = 0; index < liveUpdates.Count; index++)
                            {
                                var update = liveUpdates[index];
                                service.UpdateRendered(update.Shape, update.Source,
                                    update.WidthPt, LaTeXBlockLayoutMode.Auto,
                                    update.Render, false);
                            }
                    }
                    catch (Exception exception)
                    {
                        if (firstFailure == null) firstFailure = exception;
                    }
                    if (restoreSelection) batch.SelectionLease.TryRestore(Application);
                }, restoreSelection);
            }
            finally
            {
                try { Application.ScreenUpdating = screenUpdating; }
                catch (COMException) { }
            }
            ClearFormatBatchTargets(batch);
            ribbon?.InvalidateWidthControl();
            if (firstFailure != null)
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                    firstFailure.GetBaseException().Message, "LaTeX Blocks",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private Task AbandonFormatBatchAsync(PendingFormatBatch batch,
            Exception exception)
        {
            if (shuttingDown) return Task.FromResult(false);
            return InvokeOnWordUiAsync(() =>
            {
                if (batch.Sequence != formatBatchSequence) return;
                ClearFormatBatchTargets(batch);
                if (exception != null)
                    MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                        exception.GetBaseException().Message, "LaTeX Blocks",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }

        private void ClearFormatBatchTargets(PendingFormatBatch batch)
        {
            foreach (var request in batch.Requests)
                if (pendingFormatBatchTargets.TryGetValue(request.ShapeKey,
                        out var target) && target.Sequence == batch.Sequence)
                    pendingFormatBatchTargets.Remove(request.ShapeKey);
        }

        private bool IsBlockFrameSnapshotStillSelected(BlockFrameSnapshot snapshot)
        {
            if (snapshot == null) return false;
            try
            {
                if (snapshot.IsFloating)
                    return Blocks.TryGetSelectedFloatingBlock(out var selectedFloating, out _) &&
                        GetComIdentity(selectedFloating) == snapshot.ShapeKey;
                return Blocks.TryGetSelectedBlock(out var selectedInline, out _) &&
                    GetComIdentity(selectedInline) == snapshot.ShapeKey;
            }
            catch (COMException) { return false; }
        }

        internal void ShowInsertFormulaEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            var textColor = LaTeXBlockService.ResolveTextColor(Application.Selection);
            using (var editor = new LaTeXBlockEditorForm(Blocks, "E=mc^2", 360,
                LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt, "Insert Inline Math", false, textColor, null, null, null,
                false, LaTeXBlockKind.InlineMath))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() => InsertMathFromEditor(editor));
                }
            }
        }

        internal void ShowInsertDisplayMathEditor()
        {
            if (Application.Documents.Count == 0)
                throw new InvalidOperationException("Open a Word document first.");
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            var textColor = LaTeXBlockService.ResolveTextColor(Application.Selection);
            using (var editor = new LaTeXBlockEditorForm(Blocks, "E=mc^2", 360,
                LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt, "Insert Display Math", true, textColor, null, null, null,
                false, LaTeXBlockKind.DisplayMath))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) ==
                    DialogResult.OK)
                {
                    RunProgrammaticMutation(() => InsertMathFromEditor(editor));
                }
            }
        }

        internal void ShowInsertBlockEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var widthPt = LaTeXBlockWidthPolicy.ResolveDefaultFixedWidth();
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Fixed, 10);
            var textColor = LaTeXBlockService.ResolveTextColor(Application.Selection);
            using (var editor = new LaTeXBlockEditorForm(Blocks,
                "This is a LaTeX Block with $E=mc^2$.", widthPt,
                LaTeXBlockLayoutMode.Fixed,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt,
                null, false, textColor, null, null, null, true,
                LaTeXBlockKind.LaTeXBlock))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                        Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode,
                            editor.CurrentRender, editor.AcceptedStyle,
                            LaTeXBlockKind.LaTeXBlock));
                }
            }
        }

        internal void ShowInsertNumberedEquationEditor()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            LaTeXBlockService.ValidateNumberedEquationTarget(Application.Selection.Range);
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            var textColor = LaTeXBlockService.ResolveTextColor(Application.Selection);
            const double widthPt = 360;
            using (var editor = new LaTeXBlockEditorForm(Blocks, "E=mc^2", widthPt,
                LaTeXBlockLayoutMode.Auto,
                currentProfile ?? Renderers.DefaultAvailableProfile, SetCurrentProfile, false,
                fontSizePt, "Insert Numbered Math", true, textColor, null, null, null,
                false, LaTeXBlockKind.NumberedMath))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() => InsertMathFromEditor(editor));
                }
            }
        }

        private void InsertMathFromEditor(LaTeXBlockEditorForm editor)
        {
            if (editor.Kind == LaTeXBlockKind.NumberedMath)
                Blocks.InsertNumberedRendered(editor.Source, editor.WidthPt, editor.Mode,
                    editor.CurrentRender);
            else
                Blocks.InsertRendered(editor.Source, editor.WidthPt, editor.Mode,
                    editor.CurrentRender, null, editor.Kind);
        }

        internal void ShowInsertEquationReference()
        {
            if (Application.Documents.Count == 0) throw new InvalidOperationException("Open a Word document first.");
            var references = Blocks.GetEquationReferenceTargets(Application.ActiveDocument);
            if (references.Count == 0)
                throw new InvalidOperationException(
                    "This document has no intact numbered LaTeX equations to reference.");

            using (var picker = new EquationReferenceForm(references))
            {
                if (picker.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) != DialogResult.OK)
                    return;
                var reference = picker.SelectedReference;
                if (reference == null) return;
                RunProgrammaticMutation(() => Blocks.InsertEquationReference(reference));
                Application.StatusBar = "Inserted equation reference (" + reference.Number + ").";
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

        internal void CopySelectionAsLaTeX()
        {
            if (Application.Documents.Count == 0)
                throw new InvalidOperationException("Open a Word document first.");
            var latex = WordSelectionLaTeXExporter.Export(Application.Selection.Range.Duplicate);
            if (string.IsNullOrEmpty(latex))
                throw new InvalidOperationException("The selection does not contain exportable text or LaTeX Blocks.");
            Clipboard.SetText(latex, TextDataFormat.UnicodeText);
            Application.StatusBar = "Copied " + latex.Length + " characters as LaTeX.";
        }

        internal void PasteFromLaTeX()
        {
            if (Application.Documents.Count == 0)
                throw new InvalidOperationException("Open a Word document first.");
            if (!Clipboard.ContainsText(TextDataFormat.UnicodeText) &&
                !Clipboard.ContainsText(TextDataFormat.Text))
                throw new InvalidOperationException("The clipboard does not contain LaTeX text.");

            var source = Clipboard.GetText(TextDataFormat.UnicodeText);
            if (string.IsNullOrEmpty(source)) source = Clipboard.GetText(TextDataFormat.Text);
            var segments = LaTeXMixedContentParser.Parse(source);
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var fontSizePt = LaTeXBlockService.ResolveFontSize(Application.Selection,
                LaTeXBlockLayoutMode.Auto, 10);
            var textColor = LaTeXBlockService.ResolveTextColor(Application.Selection);
            var baseTextFormat = WordTextFormatSnapshot.Capture(Application.Selection);
            var prepared = new List<PreparedLaTeXImportSegment>();
            foreach (var segment in segments)
            {
                if (segment.Kind == LaTeXContentKind.Text)
                {
                    prepared.Add(new PreparedLaTeXImportSegment(segment, null, 0));
                    continue;
                }
                // Display math is still a natural-size formula.  Its independent
                // line and paragraph placement belong to Word; only an explicit
                // Block owns a fixed outer frame.
                var mode = LaTeXBlockService.ResolveImportedFormulaMode(segment.Kind);
                const double width = 360;
                var render = Blocks.RenderPreview(segment.Source, width, mode, profile,
                    fontSizePt, segment.Kind == LaTeXContentKind.DisplayMath, textColor);
                prepared.Add(new PreparedLaTeXImportSegment(segment, render, width));
            }

            RunProgrammaticMutation(() =>
            {
                var target = Application.Selection.Range.Duplicate;
                target.Text = string.Empty;
                target.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
                target.Select();
                foreach (var item in prepared)
                {
                    if (item.Segment.Kind == LaTeXContentKind.Text)
                    {
                        baseTextFormat.Apply(Application.Selection);
                        var textStart = Application.Selection.Start;
                        Application.Selection.TypeText(
                            LaTeXMixedContentParser.ToWordText(item.Segment.Source));
                        var textEnd = Application.Selection.Start;
                        if (textEnd > textStart &&
                            (item.Segment.Bold || item.Segment.Italic ||
                             item.Segment.FontFamily != LaTeXTextFontFamily.Inherited))
                        {
                            var insertedText = Application.ActiveDocument.Range(textStart, textEnd);
                            if (item.Segment.Bold) insertedText.Font.Bold = -1;
                            if (item.Segment.Italic) insertedText.Font.Italic = -1;
                            var fonts = ResolveImportedTextFonts(item.Segment.FontFamily, profile);
                            if (!string.IsNullOrEmpty(fonts.Western))
                            {
                                insertedText.Font.Name = fonts.Western;
                                insertedText.Font.NameAscii = fonts.Western;
                            }
                            if (!string.IsNullOrEmpty(fonts.FarEast))
                                insertedText.Font.NameFarEast = fonts.FarEast;
                            // Formatting an inserted Range also changes Word's live
                            // insertion format at its trailing edge. Restore the
                            // caller's format so a scoped LaTeX command cannot leak
                            // into the following plain-text segment.
                            baseTextFormat.Apply(Application.Selection);
                        }
                        continue;
                    }
                    var mode = LaTeXBlockService.ResolveImportedFormulaMode(item.Segment.Kind);
                    // Restore ordinary insertion formatting before each drawing so
                    // formula-owned run properties cannot leak into following text.
                    // The drawing baseline itself depends only on its TeX depth.
                    baseTextFormat.Apply(Application.Selection);
                    Blocks.InsertRendered(item.Segment.Source, item.WidthPt,
                        mode, item.Render, null,
                        item.Segment.Kind == LaTeXContentKind.DisplayMath
                            ? LaTeXBlockKind.DisplayMath
                            : LaTeXBlockKind.InlineMath);
                }
            });
            Application.StatusBar = "Pasted LaTeX text with " +
                prepared.FindAll(item => item.Render != null).Count + " formula Blocks.";
        }

        private sealed class PreparedLaTeXImportSegment
        {
            internal PreparedLaTeXImportSegment(LaTeXContentSegment segment,
                LaTeXBlockRender render, double widthPt)
            {
                Segment = segment;
                Render = render;
                WidthPt = widthPt;
            }
            internal LaTeXContentSegment Segment { get; }
            internal LaTeXBlockRender Render { get; }
            internal double WidthPt { get; }
        }

        private static ImportedTextFonts ResolveImportedTextFonts(LaTeXTextFontFamily family,
            string profile)
        {
            if (family == LaTeXTextFontFamily.Inherited) return default(ImportedTextFonts);
            var arialCjk = (profile ?? string.Empty).IndexOf("arial_lete_simhei",
                StringComparison.OrdinalIgnoreCase) >= 0;
            var cjk = arialCjk || (profile ?? string.Empty).IndexOf("cjk",
                StringComparison.OrdinalIgnoreCase) >= 0;
            switch (family)
            {
                case LaTeXTextFontFamily.SansSerif:
                    return new ImportedTextFonts("Arial", cjk ? "SimHei" : null);
                case LaTeXTextFontFamily.Monospace:
                    return new ImportedTextFonts("Consolas", cjk ? (arialCjk ? "SimHei" : "SimSun") : null);
                default:
                    return new ImportedTextFonts(arialCjk ? "Arial" : "Times New Roman",
                        cjk ? (arialCjk ? "SimHei" : "SimSun") : null);
            }
        }

        private struct ImportedTextFonts
        {
            internal ImportedTextFonts(string western, string farEast)
            {
                Western = western;
                FarEast = farEast;
            }
            internal string Western { get; }
            internal string FarEast { get; }
        }

        private sealed class WordTextFormatSnapshot
        {
            private string name;
            private string nameAscii;
            private string nameFarEast;
            private int bold;
            private int italic;
            private int position;
            private float size;
            private WordInterop.WdColor color;

            internal static WordTextFormatSnapshot Capture(WordInterop.Selection selection)
            {
                return new WordTextFormatSnapshot
                {
                    name = selection.Font.Name,
                    nameAscii = selection.Font.NameAscii,
                    nameFarEast = selection.Font.NameFarEast,
                    bold = selection.Font.Bold,
                    italic = selection.Font.Italic,
                    position = selection.Font.Position,
                    size = selection.Font.Size,
                    color = selection.Font.Color
                };
            }

            internal void Apply(WordInterop.Selection selection)
            {
                if (!string.IsNullOrEmpty(name)) selection.Font.Name = name;
                if (!string.IsNullOrEmpty(nameAscii)) selection.Font.NameAscii = nameAscii;
                if (!string.IsNullOrEmpty(nameFarEast)) selection.Font.NameFarEast = nameFarEast;
                selection.Font.Bold = bold;
                selection.Font.Italic = italic;
                selection.Font.Position = position;
                if (size > 0) selection.Font.Size = size;
                selection.Font.Color = color;
            }

            internal int Position => position;
        }

        internal void ShowEditEditor()
        {
            WordInterop.InlineShape inlineShape = null;
            WordInterop.Shape floatingShape = null;
            LaTeXBlockMetadata metadata;
            string source;
            int textColor;
            LaTeXBlockStyle style;
            double? outerWidthPt;
            double? outerHeightPt;
            if (Blocks.TryGetSelectedBlock(out inlineShape, out metadata))
            {
                source = LaTeXBlockMetadata.ReadSource(inlineShape.AlternativeText);
                textColor = LaTeXBlockService.ResolveTextColor(inlineShape.Range);
                style = metadata.HasExplicitStyle ? metadata.Style : null;
                outerWidthPt = inlineShape.Width;
                outerHeightPt = inlineShape.Height;
            }
            else if (Blocks.TryGetSelectedFloatingBlock(out floatingShape, out metadata))
            {
                if (metadata.Role != LaTeXBlockRole.Content || metadata.Mode != LaTeXBlockLayoutMode.Fixed)
                    throw new InvalidOperationException(
                        "Only fixed-width LaTeX Blocks can remain floating. Keep inline formulas and numbered equations \"In Line with Text\".");
                source = LaTeXBlockMetadata.ReadSource(floatingShape.AlternativeText);
                textColor = LaTeXBlockService.ResolveTextColor(floatingShape.Anchor);
                style = metadata.HasExplicitStyle ? metadata.Style : null;
                outerWidthPt = floatingShape.Width;
                outerHeightPt = floatingShape.Height;
            }
            else
            {
                throw new InvalidOperationException("Select a LaTeX Block first.");
            }

            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt,
                metadata.Mode, currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, true, metadata.FontSizePt, null,
                metadata.Kind == LaTeXBlockKind.DisplayMath ||
                metadata.Kind == LaTeXBlockKind.NumberedMath, textColor, style,
                outerHeightPt, outerWidthPt, floatingShape != null, metadata.Kind))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.WordWindow(Application)) == DialogResult.OK)
                {
                    RunProgrammaticMutation(() =>
                    {
                        var acceptedStyle = editor.Kind == LaTeXBlockKind.LaTeXBlock
                            ? editor.AcceptedStyle
                            : null;
                        if (inlineShape != null)
                        {
                            if (metadata.Kind != LaTeXBlockKind.LaTeXBlock &&
                                editor.Kind != metadata.Kind)
                            {
                                Blocks.ConvertMathRendered(inlineShape, editor.Source,
                                    editor.WidthPt, editor.CurrentRender, editor.Kind);
                                return;
                            }
                            // In Line with Text is only a Word layout choice. A fixed
                            // Content Block can still have an author-sized outer frame,
                            // so editing TeX must preserve that frame. Styled previews
                            // already used this exact frame for their TeX content box;
                            // only a legacy unstyled result needs SVG reframing here.
                            var inlineRender = metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                                metadata.Role == LaTeXBlockRole.Content &&
                                acceptedStyle == null
                                ? Blocks.FrameFloatingRender(editor.CurrentRender,
                                    inlineShape.Width, inlineShape.Height)
                                : editor.CurrentRender;
                            Blocks.UpdateRendered(inlineShape, editor.Source, editor.WidthPt,
                                editor.Mode, inlineRender, true, acceptedStyle, editor.Kind);
                        }
                        else
                        {
                            // A floating Block owns an outer Word frame.  Editing its
                            // TeX changes the inner layout, not the frame the user
                            // placed and sized. The editor preview already rendered
                            // against that exact frame and composed its SVG shell once.
                            var framedRender = editor.CurrentRender;
                            Blocks.UpdateFloatingRendered(floatingShape, editor.Source,
                                editor.WidthPt, editor.Mode, framedRender, true,
                                acceptedStyle);
                        }
                    });
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
                throw new ArgumentException("Enter a typesetting width from 30 to 2000 pt.",
                    nameof(text));
            if (!TryGetSelectedFixedContentBlock(out var selectedShape, out var metadata))
                throw new InvalidOperationException("Select one fixed-width LaTeX Block first.");
            var source = LaTeXBlockMetadata.ReadSource(selectedShape.AlternativeText);
            QueueFormatRefresh(new List<FormatRefreshRequest>
            {
                new FormatRefreshRequest(selectedShape, source, metadata, requestedWidthPt,
                    metadata.FontSizePt, LaTeXBlockService.ResolveTextColor(selectedShape.Range))
            });
        }

        internal bool HasSelectedBlockFrame()
        {
            return CaptureBlockFrameSnapshot() != null;
        }

        internal void ReflowSelectedBlockFrame()
        {
            var request = CreateBlockFrameReflowRequest(CaptureBlockFrameSnapshot(), true, true);
            if (request == null)
                throw new InvalidOperationException(
                    "Select one fixed-width LaTeX Block first.");
            QueueBlockFrameReflow(request);
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

        private bool TryGetSelectedFloatingFixedContentBlock(out WordInterop.Shape shape,
            out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            if (!Blocks.TryGetSelectedFloatingBlock(out var candidate, out var candidateMetadata) ||
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
            var selectionLease = SelectionRangeLease.Capture(selection);
            QueueFormatRefresh(CaptureAutoBlocks(selection.Range, fontSizePt, false,
                selectionLease));
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
                // CommandBarComboBox.Change can arrive before Word has propagated
                // the chosen size through a Ctrl+A range. Capture the committed UI
                // value now, but reconcile the selection after Word's native command
                // has finished. Reading Selection.Font.Size synchronously here can
                // otherwise return the old or mixed value and suppress the batch.
                double parsedFontSizePt;
                var requestedFontSizePt = control != null &&
                    double.TryParse(control.Text, NumberStyles.Float,
                        CultureInfo.CurrentCulture, out parsedFontSizePt) &&
                    parsedFontSizePt >= 1 && parsedFontSizePt <= 200
                        ? (double?)parsedFontSizePt
                        : null;
                _ = CommitFontSizeInteractionAfterHostAsync(requestedFontSizePt);
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message,
                    "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                refreshingNativeFontSize = false;
                ribbon?.InvalidateWidthControl();
            }
        }

        private void RefreshSelectedFontSize(double? requestedFontSizePt = null)
        {
            if (shuttingDown || !hostEventProcessingEnabled ||
                Application.Documents.Count == 0)
                return;
            var selection = Application.Selection;
            var size = requestedFontSizePt ?? (double)selection.Font.Size;
            if (size < 1 || size > 200) return;
            var selectionLease = SelectionRangeLease.Capture(selection);
            var now = DateTime.UtcNow.Ticks;
            if (lastFontSizeCommitLease != null &&
                now - lastFontSizeCommitUtcTicks <=
                    TimeSpan.FromMilliseconds(500).Ticks &&
                Math.Abs(lastFontSizeCommitPt - size) < 0.001 &&
                lastFontSizeCommitLease.Matches(selection))
                return;
            lastFontSizeCommitUtcTicks = now;
            lastFontSizeCommitPt = size;
            lastFontSizeCommitLease = selectionLease;
            QueueFormatRefresh(CaptureAutoBlocks(selection.Range, size, true,
                selectionLease));
            RememberSelection(selection);
        }

        private void Application_WindowSelectionChange(WordInterop.Selection selection)
        {
            if (shuttingDown || !hostEventProcessingEnabled || refreshingNativeFontSize ||
                programmaticMutationDepth > 0)
                return;

            try
            {
                refreshingNativeFontSize = true;
                // A native format transaction is bound to the selection that was
                // active when its UI opened. A genuine host selection change makes
                // that transaction stale; programmatic replacements are filtered by
                // programmaticMutationDepth before reaching this point.
                pendingFontColorInteraction = null;
                // The process-scoped mouse-capture hook normally commits at mouse-up.
                // Keep selection change as a no-polling fallback for locked-down
                // desktops where Windows refuses the hook, and for a Block changed
                // through a non-mouse Office command.  Geometry alone matters;
                // translation and rotation deliberately do not request a render.
                QueueBlockFrameReflow(
                    CreateBlockFrameReflowRequest(previousBlockFrameSnapshot, false, false));
                // Word exposes no general formatting-changed event. If native size or
                // color was applied through a shortcut, palette, style, paste, or macro,
                // validate the range when the user next leaves it. This is event driven,
                // not polling and does not rely on a particular Ribbon control.
                QueueFormatRefresh(CaptureBlocksWhoseHostFormatChanged(
                    previousSelectionFontSnapshots));
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message,
                    "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    RememberSelection(Application.Selection);
                }
                finally
                {
                    refreshingNativeFontSize = false;
                }
                ribbon?.InvalidateWidthControl();
            }
        }

        private List<FormatRefreshRequest> CaptureAutoBlocks(WordInterop.Range range,
            double fontSizePt, bool onlyChanged,
            SelectionRangeLease selectionLease = null)
        {
            var requests = new List<FormatRefreshRequest>();
            foreach (WordInterop.InlineShape shape in range.InlineShapes)
            {
                if (LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) &&
                    metadata.Mode == LaTeXBlockLayoutMode.Auto &&
                    (!onlyChanged || Math.Abs(metadata.FontSizePt - fontSizePt) > 0.001 ||
                     HasPendingFormatTarget(shape, fontSizePt,
                         LaTeXBlockService.ResolveTextColor(shape.Range))))
                    requests.Add(new FormatRefreshRequest(shape, source, metadata, fontSizePt,
                        LaTeXBlockService.ResolveTextColor(shape.Range), true, false,
                        selectionLease));
            }
            return requests;
        }

        private List<FormatRefreshRequest> CaptureBlocksWhoseHostFormatChanged(
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
                        !LaTeXBlockService.TryReadContract(shape, out var metadata, out var source)) continue;
                    var size = metadata.Mode == LaTeXBlockLayoutMode.Auto
                        ? (double)shape.Range.Font.Size
                        : metadata.FontSizePt;
                    var textColor = LaTeXBlockService.ResolveTextColor(shape.Range);
                    var hostFormatChanged =
                        LaTeXBlockService.TryClassifyHostFormatChange(metadata.Mode,
                            snapshot.HostFontSizePt, snapshot.HostTextColor, size,
                            textColor, metadata.FontSizePt, out var fontSizeChanged,
                            out var textColorChanged);
                    if (!hostFormatChanged &&
                        !HasPendingFormatTarget(shape, size, textColor)) continue;
                    requests.Add(new FormatRefreshRequest(shape, source, metadata, size,
                        textColor, fontSizeChanged, textColorChanged,
                        previousTextColor: snapshot.HostTextColor));
                }
                catch (COMException)
                {
                    // The block may have been deleted or moved across a story while
                    // the selection changed. Size refresh is opportunistic.
                }
            }
            return requests;
        }

        private List<FormatRefreshRequest> CaptureTextColorRefreshes(
            IList<SelectionFontSnapshot> snapshots, SelectionRangeLease selectionLease,
            int targetTextColor)
        {
            var requests = new List<FormatRefreshRequest>();
            if (snapshots == null || snapshots.Count == 0 || selectionLease == null)
                return requests;
            foreach (var snapshot in snapshots)
            {
                try
                {
                    var shape = snapshot.Shape;
                    if (shape == null ||
                        !LaTeXBlockService.TryReadContract(shape, out var metadata,
                            out var source))
                        continue;
                    var fontSizePt = metadata.Mode == LaTeXBlockLayoutMode.Auto
                        ? (double)shape.Range.Font.Size
                        : metadata.FontSizePt;
                    if (fontSizePt < 1 || fontSizePt > 200) continue;
                    requests.Add(new FormatRefreshRequest(shape, source, metadata,
                        fontSizePt, targetTextColor, false, true, selectionLease,
                        snapshot.HostTextColor));
                }
                catch (COMException)
                {
                    // The selected formula may have been deleted while the native
                    // palette was open. Never redirect the transaction to a new object.
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
                        metadata.Kind == LaTeXBlockKind.LaTeXBlock)
                        continue;
                    var size = metadata.Mode == LaTeXBlockLayoutMode.Auto
                        ? (double)shape.Range.Font.Size
                        : metadata.FontSizePt;
                    if (size >= 1 && size <= 200)
                        snapshots.Add(new SelectionFontSnapshot(shape, size,
                            LaTeXBlockService.ResolveTextColor(shape.Range)));
                }
            }
            catch (COMException) { }
            return snapshots;
        }

        private BlockFrameSnapshot CaptureBlockFrameSnapshot()
        {
            try
            {
                // Fixed Content Blocks are one semantic object whether Word exposes
                // them as an InlineShape (In Line with Text) or a floating Shape
                // under any WrapFormat.  Do not let layout participation decide
                // whether their native resize is treated as a frame request.
                if (TryGetSelectedFixedContentBlock(out var inlineShape, out var inlineMetadata))
                    return CaptureBlockFrameSnapshot(inlineShape, null, inlineMetadata,
                        LaTeXBlockMetadata.ReadSource(inlineShape.AlternativeText));
                if (TryGetSelectedFloatingFixedContentBlock(out var floatingShape,
                        out var floatingMetadata))
                    return CaptureBlockFrameSnapshot(null, floatingShape, floatingMetadata,
                        LaTeXBlockMetadata.ReadSource(floatingShape.AlternativeText));
            }
            catch (COMException)
            {
                // The selected object can be deleted while Word is raising a native
                // input or selection event.  There is simply no frame to preserve.
            }
            return null;
        }

        private BlockFrameSnapshot CaptureBlockFrameSnapshot(WordInterop.InlineShape inlineShape,
            WordInterop.Shape floatingShape, LaTeXBlockMetadata metadata, string source)
        {
            if ((inlineShape == null && floatingShape == null) || metadata == null) return null;
            var shape = (object)inlineShape ?? floatingShape;
            var shapeKey = GetComIdentity(shape);
            if (shapeKey == 0) return null;
            var frameWidthPt = NormalizeFrameExtent(inlineShape != null
                ? inlineShape.Width : floatingShape.Width);
            var frameHeightPt = NormalizeFrameExtent(inlineShape != null
                ? inlineShape.Height : floatingShape.Height);
            return new BlockFrameSnapshot(inlineShape, floatingShape, shapeKey, metadata, source,
                frameWidthPt, frameHeightPt, ResolveExpectedBlockLayoutWidth(shapeKey, metadata,
                    frameWidthPt));
        }

        // A resize may finish while an earlier render is still in flight.  Treat the
        // pending target as the current TeX measure and carry only the newest native
        // width delta forward.  This makes two rapid drags compose instead of making
        // the second one start again from stale document metadata.
        private double ResolveExpectedBlockLayoutWidth(long shapeKey,
            LaTeXBlockMetadata metadata, double observedFrameWidthPt)
        {
            if (metadata == null) return LaTeXBlockWidthPolicy.MinimumWidthPt;
            if (shapeKey != 0 && pendingBlockFrameReflows.TryGetValue(shapeKey,
                    out var pending) && SameBlockFrameState(pending.BaseMetadata, metadata))
                return ClampBlockLayoutWidth(pending.TargetWidthPt + observedFrameWidthPt -
                    pending.TargetFrameWidthPt);
            return metadata.WidthPt;
        }

        private static double NormalizeFrameExtent(double extentPt)
        {
            // This is no longer an upper clamp.  Word owns the physical frame; the
            // service only supplies a positive finite SVG viewport for it.
            return LaTeXBlockService.ClampFloatingFrameExtent(extentPt);
        }

        private BlockFrameReflowRequest CreateBlockFrameReflowRequest(
            BlockFrameSnapshot snapshot, bool force, bool restoreSelection)
        {
            if (snapshot == null) return null;
            try
            {
                if (snapshot.IsFloating)
                {
                    if (!LaTeXBlockService.TryReadContract(snapshot.FloatingShape, out var metadata,
                            out var source) ||
                        !SameBlockFrameState(metadata, snapshot.Metadata) ||
                        metadata.Mode != LaTeXBlockLayoutMode.Fixed ||
                        metadata.Role != LaTeXBlockRole.Content ||
                        !string.Equals(LaTeXBlockService.NormalizeSourceText(source),
                            snapshot.Source, StringComparison.Ordinal))
                        return null;
                    return CreateBlockFrameReflowRequest(null, snapshot.FloatingShape,
                        snapshot.ShapeKey, metadata, source, force, restoreSelection,
                        snapshot.LayoutWidthPt, snapshot.FrameWidthPt, snapshot.FrameHeightPt,
                        LaTeXBlockService.ResolveTextColor(snapshot.FloatingShape.Anchor));
                }

                if (!LaTeXBlockService.TryReadContract(snapshot.InlineShape, out var inlineMetadata,
                        out var inlineSource) ||
                    !SameBlockFrameState(inlineMetadata, snapshot.Metadata) ||
                    inlineMetadata.Mode != LaTeXBlockLayoutMode.Fixed ||
                    inlineMetadata.Role != LaTeXBlockRole.Content ||
                    !string.Equals(LaTeXBlockService.NormalizeSourceText(inlineSource),
                        snapshot.Source, StringComparison.Ordinal))
                    return null;
                return CreateBlockFrameReflowRequest(snapshot.InlineShape, null,
                    snapshot.ShapeKey, inlineMetadata, inlineSource, force, restoreSelection,
                    snapshot.LayoutWidthPt, snapshot.FrameWidthPt, snapshot.FrameHeightPt,
                    LaTeXBlockService.ResolveTextColor(snapshot.InlineShape.Range));
            }
            catch (COMException) { return null; }
        }

        private BlockFrameReflowRequest CreateBlockFrameReflowRequest(
            WordInterop.InlineShape inlineShape, WordInterop.Shape floatingShape,
            long expectedShapeKey, LaTeXBlockMetadata metadata, string source, bool force,
            bool restoreSelection, double previousLayoutWidthPt, double previousFrameWidthPt,
            double previousFrameHeightPt, int textColor)
        {
            if ((inlineShape == null && floatingShape == null) ||
                (inlineShape != null && floatingShape != null) || metadata == null)
                return null;
            var shapeKey = GetComIdentity((object)inlineShape ?? floatingShape);
            if (shapeKey == 0 || shapeKey != expectedShapeKey) return null;
            var targetFrameWidthPt = NormalizeFrameExtent(
                inlineShape != null ? inlineShape.Width : floatingShape.Width);
            var targetFrameHeightPt = NormalizeFrameExtent(
                inlineShape != null ? inlineShape.Height : floatingShape.Height);
            // Compare only with the geometry captured before this gesture.  Metadata
            // describes the last committed SVG, and is deliberately not used here:
            // while TeX renders, a move/rotation must not look like another resize.
            var widthChanged = !FrameExtentsEqual(targetFrameWidthPt, previousFrameWidthPt);
            if (!force && !LaTeXBlockService.HasNativeFrameGeometryChanged(
                    previousFrameWidthPt, previousFrameHeightPt, targetFrameWidthPt,
                    targetFrameHeightPt))
                return null;

            // Keep the measured SVG edge allowance additive rather than treating
            // the host change as a scale factor. A height-only resize still gets a
            // fresh SVG root at the current TeX measure; it clips or supplies space
            // vertically according to the user's frame rather than scaling glyphs.
            var targetWidthPt = widthChanged
                ? LaTeXBlockService.ComposeNativeFrameLayoutWidth(previousLayoutWidthPt,
                    previousFrameWidthPt, targetFrameWidthPt)
                : previousLayoutWidthPt;
            return new BlockFrameReflowRequest(inlineShape, floatingShape, shapeKey, metadata, source,
                targetWidthPt, targetFrameWidthPt, targetFrameHeightPt, textColor,
                restoreSelection);
        }

        private static bool FrameExtentsEqual(double left, double right)
        {
            return !LaTeXBlockService.HasNativeFrameGeometryChanged(left, 0, right, 0);
        }

        private static double ClampBlockLayoutWidth(double widthPt)
        {
            if (double.IsNaN(widthPt) || double.IsInfinity(widthPt))
                return LaTeXBlockWidthPolicy.MinimumWidthPt;
            // Native outer frames can be wider than the editor's historical default
            // range.  Their TeX measure must remain large enough to describe that
            // user-owned geometry; the block editor receives the same policy limit.
            return Math.Max(LaTeXBlockWidthPolicy.MinimumWidthPt,
                Math.Min(LaTeXBlockWidthPolicy.MaximumWidthPt, widthPt));
        }

        private void QueueFormatRefresh(List<FormatRefreshRequest> requests)
        {
            if (shuttingDown || requests == null || requests.Count == 0) return;
            var autoInlineRequests = new List<FormatRefreshRequest>();
            var externalColorRequests = new List<FormatRefreshRequest>();
            var individualRequests = new List<FormatRefreshRequest>();
            SelectionRangeLease autoInlineSelectionLease = null;
            foreach (var request in requests)
            {
                var externalColorOnly = request.ChangesTextColor &&
                    !request.ChangesFontSize && !request.ChangesWidth &&
                    request.PreviousTextColor.HasValue;
                if (externalColorOnly)
                {
                    externalColorRequests.Add(request);
                }
                else if (LaTeXBlockService.CanShareAutoInlineFormatBatch(request.Metadata,
                        request.ChangesWidth))
                {
                    if (autoInlineRequests.Count == 0)
                        autoInlineSelectionLease = request.SelectionLease;
                    else if (!ReferenceEquals(autoInlineSelectionLease,
                                 request.SelectionLease))
                        autoInlineSelectionLease = null;
                    autoInlineRequests.Add(request);
                }
                else
                    individualRequests.Add(request);
            }
            // A Word format command is one visual operation even when it changes
            // several renderer inputs. Colour, font size, and a future combination
            // of the two therefore share one render/commit batch. A numbered formula
            // is still an Auto InlineShape; its SEQ field and bookmark live outside
            // the drawing run and survive the same media batch. Fixed Blocks alone
            // keep their independent physical-frame replacement contract.
            if (externalColorRequests.Count > 0)
                TryApplyExternalColorRequests(externalColorRequests);
            if (autoInlineRequests.Count > 0)
                QueueFormatBatch(autoInlineRequests, autoInlineSelectionLease);
            foreach (var request in individualRequests) QueueFormatRefresh(request);
        }

        private void QueueFormatRefresh(FormatRefreshRequest request)
        {
            if (shuttingDown || request == null || request.ShapeKey == 0) return;
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var targetWidthPt = request.Metadata.WidthPt;
            var targetFontSizePt = request.Metadata.FontSizePt;
            var targetTextColor = LaTeXBlockService.ResolveTextColor(request.Shape.Range);
            if (pendingFormatRefreshes.TryGetValue(request.ShapeKey, out var existing) &&
                SameBaseState(existing, request.Metadata, request.Source, profile))
            {
                targetWidthPt = existing.TargetWidthPt;
                targetFontSizePt = existing.TargetFontSizePt;
                targetTextColor = existing.TargetTextColor;
            }
            if (request.ChangesWidth)
                targetWidthPt = request.WidthPt;
            if (request.ChangesFontSize)
                targetFontSizePt = request.FontSizePt;
            if (request.ChangesTextColor)
                targetTextColor = request.TextColor;

            if (Math.Abs(targetWidthPt - request.Metadata.WidthPt) < 0.001 &&
                Math.Abs(targetFontSizePt - request.Metadata.FontSizePt) < 0.001 &&
                !request.ChangesTextColor)
            {
                pendingFormatRefreshes.Remove(request.ShapeKey);
                ribbon?.InvalidateWidthControl();
                return;
            }

            // Fixed Content owns a physical SVG viewport.  It must never take the
            // ordinary UpdateRendered path, because that path would replace a
            // user-resized frame with an unframed SVG while a color/width refresh is
            // in flight.  Route it through the same framed commit as native resize.
            if (IsFixedContentBlock(request.Metadata))
            {
                pendingFormatRefreshes.Remove(request.ShapeKey);
                QueueFixedContentFormatRefresh(request, targetWidthPt, targetTextColor);
                return;
            }

            var sequence = Interlocked.Increment(ref formatRefreshSequence);
            var pending = new PendingFormatRefresh(request.ShapeKey, request.Shape,
                request.Metadata, request.Source, profile, targetWidthPt,
                targetFontSizePt, targetTextColor, request.SelectionLease, sequence);
            pendingFormatRefreshes[request.ShapeKey] = pending;
            StartNextFormatRefresh(request.ShapeKey);
        }

        private void QueueFixedContentFormatRefresh(FormatRefreshRequest request,
            double targetWidthPt, int targetTextColor)
        {
            if (request?.Shape == null) return;
            try
            {
                if (!LaTeXBlockService.TryReadContract(request.Shape, out var metadata,
                        out var source) || !IsFixedContentBlock(metadata) ||
                    !SameBlockFrameState(metadata, request.Metadata))
                    return;
                var snapshot = CaptureBlockFrameSnapshot(request.Shape, null, metadata, source);
                // A native Font command can leave the Block selected while the
                // replacement SVG renders. The completion path verifies that the
                // same object is still selected before restoring the selection, so
                // carrying this intent cannot steal focus after the user moves on.
                var frameRequest = CreateBlockFrameReflowRequest(snapshot, true, true);
                if (frameRequest == null) return;
                var requestedLayoutWidthPt = request.ChangesWidth
                    ? ClampBlockLayoutWidth(targetWidthPt)
                    : frameRequest.TargetWidthPt;
                QueueBlockFrameReflow(frameRequest.WithFormat(requestedLayoutWidthPt,
                    targetTextColor));
            }
            catch (COMException)
            {
                // A native format command can race an Undo/Delete.  No stale frame
                // should be resurrected just to complete a cosmetic refresh.
            }
        }

        private static bool IsFixedContentBlock(LaTeXBlockMetadata metadata)
        {
            return metadata != null && metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                   metadata.Role == LaTeXBlockRole.Content;
        }

        private bool QueueBlockFrameReflow(BlockFrameReflowRequest request)
        {
            if (shuttingDown || request == null || !request.HasShape || request.ShapeKey == 0)
                return false;
            var source = LaTeXBlockService.NormalizeSourceText(request.Source);
            if (string.IsNullOrWhiteSpace(source)) return false;
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var key = request.ShapeKey;
            // An already-rendering normal format refresh is allowed to finish its
            // worker task, but no longer owns this COM object. Its UI completion
            // observes the missing pending entry and cannot write an unframed SVG.
            pendingFormatRefreshes.Remove(key);
            var restoreSelection = request.RestoreSelection;
            if (pendingBlockFrameReflows.TryGetValue(key, out var existing))
                restoreSelection = restoreSelection || existing.RestoreSelection;
            var pending = new PendingBlockFrameReflow(key, request.InlineShape,
                request.FloatingShape, request.Metadata, source, profile, request.TargetWidthPt,
                request.TargetFrameWidthPt, request.TargetFrameHeightPt, request.TextColor,
                restoreSelection, Interlocked.Increment(ref blockFrameReflowSequence));
            pendingBlockFrameReflows[key] = pending;
            RememberExpectedBlockFrame(request);
            StartNextBlockFrameReflow(key);
            return true;
        }

        private void RememberExpectedBlockFrame(BlockFrameReflowRequest request)
        {
            // Advance the gesture baseline as soon as the request is accepted, not
            // only after the renderer returns. A following resize therefore composes
            // from the latest intended TeX width; a move/rotation sees identical
            // pre/post dimensions and does nothing.
            previousBlockFrameSnapshot = new BlockFrameSnapshot(request.InlineShape,
                request.FloatingShape, request.ShapeKey, request.Metadata, request.Source,
                request.TargetFrameWidthPt, request.TargetFrameHeightPt,
                request.TargetWidthPt);
        }

        private void StartNextBlockFrameReflow(long key)
        {
            if (shuttingDown || blockFrameReflowsInFlight.Contains(key) ||
                !pendingBlockFrameReflows.TryGetValue(key, out var pending)) return;
            blockFrameReflowsInFlight.Add(key);
            _ = ReflowBlockFrameAsync(Blocks, pending);
        }

        private async Task ReflowBlockFrameAsync(LaTeXBlockService service,
            PendingBlockFrameReflow pending)
        {
            try
            {
                var style = pending.BaseMetadata.HasExplicitStyle
                    ? pending.BaseMetadata.Style
                    : null;
                var textColor = style != null
                    ? LaTeXBlockService.ToWordColor(style.TextColor)
                    : pending.TextColor;
                var render = await service.RenderCommittedAsync(pending.Source,
                    pending.TargetWidthPt, LaTeXBlockLayoutMode.Fixed, pending.Profile,
                    pending.BaseMetadata.FontSizePt, false,
                    textColor, style,
                    style != null ? pending.TargetFrameHeightPt : (double?)null,
                    style != null ? pending.TargetFrameWidthPt : (double?)null,
                    pending.BaseMetadata.Kind)
                    .ConfigureAwait(false);
                // A styled render has already received the exact outer frame, so TeX
                // made and aligned the corresponding (frame - 2*padding) content box.
                // The SVG shell was then added once. Legacy unstyled blocks still need
                // their transparent host viewport composed here.
                var framedRender = style != null
                    ? render
                    : service.FrameFloatingRender(render, pending.TargetFrameWidthPt,
                        pending.TargetFrameHeightPt);
                await InvokeOnWordUiAsync(() => CompleteBlockFrameReflow(service,
                    pending, framedRender)).ConfigureAwait(false);
            }
            catch (TaskCanceledException)
            {
                await AbandonBlockFrameReflowAsync(pending, null).ConfigureAwait(false);
            }
            catch (ObjectDisposedException)
            {
                await AbandonBlockFrameReflowAsync(pending, null).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                await AbandonBlockFrameReflowAsync(pending, exception).ConfigureAwait(false);
            }
        }

        private void CompleteBlockFrameReflow(LaTeXBlockService service,
            PendingBlockFrameReflow pending, LaTeXBlockRender render)
        {
            blockFrameReflowsInFlight.Remove(pending.ShapeKey);
            if (shuttingDown) return;
            if (!IsCurrentBlockFrameReflow(pending))
            {
                StartNextBlockFrameReflow(pending.ShapeKey);
                return;
            }
            try
            {
                LaTeXBlockMetadata currentMetadata;
                string currentSource;
                var hasCurrentContract = pending.IsFloating
                    ? LaTeXBlockService.TryReadContract(pending.FloatingShape, out currentMetadata,
                        out currentSource)
                    : LaTeXBlockService.TryReadContract(pending.InlineShape, out currentMetadata,
                        out currentSource);
                if (!hasCurrentContract ||
                    !SameBlockFrameState(currentMetadata, pending.BaseMetadata) ||
                    !string.Equals(LaTeXBlockService.NormalizeSourceText(currentSource),
                        pending.Source, StringComparison.Ordinal) ||
                    !string.Equals(pending.Profile, currentProfile,
                        StringComparison.OrdinalIgnoreCase))
                {
                    pendingBlockFrameReflows.Remove(pending.ShapeKey);
                    return;
                }

                // A second native drag can happen while TeX is rendering. Never
                // replace that newer user-owned frame with an older result.
                var currentFrameWidthPt = NormalizeFrameExtent(
                    pending.IsFloating ? pending.FloatingShape.Width : pending.InlineShape.Width);
                var currentFrameHeightPt = NormalizeFrameExtent(
                    pending.IsFloating ? pending.FloatingShape.Height : pending.InlineShape.Height);
                if (!FrameExtentsEqual(currentFrameWidthPt, pending.TargetFrameWidthPt) ||
                    !FrameExtentsEqual(currentFrameHeightPt, pending.TargetFrameHeightPt))
                {
                    pendingBlockFrameReflows.Remove(pending.ShapeKey);
                    return;
                }

                // A legacy unstyled Fixed Block takes its foreground from Word's
                // native Font.Color. Styled Blocks instead take it from durable style
                // metadata. For the legacy case, apply the same live-value guard as
                // Auto formulas so an older in-flight SVG cannot repaint a newer
                // colour command.
                if (!currentMetadata.HasExplicitStyle)
                {
                    var liveTextColor = pending.IsFloating
                        ? LaTeXBlockService.ResolveTextColor(pending.FloatingShape.Anchor)
                        : LaTeXBlockService.ResolveTextColor(pending.InlineShape.Range);
                    if (!LaTeXBlockService.TextColorsEqual(liveTextColor,
                            pending.TextColor))
                    {
                        pendingBlockFrameReflows.Remove(pending.ShapeKey);
                        QueueBlockFrameReflow(new BlockFrameReflowRequest(
                            pending.InlineShape, pending.FloatingShape, pending.ShapeKey,
                            currentMetadata, currentSource, pending.TargetWidthPt,
                            pending.TargetFrameWidthPt, pending.TargetFrameHeightPt,
                            liveTextColor, pending.RestoreSelection));
                        return;
                    }
                }

                var restoreSelection = pending.RestoreSelection &&
                    IsPendingBlockStillSelected(pending);
                RunProgrammaticMutation(() =>
                {
                    if (pending.IsFloating)
                    {
                        var replacement = service.UpdateFloatingRendered(pending.FloatingShape,
                            pending.Source, pending.TargetWidthPt, LaTeXBlockLayoutMode.Fixed,
                            render, false);
                        if (restoreSelection) replacement.Select();
                    }
                    else
                    {
                        var replacement = service.UpdateRendered(pending.InlineShape,
                            pending.Source, pending.TargetWidthPt, LaTeXBlockLayoutMode.Fixed,
                            render, false);
                        if (restoreSelection) replacement.Range.Select();
                    }
                }, restoreSelection);
                pendingBlockFrameReflows.Remove(pending.ShapeKey);
                ribbon?.InvalidateWidthControl();
            }
            catch (Exception exception)
            {
                AbandonBlockFrameReflow(pending, exception);
            }
            finally
            {
                StartNextBlockFrameReflow(pending.ShapeKey);
            }
        }

        private bool IsPendingBlockStillSelected(PendingBlockFrameReflow pending)
        {
            try
            {
                if (pending.IsFloating)
                    return Blocks.TryGetSelectedFloatingBlock(out var selectedFloating, out _) &&
                        GetComIdentity(selectedFloating) == pending.ShapeKey;
                return IsInlineShapeStillExactlySelected(pending.ShapeKey);
            }
            catch (COMException) { return false; }
        }

        private Task AbandonBlockFrameReflowAsync(PendingBlockFrameReflow pending,
            Exception exception)
        {
            if (shuttingDown) return Task.FromResult(false);
            return InvokeOnWordUiAsync(() => AbandonBlockFrameReflow(pending, exception));
        }

        private void AbandonBlockFrameReflow(PendingBlockFrameReflow pending,
            Exception exception)
        {
            blockFrameReflowsInFlight.Remove(pending.ShapeKey);
            if (shuttingDown) return;
            if (IsCurrentBlockFrameReflow(pending))
            {
                pendingBlockFrameReflows.Remove(pending.ShapeKey);
                ribbon?.InvalidateWidthControl();
                if (exception != null)
                    MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application),
                        exception.GetBaseException().Message, "LaTeX Blocks",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            StartNextBlockFrameReflow(pending.ShapeKey);
        }

        private bool IsCurrentBlockFrameReflow(PendingBlockFrameReflow pending)
        {
            return pendingBlockFrameReflows.TryGetValue(pending.ShapeKey,
                       out var current) && current.Sequence == pending.Sequence;
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
                    pending.BaseMetadata.Kind == LaTeXBlockKind.DisplayMath ||
                    pending.BaseMetadata.Kind == LaTeXBlockKind.NumberedMath,
                    pending.TargetTextColor, renderKind: pending.BaseMetadata.Kind)
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
            var shape = ResolveCurrentPendingFormatShape(pending);
            FormatRefreshRequest currentFormatRefresh = null;
            if (shape != null && LaTeXBlockService.TryReadContract(shape,
                    out var currentMetadata, out var currentSource) &&
                currentSource == pending.Source &&
                SameMetadataState(currentMetadata, pending.BaseMetadata) &&
                string.Equals(pending.Profile, currentProfile,
                    StringComparison.OrdinalIgnoreCase))
            {
                if (IsFixedContentBlock(currentMetadata))
                {
                    // This path is reachable only for a format render that began
                    // before the fixed-Content routing above (or during a host
                    // event race). Discard that old result and always take the
                    // normal frame-reflow route. In particular, it is not enough
                    // to put a frame around the already-rendered SVG here: a styled
                    // Block's TeX width, leading and foreground colour all depend
                    // on its durable style metadata. Reflow is the one path that
                    // derives all of those values from that metadata before it
                    // composes the physical SVG frame.
                    pendingFormatRefreshes.Remove(pending.ShapeKey);
                    var fixedTextColor = LaTeXBlockService.ResolveTextColor(shape.Range);
                    QueueFixedContentFormatRefresh(new FormatRefreshRequest(shape,
                        currentSource, currentMetadata, pending.TargetWidthPt,
                        currentMetadata.FontSizePt, fixedTextColor), pending.TargetWidthPt,
                        fixedTextColor);
                    ribbon?.InvalidateWidthControl();
                    StartNextFormatRefresh(pending.ShapeKey);
                    return;
                }

                // Native size and colour remain Word's source of truth while an SVG
                // render is in flight. Do not let an older bitmap-like result overwrite
                // either half of a newer format command. A later host event normally
                // creates a new pending sequence, but the live comparison also closes
                // the interval between Word mutating the run and that event arriving.
                var liveTextColor = LaTeXBlockService.ResolveTextColor(shape.Range);
                var liveFontSizePt = currentMetadata.Mode == LaTeXBlockLayoutMode.Auto
                    ? (double)shape.Range.Font.Size
                    : currentMetadata.FontSizePt;
                var liveFontSizeIsValid = liveFontSizePt >= 1 && liveFontSizePt <= 200;
                var fontSizeMatches = currentMetadata.Mode != LaTeXBlockLayoutMode.Auto ||
                    liveFontSizeIsValid &&
                    Math.Abs(liveFontSizePt - pending.TargetFontSizePt) < 0.001;
                var textColorMatches = LaTeXBlockService.TextColorsEqual(liveTextColor,
                    pending.TargetTextColor);
                if (fontSizeMatches && textColorMatches)
                {
                    // Updating an SVG replaces its Word InlineShape. If the old
                    // object is still selected, deleting it would otherwise collapse
                    // the selection and make the user's highlight disappear. Restore
                    // only while the pending object remains selected; never pull the
                    // selection back after the user has moved elsewhere during render.
                    var restoreRangeSelection = pending.SelectionLease != null &&
                        pending.SelectionLease.Matches(Application.Selection);
                    var restoreExactSelection = pending.SelectionLease == null &&
                        IsPendingFormatTargetStillSelected(pending);
                    RunProgrammaticMutation(() =>
                    {
                        var replacement = service.UpdateRendered(shape, pending.Source,
                            pending.TargetWidthPt, pending.BaseMetadata.Mode, render, false);
                        if (restoreRangeSelection)
                            pending.SelectionLease.TryRestore(Application);
                        else if (restoreExactSelection)
                            replacement.Range.Select();
                    }, restoreRangeSelection || restoreExactSelection);
                }
                else
                {
                    // A one-character formula range should always have one concrete
                    // size. If Word transiently reports wdUndefined, retain the last
                    // authoritative target rather than feeding an invalid design size
                    // to StemTeX; the next selection event will observe the final value.
                    var targetFontSizePt = liveFontSizeIsValid
                        ? liveFontSizePt
                        : pending.TargetFontSizePt;
                    currentFormatRefresh = new FormatRefreshRequest(shape, currentSource,
                        currentMetadata, targetFontSizePt, liveTextColor,
                        currentMetadata.Mode == LaTeXBlockLayoutMode.Auto &&
                        Math.Abs(targetFontSizePt - currentMetadata.FontSizePt) > 0.001,
                        !textColorMatches, pending.SelectionLease);
                }
            }
            pendingFormatRefreshes.Remove(pending.ShapeKey);
            if (currentFormatRefresh != null) QueueFormatRefresh(currentFormatRefresh);
            ribbon?.InvalidateWidthControl();
            StartNextFormatRefresh(pending.ShapeKey);
        }

        private WordInterop.InlineShape ResolveCurrentPendingFormatShape(
            PendingFormatRefresh pending)
        {
            if (pending == null) return null;
            try
            {
                if (pending.Shape != null &&
                    LaTeXBlockService.TryReadContract(pending.Shape,
                        out var metadata, out _) &&
                    metadata.Id == pending.BaseMetadata.Id)
                    return pending.Shape;
            }
            catch (COMException) { }

            // A direct-media batch reconstructs one Flat OPC envelope. Drawings
            // that are not batch targets survive in the document, but their former
            // COM wrappers are deleted. The persisted formula GUID is the stable
            // identity, so reacquire the current InlineShape before completing an
            // independent render such as an interleaved numbered display equation.
            try
            {
                foreach (WordInterop.Document document in Application.Documents)
                {
                    var current = LaTeXBlockService.FindInlineShapeById(document,
                        pending.BaseMetadata.Id);
                    if (current != null) return current;
                }
            }
            catch (COMException) { }
            return null;
        }

        private bool QueueObservedBlockFrameResize(BlockFrameSnapshot snapshot, long shapeKey)
        {
            if (snapshot == null || snapshot.ShapeKey != shapeKey) return false;
            var request = CreateBlockFrameReflowRequest(snapshot, false, false);
            return request != null && QueueBlockFrameReflow(request);
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

        private bool IsPendingFormatTargetStillSelected(PendingFormatRefresh pending)
        {
            if (pending == null) return false;
            try
            {
                var selection = Application.Selection;
                if (selection == null || selection.InlineShapes.Count != 1 ||
                    GetComIdentity(selection.InlineShapes[1]) != pending.ShapeKey)
                    return false;
                return selection.Type == WordInterop.WdSelectionType.wdSelectionInlineShape;
            }
            catch (COMException) { return false; }
        }

        private bool IsInlineShapeStillExactlySelected(long shapeKey)
        {
            try
            {
                var selection = Application.Selection;
                return selection != null &&
                       selection.Type == WordInterop.WdSelectionType.wdSelectionInlineShape &&
                       selection.InlineShapes.Count == 1 &&
                       GetComIdentity(selection.InlineShapes[1]) == shapeKey;
            }
            catch (COMException) { return false; }
        }

        private bool HasPendingFormatTarget(WordInterop.InlineShape shape, double fontSizePt,
            int textColor)
        {
            if (shape == null) return false;
            var shapeKey = GetComIdentity(shape);
            if (pendingFormatRefreshes.TryGetValue(shapeKey, out var pending) &&
                (Math.Abs(pending.TargetFontSizePt - fontSizePt) > 0.001 ||
                 !LaTeXBlockService.TextColorsEqual(pending.TargetTextColor, textColor)))
                return true;
            return pendingFormatBatchTargets.TryGetValue(shapeKey, out var batch) &&
                   (Math.Abs(batch.FontSizePt - fontSizePt) > 0.001 ||
                    !LaTeXBlockService.TextColorsEqual(batch.TextColor, textColor));
        }

        private bool HasPendingTextColorTarget(WordInterop.InlineShape shape, int textColor)
        {
            if (shape == null) return false;
            var shapeKey = GetComIdentity(shape);
            if (pendingFormatRefreshes.TryGetValue(shapeKey, out var pending) &&
                !LaTeXBlockService.TextColorsEqual(pending.TargetTextColor, textColor))
                return true;
            return pendingFormatBatchTargets.TryGetValue(shapeKey, out var batch) &&
                   !LaTeXBlockService.TextColorsEqual(batch.TextColor, textColor);
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
            return LaTeXBlockService.SameRefreshMetadataState(left, right);
        }

        private static bool SameBlockFrameState(LaTeXBlockMetadata left,
            LaTeXBlockMetadata right)
        {
            return SameMetadataState(left, right) &&
                   Math.Abs(left.DepthPt - right.DepthPt) < 0.001;
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
                try
                {
                    if (outermost && recaptureSelectionSnapshot && !shuttingDown &&
                        Application.Documents.Count > 0)
                    {
                        try
                        {
                            RememberSelection(Application.Selection);
                        }
                        catch { ClearPreviousSelectionSnapshot(); }
                    }
                }
                finally
                {
                    programmaticMutationDepth--;
                }
            }
        }

        private void RememberSelection(WordInterop.Selection selection)
        {
            previousSelectionFontSnapshots = CaptureSelectionFontSnapshots(selection);
            previousSelectionRangeLease = SelectionRangeLease.Capture(selection);
            previousBlockFrameSnapshot = CaptureBlockFrameSnapshot();
            UpdateNativeFormatMonitorContext(selection);
        }

        private void UpdateNativeFormatMonitorContext(WordInterop.Selection selection)
        {
            var monitor = wordFormatInteractionSource as WordFontColorMonitor;
            if (monitor == null) return;
            var enabled = false;
            try
            {
                // The same native adapter observes both Font Color and Font Size.
                // An exact InlineShape selection must remain enabled so a display or
                // inline formula receives Font Size commits. CaptureFontColorInteraction
                // independently rejects that exact selection, leaving its foreground
                // entirely on Office Graphics Fill.
                enabled = selection != null &&
                    previousSelectionFontSnapshots.Count > 0;
            }
            catch (COMException) { enabled = false; }
            monitor.SetInteractionContext(enabled);
        }

        private void ClearPreviousSelectionSnapshot()
        {
            previousSelectionFontSnapshots = new List<SelectionFontSnapshot>();
            previousSelectionRangeLease = null;
            previousBlockFrameSnapshot = null;
        }

        private sealed class SelectionRangeLease
        {
            private SelectionRangeLease(long documentKey,
                WordInterop.WdStoryType storyType, int start, int end,
                WordInterop.WdSelectionType selectionType)
            {
                DocumentKey = documentKey;
                StoryType = storyType;
                Start = start;
                End = end;
                SelectionType = selectionType;
            }

            internal long DocumentKey { get; }
            internal WordInterop.WdStoryType StoryType { get; }
            internal int Start { get; }
            internal int End { get; }
            internal WordInterop.WdSelectionType SelectionType { get; }

            internal static SelectionRangeLease Capture(WordInterop.Selection selection)
            {
                if (selection == null) return null;
                try
                {
                    var selectedRange = selection.Range;
                    return new SelectionRangeLease(
                        GetComIdentity(selectedRange.Document), selectedRange.StoryType,
                        selectedRange.Start, selectedRange.End, selection.Type);
                }
                catch (COMException) { return null; }
            }

            internal SelectionRangeLease Clone()
            {
                return new SelectionRangeLease(DocumentKey, StoryType, Start, End,
                    SelectionType);
            }

            internal bool Matches(WordInterop.Selection selection)
            {
                if (selection == null) return false;
                try
                {
                    var current = selection.Range;
                    return GetComIdentity(current.Document) == DocumentKey &&
                           current.StoryType == StoryType && current.Start == Start &&
                           current.End == End && selection.Type == SelectionType;
                }
                catch (COMException) { return false; }
            }

            internal bool TryRestore(WordInterop.Application application)
            {
                if (application == null) return false;
                try
                {
                    // InlineShape replacement is a one-character-for-one-character
                    // edit. Recreate the range from scalar coordinates so a lease
                    // never roots a Word Range COM object across document close.
                    var document = application.ActiveDocument;
                    if (GetComIdentity(document) != DocumentKey) return false;
                    var range = document.StoryRanges[StoryType].Duplicate;
                    range.SetRange(Start, End);
                    range.Select();
                    return true;
                }
                catch (COMException) { return false; }
            }
        }

        private sealed class PendingFontColorInteraction
        {
            internal PendingFontColorInteraction(long interactionId,
                SelectionRangeLease selectionLease,
                List<SelectionFontSnapshot> formulas)
            {
                InteractionId = interactionId;
                SelectionLease = selectionLease ??
                    throw new ArgumentNullException(nameof(selectionLease));
                Formulas = formulas ?? new List<SelectionFontSnapshot>();
            }

            internal long InteractionId { get; }
            internal SelectionRangeLease SelectionLease { get; }
            internal List<SelectionFontSnapshot> Formulas { get; }
        }

        private sealed class SelectionFontSnapshot
        {
            internal SelectionFontSnapshot(WordInterop.InlineShape shape,
                double hostFontSizePt, int hostTextColor)
            {
                Shape = shape;
                ShapeKey = GetComIdentity(shape);
                HostFontSizePt = hostFontSizePt;
                HostTextColor = hostTextColor;
            }
            internal WordInterop.InlineShape Shape { get; }
            internal long ShapeKey { get; }
            internal double HostFontSizePt { get; }
            internal int HostTextColor { get; }
        }

        private sealed class BlockFrameSnapshot
        {
            internal BlockFrameSnapshot(WordInterop.InlineShape inlineShape,
                WordInterop.Shape floatingShape, long shapeKey, LaTeXBlockMetadata metadata,
                string source, double frameWidthPt, double frameHeightPt,
                double layoutWidthPt)
            {
                if ((inlineShape == null && floatingShape == null) ||
                    (inlineShape != null && floatingShape != null))
                    throw new ArgumentException("A Block frame snapshot needs exactly one Word shape.");
                InlineShape = inlineShape;
                FloatingShape = floatingShape;
                ShapeKey = shapeKey;
                Metadata = metadata;
                Source = LaTeXBlockService.NormalizeSourceText(source);
                FrameWidthPt = frameWidthPt;
                FrameHeightPt = frameHeightPt;
                LayoutWidthPt = layoutWidthPt;
            }

            internal WordInterop.InlineShape InlineShape { get; }
            internal WordInterop.Shape FloatingShape { get; }
            internal bool IsFloating => FloatingShape != null;
            internal long ShapeKey { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal string Source { get; }
            internal double FrameWidthPt { get; }
            internal double FrameHeightPt { get; }
            internal double LayoutWidthPt { get; }
        }

        private sealed class BlockFrameReflowRequest
        {
            internal BlockFrameReflowRequest(WordInterop.InlineShape inlineShape,
                WordInterop.Shape floatingShape, long shapeKey, LaTeXBlockMetadata metadata,
                string source, double targetWidthPt, double targetFrameWidthPt,
                double targetFrameHeightPt, int textColor, bool restoreSelection)
            {
                InlineShape = inlineShape;
                FloatingShape = floatingShape;
                ShapeKey = shapeKey;
                Metadata = metadata;
                Source = LaTeXBlockService.NormalizeSourceText(source);
                TargetWidthPt = targetWidthPt;
                TargetFrameWidthPt = targetFrameWidthPt;
                TargetFrameHeightPt = targetFrameHeightPt;
                TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
                RestoreSelection = restoreSelection;
            }

            internal WordInterop.InlineShape InlineShape { get; }
            internal WordInterop.Shape FloatingShape { get; }
            internal bool IsFloating => FloatingShape != null;
            internal bool HasShape => (InlineShape != null) != (FloatingShape != null);
            internal long ShapeKey { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal string Source { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFrameWidthPt { get; }
            internal double TargetFrameHeightPt { get; }
            internal int TextColor { get; }
            internal bool RestoreSelection { get; }

            internal BlockFrameReflowRequest WithFormat(double targetWidthPt, int textColor)
            {
                return new BlockFrameReflowRequest(InlineShape, FloatingShape, ShapeKey,
                    Metadata, Source, targetWidthPt, TargetFrameWidthPt, TargetFrameHeightPt,
                    textColor, RestoreSelection);
            }
        }

        private sealed class FormatRefreshRequest
        {
            internal FormatRefreshRequest(WordInterop.InlineShape shape, string source,
                LaTeXBlockMetadata metadata, double fontSizePt, int textColor,
                bool changesFontSize = true, bool changesTextColor = true,
                SelectionRangeLease selectionLease = null,
                int? previousTextColor = null)
            {
                Shape = shape;
                ShapeKey = GetComIdentity(shape);
                Source = source;
                Metadata = metadata;
                WidthPt = metadata.WidthPt;
                FontSizePt = fontSizePt;
                TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
                ChangesFontSize = changesFontSize;
                ChangesTextColor = changesTextColor;
                SelectionLease = selectionLease;
                PreviousTextColor = previousTextColor.HasValue
                    ? LaTeXBlockService.NormalizeTextColor(
                        previousTextColor.Value)
                    : (int?)null;
            }
            internal FormatRefreshRequest(WordInterop.InlineShape shape, string source,
                LaTeXBlockMetadata metadata, double widthPt, double fontSizePt, int textColor,
                SelectionRangeLease selectionLease = null)
            {
                Shape = shape;
                ShapeKey = GetComIdentity(shape);
                Source = source;
                Metadata = metadata;
                WidthPt = widthPt;
                FontSizePt = fontSizePt;
                TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
                ChangesWidth = true;
                SelectionLease = selectionLease;
            }
            internal WordInterop.InlineShape Shape { get; }
            internal long ShapeKey { get; }
            internal string Source { get; }
            internal LaTeXBlockMetadata Metadata { get; }
            internal double WidthPt { get; }
            internal double FontSizePt { get; }
            internal int TextColor { get; }
            internal bool ChangesWidth { get; }
            internal bool ChangesFontSize { get; }
            internal bool ChangesTextColor { get; }
            internal int? PreviousTextColor { get; }
            internal SelectionRangeLease SelectionLease { get; }
        }

        private sealed class PendingFormatBatch
        {
            internal PendingFormatBatch(long sequence, string profile,
                SelectionRangeLease selectionLease,
                List<FormatRefreshRequest> requests)
            {
                Sequence = sequence;
                Profile = profile;
                SelectionLease = selectionLease;
                Requests = requests ?? new List<FormatRefreshRequest>();
            }

            internal long Sequence { get; }
            internal string Profile { get; }
            internal SelectionRangeLease SelectionLease { get; }
            internal List<FormatRefreshRequest> Requests { get; }
        }

        private sealed class PendingFormatBatchTarget
        {
            internal PendingFormatBatchTarget(long sequence, int textColor,
                double fontSizePt)
            {
                Sequence = sequence;
                TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
                FontSizePt = fontSizePt;
            }

            internal long Sequence { get; }
            internal int TextColor { get; }
            internal double FontSizePt { get; }
        }

        private sealed class PendingFormatRefresh
        {
            internal PendingFormatRefresh(long shapeKey, WordInterop.InlineShape shape,
                LaTeXBlockMetadata baseMetadata, string source, string profile,
                double targetWidthPt, double targetFontSizePt, int targetTextColor,
                SelectionRangeLease selectionLease, long sequence)
            {
                ShapeKey = shapeKey;
                Shape = shape;
                BaseMetadata = baseMetadata;
                Source = source;
                Profile = profile;
                TargetWidthPt = targetWidthPt;
                TargetFontSizePt = targetFontSizePt;
                TargetTextColor = LaTeXBlockService.NormalizeTextColor(targetTextColor);
                SelectionLease = selectionLease;
                Sequence = sequence;
            }

            internal long ShapeKey { get; }
            internal WordInterop.InlineShape Shape { get; }
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal string Source { get; }
            internal string Profile { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFontSizePt { get; }
            internal int TargetTextColor { get; }
            internal SelectionRangeLease SelectionLease { get; }
            internal long Sequence { get; }
        }

        private sealed class PendingBlockFrameReflow
        {
            internal PendingBlockFrameReflow(long shapeKey, WordInterop.InlineShape inlineShape,
                WordInterop.Shape floatingShape, LaTeXBlockMetadata baseMetadata,
                string source, string profile, double targetWidthPt, double targetFrameWidthPt,
                double targetFrameHeightPt, int textColor, bool restoreSelection, long sequence)
            {
                ShapeKey = shapeKey;
                InlineShape = inlineShape;
                FloatingShape = floatingShape;
                BaseMetadata = baseMetadata;
                Source = source;
                Profile = profile;
                TargetWidthPt = targetWidthPt;
                TargetFrameWidthPt = targetFrameWidthPt;
                TargetFrameHeightPt = targetFrameHeightPt;
                TextColor = LaTeXBlockService.NormalizeTextColor(textColor);
                RestoreSelection = restoreSelection;
                Sequence = sequence;
            }

            internal long ShapeKey { get; }
            internal WordInterop.InlineShape InlineShape { get; }
            internal WordInterop.Shape FloatingShape { get; }
            internal bool IsFloating => FloatingShape != null;
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal string Source { get; }
            internal string Profile { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFrameWidthPt { get; }
            internal double TargetFrameHeightPt { get; }
            internal int TextColor { get; }
            internal bool RestoreSelection { get; }
            internal long Sequence { get; }
        }

        private void Application_WindowBeforeDoubleClick(WordInterop.Selection selection, ref bool cancel)
        {
            if (shuttingDown || !hostEventProcessingEnabled) return;
            try
            {
                if (!Blocks.TryGetSelectedBlock(out _, out _) &&
                    (!Blocks.TryGetSelectedFloatingBlock(out _, out var floatingMetadata) ||
                     floatingMetadata.Role != LaTeXBlockRole.Content ||
                     floatingMetadata.Mode != LaTeXBlockLayoutMode.Fixed))
                    return;
                cancel = true;
                ShowEditEditor();
            }
            catch (Exception exception)
            {
                MessageBox.Show(new LaTeXBlocksRibbon.WordWindow(Application), exception.Message, "LaTeX Blocks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string LoadCurrentProfile(IStemTeXBackend pool)
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
            pendingFormatBatchTargets.Clear();
            Interlocked.Increment(ref formatBatchSequence);
            pendingBlockFrameReflows.Clear();
            Interlocked.Increment(ref blockFrameReflowSequence);
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
