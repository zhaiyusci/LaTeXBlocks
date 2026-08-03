using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace LaTeXBlocks.PowerPoint
{
    public partial class ThisAddIn
    {
        // A profile is a host-level preference: selecting a PowerPoint profile
        // must not change the one Word starts with.
        private const string SettingsKey = @"Software\LaTeXBlocks\PowerPoint";
        private const string LegacySettingsKey = @"Software\LaTeXBlocks";
        private StemTeXBackend rendererPool;
        private PowerPointBlockService blocks;
        private string currentProfile;
        private string backendStartupError;
        private Control powerPointUiDispatcher;
        private LaTeXBlocksRibbon ribbon;
        private readonly Dictionary<PowerPointShapeKey, PendingBlockFormat> pendingBlockFormats =
            new Dictionary<PowerPointShapeKey, PendingBlockFormat>();
        private readonly Dictionary<PowerPointShapeKey, long> blockFormatsInFlight =
            new Dictionary<PowerPointShapeKey, long>();
        private readonly Dictionary<PowerPointShapeKey, PendingNativeFrameGesture> pendingNativeFrameGestures =
            new Dictionary<PowerPointShapeKey, PendingNativeFrameGesture>();
        private long blockFormatSequence;
        private long nativeFrameGestureSequence;
        private int programmaticShapeMutationDepth;
        private bool shuttingDown;
        // Automatic fitting must never persist a width the user cannot inspect or
        // edit with the PowerPoint controls.
        private const double NativeReflowMinimumWidthPt = BlockLayoutWidthPolicy.MinimumPt;
        private const double NativeReflowMaximumWidthPt = BlockLayoutWidthPolicy.MaximumPt;
        private const double FrameFitTolerancePt = 0.05;
        private const int MaximumNativeReflowAttempts = 3;

        internal PowerPointInterop.Application PowerPointApplication => Application;
        private StemTeXBackend Renderers =>
            rendererPool ?? (rendererPool = new StemTeXBackend());
        private PowerPointBlockService Blocks =>
            blocks ?? (blocks = new PowerPointBlockService(Application, Renderers));

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                powerPointUiDispatcher = new Control();
                powerPointUiDispatcher.CreateControl();
                Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                Application.AfterShapeSizeChange += Application_AfterShapeSizeChange;
                var pool = Renderers;
                currentProfile = LoadCurrentProfile(pool);
                pool.SwitchProfile(currentProfile);
            }
            catch (Exception exception)
            {
                backendStartupError = exception.Message;
                // Startup can fail after Office event sinks or the UI dispatcher have
                // already been created. Leave the add-in disabled-but-clean instead
                // of retaining callbacks into a half-initialized host object.
                ReleaseHostResources();
            }
        }

        private void ReleaseHostResources()
        {
            try { Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick; } catch { }
            try { Application.WindowSelectionChange -= Application_WindowSelectionChange; } catch { }
            try { Application.AfterShapeSizeChange -= Application_AfterShapeSizeChange; } catch { }
            try { rendererPool?.Dispose(); } catch { }
            rendererPool = null;
            blocks = null;
            try { powerPointUiDispatcher?.Dispose(); } catch { }
            powerPointUiDispatcher = null;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            shuttingDown = true;
            Interlocked.Increment(ref blockFormatSequence);
            pendingBlockFormats.Clear();
            blockFormatsInFlight.Clear();
            CancelAllPendingNativeFrameGestures();
            try
            {
                // Office can disconnect event sinks before VSTO raises Shutdown. Each
                // unhook is therefore best effort; none may prevent the renderer's
                // non-blocking disposal/reaper from running.
                try { Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick; } catch { }
                try { Application.WindowSelectionChange -= Application_WindowSelectionChange; } catch { }
                try { Application.AfterShapeSizeChange -= Application_AfterShapeSizeChange; } catch { }
            }
            finally
            {
                ReleaseHostResources();
            }
        }

        internal void ShowInsertBlockEditor()
        {
            EnsureBackendAvailable();
            if (Application.Presentations.Count == 0 || Application.ActiveWindow == null)
                throw new InvalidOperationException(
                    "Open a PowerPoint presentation before inserting a LaTeX Block.");
            var fontSizePt = PowerPointBlockService.ResolveFontSize(Application, 18);
            var widthPt = Blocks.ResolveInitialWidth(360);
            using (var editor = new LaTeXBlockEditorForm(Blocks, "\\[\nE=mc^2\n\\]",
                widthPt, fontSizePt,
                currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, false, LaTeXBlockStyle.Default))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.PowerPointWindow(Application)) ==
                    DialogResult.OK)
                {
                    if (editor.AcceptedRender == null)
                        throw new InvalidOperationException(
                            "The accepted LaTeX preview is unavailable.");
                    RunShapeMutation(() => Blocks.InsertRendered(editor.AcceptedSource,
                        editor.AcceptedWidthPt, editor.AcceptedRender,
                        editor.AcceptedStyle));
                }
            }
        }

        internal void ShowEditBlockEditor()
        {
            EnsureBackendAvailable();
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select a LaTeX Block first.");
            var source = shape.AlternativeText;
            var style = PowerPointBlockService.ReadStyle(shape);
            var frameHeightPt = PowerPointBlockService.ReadFrameHeightPt(shape);
            var frameWidthPt = PowerPointBlockService.ReadFrameWidthPt(shape);
            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt,
                metadata.FontSizePt,
                currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, true, style, frameHeightPt, frameWidthPt))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.PowerPointWindow(Application)) ==
                    DialogResult.OK)
                {
                    if (editor.AcceptedRender == null)
                        throw new InvalidOperationException(
                            "The accepted LaTeX preview is unavailable.");
                    RunShapeMutation(() => Blocks.UpdateRendered(shape,
                        editor.AcceptedSource, editor.AcceptedWidthPt,
                        editor.AcceptedRender, true, null, null,
                        editor.AcceptedStyle));
                }
            }
        }

        internal bool HasSelectedBlockLayoutWidth()
        {
            return TryGetSelectedBlockContract(out _, out _);
        }

        internal string GetSelectedBlockLayoutWidthText()
        {
            return TryGetSelectedBlockContract(out _, out var metadata)
                ? metadata.WidthPt.ToString("0.0", CultureInfo.CurrentCulture)
                : string.Empty;
        }

        internal void ApplySelectedBlockLayoutWidth(string text)
        {
            EnsureBackendAvailable();
            if (!PowerPointRibbonContract.TryParseLayoutWidthPt(text, out var widthPt))
                throw new ArgumentException(
                    "Enter a typesetting width from 30 to 450 pt.", nameof(text));
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select one LaTeX Block first.");
            QueueBlockFormat(shape, metadata, widthPt, null);
        }

        internal bool HasSelectedBlockFontSize()
        {
            return TryGetSelectedBlockContract(out _, out _);
        }

        internal string GetSelectedBlockFontSizeText()
        {
            return TryGetSelectedBlockContract(out _, out var metadata)
                ? metadata.FontSizePt.ToString("0.###", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        internal void ApplySelectedBlockFontSize(string text)
        {
            EnsureBackendAvailable();
            if (!PowerPointRibbonContract.TryParseFontSize(text, out var fontSizePt))
                throw new ArgumentException("Enter a TeX font size from 1 to 200 pt.", nameof(text));
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select one LaTeX Block first.");
            QueueBlockFormat(shape, metadata, null, fontSizePt);
        }

        private void QueueBlockFormat(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata, double? widthPt, double? fontSizePt,
            double? frameHeightPt = null, double? frameWidthPt = null,
            long? nativeFrameGestureSequence = null)
        {
            var key = PowerPointBlockService.GetShapeKey(shape);
            var source = PowerPointBlockService.NormalizeSourceText(shape.AlternativeText);
            var style = PowerPointBlockService.ReadStyle(shape);
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var targetWidthPt = metadata.WidthPt;
            var targetFontSizePt = metadata.FontSizePt;
            var currentFrameWidthPt = PowerPointBlockService.ReadFrameWidthPt(shape);
            var currentFrameHeightPt = PowerPointBlockService.ReadFrameHeightPt(shape);
            // During a native resize, PowerPoint has already changed Shape.Width /
            // Height but the embedded SVG still advertises its previous root box.
            // Compare against that root box when deciding whether a frame update is
            // meaningful; comparing against the temporary host scale would wrongly
            // discard every native request as a no-op.
            var storedFrameWidthPt = PowerPointBlockService.ReadPositiveTag(shape,
                PowerPointBlockService.SvgWidthTag, currentFrameWidthPt);
            var storedFrameHeightPt = PowerPointBlockService.ReadPositiveTag(shape,
                PowerPointBlockService.SvgHeightTag, currentFrameHeightPt);
            var targetFrameWidthPt = currentFrameWidthPt;
            var targetFrameHeightPt = currentFrameHeightPt;
            // A native resize supplies one frame-fitting request. These flags say
            // which axes the user actually constrained; they do not create
            // different side/corner resize modes. A horizontal drag, for example,
            // may let a reflowed block grow naturally in height.
            var constrainFrameWidth = frameWidthPt.HasValue;
            var constrainFrameHeight = frameHeightPt.HasValue;
            var targetNativeFrameGestureSequence = nativeFrameGestureSequence;
            var hasExplicitFormatIntent = widthPt.HasValue || fontSizePt.HasValue;
            var autoReflowAttempts = 0;
            if (pendingBlockFormats.TryGetValue(key, out var existing) &&
                SameBaseState(existing, metadata, source, style, profile))
            {
                targetWidthPt = existing.TargetWidthPt;
                targetFontSizePt = existing.TargetFontSizePt;
                targetFrameWidthPt = existing.TargetFrameWidthPt;
                targetFrameHeightPt = existing.TargetFrameHeightPt;
                constrainFrameWidth = existing.ConstrainFrameWidth;
                constrainFrameHeight = existing.ConstrainFrameHeight;
                // A format edit and an external-frame intent can coexist. Track
                // them separately: a later return to the stored frame must discard
                // only the native frame intent, never a queued TeX size/width edit.
                targetNativeFrameGestureSequence = nativeFrameGestureSequence ??
                    existing.NativeFrameGestureSequence;
                hasExplicitFormatIntent |= existing.HasExplicitFormatIntent;
                autoReflowAttempts = existing.AutoReflowAttempts;
            }
            if (widthPt.HasValue) targetWidthPt = widthPt.Value;
            if (fontSizePt.HasValue) targetFontSizePt = fontSizePt.Value;
            if (frameHeightPt.HasValue)
            {
                targetFrameHeightPt = frameHeightPt.Value;
                constrainFrameHeight = true;
            }
            if (frameWidthPt.HasValue)
            {
                targetFrameWidthPt = frameWidthPt.Value;
                constrainFrameWidth = true;
            }
            // A new native frame or a deliberate TeX formatting change deserves a
            // fresh fit pass. The renderer result may have a different natural box.
            if (nativeFrameGestureSequence.HasValue || widthPt.HasValue || fontSizePt.HasValue)
            {
                autoReflowAttempts = 0;
            }

            // A native size change is always a genuine TeX layout request, even
            // when the previous SVG happened to fit inside the new outer frame.
            // Its first width estimate comes from the previous real SVG root, not
            // from PowerPoint's temporary host-side scale. A height-only request
            // rerenders the same width; that is still a fresh TeX layout pass.
            if (nativeFrameGestureSequence.HasValue && frameWidthPt.HasValue)
                targetWidthPt = EstimateNativeLayoutWidth(metadata.WidthPt,
                    storedFrameWidthPt, targetFrameWidthPt);

            var comparisonFrameWidthPt = nativeFrameGestureSequence.HasValue
                ? storedFrameWidthPt : currentFrameWidthPt;
            var comparisonFrameHeightPt = nativeFrameGestureSequence.HasValue
                ? storedFrameHeightPt : currentFrameHeightPt;
            if (Math.Abs(targetWidthPt - metadata.WidthPt) < 0.01 &&
                Math.Abs(targetFontSizePt - metadata.FontSizePt) < 0.001 &&
                Math.Abs(targetFrameWidthPt - comparisonFrameWidthPt) < 0.01 &&
                Math.Abs(targetFrameHeightPt - comparisonFrameHeightPt) < 0.01)
            {
                pendingBlockFormats.Remove(key);
                ribbon?.InvalidateBlockControls();
                return;
            }

            var sequence = Interlocked.Increment(ref blockFormatSequence);
            var pending = new PendingBlockFormat(key, shape, metadata, source, profile,
                targetWidthPt, targetFontSizePt, targetFrameWidthPt, targetFrameHeightPt,
                constrainFrameWidth, constrainFrameHeight, sequence,
                targetNativeFrameGestureSequence, hasExplicitFormatIntent,
                autoReflowAttempts, style);
            pendingBlockFormats[key] = pending;
            StartNextBlockFormat(key);
        }

        private void StartNextBlockFormat(PowerPointShapeKey key)
        {
            if (shuttingDown || blockFormatsInFlight.ContainsKey(key) ||
                !pendingBlockFormats.TryGetValue(key, out var pending)) return;
            blockFormatsInFlight[key] = pending.Sequence;
            _ = RenderBlockFormatAsync(Blocks, pending);
        }

        private async Task RenderBlockFormatAsync(PowerPointBlockService service,
            PendingBlockFormat pending)
        {
            try
            {
                var render = await service.RenderCommittedAsync(pending.Source,
                    pending.TargetWidthPt, pending.Profile,
                    pending.TargetFontSizePt, pending.Style,
                    pending.TargetFrameHeightPt, pending.TargetFrameWidthPt)
                    .ConfigureAwait(false);
                PostToPowerPointUi(() => CompleteBlockFormat(service, pending, render));
            }
            catch (TaskCanceledException)
            {
                PostToPowerPointUi(() => AbandonBlockFormat(pending, null));
            }
            catch (ObjectDisposedException)
            {
                PostToPowerPointUi(() => AbandonBlockFormat(pending, null));
            }
            catch (Exception exception)
            {
                PostToPowerPointUi(() => AbandonBlockFormat(pending, exception));
            }
        }

        private void CompleteBlockFormat(PowerPointBlockService service,
            PendingBlockFormat pending, LaTeXBlockRender render)
        {
            ReleaseInFlight(pending);
            if (shuttingDown) return;
            if (!IsCurrentPending(pending))
            {
                StartNextBlockFormat(pending.Key);
                return;
            }
            try
            {
                var shape = pending.Shape;
                if (shape == null || !PowerPointBlockService.TryReadContract(shape,
                        out var current, out var currentSource) ||
                    !SameMetadataState(current, pending.BaseMetadata) ||
                    currentSource != pending.Source ||
                    !pending.Style.Equals(PowerPointBlockService.ReadStyle(shape)) ||
                    !string.Equals(pending.Profile, currentProfile,
                        StringComparison.OrdinalIgnoreCase))
                {
                    pendingBlockFormats.Remove(pending.Key);
                    return;
                }
                if (TryCreateAutoReflowPending(pending, render, out var reflowPending))
                {
                    // Every native size change has already started a real TeX
                    // layout pass. If its first measured result still misses a
                    // constrained edge, retry with a corrected typesetting width.
                    // This is reflow, never SVG scale or crop operation.
                    pendingBlockFormats[pending.Key] = reflowPending;
                    return;
                }
                var keepSelected = service.TryGetSelectedBlock(out var selectedShape,
                    out _) && PowerPointBlockService.GetShapeKey(selectedShape)
                    .Equals(pending.Key);
                RunShapeMutation(() => service.UpdateRendered(shape, pending.Source,
                    pending.TargetWidthPt, render, keepSelected, pending.TargetFrameHeightPt,
                    pending.TargetFrameWidthPt, pending.Style));
                pendingBlockFormats.Remove(pending.Key);
                ribbon?.InvalidateBlockControls();
            }
            catch (Exception exception)
            {
                AbandonBlockFormat(pending, exception);
            }
            finally
            {
                StartNextBlockFormat(pending.Key);
            }
        }

        private void AbandonBlockFormat(PendingBlockFormat pending, Exception exception)
        {
            ReleaseInFlight(pending);
            if (shuttingDown) return;
            if (IsCurrentPending(pending))
            {
                pendingBlockFormats.Remove(pending.Key);
                ribbon?.InvalidateBlockControls();
                if (exception != null) ShowBlockFormattingError(exception);
            }
            StartNextBlockFormat(pending.Key);
        }

        private bool IsCurrentPending(PendingBlockFormat pending)
        {
            return pendingBlockFormats.TryGetValue(pending.Key,
                       out var current) && current.Sequence == pending.Sequence;
        }

        private void ReleaseInFlight(PendingBlockFormat pending)
        {
            if (blockFormatsInFlight.TryGetValue(pending.Key, out var sequence) &&
                sequence == pending.Sequence)
                blockFormatsInFlight.Remove(pending.Key);
        }

        private bool TryCreateAutoReflowPending(PendingBlockFormat pending,
            LaTeXBlockRender render, out PendingBlockFormat reflowPending)
        {
            reflowPending = null;
            if (!pending.NativeFrameGestureSequence.HasValue ||
                pending.AutoReflowAttempts >= MaximumNativeReflowAttempts)
                return false;

            var renderedWidthPt = PowerPointBlockService.ReadSvgWidthPt(render.SvgBytes);
            var renderedHeightPt = PowerPointBlockService.ReadSvgHeightPt(render.SvgBytes);
            var widthOverflows = pending.ConstrainFrameWidth &&
                renderedWidthPt > pending.TargetFrameWidthPt + FrameFitTolerancePt;
            var heightOverflows = pending.ConstrainFrameHeight &&
                renderedHeightPt > pending.TargetFrameHeightPt + FrameFitTolerancePt;
            if (!widthOverflows && !heightOverflows) return false;

            var reflowWidthPt = ResolveNativeReflowWidth(pending, renderedWidthPt,
                widthOverflows, heightOverflows);
            if (double.IsNaN(reflowWidthPt)) return false;
            if (Math.Abs(reflowWidthPt - pending.TargetWidthPt) < FrameFitTolerancePt)
                return false;

            reflowPending = new PendingBlockFormat(pending.Key, pending.Shape,
                pending.BaseMetadata, pending.Source, pending.Profile, reflowWidthPt,
                pending.TargetFontSizePt, pending.TargetFrameWidthPt,
                pending.TargetFrameHeightPt, pending.ConstrainFrameWidth,
                pending.ConstrainFrameHeight, Interlocked.Increment(ref blockFormatSequence),
                pending.NativeFrameGestureSequence, pending.HasExplicitFormatIntent,
                pending.AutoReflowAttempts + 1, pending.Style);
            return true;
        }

        private static double ResolveNativeReflowWidth(PendingBlockFormat pending,
            double renderedWidthPt, bool widthOverflows, bool heightOverflows)
        {
            if (pending.ConstrainFrameWidth && renderedWidthPt > 0)
            {
                // A fixed-width StemTeX block normally has a near-constant SVG
                // edge allowance: natural width is approximately layout width plus
                // that allowance. Use the measured additive error as one bounded
                // estimate, then verify it with a real TeX render. This is not SVG
                // scaling and it tolerates a source whose actual line breaks differ
                // from that local estimate.
                var correctedWidthPt = ClampNativeReflowWidth(pending.TargetWidthPt +
                    pending.TargetFrameWidthPt - renderedWidthPt);
                if (widthOverflows) return correctedWidthPt;
                if (heightOverflows && correctedWidthPt > pending.TargetWidthPt +
                    FrameFitTolerancePt)
                    return correctedWidthPt;
                return double.NaN;
            }
            if (heightOverflows)
            {
                // A height-only resize leaves width unconstrained. Try the widest
                // user-editable TeX measure once; it gives text the best chance to
                // reduce line count without introducing a host-side scale.
                return NativeReflowMaximumWidthPt;
            }
            return double.NaN;
        }

        private static double ClampNativeReflowWidth(double widthPt)
        {
            if (double.IsNaN(widthPt) || double.IsInfinity(widthPt))
                return NativeReflowMinimumWidthPt;
            return Math.Max(NativeReflowMinimumWidthPt,
                Math.Min(NativeReflowMaximumWidthPt, widthPt));
        }

        private static double EstimateNativeLayoutWidth(double currentLayoutWidthPt,
            double storedSvgWidthPt, double targetFrameWidthPt)
        {
            if (!(storedSvgWidthPt > 0)) return ClampNativeReflowWidth(targetFrameWidthPt);
            // In a fixed-width StemTeX block the root width is the layout width
            // plus its measured edge allowance. Preserve that allowance instead
            // of treating the PowerPoint frame as an image-scale factor.
            return ClampNativeReflowWidth(currentLayoutWidthPt + targetFrameWidthPt -
                storedSvgWidthPt);
        }

        private static bool SameBaseState(PendingBlockFormat pending,
            LaTeXBlockMetadata metadata, string source, LaTeXBlockStyle style,
            string profile)
        {
            return SameMetadataState(pending.BaseMetadata, metadata) &&
                   pending.Source == source && pending.Style.Equals(style) &&
                   string.Equals(pending.Profile, profile,
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

        private void Application_WindowSelectionChange(PowerPointInterop.Selection selection)
        {
            ribbon?.InvalidateBlockControls();
        }

        private void Application_AfterShapeSizeChange(PowerPointInterop.Shape shape)
        {
            if (shuttingDown || programmaticShapeMutationDepth > 0 || shape == null) return;
            try
            {
                if (!PowerPointBlockService.TryReadContract(shape, out var metadata, out _))
                    return;
                var key = PowerPointBlockService.GetShapeKey(shape);
                var frameUpdate = PowerPointBlockService.CaptureFrameResize(shape, metadata);
                if (frameUpdate.HasChange)
                {
                    // PowerPoint reports a corner drag as adjacent width and height
                    // notifications. Wait for a short quiet period before replacing
                    // the temporary scaled picture, so every handle delivers the
                    // same final outer-frame operation.
                    ScheduleNativeFrameGesture(shape, metadata);
                    ribbon?.InvalidateBlockControls();
                }
                else
                {
                    // A drag can end back at the stored frame after a previous native
                    // frame render has already begun. The original geometry is the
                    // latest user intent, so cancel only frame-only pending work.
                    CancelNativeFrameWork(key, shape);
                    ribbon?.InvalidateBlockControls();
                }
            }
            catch (COMException) { }
            catch (Exception exception)
            {
                ShowBlockFormattingError(exception);
            }
        }

        private void ScheduleNativeFrameGesture(PowerPointInterop.Shape shape,
            LaTeXBlockMetadata metadata)
        {
            var key = PowerPointBlockService.GetShapeKey(shape);
            // A new gesture makes any older, frame-only render obsolete. Keep a
            // genuine Ribbon format change intact; it is not a resize operation.
            CancelNativeFrameWork(key, shape);
            var pending = new PendingNativeFrameGesture(key, shape, metadata,
                Interlocked.Increment(ref nativeFrameGestureSequence));
            pendingNativeFrameGestures[key] = pending;
            _ = CommitNativeFrameGestureAfterIdleAsync(pending);
        }

        private async Task CommitNativeFrameGestureAfterIdleAsync(
            PendingNativeFrameGesture pending)
        {
            // This is a gesture debounce, not resize polling. It only joins the
            // individual Office notifications emitted for one native manipulation.
            try
            {
                await Task.Delay(120, pending.Cancellation.Token).ConfigureAwait(false);
                if (!pending.Cancellation.IsCancellationRequested)
                    PostToPowerPointUi(() => CommitNativeFrameGesture(pending));
            }
            catch (OperationCanceledException) { }
            finally
            {
                pending.DisposeCancellation();
            }
        }

        private void CommitNativeFrameGesture(PendingNativeFrameGesture pending)
        {
            if (shuttingDown || !pendingNativeFrameGestures.TryGetValue(pending.Key,
                    out var current) || current.Sequence != pending.Sequence)
                return;
            pendingNativeFrameGestures.Remove(pending.Key);
            try
            {
                if (pending.Shape == null || !PowerPointBlockService.TryReadContract(
                        pending.Shape, out var metadata, out _) ||
                    !SameMetadataState(metadata, pending.BaseMetadata))
                    return;

                var frameUpdate = PowerPointBlockService.CaptureFrameResize(pending.Shape,
                    metadata);
                if (!frameUpdate.HasChange) return;

                // The temporary PowerPoint-scaled picture stays in place while the
                // SVG is rendered. Replacing it directly with the same final frame
                // avoids a second, programmatic size-change event and keeps every
                // native handle on the identical outer-frame path.
                QueueBlockFormat(pending.Shape, metadata, null, null,
                    frameUpdate.HeightChanged ? frameUpdate.FrameHeightPt : (double?)null,
                    frameUpdate.WidthChanged ? frameUpdate.FrameWidthPt : (double?)null,
                    pending.Sequence);
                ribbon?.InvalidateBlockControls();
            }
            catch (COMException) { }
            catch (Exception exception)
            {
                ShowBlockFormattingError(exception);
            }
        }

        private bool TryGetSelectedBlockContract(out PowerPointInterop.Shape shape,
            out LaTeXBlockMetadata metadata)
        {
            shape = null;
            metadata = null;
            try
            {
                var selection = Application.ActiveWindow?.Selection;
                if (selection == null ||
                    selection.Type != PowerPointInterop.PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count != 1) return false;
                var candidate = selection.ShapeRange[1];
                if (!PowerPointBlockService.TryReadContract(candidate, out metadata, out _))
                    return false;
                shape = candidate;
                return true;
            }
            catch (COMException) { return false; }
        }

        private bool PostToPowerPointUi(Action action)
        {
            var dispatcher = powerPointUiDispatcher;
            if (dispatcher == null || dispatcher.IsDisposed || shuttingDown) return false;
            try
            {
                dispatcher.BeginInvoke(action);
                return true;
            }
            catch (InvalidOperationException) { return false; }
        }

        private void ShowBlockFormattingError(Exception exception)
        {
            if (shuttingDown) return;
            MessageBox.Show(new LaTeXBlocksRibbon.PowerPointWindow(Application),
                exception.GetBaseException().Message, "LaTeX Blocks",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            ribbon?.InvalidateBlockControls();
        }

        private T RunShapeMutation<T>(Func<T> action)
        {
            programmaticShapeMutationDepth++;
            try { return action(); }
            finally { programmaticShapeMutationDepth--; }
        }

        private void RunShapeMutation(Action action)
        {
            programmaticShapeMutationDepth++;
            try { action(); }
            finally { programmaticShapeMutationDepth--; }
        }

        private void CancelNativeFrameWork(PowerPointShapeKey key,
            PowerPointInterop.Shape shape = null)
        {
            if (pendingNativeFrameGestures.TryGetValue(key, out var pendingGesture))
            {
                pendingNativeFrameGestures.Remove(key);
                pendingGesture.Cancel();
            }
            if (pendingBlockFormats.TryGetValue(key, out var pendingFormat) &&
                pendingFormat.NativeFrameGestureSequence.HasValue)
            {
                if (!pendingFormat.HasExplicitFormatIntent)
                {
                    pendingBlockFormats.Remove(key);
                    return;
                }

                // A queued Ribbon edit remains valid, but its native frame intent
                // no longer is. Publish a fresh pending item so an in-flight old
                // render cannot later restore the obsolete frame.
                if (shape == null) return;
                var frameWidthPt = PowerPointBlockService.ReadFrameWidthPt(shape);
                var frameHeightPt = PowerPointBlockService.ReadFrameHeightPt(shape);
                var replacement = new PendingBlockFormat(key, shape,
                    pendingFormat.BaseMetadata, pendingFormat.Source,
                    pendingFormat.Profile, pendingFormat.TargetWidthPt,
                    pendingFormat.TargetFontSizePt, frameWidthPt, frameHeightPt,
                    false, false, Interlocked.Increment(ref blockFormatSequence),
                    null, true, 0, pendingFormat.Style);
                pendingBlockFormats[key] = replacement;
                StartNextBlockFormat(key);
            }
        }

        private void CancelAllPendingNativeFrameGestures()
        {
            foreach (var pending in pendingNativeFrameGestures.Values)
                pending.Cancel();
            pendingNativeFrameGestures.Clear();
        }

        private void Application_WindowBeforeDoubleClick(PowerPointInterop.Selection selection,
            ref bool cancel)
        {
            try
            {
                if (selection == null ||
                    selection.Type != PowerPointInterop.PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count != 1 ||
                    !PowerPointBlockService.TryReadContract(selection.ShapeRange[1], out _, out _))
                    return;
                // Do not cancel PowerPoint's native edit until we know the deferred
                // editor invocation was actually queued. A disposed dispatcher is a
                // normal shutdown race, not an exception that may escape the event.
                if (!PostToPowerPointUi(new Action(() =>
                {
                    if (shuttingDown) return;
                    try { ShowEditBlockEditor(); }
                    catch (Exception exception)
                    {
                        MessageBox.Show(new LaTeXBlocksRibbon.PowerPointWindow(Application),
                            exception.Message, "LaTeX Blocks", MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }))) return;
                cancel = true;
            }
            catch (COMException) { }
            catch (InvalidOperationException) { }
        }

        private void EnsureBackendAvailable()
        {
            if (!string.IsNullOrEmpty(backendStartupError) && rendererPool == null)
                throw new InvalidOperationException(
                    "StemTeX could not start: " + backendStartupError);
            var ignored = Renderers;
            if (string.IsNullOrEmpty(currentProfile))
                currentProfile = LoadCurrentProfile(Renderers);
        }

        private static string LoadCurrentProfile(StemTeXBackend pool)
        {
            string saved = null;
            using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey))
                saved = key?.GetValue("Profile") as string;
            // Upgrade from the former shared preference once. Subsequent saves
            // always go to PowerPoint's own key and can no longer affect Word.
            if (string.IsNullOrWhiteSpace(saved))
                using (var key = Registry.CurrentUser.OpenSubKey(LegacySettingsKey))
                    saved = key?.GetValue("Profile") as string;
            foreach (var profile in pool.Profiles)
                if (string.Equals(profile, saved, StringComparison.OrdinalIgnoreCase))
                    return profile;
            return pool.DefaultAvailableProfile;
        }

        private void SetCurrentProfile(string profile)
        {
            var valid = false;
            foreach (var candidate in Renderers.Profiles)
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase))
                {
                    profile = candidate;
                    valid = true;
                    break;
                }
            if (!valid)
                throw new ArgumentException("Unknown StemTeX profile: " + profile,
                    nameof(profile));
            if (shuttingDown) throw new ObjectDisposedException(nameof(ThisAddIn));
            if (string.Equals(currentProfile, profile, StringComparison.OrdinalIgnoreCase)) return;

            // Persist before publishing the new in-memory profile. A registry failure
            // must leave the current profile and all queued document updates intact.
            var previousProfile = currentProfile;
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
                    key.SetValue("Profile", profile, RegistryValueKind.String);
                Renderers.SwitchProfile(profile);
            }
            catch
            {
                RestorePersistedProfile(previousProfile);
                throw;
            }
            currentProfile = profile;
            foreach (var pending in new List<PendingBlockFormat>(pendingBlockFormats.Values))
                AbandonBlockFormat(pending, null);
        }

        private static void RestorePersistedProfile(string profile)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
                {
                    if (string.IsNullOrWhiteSpace(profile)) key.DeleteValue("Profile", false);
                    else key.SetValue("Profile", profile, RegistryValueKind.String);
                }
            }
            catch { }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return ribbon = new LaTeXBlocksRibbon(this);
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private sealed class PendingNativeFrameGesture
        {
            internal PendingNativeFrameGesture(PowerPointShapeKey key,
                PowerPointInterop.Shape shape, LaTeXBlockMetadata baseMetadata,
                long sequence)
            {
                Key = key;
                Shape = shape;
                BaseMetadata = baseMetadata;
                Sequence = sequence;
                Cancellation = new CancellationTokenSource();
            }

            internal PowerPointShapeKey Key { get; }
            internal PowerPointInterop.Shape Shape { get; }
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal long Sequence { get; }
            internal CancellationTokenSource Cancellation { get; }

            internal void Cancel()
            {
                try { Cancellation.Cancel(); } catch (ObjectDisposedException) { }
            }

            internal void DisposeCancellation()
            {
                try { Cancellation.Dispose(); } catch (ObjectDisposedException) { }
            }
        }

        private sealed class PendingBlockFormat
        {
            internal PendingBlockFormat(PowerPointShapeKey key,
                PowerPointInterop.Shape shape, LaTeXBlockMetadata baseMetadata,
                string source, string profile, double targetWidthPt,
                double targetFontSizePt, double targetFrameWidthPt,
                double targetFrameHeightPt, bool constrainFrameWidth,
                bool constrainFrameHeight, long sequence,
                long? nativeFrameGestureSequence, bool hasExplicitFormatIntent,
                int autoReflowAttempts, LaTeXBlockStyle style)
            {
                Key = key;
                Shape = shape;
                BaseMetadata = baseMetadata;
                Source = source;
                Profile = profile;
                TargetWidthPt = targetWidthPt;
                TargetFontSizePt = targetFontSizePt;
                TargetFrameWidthPt = targetFrameWidthPt;
                TargetFrameHeightPt = targetFrameHeightPt;
                ConstrainFrameWidth = constrainFrameWidth;
                ConstrainFrameHeight = constrainFrameHeight;
                Sequence = sequence;
                NativeFrameGestureSequence = nativeFrameGestureSequence;
                HasExplicitFormatIntent = hasExplicitFormatIntent;
                AutoReflowAttempts = autoReflowAttempts;
                Style = style ?? LaTeXBlockStyle.Default;
            }

            internal PowerPointShapeKey Key { get; }
            internal PowerPointInterop.Shape Shape { get; }
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal string Source { get; }
            internal string Profile { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFontSizePt { get; }
            internal double TargetFrameWidthPt { get; }
            internal double TargetFrameHeightPt { get; }
            internal bool ConstrainFrameWidth { get; }
            internal bool ConstrainFrameHeight { get; }
            internal long Sequence { get; }
            internal long? NativeFrameGestureSequence { get; }
            internal bool HasExplicitFormatIntent { get; }
            internal int AutoReflowAttempts { get; }
            internal LaTeXBlockStyle Style { get; }
        }
    }
}
