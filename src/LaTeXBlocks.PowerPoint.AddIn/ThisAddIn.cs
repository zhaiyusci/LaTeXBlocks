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
        private readonly HashSet<PowerPointShapeKey> deferredSizeEventSuppression =
            new HashSet<PowerPointShapeKey>();
        private long blockFormatSequence;
        private int programmaticShapeMutationDepth;
        private bool shuttingDown;

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
            deferredSizeEventSuppression.Clear();
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
                SetCurrentProfile, false))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.PowerPointWindow(Application)) ==
                    DialogResult.OK)
                {
                    if (editor.AcceptedRender == null)
                        throw new InvalidOperationException(
                            "The accepted LaTeX preview is unavailable.");
                    RunShapeMutation(() => Blocks.InsertRendered(editor.AcceptedSource,
                        editor.AcceptedWidthPt, editor.AcceptedRender));
                }
            }
        }

        internal void ShowEditBlockEditor()
        {
            EnsureBackendAvailable();
            if (!Blocks.TryGetSelectedBlock(out var shape, out var metadata))
                throw new InvalidOperationException("Select a LaTeX Block first.");
            var source = shape.AlternativeText;
            using (var editor = new LaTeXBlockEditorForm(Blocks, source, metadata.WidthPt,
                metadata.FontSizePt,
                currentProfile ?? Renderers.DefaultAvailableProfile,
                SetCurrentProfile, true))
            {
                if (editor.ShowDialog(new LaTeXBlocksRibbon.PowerPointWindow(Application)) ==
                    DialogResult.OK)
                {
                    if (editor.AcceptedRender == null)
                        throw new InvalidOperationException(
                            "The accepted LaTeX preview is unavailable.");
                    RunShapeMutation(() => Blocks.UpdateRendered(shape,
                        editor.AcceptedSource, editor.AcceptedWidthPt,
                        editor.AcceptedRender));
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
            LaTeXBlockMetadata metadata, double? widthPt, double? fontSizePt)
        {
            var key = PowerPointBlockService.GetShapeKey(shape);
            var source = PowerPointBlockService.NormalizeSourceText(shape.AlternativeText);
            var profile = currentProfile ?? Renderers.DefaultAvailableProfile;
            var targetWidthPt = metadata.WidthPt;
            var targetFontSizePt = metadata.FontSizePt;
            if (pendingBlockFormats.TryGetValue(key, out var existing) &&
                SameBaseState(existing, metadata, source, profile))
            {
                targetWidthPt = existing.TargetWidthPt;
                targetFontSizePt = existing.TargetFontSizePt;
            }
            if (widthPt.HasValue) targetWidthPt = widthPt.Value;
            if (fontSizePt.HasValue) targetFontSizePt = fontSizePt.Value;

            if (Math.Abs(targetWidthPt - metadata.WidthPt) < 0.01 &&
                Math.Abs(targetFontSizePt - metadata.FontSizePt) < 0.001)
            {
                pendingBlockFormats.Remove(key);
                ribbon?.InvalidateBlockControls();
                return;
            }

            var sequence = Interlocked.Increment(ref blockFormatSequence);
            var pending = new PendingBlockFormat(key, shape, metadata, source, profile,
                targetWidthPt, targetFontSizePt, sequence);
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
                    pending.TargetFontSizePt).ConfigureAwait(false);
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
                    !string.Equals(pending.Profile, currentProfile,
                        StringComparison.OrdinalIgnoreCase))
                {
                    pendingBlockFormats.Remove(pending.Key);
                    return;
                }
                var keepSelected = service.TryGetSelectedBlock(out var selectedShape,
                    out _) && PowerPointBlockService.GetShapeKey(selectedShape)
                    .Equals(pending.Key);
                RunShapeMutation(() => service.UpdateRendered(shape, pending.Source,
                    pending.TargetWidthPt, render, keepSelected));
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

        private static bool SameBaseState(PendingBlockFormat pending,
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
                if (deferredSizeEventSuppression.Contains(key)) return;
                var resize = PowerPointBlockService.ClassifyResize(shape, metadata);
                if (resize.Kind == PowerPointResizeKind.VisualScale)
                {
                    RunDeferredSizeMutation(shape, () =>
                        PowerPointBlockService.NormalizeVisualScale(shape, resize.VisualScale));
                    ribbon?.InvalidateBlockControls();
                }
                else if (resize.Kind == PowerPointResizeKind.LayoutWidth)
                {
                    // PowerPoint has already distorted the old SVG to report the side
                    // handle's requested width. Restore valid vector geometry
                    // immediately; the asynchronously rendered replacement will reuse
                    // this (possibly moved) upper-left anchor.
                    RunDeferredSizeMutation(shape, () =>
                        PowerPointBlockService.RestoreStoredGeometry(shape, metadata));
                    QueueBlockFormat(shape, metadata, resize.LayoutWidthPt, null);
                }
                else if (pendingBlockFormats.TryGetValue(key, out var pending) &&
                         Math.Abs(pending.TargetWidthPt - metadata.WidthPt) > 0.01)
                {
                    // The user dragged the side handle back to the stored width while
                    // an older reflow was still running. Cancel that desired width;
                    // an independently pending TeX-size change remains intact.
                    QueueBlockFormat(shape, metadata, metadata.WidthPt, null);
                }
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

        private void RunDeferredSizeMutation(PowerPointInterop.Shape shape, Action action)
        {
            var key = PowerPointBlockService.GetShapeKey(shape);
            deferredSizeEventSuppression.Add(key);
            try { RunShapeMutation(action); }
            finally
            {
                PostToPowerPointUi(() => deferredSizeEventSuppression.Remove(key));
            }
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

        private sealed class PendingBlockFormat
        {
            internal PendingBlockFormat(PowerPointShapeKey key,
                PowerPointInterop.Shape shape, LaTeXBlockMetadata baseMetadata,
                string source, string profile, double targetWidthPt,
                double targetFontSizePt, long sequence)
            {
                Key = key;
                Shape = shape;
                BaseMetadata = baseMetadata;
                Source = source;
                Profile = profile;
                TargetWidthPt = targetWidthPt;
                TargetFontSizePt = targetFontSizePt;
                Sequence = sequence;
            }

            internal PowerPointShapeKey Key { get; }
            internal PowerPointInterop.Shape Shape { get; }
            internal LaTeXBlockMetadata BaseMetadata { get; }
            internal string Source { get; }
            internal string Profile { get; }
            internal double TargetWidthPt { get; }
            internal double TargetFontSizePt { get; }
            internal long Sequence { get; }
        }
    }
}
