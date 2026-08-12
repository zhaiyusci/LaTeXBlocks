using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Accessibility;

namespace LaTeXBlocks.Word
{
    internal enum WordFormatInteractionPhase
    {
        Began,
        Committed,
        Canceled
    }

    internal enum WordFormatProperty
    {
        TextColor,
        FontSize
    }

    internal enum WordFormatInteractionOrigin
    {
        FontColorMainButton,
        FontColorPalette,
        FontColorMoreColorsDialog,
        FontSizeControl
    }

    /// <summary>
    /// Describes the lifecycle of one native Word formatting interaction. The signal
    /// deliberately carries no value: the Word integration reconciles the affected
    /// host ranges after a committed choice and remains the sole owner of Word COM.
    /// </summary>
    internal sealed class WordFormatInteractionEventArgs : EventArgs
    {
        internal WordFormatInteractionEventArgs(long interactionId,
            WordFormatInteractionPhase phase, WordFormatProperty property,
            WordFormatInteractionOrigin origin)
        {
            if (interactionId <= 0)
                throw new ArgumentOutOfRangeException(nameof(interactionId));
            InteractionId = interactionId;
            Phase = phase;
            Property = property;
            Origin = origin;
        }

        internal long InteractionId { get; }
        internal WordFormatInteractionPhase Phase { get; }
        internal WordFormatProperty Property { get; }
        internal WordFormatInteractionOrigin Origin { get; }
    }

    internal interface IWordFormatInteractionSource : IDisposable
    {
        event EventHandler<WordFormatInteractionEventArgs> FormatInteraction;
        void Start();
    }

    /// <summary>
    /// Value-free lifecycle state shared by the native adapters and smoke tests. It
    /// is intentionally independent of UIA, MSAA, Win32 and Word COM.
    /// </summary>
    internal sealed class WordFormatTransactionState
    {
        private readonly object gate = new object();
        private long nextInteractionId;
        private long activeInteractionId;
        private WordFormatProperty activeProperty;
        private WordFormatInteractionOrigin activeOrigin;
        private string lastOperation = string.Empty;

        internal long ActiveInteractionId
        {
            get { lock (gate) return activeInteractionId; }
        }
        internal WordFormatInteractionOrigin ActiveOrigin
        {
            get { lock (gate) return activeOrigin; }
        }
        internal string DiagnosticState
        {
            get { lock (gate) return lastOperation; }
        }

        internal WordFormatInteractionEventArgs Begin(WordFormatProperty property,
            WordFormatInteractionOrigin origin,
            out WordFormatInteractionEventArgs canceledPrevious)
        {
            lock (gate)
            {
                canceledPrevious = activeInteractionId == 0
                    ? null
                    : CreateActiveSignal(WordFormatInteractionPhase.Canceled,
                        activeOrigin);
                activeInteractionId = NextInteractionId();
                activeProperty = property;
                activeOrigin = origin;
                lastOperation = "begin:" + activeInteractionId + "/" + origin;
                return CreateActiveSignal(WordFormatInteractionPhase.Began, origin);
            }
        }

        internal WordFormatInteractionEventArgs Commit(long interactionId,
            WordFormatProperty property, WordFormatInteractionOrigin origin)
        {
            lock (gate)
            {
                if (!Matches(interactionId, property))
                {
                    lastOperation = "commit-miss:" + interactionId + "/active=" +
                        activeInteractionId;
                    return null;
                }
                var signal = CreateActiveSignal(WordFormatInteractionPhase.Committed,
                    origin);
                Clear();
                lastOperation = "commit:" + interactionId + "/" + origin;
                return signal;
            }
        }

        internal WordFormatInteractionEventArgs Cancel(long interactionId,
            WordFormatProperty property, WordFormatInteractionOrigin origin)
        {
            lock (gate)
            {
                if (!Matches(interactionId, property))
                {
                    lastOperation = "cancel-miss:" + interactionId + "/active=" +
                        activeInteractionId;
                    return null;
                }
                var signal = CreateActiveSignal(WordFormatInteractionPhase.Canceled,
                    origin);
                Clear();
                lastOperation = "cancel:" + interactionId + "/" + origin;
                return signal;
            }
        }

        internal bool UpdateOrigin(long interactionId,
            WordFormatInteractionOrigin origin)
        {
            lock (gate)
            {
                if (activeInteractionId == 0 || activeInteractionId != interactionId)
                    return false;
                activeOrigin = origin;
                return true;
            }
        }

        private bool Matches(long interactionId, WordFormatProperty property)
        {
            return interactionId != 0 && activeInteractionId == interactionId &&
                   activeProperty == property;
        }

        private WordFormatInteractionEventArgs CreateActiveSignal(
            WordFormatInteractionPhase phase, WordFormatInteractionOrigin origin)
        {
            return new WordFormatInteractionEventArgs(activeInteractionId, phase,
                activeProperty, origin);
        }

        private long NextInteractionId()
        {
            var interactionId = ++nextInteractionId;
            if (interactionId > 0) return interactionId;
            nextInteractionId = 1;
            return 1;
        }

        private void Clear()
        {
            activeInteractionId = 0;
            activeProperty = default(WordFormatProperty);
            activeOrigin = default(WordFormatInteractionOrigin);
        }
    }

    internal enum WordFontColorSignal
    {
        MainButtonInvoked,
        PaletteItemCommitted,
        MoreColorsOpened,
        MoreColorsAccepted,
        MoreColorsRejected,
        MoreColorsCanceled,
        MoreColorsClosed
    }

    /// <summary>
    /// Small deterministic state machine shared by the UI Automation adapter and the
    /// smoke test. Opening a gallery, closing it, or canceling a dialog is explicitly
    /// not a colour commit.
    /// </summary>
    internal sealed class WordFontColorInteractionState
    {
        private bool moreColorsOpen;
        private bool moreColorsAccepted;

        internal bool Observe(WordFontColorSignal signal)
        {
            switch (signal)
            {
                case WordFontColorSignal.MainButtonInvoked:
                case WordFontColorSignal.PaletteItemCommitted:
                    moreColorsOpen = false;
                    moreColorsAccepted = false;
                    return true;
                case WordFontColorSignal.MoreColorsOpened:
                    moreColorsOpen = true;
                    moreColorsAccepted = false;
                    return false;
                case WordFontColorSignal.MoreColorsAccepted:
                    if (!moreColorsOpen) return false;
                    moreColorsAccepted = true;
                    return false;
                case WordFontColorSignal.MoreColorsRejected:
                    if (moreColorsOpen) moreColorsAccepted = false;
                    return false;
                case WordFontColorSignal.MoreColorsCanceled:
                    moreColorsOpen = false;
                    moreColorsAccepted = false;
                    return false;
                case WordFontColorSignal.MoreColorsClosed:
                    var committed = moreColorsOpen && moreColorsAccepted;
                    moreColorsOpen = false;
                    moreColorsAccepted = false;
                    return committed;
                default:
                    return false;
            }
        }

        internal bool IsMoreColorsOpen => moreColorsOpen;
    }

    /// <summary>
    /// Observes actual commits made by Word's built-in Font Color control. Word has
    /// no object-model event for this command, and generic mouse capture is not a
    /// command signal: merely opening or canceling the gallery also ends capture.
    /// Office exposes stable Fluent control ids through UI Automation. Gallery hover
    /// identifies a candidate, while a paired left-button down/up on that same live
    /// popup candidate (or a provider Invoke on builds that expose one) commits it.
    /// </summary>
    internal sealed class WordFontColorMonitor : IWordFormatInteractionSource
    {
        private const string FontColorPickerId = "FontColorPicker";
        private const string FontColorDropDownId = "FontColorPicker_Dropdown";
        private const string MoreColorsId = "FontColorMoreColorsDialog";
        private const string FontSizeId = "FontSize";
        private const int MaximumAncestorDepth = 16;
        private const uint EventObjectHide = 0x8003;
        private const uint EventObjectFocus = 0x8005;
        private const uint EventObjectSelection = 0x8006;
        private const uint EventObjectInvoked = 0x8013;
        private const uint WineventOutOfContext = 0;
        private const int RoleSystemListItem = 0x22;
        private const int RoleSystemMenuItem = 0x0c;
        private const int ObjIdWindow = 0;
        private const int HookCallWndProc = 4;
        private const int HookLowLevelKeyboard = 13;
        private const int HookLowLevelMouse = 14;
        private const int HookCodeAction = 0;
        private const uint GetAncestorRoot = 2;
        private const uint WindowMessageLeftButtonDown = 0x0201;
        private const uint WindowMessageLeftButtonUp = 0x0202;
        private const uint WindowMessageKeyDown = 0x0100;
        private const uint WindowMessageSystemKeyDown = 0x0104;
        private const uint VirtualKeyEscape = 0x1b;
        private const uint WindowMessageClose = 0x0010;
        private const uint WindowMessageCommand = 0x0111;
        private const int DialogResultOk = 1;
        private const int DialogResultCancel = 2;
        private const int PaletteCommitDelayMilliseconds = 1;
        private const uint PaletteEventPairWindowMilliseconds = 3000;

        private readonly Control dispatcher;
        private readonly int wordProcessId;
        private readonly WinEventDelegate winEventHandler;
        private readonly CallWndProcDelegate callWndProcHandler;
        private readonly LowLevelMouseDelegate lowLevelMouseHandler;
        private readonly LowLevelKeyboardDelegate lowLevelKeyboardHandler;
        private System.Threading.Timer paletteCommitTimer;
        private readonly WordFontColorInteractionState interactionState =
            new WordFontColorInteractionState();
        private int interactionContextEnabled;
        private readonly WordFormatTransactionState formatTransactionState =
            new WordFormatTransactionState();
        private readonly object stateGate = new object();
        private readonly Queue<WordFormatInteractionEventArgs> pendingFormatSignals =
            new Queue<WordFormatInteractionEventArgs>();
        private long mainButtonCommitDedupUntilUtcTicks;
        private long fontSizeCommitDedupUntilUtcTicks;
        private long fontSizeSessionUntilUtcTicks;
        private long paletteSessionUntilUtcTicks;
        private long paletteCandidateUntilUtcTicks;
        private IntPtr paletteCandidateHwnd;
        private int paletteCandidateObjectId;
        private int paletteCandidateChildId;
        private uint paletteCandidateEventTime;
        private long paletteCandidateInteractionId;
        private IntPtr palettePopupRootHwnd;
        private long palettePressedInteractionId;
        private IntPtr palettePressedHwnd;
        private long pendingPaletteCommitInteractionId;
        private long paletteCommitGeneration;
        private IntPtr hideHook;
        private IntPtr focusHook;
        private IntPtr selectionHook;
        private IntPtr invokedHook;
        private IntPtr lowLevelMouseHook;
        private IntPtr lowLevelKeyboardHook;
        private IntPtr dialogMessageHook;
        private IntPtr moreColorsDialogHwnd;
        private IntPtr pendingMoreColorsDialogHwnd;
        private long moreColorsInteractionId;
        private int paletteSelectionsForTest;
        private int paletteCandidatesForTest;
        private int paletteInvocationsForTest;
        private int paletteMatchesForTest;
        private int formatBeginsForTest;
        private int formatCommitsForTest;
        private int formatCancelsForTest;
        private string lastFormatSignalForTest = string.Empty;
        private int lastPaletteSelectionRoleForTest;
        private string lastPaletteCandidateTupleForTest = string.Empty;
        private string lastPaletteInvokedTupleForTest = string.Empty;
        private string lastPalettePointerClassForTest = string.Empty;
        private string lastPaletteHideForTest = string.Empty;
        private bool started;
        private bool disposed;
        private bool formatSignalDrainScheduled;

        internal WordFontColorMonitor(Control dispatcher)
            : this(dispatcher, Process.GetCurrentProcess().Id)
        {
        }

        internal WordFontColorMonitor(Control dispatcher, int targetProcessId)
        {
            this.dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            if (targetProcessId <= 0) throw new ArgumentOutOfRangeException(nameof(targetProcessId));
            wordProcessId = targetProcessId;
            winEventHandler = OnWinEvent;
            callWndProcHandler = OnCallWndProc;
            lowLevelMouseHandler = OnLowLevelMouse;
            lowLevelKeyboardHandler = OnLowLevelKeyboard;
        }

        public event EventHandler<WordFormatInteractionEventArgs> FormatInteraction;

        internal bool IsRunning => started && !disposed;
        internal string DiagnosticStateForTest =>
            "selection=" + Volatile.Read(ref paletteSelectionsForTest) +
            ", candidate=" + Volatile.Read(ref paletteCandidatesForTest) +
            ", invoked=" + Volatile.Read(ref paletteInvocationsForTest) +
            ", matched=" + Volatile.Read(ref paletteMatchesForTest) +
            ", semantic=" + Volatile.Read(ref formatBeginsForTest) + "/" +
                Volatile.Read(ref formatCommitsForTest) + "/" +
                Volatile.Read(ref formatCancelsForTest) +
            ", lastSignal=" + lastFormatSignalForTest +
            ", transaction=" + formatTransactionState.DiagnosticState +
            ", role=" + Volatile.Read(ref lastPaletteSelectionRoleForTest) +
            ", candidateTuple=" + lastPaletteCandidateTupleForTest +
            ", invokedTuple=" + lastPaletteInvokedTupleForTest +
            ", pointerClass=" + lastPalettePointerClassForTest +
            ", hide=" + lastPaletteHideForTest +
            ", session=" + IsPaletteSessionActive();

        internal void SetInteractionContext(bool enabled)
        {
            Interlocked.Exchange(ref interactionContextEnabled, enabled ? 1 : 0);
            if (enabled) return;
            StopPaletteCommitTimer();
            lock (stateGate)
            {
                pendingFormatSignals.Clear();
                formatSignalDrainScheduled = false;
                pendingPaletteCommitInteractionId = 0;
                ClearPaletteCandidateLocked();
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
            }
        }

        public void Start()
        {
            ThrowIfDisposed();
            if (started) return;
            try
            {
                // Never subscribe an in-process Office add-in to desktop-root UIA
                // events. UIAutomationCore creates callback infrastructure that can
                // keep WINWORD hot for minutes after its windows and add-in have
                // closed. Native WinEvent/MSAA hooks below provide the event stream;
                // UIA remains a short-lived, on-demand classifier only.
                hideHook = RegisterWinEvent(EventObjectHide);
                focusHook = RegisterWinEvent(EventObjectFocus);
                selectionHook = RegisterWinEvent(EventObjectSelection);
                invokedHook = RegisterWinEvent(EventObjectInvoked);
                lowLevelMouseHook = SetLowLevelMouseHook(HookLowLevelMouse,
                    lowLevelMouseHandler, GetModuleHandle(null), 0);
                if (lowLevelMouseHook == IntPtr.Zero)
                    throw new Win32Exception(Marshal.GetLastWin32Error(),
                        "Word could not register its Font Color mouse monitor.");
                lowLevelKeyboardHook = SetLowLevelKeyboardHook(
                    HookLowLevelKeyboard, lowLevelKeyboardHandler,
                    GetModuleHandle(null), 0);
                started = true;
            }
            catch
            {
                RemoveAutomationHandlers();
                throw;
            }
        }

        private IntPtr RegisterWinEvent(uint eventType)
        {
            var hook = SetWinEventHook(eventType, eventType, IntPtr.Zero, winEventHandler,
                unchecked((uint)wordProcessId), 0, WineventOutOfContext);
            if (hook == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    "Word could not register its Font Color accessibility monitor.");
            return hook;
        }

        private void OnWinEvent(IntPtr hookHandle, uint eventType, IntPtr hwnd,
            int objectId, int childId, uint eventThread, uint eventTime)
        {
            if (disposed || Volatile.Read(ref interactionContextEnabled) == 0)
                return;
            try
            {
                if (eventType == EventObjectSelection)
                {
                    Interlocked.Increment(ref paletteSelectionsForTest);
                    // The Font Color dropdown click establishes the palette session
                    // once, asynchronously. Never call UI Automation from every
                    // gallery selection event: Office services those provider calls
                    // on its UI thread, so swatch hover/selection can otherwise make
                    // the pointer visibly stall.
                    if (!IsPaletteSessionActive()) return;
                    // Record only values already carried by the native event. Even
                    // MSAA accRole/accLocation calls marshal back into Word and are
                    // therefore forbidden on the hover path.
                    Volatile.Write(ref lastPaletteSelectionRoleForTest, 0);
                    lastPalettePointerClassForTest = "native-event";
                    SetPaletteCandidate(hwnd, objectId, childId, eventTime);
                    return;
                }
                if (eventType == EventObjectInvoked)
                {
                    Interlocked.Increment(ref paletteInvocationsForTest);
                    if (TryConsumePaletteCandidate(hwnd, objectId, childId, eventTime,
                            out var paletteInteractionId) &&
                        TryGetAccessibleRole(hwnd, objectId, childId,
                            out var role) && IsPaletteItemRole(role))
                    {
                        Observe(WordFontColorSignal.PaletteItemCommitted,
                            paletteInteractionId);
                        return;
                    }
                    if (TryCapturePointerOrFocusSnapshot(out var invokedSnapshot))
                    {
                        if (invokedSnapshot.IsMoreColorsMenuItem &&
                            IsPaletteSessionActive())
                            Observe(WordFontColorSignal.MoreColorsOpened);
                        else if (invokedSnapshot.IsInsideFontColorPicker &&
                                 !invokedSnapshot.IsFontColorDropDown &&
                                 invokedSnapshot.IsMainButton)
                            ObserveMainButtonCommit();
                    }
                    return;
                }

                if (eventType == EventObjectFocus)
                {
                    if (IsWordDialogWindow(hwnd))
                        TrackOrRememberMoreColorsDialog(hwnd);
                    return;
                }

                if (eventType == EventObjectHide)
                {
                    var hiddenRoot = GetAncestor(hwnd, GetAncestorRoot);
                    lastPaletteHideForTest = hwnd.ToInt64() + "/" + objectId +
                        "/" + childId + "/root=" + hiddenRoot.ToInt64() +
                        "/paletteRoot=" + palettePopupRootHwnd.ToInt64();
                    if (TryCancelHiddenPalette(hwnd)) return;
                }

                bool moreColorsOpen;
                lock (stateGate) moreColorsOpen = interactionState.IsMoreColorsOpen;
                if (!moreColorsOpen) return;
                if (eventType == EventObjectHide && childId == 0 &&
                    objectId == ObjIdWindow && hwnd == moreColorsDialogHwnd)
                {
                    Observe(WordFontColorSignal.MoreColorsClosed,
                        GetMoreColorsInteractionId());
                }
            }
            catch
            {
                // Accessibility providers disappear with their popup. A missing or
                // malformed event fails closed and never applies a stale picker value.
            }
        }


        private void ObserveMainButtonCommit()
        {
            lock (stateGate)
            {
                var now = DateTime.UtcNow.Ticks;
                if (now <= mainButtonCommitDedupUntilUtcTicks) return;
                // The low-level mouse confirmation and an optional provider Invoke
                // describe the same split-button gesture. Accept the first signal
                // and suppress only its immediate duplicate.
                mainButtonCommitDedupUntilUtcTicks =
                    DateTime.UtcNow.AddMilliseconds(300).Ticks;
            }
            Observe(WordFontColorSignal.MainButtonInvoked);
        }

        private void ObserveFontSizeCommit()
        {
            lock (stateGate)
            {
                var now = DateTime.UtcNow.Ticks;
                if (now <= fontSizeCommitDedupUntilUtcTicks) return;
                fontSizeCommitDedupUntilUtcTicks =
                    DateTime.UtcNow.AddMilliseconds(300).Ticks;
                var began = formatTransactionState.Begin(WordFormatProperty.FontSize,
                    WordFormatInteractionOrigin.FontSizeControl,
                    out var canceledPrevious);
                QueueFormatSignalLocked(canceledPrevious);
                QueueFormatSignalLocked(began);
                QueueFormatSignalLocked(formatTransactionState.Commit(
                    began.InteractionId, WordFormatProperty.FontSize,
                    WordFormatInteractionOrigin.FontSizeControl));
            }
        }

        private void BeginFontSizeSession()
        {
            Interlocked.Exchange(ref fontSizeSessionUntilUtcTicks,
                DateTime.UtcNow.AddSeconds(30).Ticks);
        }

        private bool TryConsumeFontSizeSession()
        {
            var deadline = Interlocked.Exchange(ref fontSizeSessionUntilUtcTicks, 0);
            return deadline != 0 && DateTime.UtcNow.Ticks <= deadline;
        }

        private long GetMoreColorsInteractionId()
        {
            lock (stateGate) return moreColorsInteractionId;
        }

        private void Observe(WordFontColorSignal signal, long interactionId = 0)
        {
            bool committed;
            var closeDialog = false;
            var beginDialog = false;
            var stopPaletteCommit = false;
            var pendingDialog = IntPtr.Zero;
            lock (stateGate)
            {
                // A deferred native callback belongs only to the transaction that
                // scheduled it. A later palette/main-button interaction must never
                // be committed, cleared, or have its timer stopped by a stale token.
                if (signal == WordFontColorSignal.PaletteItemCommitted &&
                    interactionId != 0 &&
                    formatTransactionState.ActiveInteractionId != interactionId)
                    return;
                if ((signal == WordFontColorSignal.MoreColorsAccepted ||
                     signal == WordFontColorSignal.MoreColorsRejected ||
                     signal == WordFontColorSignal.MoreColorsCanceled ||
                     signal == WordFontColorSignal.MoreColorsClosed) &&
                    (interactionId == 0 ||
                     interactionId != moreColorsInteractionId))
                    return;
                if (signal == WordFontColorSignal.MoreColorsOpened)
                {
                    // Invoke and popup notifications can be duplicated or reordered.
                    // Once the transaction is open, a late duplicate must not erase
                    // the hwnd/hook that already belongs to the live dialog.
                    if (interactionState.IsMoreColorsOpen) return;
                    beginDialog = true;
                    moreColorsDialogHwnd = IntPtr.Zero;
                    pendingDialog = pendingMoreColorsDialogHwnd;
                    pendingMoreColorsDialogHwnd = IntPtr.Zero;
                }
                committed = interactionState.Observe(signal);
                switch (signal)
                {
                    case WordFontColorSignal.MainButtonInvoked:
                        var mainInteractionId = BeginFormatTransactionLocked(
                            WordFormatInteractionOrigin.FontColorMainButton);
                        QueueFormatSignalLocked(
                            formatTransactionState.Commit(mainInteractionId,
                                WordFormatProperty.TextColor,
                                WordFormatInteractionOrigin.FontColorMainButton));
                        stopPaletteCommit = true;
                        pendingPaletteCommitInteractionId = 0;
                        ClearPaletteCandidateLocked();
                        Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
                        break;
                    case WordFontColorSignal.PaletteItemCommitted:
                        if (interactionId == 0)
                            interactionId = EnsureFormatTransactionLocked(
                                WordFormatInteractionOrigin.FontColorPalette);
                        QueueFormatSignalLocked(
                            formatTransactionState.Commit(interactionId,
                                WordFormatProperty.TextColor,
                                WordFormatInteractionOrigin.FontColorPalette));
                        stopPaletteCommit = true;
                        pendingPaletteCommitInteractionId = 0;
                        ClearPaletteCandidateLocked();
                        Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
                        break;
                    case WordFontColorSignal.MoreColorsOpened:
                        interactionId = EnsureFormatTransactionLocked(
                            WordFormatInteractionOrigin.FontColorMoreColorsDialog);
                        formatTransactionState.UpdateOrigin(interactionId,
                            WordFormatInteractionOrigin.FontColorMoreColorsDialog);
                        moreColorsInteractionId = interactionId;
                        stopPaletteCommit = true;
                        pendingPaletteCommitInteractionId = 0;
                        ClearPaletteCandidateLocked();
                        break;
                    case WordFontColorSignal.MoreColorsCanceled:
                        QueueFormatSignalLocked(
                            formatTransactionState.Cancel(interactionId,
                                WordFormatProperty.TextColor,
                                WordFormatInteractionOrigin.
                                    FontColorMoreColorsDialog));
                        stopPaletteCommit = true;
                        break;
                    case WordFontColorSignal.MoreColorsClosed:
                        QueueFormatSignalLocked(committed
                            ? formatTransactionState.Commit(interactionId,
                                WordFormatProperty.TextColor,
                                WordFormatInteractionOrigin.
                                    FontColorMoreColorsDialog)
                            : formatTransactionState.Cancel(interactionId,
                                WordFormatProperty.TextColor,
                                    WordFormatInteractionOrigin.
                                    FontColorMoreColorsDialog));
                        stopPaletteCommit = true;
                        break;
                }
                if (signal == WordFontColorSignal.MoreColorsCanceled ||
                    signal == WordFontColorSignal.MoreColorsClosed)
                {
                    moreColorsDialogHwnd = IntPtr.Zero;
                    pendingMoreColorsDialogHwnd = IntPtr.Zero;
                    moreColorsInteractionId = 0;
                    pendingPaletteCommitInteractionId = 0;
                    ClearPaletteCandidateLocked();
                    closeDialog = true;
                }
            }
            if (beginDialog)
            {
                RemoveDialogMessageHook();
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
                if (pendingDialog != IntPtr.Zero)
                    TrackMoreColorsDialog(pendingDialog);
            }
            if (closeDialog) RemoveDialogMessageHook();
            if (stopPaletteCommit) StopPaletteCommitTimer();
        }

        private void BeginPaletteInteraction()
        {
            lock (stateGate)
            {
                // Native accessibility events can arrive out of order. Once a click
                // has started or been confirmed, rotating the token would discard
                // its mouse-up or generation-bound commit ticket.
                if (formatTransactionState.ActiveInteractionId != 0 &&
                    formatTransactionState.ActiveOrigin ==
                        WordFormatInteractionOrigin.FontColorPalette &&
                    (palettePressedInteractionId != 0 ||
                     pendingPaletteCommitInteractionId != 0))
                    return;
                BeginFormatTransactionLocked(
                    WordFormatInteractionOrigin.FontColorPalette);
            }
        }

        private long BeginFormatTransactionLocked(
            WordFormatInteractionOrigin origin)
        {
            var began = formatTransactionState.Begin(WordFormatProperty.TextColor,
                origin, out var canceledPrevious);
            QueueFormatSignalLocked(canceledPrevious);
            QueueFormatSignalLocked(began);
            return began.InteractionId;
        }

        private long EnsureFormatTransactionLocked(
            WordFormatInteractionOrigin origin)
        {
            var interactionId = formatTransactionState.ActiveInteractionId;
            return interactionId != 0
                ? interactionId
                : BeginFormatTransactionLocked(origin);
        }

        private void QueueFormatSignalLocked(WordFormatInteractionEventArgs signal)
        {
            if (signal == null) return;
            if (signal.Phase == WordFormatInteractionPhase.Began)
                Interlocked.Increment(ref formatBeginsForTest);
            else if (signal.Phase == WordFormatInteractionPhase.Committed)
                Interlocked.Increment(ref formatCommitsForTest);
            else if (signal.Phase == WordFormatInteractionPhase.Canceled)
                Interlocked.Increment(ref formatCancelsForTest);
            lastFormatSignalForTest = signal.InteractionId + "/" +
                signal.Phase + "/" + signal.Origin;
            pendingFormatSignals.Enqueue(signal);
            if (formatSignalDrainScheduled || disposed) return;
            try
            {
                if (dispatcher.IsDisposed || !dispatcher.IsHandleCreated) return;
                formatSignalDrainScheduled = true;
                dispatcher.BeginInvoke(new Action(DrainFormatSignals));
            }
            catch (ObjectDisposedException)
            {
                formatSignalDrainScheduled = false;
            }
            catch (InvalidOperationException)
            {
                formatSignalDrainScheduled = false;
            }
        }

        private void SchedulePaletteCommit(long interactionId)
        {
            if (interactionId == 0 || disposed) return;
            System.Threading.Timer previousTimer;
            lock (stateGate)
            {
                if (disposed || formatTransactionState.ActiveInteractionId !=
                        interactionId)
                    return;
                pendingPaletteCommitInteractionId = interactionId;
                var generation = ++paletteCommitGeneration;
                if (generation <= 0)
                {
                    paletteCommitGeneration = 1;
                    generation = 1;
                }
                var ticket = new PaletteCommitTicket(interactionId, generation);
                previousTimer = paletteCommitTimer;
                // WH_MOUSE_LL runs before Word receives WM_LBUTTONUP. Cross one
                // scheduler turn, then post the semantic commit through the Word UI
                // dispatcher. A longer fixed delay makes SVG formula paint visibly
                // trail Word's native text colour.
                paletteCommitTimer = new System.Threading.Timer(
                    PaletteCommitTimerElapsed, ticket,
                    PaletteCommitDelayMilliseconds, Timeout.Infinite);
            }
            DisposeTimer(previousTimer);
        }

        private void StopPaletteCommitTimer()
        {
            System.Threading.Timer timer;
            lock (stateGate)
            {
                pendingPaletteCommitInteractionId = 0;
                paletteCommitGeneration++;
                timer = paletteCommitTimer;
                paletteCommitTimer = null;
            }
            DisposeTimer(timer);
        }

        private void PaletteCommitTimerElapsed(object state)
        {
            var ticket = state as PaletteCommitTicket;
            if (ticket == null || disposed) return;
            System.Threading.Timer completedTimer;
            lock (stateGate)
            {
                if (disposed || ticket.Generation != paletteCommitGeneration ||
                    ticket.InteractionId != pendingPaletteCommitInteractionId ||
                    formatTransactionState.ActiveInteractionId !=
                        ticket.InteractionId)
                    return;
                completedTimer = paletteCommitTimer;
                paletteCommitTimer = null;
            }
            DisposeTimer(completedTimer);
            Observe(WordFontColorSignal.PaletteItemCommitted,
                ticket.InteractionId);
        }

        private static void DisposeTimer(System.Threading.Timer timer)
        {
            if (timer == null) return;
            try { timer.Dispose(); }
            catch (ObjectDisposedException) { }
        }

        private void DrainFormatSignals()
        {
            while (true)
            {
                WordFormatInteractionEventArgs formatSignal;
                lock (stateGate)
                {
                    if (disposed || pendingFormatSignals.Count == 0)
                    {
                        pendingFormatSignals.Clear();
                        formatSignalDrainScheduled = false;
                        return;
                    }
                    formatSignal = pendingFormatSignals.Dequeue();
                }
                try { FormatInteraction?.Invoke(this, formatSignal); }
                catch
                {
                    // Never let an integration failure escape through Word's native
                    // message pump or suppress the following transaction terminal.
                }
            }
        }

        private bool IsPaletteSessionActive()
        {
            return DateTime.UtcNow.Ticks <=
                   Interlocked.Read(ref paletteSessionUntilUtcTicks);
        }

        private static bool IsPaletteItemRole(int role)
        {
            // Office accessibility providers vary by build: the same Fluent gallery
            // swatch has been observed as ROLE_SYSTEM_LISTITEM and MENUITEM. UIA still
            // identifies the popup as NetUIGalleryButton/category in both cases.
            return role == RoleSystemListItem || role == RoleSystemMenuItem;
        }

        private void SetPaletteCandidate(IntPtr hwnd, int objectId, int childId,
            uint eventTime)
        {
            lock (stateGate)
            {
                var root = GetAncestor(hwnd, GetAncestorRoot);
                if (root != IntPtr.Zero) palettePopupRootHwnd = root;
                paletteCandidateHwnd = hwnd;
                paletteCandidateObjectId = objectId;
                paletteCandidateChildId = childId;
                paletteCandidateEventTime = eventTime;
                paletteCandidateInteractionId =
                    formatTransactionState.ActiveInteractionId;
                lastPaletteCandidateTupleForTest = hwnd.ToInt64() + "/" +
                    objectId + "/" + childId + "/tx=" +
                    paletteCandidateInteractionId;
                Interlocked.Increment(ref paletteCandidatesForTest);
                Interlocked.Exchange(ref paletteCandidateUntilUtcTicks,
                    long.MaxValue);
            }
        }

        private bool TryCancelHiddenPalette(IntPtr hwnd)
        {
            var stopCommit = false;
            lock (stateGate)
            {
                if (palettePopupRootHwnd == IntPtr.Zero ||
                    formatTransactionState.ActiveInteractionId == 0 ||
                    formatTransactionState.ActiveOrigin !=
                        WordFormatInteractionOrigin.FontColorPalette ||
                    pendingPaletteCommitInteractionId != 0 ||
                    !IsSameWindowTree(hwnd, palettePopupRootHwnd))
                    return false;
                var interactionId = formatTransactionState.ActiveInteractionId;
                QueueFormatSignalLocked(formatTransactionState.Cancel(interactionId,
                    WordFormatProperty.TextColor,
                    WordFormatInteractionOrigin.FontColorPalette));
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
                palettePopupRootHwnd = IntPtr.Zero;
                ClearPaletteCandidateLocked();
                stopCommit = true;
            }
            if (stopCommit) StopPaletteCommitTimer();
            return true;
        }

        private void ClearPaletteCandidateLocked()
        {
            paletteCandidateHwnd = IntPtr.Zero;
            paletteCandidateObjectId = 0;
            paletteCandidateChildId = 0;
                paletteCandidateEventTime = 0;
                paletteCandidateInteractionId = 0;
                palettePopupRootHwnd = IntPtr.Zero;
                ClearPalettePressLocked();
            Interlocked.Exchange(ref paletteCandidateUntilUtcTicks, 0);
        }

        private void ClearPalettePressLocked()
        {
            palettePressedInteractionId = 0;
            palettePressedHwnd = IntPtr.Zero;
        }

        private bool TryConsumePaletteCandidate(IntPtr hwnd, int objectId, int childId,
            uint eventTime, out long interactionId)
        {
            interactionId = 0;
            lock (stateGate)
            {
                lastPaletteInvokedTupleForTest = hwnd.ToInt64() + "/" +
                    objectId + "/" + childId + "/active=" +
                    formatTransactionState.ActiveInteractionId;
                // Office exposes the same swatch through two accessibility provider
                // projections: OBJECT_SELECTION identifies a gallery/menu child,
                // while OBJECT_INVOKED can identify a different proxy hwnd/object.
                // Pair the ordered events by the active semantic transaction and the
                // WinEvent generation timestamp, not by an object tuple that Office
                // does not keep stable between those two notifications.
                var eventDelta = unchecked(eventTime - paletteCandidateEventTime);
                var activeInteractionId =
                    formatTransactionState.ActiveInteractionId;
                var matches = DateTime.UtcNow.Ticks <=
                                  Interlocked.Read(ref paletteCandidateUntilUtcTicks) &&
                              eventDelta <= PaletteEventPairWindowMilliseconds &&
                              paletteCandidateInteractionId != 0 &&
                              activeInteractionId != 0 &&
                              formatTransactionState.ActiveOrigin ==
                                  WordFormatInteractionOrigin.FontColorPalette;
                if (!matches)
                {
                    if (DateTime.UtcNow.Ticks >
                        Interlocked.Read(ref paletteCandidateUntilUtcTicks))
                    {
                        ClearPaletteCandidateLocked();
                    }
                    return false;
                }
                // A delayed duplicate Expand notification can rotate the semantic
                // token after the hover event. The provider Invoke is still scoped by
                // the live palette session, role and event pair, so bind it to the
                // current palette token rather than rejecting the real gesture.
                interactionId = activeInteractionId;
                pendingPaletteCommitInteractionId = interactionId;
                ClearPaletteCandidateLocked();
                Interlocked.Increment(ref paletteMatchesForTest);
                return matches;
            }
        }

        private void TrackOrRememberMoreColorsDialog(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || disposed) return;
            var trackNow = false;
            lock (stateGate)
            {
                if (interactionState.IsMoreColorsOpen)
                    trackNow = true;
                else if (IsPaletteSessionActive())
                    pendingMoreColorsDialogHwnd = hwnd;
            }
            if (trackNow) TrackMoreColorsDialog(hwnd);
        }

        private void TrackMoreColorsDialog(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || disposed) return;
            var staleHook = IntPtr.Zero;
            lock (stateGate)
            {
                if (!interactionState.IsMoreColorsOpen) return;
                if (moreColorsDialogHwnd == hwnd &&
                    (wordProcessId != Process.GetCurrentProcess().Id ||
                     dialogMessageHook != IntPtr.Zero))
                    return;
                if (moreColorsDialogHwnd != hwnd)
                    staleHook = Interlocked.Exchange(ref dialogMessageHook,
                        IntPtr.Zero);
                moreColorsDialogHwnd = hwnd;
            }
            if (staleHook != IntPtr.Zero)
                UnhookWindowsHookEx(staleHook);
            if (wordProcessId != Process.GetCurrentProcess().Id) return;
            var threadId = GetWindowThreadProcessId(hwnd, out var processId);
            if (threadId == 0 || processId != unchecked((uint)wordProcessId)) return;
            var newHook = SetWindowsHookEx(HookCallWndProc,
                callWndProcHandler, IntPtr.Zero, threadId);
            if (newHook == IntPtr.Zero) return;
            var keepHook = false;
            lock (stateGate)
            {
                if (!disposed && interactionState.IsMoreColorsOpen &&
                    moreColorsDialogHwnd == hwnd && dialogMessageHook == IntPtr.Zero)
                {
                    dialogMessageHook = newHook;
                    keepHook = true;
                }
            }
            if (!keepHook) UnhookWindowsHookEx(newHook);
        }

        private IntPtr OnCallWndProc(int hookCode, IntPtr wParam, IntPtr lParam)
        {
            if (!disposed && hookCode >= HookCodeAction && lParam != IntPtr.Zero)
            {
                try
                {
                    var message = (CallWndProcMessage)Marshal.PtrToStructure(
                        lParam, typeof(CallWndProcMessage));
                    if (message.WindowHandle == moreColorsDialogHwnd)
                    {
                        if (message.Message == WindowMessageCommand)
                        {
                            var commandId = unchecked((int)(
                                message.WordParameter.ToInt64() & 0xffff));
                            if (commandId == DialogResultOk)
                                Observe(WordFontColorSignal.MoreColorsAccepted,
                                    GetMoreColorsInteractionId());
                            else if (commandId == DialogResultCancel)
                                Observe(WordFontColorSignal.MoreColorsRejected,
                                    GetMoreColorsInteractionId());
                        }
                        else if (message.Message == WindowMessageClose)
                        {
                            Observe(WordFontColorSignal.MoreColorsRejected,
                                GetMoreColorsInteractionId());
                        }
                    }
                }
                catch
                {
                    // A malformed/destroyed dialog message fails closed. The root
                    // HIDE event will clear the transaction without a commit.
                }
            }
            return CallNextHookEx(IntPtr.Zero, hookCode, wParam, lParam);
        }

        private IntPtr OnLowLevelMouse(int hookCode, IntPtr wParam, IntPtr lParam)
        {
            var message = unchecked((uint)wParam.ToInt64());
            if (!disposed && Volatile.Read(ref interactionContextEnabled) != 0 &&
                hookCode >= HookCodeAction && lParam != IntPtr.Zero &&
                (message == WindowMessageLeftButtonDown ||
                 message == WindowMessageLeftButtonUp))
            {
                try
                {
                    var mouse = (LowLevelMouseData)Marshal.PtrToStructure(
                        lParam, typeof(LowLevelMouseData));
                    var window = WindowFromPoint(mouse.Point);
                    GetWindowThreadProcessId(window, out var processId);
                    long interactionId = 0;
                    lock (stateGate)
                    {
                        var isTargetProcess = processId ==
                            unchecked((uint)wordProcessId);
                        var activeInteractionId =
                            formatTransactionState.ActiveInteractionId;
                        if (message == WindowMessageLeftButtonDown)
                        {
                            ClearPalettePressLocked();
                            if (isTargetProcess && activeInteractionId != 0 &&
                                formatTransactionState.ActiveOrigin ==
                                    WordFormatInteractionOrigin.FontColorPalette &&
                                DateTime.UtcNow.Ticks <= Interlocked.Read(
                                    ref paletteCandidateUntilUtcTicks) &&
                                paletteCandidateInteractionId != 0 &&
                                IsSameWindowTree(window, paletteCandidateHwnd))
                            {
                                // The latest native selection event and the click are
                                // in the same live popup. No provider hit-test or
                                // geometry query is needed.
                                palettePressedInteractionId = activeInteractionId;
                                palettePressedHwnd = paletteCandidateHwnd;
                            }
                        }
                        else
                        {
                            if (isTargetProcess &&
                                palettePressedInteractionId != 0 &&
                                activeInteractionId == palettePressedInteractionId &&
                                IsSameWindowTree(window, palettePressedHwnd))
                            {
                                interactionId = palettePressedInteractionId;
                                // Publish confirmation before releasing stateGate so a
                                // late duplicate Expanded cannot rotate the token in
                                // the tiny hook-to-timer scheduling window.
                                pendingPaletteCommitInteractionId = interactionId;
                                ClearPaletteCandidateLocked();
                                Interlocked.Increment(ref paletteMatchesForTest);
                            }
                            else
                            {
                                ClearPalettePressLocked();
                            }
                        }
                    }
                    if (interactionId != 0)
                        SchedulePaletteCommit(interactionId);
                    else if (message == WindowMessageLeftButtonUp &&
                             processId == unchecked((uint)wordProcessId) &&
                             IsOfficeFormattingWindow(window))
                        QueueFormattingControlHitTest(mouse.Point);
                }
                catch
                {
                    // A low-level hook must always return promptly. A disappearing
                    // popup fails closed and the palette cancellation path terminates
                    // the semantic transaction.
                }
            }
            return CallNextHookEx(lowLevelMouseHook, hookCode, wParam, lParam);
        }

        private IntPtr OnLowLevelKeyboard(int hookCode, IntPtr wParam, IntPtr lParam)
        {
            var message = unchecked((uint)wParam.ToInt64());
            if (!disposed && Volatile.Read(ref interactionContextEnabled) != 0 &&
                hookCode >= HookCodeAction && lParam != IntPtr.Zero &&
                (message == WindowMessageKeyDown ||
                 message == WindowMessageSystemKeyDown))
            {
                try
                {
                    var keyboard = (LowLevelKeyboardData)Marshal.PtrToStructure(
                        lParam, typeof(LowLevelKeyboardData));
                    if (keyboard.VirtualKey == VirtualKeyEscape)
                        CancelActivePaletteFromKeyboard();
                }
                catch
                {
                    // Never block global keyboard delivery. A malformed callback
                    // simply leaves native Word to close the popup normally.
                }
            }
            return CallNextHookEx(lowLevelKeyboardHook, hookCode, wParam, lParam);
        }

        private void CancelActivePaletteFromKeyboard()
        {
            lock (stateGate)
            {
                var interactionId = formatTransactionState.ActiveInteractionId;
                if (interactionId == 0 ||
                    formatTransactionState.ActiveOrigin !=
                        WordFormatInteractionOrigin.FontColorPalette ||
                    pendingPaletteCommitInteractionId != 0)
                    return;
                QueueFormatSignalLocked(formatTransactionState.Cancel(interactionId,
                    WordFormatProperty.TextColor,
                    WordFormatInteractionOrigin.FontColorPalette));
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
                palettePopupRootHwnd = IntPtr.Zero;
                ClearPaletteCandidateLocked();
            }
            StopPaletteCommitTimer();
        }

        private void QueueFormattingControlHitTest(NativePoint point)
        {
            // EVENT_OBJECT_INVOKED is optional for the main half of Office's split
            // button. Confirm the released control independently. Run UIA away from
            // the low-level hook so input delivery is never held up by a provider.
            ThreadPool.QueueUserWorkItem(_ =>
            {
                if (disposed) return;
                try
                {
                    var element = AutomationElement.FromPoint(
                        new System.Windows.Point(point.X, point.Y));
                    if (element == null) return;
                    var snapshot = ElementSnapshot.Capture(element, wordProcessId);
                    if (snapshot.IsInsideFontColorPicker &&
                        snapshot.IsFontColorDropDown)
                    {
                        Interlocked.Exchange(ref paletteSessionUntilUtcTicks,
                            long.MaxValue);
                        BeginPaletteInteraction();
                    }
                    else if (snapshot.IsInsideFontColorPicker &&
                        !snapshot.IsFontColorDropDown && snapshot.IsMainButton)
                        ObserveMainButtonCommit();
                    else if (snapshot.IsInsideFontSize &&
                             !snapshot.IsPopupChoice)
                        BeginFontSizeSession();
                    else if (snapshot.IsInsideFontSize && snapshot.IsPopupChoice)
                    {
                        TryConsumeFontSizeSession();
                        ObserveFontSizeCommit();
                    }
                    else if (TryConsumeFontSizeSession())
                        ObserveFontSizeCommit();
                }
                catch (ElementNotAvailableException) { }
                catch (InvalidOperationException) { }
            });
        }

        private static bool IsOfficeFormattingWindow(IntPtr window)
        {
            if (window == IntPtr.Zero) return false;
            var className = new StringBuilder(128);
            if (GetClassName(window, className, className.Capacity) <= 0)
                return false;
            // Ribbon controls and their galleries use NetUI windows. In particular,
            // never start UI Automation for clicks in Word's document canvas: one
            // provider call per ordinary click can build a backlog that delays both
            // later colour menus and WINWORD process exit.
            return className.ToString().StartsWith("NetUI",
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSameWindowTree(IntPtr first, IntPtr second)
        {
            if (first == IntPtr.Zero || second == IntPtr.Zero) return false;
            var firstRoot = GetAncestor(first, GetAncestorRoot);
            var secondRoot = GetAncestor(second, GetAncestorRoot);
            return firstRoot != IntPtr.Zero && firstRoot == secondRoot;
        }

        private void RemoveDialogMessageHook()
        {
            var hook = Interlocked.Exchange(ref dialogMessageHook, IntPtr.Zero);
            if (hook == IntPtr.Zero) return;
            try { UnhookWindowsHookEx(hook); }
            catch { }
        }

        public void Dispose()
        {
            DisposeCore();
        }

        internal void DisposeForHostShutdown()
        {
            DisposeCore();
        }

        private void DisposeCore()
        {
            if (disposed) return;
            disposed = true;
            StopPaletteCommitTimer();
            lock (stateGate)
            {
                pendingFormatSignals.Clear();
                formatSignalDrainScheduled = false;
                pendingPaletteCommitInteractionId = 0;
                ClearPaletteCandidateLocked();
            }
            RemoveAutomationHandlers();
        }

        private void RemoveAutomationHandlers()
        {
            RemoveDialogMessageHook();
            var mouseHook = Interlocked.Exchange(ref lowLevelMouseHook, IntPtr.Zero);
            if (mouseHook != IntPtr.Zero)
                try { UnhookWindowsHookEx(mouseHook); } catch { }
            var keyboardHook = Interlocked.Exchange(ref lowLevelKeyboardHook,
                IntPtr.Zero);
            if (keyboardHook != IntPtr.Zero)
                try { UnhookWindowsHookEx(keyboardHook); } catch { }
            Unhook(ref invokedHook);
            Unhook(ref selectionHook);
            Unhook(ref focusHook);
            Unhook(ref hideHook);
            started = false;
            // No global UIA handlers or retained desktop-root providers exist, so
            // shutdown never needs a UIA provider/RPC teardown.
        }

        private bool TryCapturePointerOrFocusSnapshot(out ElementSnapshot snapshot)
        {
            snapshot = null;
            try
            {
                var cursor = Cursor.Position;
                var element = AutomationElement.FromPoint(
                    new System.Windows.Point(cursor.X, cursor.Y)) ??
                    AutomationElement.FocusedElement;
                if (element == null) return false;
                snapshot = ElementSnapshot.Capture(element, wordProcessId);
                return snapshot.IsWordProcess;
            }
            catch (ElementNotAvailableException) { return false; }
            catch (InvalidOperationException) { return false; }
        }

        private static void Unhook(ref IntPtr hook)
        {
            if (hook == IntPtr.Zero) return;
            try { UnhookWinEvent(hook); }
            finally { hook = IntPtr.Zero; }
        }

        private static bool TryGetAccessibleRole(IntPtr hwnd, int objectId,
            int childId, out int role)
        {
            return TryGetAccessibleRoleAndBounds(hwnd, objectId, childId,
                out role, out _);
        }

        private static bool TryGetAccessibleRoleAndBounds(IntPtr hwnd, int objectId,
            int childId, out int role, out System.Windows.Rect bounds)
        {
            role = 0;
            bounds = System.Windows.Rect.Empty;
            IAccessible accessible = null;
            object child = null;
            try
            {
                if (AccessibleObjectFromEvent(hwnd, unchecked((uint)objectId),
                        unchecked((uint)childId), out accessible, out child) != 0 ||
                    accessible == null)
                    return false;
                role = Convert.ToInt32(accessible.get_accRole(child));
                try
                {
                    accessible.accLocation(out var left, out var top,
                        out var width, out var height, child);
                    if (width > 0 && height > 0)
                        bounds = new System.Windows.Rect(left, top, width, height);
                }
                catch (COMException)
                {
                    // Role is sufficient to recognize the event. The UIA hit-test
                    // bounds remain the safe fallback when MSAA omits accLocation.
                }
                return true;
            }
            catch (COMException) { return false; }
            catch (InvalidCastException) { return false; }
            finally
            {
                if (child != null && !ReferenceEquals(child, accessible) &&
                    Marshal.IsComObject(child))
                {
                    try { Marshal.ReleaseComObject(child); }
                    catch (COMException) { }
                    catch (InvalidComObjectException) { }
                }
                if (accessible != null && Marshal.IsComObject(accessible))
                {
                    try { Marshal.ReleaseComObject(accessible); }
                    catch (COMException) { }
                    catch (InvalidComObjectException) { }
                }
            }
        }

        private static bool IsWordDialogWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;
            var className = new StringBuilder(128);
            return GetClassName(hwnd, className, className.Capacity) > 0 &&
                   string.Equals(className.ToString(), "bosa_sdm_msword",
                       StringComparison.OrdinalIgnoreCase);
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException(nameof(WordFontColorMonitor));
        }

        private sealed class PaletteCommitTicket
        {
            internal PaletteCommitTicket(long interactionId, long generation)
            {
                InteractionId = interactionId;
                Generation = generation;
            }

            internal long InteractionId { get; }
            internal long Generation { get; }
        }

        private sealed class ElementSnapshot
        {
            private ElementSnapshot(bool isWordProcess,
                bool isInsideFontColorPicker, bool isFontColorDropDown,
                bool isMoreColorsButton, bool isInsideFontSize, string className,
                int controlTypeId)
            {
                IsWordProcess = isWordProcess;
                IsInsideFontColorPicker = isInsideFontColorPicker;
                IsFontColorDropDown = isFontColorDropDown;
                IsMoreColorsButton = isMoreColorsButton;
                IsInsideFontSize = isInsideFontSize;
                ClassName = className;
                ControlTypeId = controlTypeId;
            }

            internal bool IsWordProcess { get; }
            internal bool IsInsideFontColorPicker { get; }
            internal bool IsFontColorDropDown { get; }
            internal bool IsMoreColorsButton { get; }
            internal bool IsInsideFontSize { get; }
            internal string ClassName { get; }
            internal int ControlTypeId { get; }
            internal bool IsMainButton =>
                ControlTypeId == ControlType.Button.Id &&
                string.Equals(ClassName, "NetUIRibbonButton",
                    StringComparison.OrdinalIgnoreCase);
            internal bool IsPopupChoice =>
                ControlTypeId == ControlType.ListItem.Id ||
                ControlTypeId == ControlType.MenuItem.Id;
            internal bool IsMoreColorsMenuItem => IsMoreColorsButton ||
                (ControlTypeId == ControlType.MenuItem.Id &&
                 string.Equals(ClassName, "NetUITWBtnMenuItem",
                     StringComparison.OrdinalIgnoreCase));

            internal static ElementSnapshot Capture(AutomationElement element,
                int wordProcessId)
            {
                var processId = element.Current.ProcessId;
                var automationId = element.Current.AutomationId ?? string.Empty;
                var className = element.Current.ClassName ?? string.Empty;
                var controlType = element.Current.ControlType;
                var controlTypeId = controlType == null ? 0 : controlType.Id;
                if (processId != wordProcessId)
                    return new ElementSnapshot(false, false, false, false, false,
                        className, controlTypeId);

                var insidePicker = automationId == FontColorPickerId;
                var isDropDown = automationId == FontColorDropDownId;
                var isMoreColors = automationId == MoreColorsId;
                var isFontSize = automationId == FontSizeId;
                var current = element;
                for (var depth = 0; depth < MaximumAncestorDepth && current != null; depth++)
                {
                    string ancestorId;
                    try { ancestorId = current.Current.AutomationId ?? string.Empty; }
                    catch (ElementNotAvailableException) { break; }
                    insidePicker = insidePicker || ancestorId == FontColorPickerId;
                    isDropDown = isDropDown || ancestorId == FontColorDropDownId;
                    isMoreColors = isMoreColors || ancestorId == MoreColorsId;
                    isFontSize = isFontSize || ancestorId == FontSizeId;
                    try { current = TreeWalker.RawViewWalker.GetParent(current); }
                    catch (ElementNotAvailableException) { break; }
                }
                return new ElementSnapshot(true, insidePicker, isDropDown,
                    isMoreColors, isFontSize, className, controlTypeId);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CallWndProcMessage
        {
            internal IntPtr LongParameter;
            internal IntPtr WordParameter;
            internal uint Message;
            internal IntPtr WindowHandle;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativePoint
        {
            internal int X;
            internal int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LowLevelMouseData
        {
            internal NativePoint Point;
            internal uint MouseData;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LowLevelKeyboardData
        {
            internal uint VirtualKey;
            internal uint ScanCode;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint idEventThread,
            uint dwmsEventTime);

        private delegate IntPtr CallWndProcDelegate(int hookCode, IntPtr wordParameter,
            IntPtr longParameter);

        private delegate IntPtr LowLevelMouseDelegate(int hookCode,
            IntPtr wordParameter, IntPtr longParameter);

        private delegate IntPtr LowLevelKeyboardDelegate(int hookCode,
            IntPtr wordParameter, IntPtr longParameter);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
            uint idThread, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromEvent(IntPtr hwnd, uint dwId,
            uint dwChildId, out IAccessible accessible,
            [MarshalAs(UnmanagedType.Struct)] out object child);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder className,
            int maximumCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookId,
            CallWndProcDelegate hookProcedure, IntPtr moduleHandle, uint threadId);

        [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW",
            SetLastError = true)]
        private static extern IntPtr SetLowLevelMouseHook(int hookId,
            LowLevelMouseDelegate hookProcedure, IntPtr moduleHandle, uint threadId);

        [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW",
            SetLastError = true)]
        private static extern IntPtr SetLowLevelKeyboardHook(int hookId,
            LowLevelKeyboardDelegate hookProcedure, IntPtr moduleHandle,
            uint threadId);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int hookCode,
            IntPtr wordParameter, IntPtr longParameter);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd,
            out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(NativePoint point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetModuleHandle(string moduleName);
    }
}
