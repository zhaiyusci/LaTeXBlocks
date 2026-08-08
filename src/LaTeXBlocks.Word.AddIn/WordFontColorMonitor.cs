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
        void SuppressProgrammaticInvocations();
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
        private const int HookLowLevelMouse = 14;
        private const int HookCodeAction = 0;
        private const uint GetAncestorRoot = 2;
        private const uint WindowMessageLeftButtonDown = 0x0201;
        private const uint WindowMessageLeftButtonUp = 0x0202;
        private const uint WindowMessageClose = 0x0010;
        private const uint WindowMessageCommand = 0x0111;
        private const int DialogResultOk = 1;
        private const int DialogResultCancel = 2;
        private const int PaletteCloseGraceMilliseconds = 1500;
        private const int PaletteCommitDelayMilliseconds = 100;
        private const uint PaletteEventPairWindowMilliseconds = 3000;

        private readonly Control dispatcher;
        private readonly int wordProcessId;
        private readonly AutomationEventHandler automationHandler;
        private readonly AutomationPropertyChangedEventHandler propertyChangedHandler;
        private readonly WinEventDelegate winEventHandler;
        private readonly CallWndProcDelegate callWndProcHandler;
        private readonly LowLevelMouseDelegate lowLevelMouseHandler;
        private readonly System.Threading.Timer paletteCancelTimer;
        private System.Threading.Timer paletteCommitTimer;
        private readonly WordFontColorInteractionState interactionState =
            new WordFontColorInteractionState();
        private int interactionContextEnabled;
        private readonly WordFormatTransactionState formatTransactionState =
            new WordFormatTransactionState();
        private readonly object stateGate = new object();
        private readonly Queue<WordFormatInteractionEventArgs> pendingFormatSignals =
            new Queue<WordFormatInteractionEventArgs>();
        private AutomationElement automationRoot;
        private long suppressInvocationsUntilUtcTicks;
        private int suppressNextPickerInvocation;
        private long mainButtonCommitDedupUntilUtcTicks;
        private long fontSizeCommitDedupUntilUtcTicks;
        private long fontSizeSessionUntilUtcTicks;
        private long paletteSessionUntilUtcTicks;
        private long paletteCandidateUntilUtcTicks;
        private IntPtr paletteCandidateHwnd;
        private int paletteCandidateObjectId;
        private int paletteCandidateChildId;
        private uint paletteCandidateEventTime;
        private double paletteCandidateLeft;
        private double paletteCandidateTop;
        private double paletteCandidateRight;
        private double paletteCandidateBottom;
        private long paletteCandidateInteractionId;
        private long palettePressedInteractionId;
        private double palettePressedLeft;
        private double palettePressedTop;
        private double palettePressedRight;
        private double palettePressedBottom;
        private long pendingPaletteCommitInteractionId;
        private long paletteCommitGeneration;
        private IntPtr hideHook;
        private IntPtr focusHook;
        private IntPtr selectionHook;
        private IntPtr invokedHook;
        private IntPtr lowLevelMouseHook;
        private IntPtr dialogMessageHook;
        private IntPtr moreColorsDialogHwnd;
        private IntPtr pendingMoreColorsDialogHwnd;
        private long moreColorsInteractionId;
        private long paletteCancellationInteractionId;
        private int palettePropertyChangesForTest;
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
            automationHandler = OnAutomationEvent;
            propertyChangedHandler = OnAutomationPropertyChanged;
            winEventHandler = OnWinEvent;
            callWndProcHandler = OnCallWndProc;
            lowLevelMouseHandler = OnLowLevelMouse;
            paletteCancelTimer = new System.Threading.Timer(
                PaletteCancelTimerElapsed, null, Timeout.Infinite, Timeout.Infinite);
        }

        public event EventHandler<WordFormatInteractionEventArgs> FormatInteraction;

        internal bool IsRunning => started && !disposed;
        internal string DiagnosticStateForTest =>
            "property=" + Volatile.Read(ref palettePropertyChangesForTest) +
            ", selection=" + Volatile.Read(ref paletteSelectionsForTest) +
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
            ", session=" + IsPaletteSessionActive();

        public void SuppressProgrammaticInvocations()
        {
            // ExecuteMso used by the collapsed-caret probe is normally silent in UIA,
            // but Office builds are allowed to surface an Invoke event. Consume only
            // that immediate echo; the expiry prevents a missing echo from swallowing
            // the user's next real command.
            Interlocked.Exchange(ref suppressNextPickerInvocation, 1);
            Interlocked.Exchange(ref suppressInvocationsUntilUtcTicks,
                DateTime.UtcNow.AddMilliseconds(100).Ticks);
        }

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
                paletteCancellationInteractionId = 0;
                ClearPaletteCandidateLocked();
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
            }
        }

        public void Start()
        {
            ThrowIfDisposed();
            if (started) return;
            var root = AutomationElement.RootElement;
            automationRoot = root;
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
                    var hasRole = TryGetAccessibleRoleAndBounds(hwnd, objectId,
                        childId, out var role, out var accessibleBounds);
                    Volatile.Write(ref lastPaletteSelectionRoleForTest, role);
                    var isGalleryItem = IsGalleryItemAtPointerOrFocus(
                        out var pointerClass, out var candidateBounds);
                    if (!accessibleBounds.IsEmpty && accessibleBounds.Width > 0 &&
                        accessibleBounds.Height > 0)
                        candidateBounds = accessibleBounds;
                    lastPalettePointerClassForTest = pointerClass;
                    var isFontColorSession = IsPaletteSessionActive();
                    if (!isFontColorSession && hasRole &&
                        IsPaletteItemRole(role) && isGalleryItem)
                        isFontColorSession = TryBeginVisiblyExpandedPaletteInteraction();
                    if (isFontColorSession && hasRole &&
                        IsPaletteItemRole(role) && isGalleryItem)
                        SetPaletteCandidate(hwnd, objectId, childId, eventTime,
                            candidateBounds);
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
                                 invokedSnapshot.IsMainButton &&
                                 !TryConsumeProgrammaticInvocationSuppression())
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

        private void OnAutomationPropertyChanged(object sender,
            AutomationPropertyChangedEventArgs eventArgs)
        {
            if (disposed || eventArgs == null ||
                eventArgs.Property != ExpandCollapsePattern.ExpandCollapseStateProperty ||
                !(sender is AutomationElement element))
                return;
            ElementSnapshot snapshot;
            try { snapshot = ElementSnapshot.Capture(element, wordProcessId); }
            catch (ElementNotAvailableException) { return; }
            catch (InvalidOperationException) { return; }
            if (!snapshot.IsWordProcess ||
                (snapshot.AutomationId != FontColorPickerId &&
                 snapshot.AutomationId != FontColorDropDownId))
                return;
            Interlocked.Increment(ref palettePropertyChangesForTest);
            int expandState;
            try { expandState = Convert.ToInt32(eventArgs.NewValue); }
            catch (FormatException) { return; }
            catch (InvalidCastException) { return; }
            // UIA property callbacks are asynchronous and can arrive after the same
            // control has already moved to a later state. Re-read the live picker so
            // a stale Expanded cannot create a phantom transaction after commit, and
            // a stale Collapsed cannot shorten a newly reopened session.
            if (TryGetCurrentExpandCollapseState(element, out var currentState) &&
                currentState != (ExpandCollapseState)expandState)
                return;
            if (expandState == (int)ExpandCollapseState.Expanded)
            {
                var previousSession = Interlocked.Exchange(
                    ref paletteSessionUntilUtcTicks, long.MaxValue);
                if (previousSession != long.MaxValue)
                    BeginPaletteInteraction();
                return;
            }
            // Office collapses the popup just before it reports the selected swatch
            // or More Colors invocation. Keep the hovered candidate alive during the
            // short close-ordering grace: some builds publish Collapsed before the
            // low-level mouse-up that confirms the actual click.
            if (Interlocked.Read(ref paletteSessionUntilUtcTicks) != 0)
            {
                var closeDeadline = DateTime.UtcNow.AddMilliseconds(
                    PaletteCloseGraceMilliseconds).Ticks;
                lock (stateGate)
                {
                    if (paletteCandidateInteractionId != 0)
                        Interlocked.Exchange(ref paletteCandidateUntilUtcTicks,
                            closeDeadline);
                }
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, closeDeadline);
                SchedulePaletteCancellation();
            }
        }

        private void OnAutomationEvent(object sender, AutomationEventArgs eventArgs)
        {
            if (disposed || !(sender is AutomationElement element) || eventArgs == null)
                return;

            ElementSnapshot snapshot;
            try
            {
                snapshot = ElementSnapshot.Capture(element, wordProcessId);
            }
            catch (ElementNotAvailableException)
            {
                // The native OBJECT_HIDE hook owns dialog close because UIA commonly
                // invalidates a WindowClosed sender before its properties can be read.
                return;
            }
            catch (InvalidOperationException)
            {
                return;
            }
            if (!snapshot.IsWordProcess) return;

            if (eventArgs.EventId == WindowPattern.WindowOpenedEvent)
            {
                if (string.Equals(snapshot.ClassName, "bosa_sdm_msword",
                        StringComparison.OrdinalIgnoreCase) &&
                    snapshot.NativeWindowHandle != IntPtr.Zero)
                    TrackOrRememberMoreColorsDialog(snapshot.NativeWindowHandle);
                return;
            }

            if (eventArgs.EventId == WindowPattern.WindowClosedEvent)
            {
                if (string.Equals(snapshot.ClassName, "bosa_sdm_msword",
                        StringComparison.OrdinalIgnoreCase) &&
                    moreColorsDialogHwnd != IntPtr.Zero &&
                    snapshot.NativeWindowHandle == moreColorsDialogHwnd)
                    Observe(WordFontColorSignal.MoreColorsClosed,
                        GetMoreColorsInteractionId());
                return;
            }

            if (eventArgs.EventId != InvokePattern.InvokedEvent) return;
            if (snapshot.IsInsideFontColorPicker &&
                TryConsumeProgrammaticInvocationSuppression())
                return;
            if (snapshot.IsMoreColorsMenuItem && IsPaletteSessionActive())
            {
                Observe(WordFontColorSignal.MoreColorsOpened);
                return;
            }
            // Fluent gallery popups live in a separate NetUIToolWindow and are not
            // descendants of the Ribbon's FontColorPicker element. The active
            // transaction is the scope boundary for those external popup items;
            // requiring IsInsideFontColorPicker here would discard the actual swatch
            // Invoke while still observing its hover-only MSAA selection event.
            if (snapshot.IsPaletteItem && IsPaletteSessionActive())
            {
                ObserveCurrentPaletteCommit();
                return;
            }
            if (!snapshot.IsInsideFontColorPicker || snapshot.IsFontColorDropDown)
                return;

            // Current Office builds expose the split-button anchor as
            // AutomationId=FontColorPicker but put InvokePattern on its empty-id
            // NetUIRibbonButton child. Palette swatches are NetUIGalleryButton
            // ListItems. More Colors is an empty-id NetUITWBtnMenuItem rather than
            // exposing its idMso through UIA, so classify by control type/class too.
            if (snapshot.IsMainButton)
                ObserveMainButtonCommit();
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

        private void ObserveCurrentPaletteCommit()
        {
            long interactionId = 0;
            lock (stateGate)
            {
                var activeInteractionId =
                    formatTransactionState.ActiveInteractionId;
                if (activeInteractionId != 0 &&
                    formatTransactionState.ActiveOrigin ==
                        WordFormatInteractionOrigin.FontColorPalette)
                {
                    interactionId = activeInteractionId;
                    pendingPaletteCommitInteractionId = interactionId;
                }
            }
            if (interactionId != 0)
                Observe(WordFontColorSignal.PaletteItemCommitted, interactionId);
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
            var stopPaletteCancellation = false;
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
                        stopPaletteCancellation = true;
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
                        stopPaletteCancellation = true;
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
                        stopPaletteCancellation = true;
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
                        stopPaletteCancellation = true;
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
                        stopPaletteCancellation = true;
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
            if (stopPaletteCancellation) StopPaletteCancellationTimer();
            if (stopPaletteCommit) StopPaletteCommitTimer();
        }

        private void BeginPaletteInteraction()
        {
            lock (stateGate)
            {
                // UIA may report a late duplicate Expanded after the popup has
                // already collapsed for a real click. Once that click has started or
                // been confirmed, rotating the token would discard its mouse-up or
                // generation-bound commit ticket. Treat only that window as a
                // duplicate; a genuine reopen after Escape has neither marker.
                if (formatTransactionState.ActiveInteractionId != 0 &&
                    formatTransactionState.ActiveOrigin ==
                        WordFormatInteractionOrigin.FontColorPalette &&
                    (palettePressedInteractionId != 0 ||
                     pendingPaletteCommitInteractionId != 0))
                    return;
                BeginFormatTransactionLocked(
                    WordFormatInteractionOrigin.FontColorPalette);
            }
            StopPaletteCancellationTimer();
        }

        private long BeginFormatTransactionLocked(
            WordFormatInteractionOrigin origin)
        {
            var began = formatTransactionState.Begin(WordFormatProperty.TextColor,
                origin, out var canceledPrevious);
            QueueFormatSignalLocked(canceledPrevious);
            QueueFormatSignalLocked(began);
            paletteCancellationInteractionId = 0;
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

        private void SchedulePaletteCancellation()
        {
            lock (stateGate)
            {
                if (disposed || formatTransactionState.ActiveInteractionId == 0 ||
                    interactionState.IsMoreColorsOpen)
                    return;
                paletteCancellationInteractionId =
                    formatTransactionState.ActiveInteractionId;
            }
            try
            {
                paletteCancelTimer.Change(PaletteCloseGraceMilliseconds,
                    Timeout.Infinite);
            }
            catch (ObjectDisposedException) { }
        }

        private void StopPaletteCancellationTimer()
        {
            try { paletteCancelTimer.Change(Timeout.Infinite, Timeout.Infinite); }
            catch (ObjectDisposedException) { }
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
                // WH_MOUSE_LL runs before Word receives WM_LBUTTONUP. Defer the
                // semantic commit so the Word UI thread first applies the native
                // colour/MRU update; the dispatcher drain then reconciles formulas.
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

        private void PaletteCancelTimerElapsed(object state)
        {
            if (disposed) return;
            var sessionDeadline = Interlocked.Read(ref paletteSessionUntilUtcTicks);
            var remainingTicks = sessionDeadline == long.MaxValue
                ? long.MaxValue
                : sessionDeadline - DateTime.UtcNow.Ticks;
            if (remainingTicks > 0)
            {
                if (remainingTicks == long.MaxValue) return;
                var remainingMilliseconds = Math.Max(1,
                    (int)Math.Min(int.MaxValue, remainingTicks / TimeSpan.TicksPerMillisecond + 1));
                try { paletteCancelTimer.Change(remainingMilliseconds, Timeout.Infinite); }
                catch (ObjectDisposedException) { }
                return;
            }

            lock (stateGate)
            {
                if (disposed || paletteCancellationInteractionId == 0 ||
                    formatTransactionState.ActiveInteractionId !=
                        paletteCancellationInteractionId ||
                    interactionState.IsMoreColorsOpen)
                    return;
                QueueFormatSignalLocked(
                    formatTransactionState.Cancel(
                        paletteCancellationInteractionId,
                        WordFormatProperty.TextColor,
                        WordFormatInteractionOrigin.FontColorPalette));
                paletteCancellationInteractionId = 0;
                pendingPaletteCommitInteractionId = 0;
                ClearPaletteCandidateLocked();
                Interlocked.Exchange(ref paletteSessionUntilUtcTicks, 0);
            }
            StopPaletteCommitTimer();
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

        private static bool TryGetCurrentExpandCollapseState(
            AutomationElement element, out ExpandCollapseState state)
        {
            state = ExpandCollapseState.Collapsed;
            var current = element;
            for (var depth = 0;
                 depth < MaximumAncestorDepth && current != null;
                 depth++)
            {
                try
                {
                    if (current.TryGetCurrentPattern(
                            ExpandCollapsePattern.Pattern, out var pattern))
                    {
                        state = ((ExpandCollapsePattern)pattern).Current.
                            ExpandCollapseState;
                        return true;
                    }
                    current = TreeWalker.RawViewWalker.GetParent(current);
                }
                catch (ElementNotAvailableException) { return false; }
                catch (InvalidOperationException) { return false; }
            }
            return false;
        }

        private bool TryBeginVisiblyExpandedPaletteInteraction()
        {
            // WINEVENT_OUTOFCONTEXT and UIA property notifications are both
            // asynchronous. Some Office builds deliver OBJECT_SELECTION for a
            // swatch before the earlier ExpandCollapse property callback reaches us.
            // Confirm the stable FontColorPicker itself is visibly expanded, then
            // establish the same semantic transaction the delayed callback would
            // have established. This is deliberately narrower than accepting an
            // arbitrary NetUIGalleryButton event.
            AutomationElement picker = null;
            try
            {
                var root = automationRoot;
                if (root == null) return false;
                picker = root.FindFirst(TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ProcessIdProperty,
                            wordProcessId),
                        new PropertyCondition(AutomationElement.AutomationIdProperty,
                            FontColorPickerId)));
                if (picker == null || !picker.TryGetCurrentPattern(
                        ExpandCollapsePattern.Pattern, out var pattern) ||
                    ((ExpandCollapsePattern)pattern).Current.ExpandCollapseState !=
                        ExpandCollapseState.Expanded)
                    return false;
            }
            catch (ElementNotAvailableException) { return false; }
            catch (InvalidOperationException) { return false; }

            Interlocked.Exchange(ref paletteSessionUntilUtcTicks, long.MaxValue);
            BeginPaletteInteraction();
            return true;
        }

        private static bool IsPaletteItemRole(int role)
        {
            // Office accessibility providers vary by build: the same Fluent gallery
            // swatch has been observed as ROLE_SYSTEM_LISTITEM and MENUITEM. UIA still
            // identifies the popup as NetUIGalleryButton/category in both cases.
            return role == RoleSystemListItem || role == RoleSystemMenuItem;
        }

        private bool TryConsumeProgrammaticInvocationSuppression()
        {
            if (DateTime.UtcNow.Ticks >
                Interlocked.Read(ref suppressInvocationsUntilUtcTicks))
            {
                Interlocked.Exchange(ref suppressNextPickerInvocation, 0);
                return false;
            }
            return Interlocked.Exchange(ref suppressNextPickerInvocation, 0) == 1;
        }

        private void SetPaletteCandidate(IntPtr hwnd, int objectId, int childId,
            uint eventTime, System.Windows.Rect bounds)
        {
            lock (stateGate)
            {
                paletteCandidateHwnd = hwnd;
                paletteCandidateObjectId = objectId;
                paletteCandidateChildId = childId;
                paletteCandidateEventTime = eventTime;
                paletteCandidateLeft = bounds.Left;
                paletteCandidateTop = bounds.Top;
                paletteCandidateRight = bounds.Right;
                paletteCandidateBottom = bounds.Bottom;
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

        private void ClearPaletteCandidateLocked()
        {
            paletteCandidateHwnd = IntPtr.Zero;
            paletteCandidateObjectId = 0;
            paletteCandidateChildId = 0;
            paletteCandidateEventTime = 0;
            paletteCandidateInteractionId = 0;
            paletteCandidateLeft = 0;
            paletteCandidateTop = 0;
            paletteCandidateRight = 0;
            paletteCandidateBottom = 0;
            ClearPalettePressLocked();
            Interlocked.Exchange(ref paletteCandidateUntilUtcTicks, 0);
        }

        private void ClearPalettePressLocked()
        {
            palettePressedInteractionId = 0;
            palettePressedLeft = 0;
            palettePressedTop = 0;
            palettePressedRight = 0;
            palettePressedBottom = 0;
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
                                IsSameWindowTree(window, paletteCandidateHwnd) &&
                                mouse.Point.X >= paletteCandidateLeft &&
                                mouse.Point.X <= paletteCandidateRight &&
                                mouse.Point.Y >= paletteCandidateTop &&
                                mouse.Point.Y <= paletteCandidateBottom)
                            {
                                // As with provider Invoke, a late duplicate Expand may
                                // rotate the token between hover and mouse-down. The
                                // same live popup window is the gesture boundary.
                                palettePressedInteractionId = activeInteractionId;
                                palettePressedLeft = paletteCandidateLeft;
                                palettePressedTop = paletteCandidateTop;
                                palettePressedRight = paletteCandidateRight;
                                palettePressedBottom = paletteCandidateBottom;
                            }
                        }
                        else
                        {
                            if (isTargetProcess &&
                                palettePressedInteractionId != 0 &&
                                activeInteractionId == palettePressedInteractionId &&
                                mouse.Point.X >= palettePressedLeft &&
                                mouse.Point.X <= palettePressedRight &&
                                mouse.Point.Y >= palettePressedTop &&
                                mouse.Point.Y <= palettePressedBottom)
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
                        !snapshot.IsFontColorDropDown && snapshot.IsMainButton &&
                        !TryConsumeProgrammaticInvocationSuppression())
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
            try { paletteCancelTimer.Dispose(); }
            catch (ObjectDisposedException) { }
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
            Unhook(ref invokedHook);
            Unhook(ref selectionHook);
            Unhook(ref focusHook);
            Unhook(ref hideHook);
            automationRoot = null;
            started = false;
            // No global UIA handlers are registered. Clearing the root reference is
            // sufficient and, unlike RemoveAutomationEventHandler, cannot start a
            // provider/RPC teardown during Word exit.
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

        private static bool IsGalleryItemAtPointerOrFocus(out string className,
            out System.Windows.Rect bounds)
        {
            className = string.Empty;
            bounds = System.Windows.Rect.Empty;
            AutomationElement element = null;
            try
            {
                var cursor = Cursor.Position;
                element = AutomationElement.FromPoint(
                    new System.Windows.Point(cursor.X, cursor.Y));
                if (IsGalleryItem(element, out className, out bounds)) return true;
                element = AutomationElement.FocusedElement;
                return IsGalleryItem(element, out className, out bounds);
            }
            catch (ElementNotAvailableException) { return false; }
            catch (InvalidOperationException) { return false; }
        }

        private static bool IsGalleryItem(AutomationElement element,
            out string className, out System.Windows.Rect bounds)
        {
            className = string.Empty;
            bounds = System.Windows.Rect.Empty;
            if (element == null) return false;
            try
            {
                className = element.Current.ClassName ?? string.Empty;
                var isGalleryItem = string.Equals(className,
                                        "NetUIGalleryButton",
                                        StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(className,
                                        "NetUIGalleryCategoryContainer",
                                        StringComparison.OrdinalIgnoreCase);
                if (isGalleryItem) bounds = element.Current.BoundingRectangle;
                return isGalleryItem && !bounds.IsEmpty &&
                       bounds.Width > 0 && bounds.Height > 0;
            }
            catch (ElementNotAvailableException) { return false; }
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
            private ElementSnapshot(bool isWordProcess, string automationId,
                bool isInsideFontColorPicker, bool isFontColorDropDown,
                bool isMoreColorsButton, bool isInsideFontSize, string className,
                int controlTypeId, IntPtr nativeWindowHandle)
            {
                IsWordProcess = isWordProcess;
                AutomationId = automationId;
                IsInsideFontColorPicker = isInsideFontColorPicker;
                IsFontColorDropDown = isFontColorDropDown;
                IsMoreColorsButton = isMoreColorsButton;
                IsInsideFontSize = isInsideFontSize;
                ClassName = className;
                ControlTypeId = controlTypeId;
                NativeWindowHandle = nativeWindowHandle;
            }

            internal bool IsWordProcess { get; }
            internal string AutomationId { get; }
            internal bool IsInsideFontColorPicker { get; }
            internal bool IsFontColorDropDown { get; }
            internal bool IsMoreColorsButton { get; }
            internal bool IsInsideFontSize { get; }
            internal string ClassName { get; }
            internal int ControlTypeId { get; }
            internal IntPtr NativeWindowHandle { get; }
            internal bool IsMainButton =>
                ControlTypeId == ControlType.Button.Id &&
                string.Equals(ClassName, "NetUIRibbonButton",
                    StringComparison.OrdinalIgnoreCase);
            internal bool IsPaletteItem =>
                ControlTypeId == ControlType.ListItem.Id ||
                string.Equals(ClassName, "NetUIGalleryButton",
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
                var nativeWindowHandle = new IntPtr(
                    element.Current.NativeWindowHandle);
                if (processId != wordProcessId)
                    return new ElementSnapshot(false, automationId, false, false,
                        false, false, className, controlTypeId, nativeWindowHandle);

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
                return new ElementSnapshot(true, automationId, insidePicker,
                    isDropDown, isMoreColors, isFontSize, className, controlTypeId,
                    nativeWindowHandle);
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

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint idEventThread,
            uint dwmsEventTime);

        private delegate IntPtr CallWndProcDelegate(int hookCode, IntPtr wordParameter,
            IntPtr longParameter);

        private delegate IntPtr LowLevelMouseDelegate(int hookCode,
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
