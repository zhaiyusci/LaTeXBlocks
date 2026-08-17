using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LaTeXBlocks.Word
{
    /// <summary>
    /// Observes the end of Word's native mouse-capture gesture without subclassing
    /// an Office window. Word has no object-model event for a completed InlineShape
    /// resize while that object remains selected. Ribbon commands deliberately do not
    /// use this signal: opening or canceling a gallery can end capture without having
    /// committed anything.
    ///
    /// The WinEvent callback intentionally does no COM work.  It only queues a
    /// one-shot UI timer on the VSTO thread; Word then has completed its own input
    /// handling before the subscriber reads Shape geometry or starts rendering.
    /// </summary>
    internal sealed class WordMouseCaptureMonitor : IDisposable
    {
        private const uint EventSystemCaptureStart = 0x0008;
        private const uint EventSystemCaptureEnd = 0x0009;
        private const uint WineventOutOfContext = 0;
        private const int WhMouseLl = 14;
        private const int WmLButtonDown = 0x0201;
        private const int WmLButtonUp = 0x0202;

        private readonly Control dispatcher;
        private readonly WinEventDelegate callback;
        private readonly LowLevelMouseDelegate mouseCallback;
        private readonly Timer captureEndTimer;
        private readonly uint observedProcessId;
        private IntPtr hook;
        private IntPtr mouseHook;
        private bool mouseGestureInObservedProcess;
        private bool disposed;

        internal WordMouseCaptureMonitor(Control dispatcher, int observedProcessId = 0)
        {
            this.dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            this.observedProcessId = unchecked((uint)(observedProcessId > 0
                ? observedProcessId
                : Process.GetCurrentProcess().Id));
            callback = OnWinEvent;
            mouseCallback = OnLowLevelMouse;
            captureEndTimer = new Timer { Interval = 20 };
            captureEndTimer.Tick += CaptureEndTimer_Tick;
        }

        /// <summary>
        /// Raised on the Word UI thread after Word begins any native mouse-capture
        /// gesture.  Subscribers may snapshot their current selected object here.
        /// </summary>
        internal event EventHandler CaptureStarted;

        /// <summary>
        /// Raised once on the Word UI thread after a native capture ends.  The
        /// one-shot delay is a message-queue boundary, not a geometry poll.
        /// </summary>
        internal event EventHandler CaptureEnded;

        internal bool IsRunning => (hook != IntPtr.Zero || mouseHook != IntPtr.Zero) && !disposed;

        // Exposed to the isolated Word smoke so a passing accessibility hook
        // cannot mask failure of the mouse-up fallback this monitor now needs.
        internal bool IsMouseFallbackRunning => mouseHook != IntPtr.Zero && !disposed;

        internal void Start()
        {
            ThrowIfDisposed();
            if (hook != IntPtr.Zero || mouseHook != IntPtr.Zero) return;

            // A VSTO add-in is loaded into WINWORD.EXE, so this scoping includes
            // every Word UI thread but excludes mouse capture in other programs.
            hook = SetWinEventHook(EventSystemCaptureStart, EventSystemCaptureEnd,
                IntPtr.Zero, callback, observedProcessId, 0,
                WineventOutOfContext);

            // Word's modern drawing layer does not consistently expose every resize
            // drag through EVENT_SYSTEM_CAPTUREEND. A low-level mouse hook is only a
            // completion signal: it never reads Office objects and is filtered to a
            // foreground window owned by this Word process. The subscriber still
            // compares exact pre/post geometry, so normal clicks and Ribbon actions
            // do not render anything.
            // WH_MOUSE_LL runs the callback on this UI thread, but Windows still
            // requires a module handle for a desktop-wide (threadId == 0) hook.
            // The host executable is WINWORD.EXE in production and the smoke-test
            // executable under test, so the current process module is the correct
            // lifetime anchor in both cases.
            mouseHook = SetWindowsHookEx(WhMouseLl, mouseCallback,
                GetModuleHandle(null), 0);
            if (hook == IntPtr.Zero && mouseHook == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    "Word could not register a native resize-completion monitor.");
        }

        private IntPtr OnLowLevelMouse(int code, IntPtr message, IntPtr data)
        {
            if (code >= 0 && !disposed)
            {
                var messageId = unchecked((int)message.ToInt64());
                if (messageId == WmLButtonDown)
                {
                    mouseGestureInObservedProcess = IsObservedProcessForeground();
                    if (mouseGestureInObservedProcess)
                        Post(RaiseCaptureStarted);
                }
                else if (messageId == WmLButtonUp && mouseGestureInObservedProcess)
                {
                    mouseGestureInObservedProcess = false;
                    Post(ArmCaptureEndTimer);
                }
            }
            return CallNextHookEx(mouseHook, code, message, data);
        }

        private bool IsObservedProcessForeground()
        {
            var window = GetForegroundWindow();
            if (window == IntPtr.Zero) return false;
            GetWindowThreadProcessId(window, out var processId);
            return processId == observedProcessId;
        }

        private void Post(Action action)
        {
            try
            {
                if (!dispatcher.IsDisposed && dispatcher.IsHandleCreated)
                    dispatcher.BeginInvoke(action);
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void OnWinEvent(IntPtr hookHandle, uint eventType, IntPtr hwnd,
            int objectId, int childId, uint eventThread, uint eventTime)
        {
            // This callback can arrive at an arbitrary point in Word's own input
            // processing.  In particular, do not touch Application, Selection,
            // Shape, or any other COM object here.
            if (disposed) return;
            try
            {
                if (dispatcher.IsDisposed || !dispatcher.IsHandleCreated) return;
                if (eventType == EventSystemCaptureStart)
                    dispatcher.BeginInvoke(new Action(RaiseCaptureStarted));
                else if (eventType == EventSystemCaptureEnd)
                    dispatcher.BeginInvoke(new Action(ArmCaptureEndTimer));
            }
            catch (ObjectDisposedException)
            {
                // See above.  WinEvent callbacks can race UnhookWinEvent.
            }
            catch (InvalidOperationException)
            {
                // Word is tearing down its hidden VSTO dispatcher window.
            }
        }

        private void RaiseCaptureStarted()
        {
            if (disposed) return;
            try { CaptureStarted?.Invoke(this, EventArgs.Empty); }
            catch
            {
                // A host integration failure must never escape through the WinForms
                // message pump or interfere with Word's native gesture.
            }
        }

        private void ArmCaptureEndTimer()
        {
            if (disposed) return;
            // Coalesce any duplicate capture-end accessibility notifications into a
            // single UI turn.  This is a one-shot debounce, never a repeating poll.
            captureEndTimer.Stop();
            captureEndTimer.Start();
        }

        private void CaptureEndTimer_Tick(object sender, EventArgs e)
        {
            captureEndTimer.Stop();
            if (disposed) return;
            try { CaptureEnded?.Invoke(this, EventArgs.Empty); }
            catch
            {
                // The add-in's own failure path handles render errors on the UI
                // thread.  Do not let an observer tear down Word's message loop.
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            captureEndTimer.Stop();
            captureEndTimer.Tick -= CaptureEndTimer_Tick;
            captureEndTimer.Dispose();
            if (hook != IntPtr.Zero)
            {
                try { UnhookWinEvent(hook); }
                finally { hook = IntPtr.Zero; }
            }
            if (mouseHook != IntPtr.Zero)
            {
                try { UnhookWindowsHookEx(mouseHook); }
                finally { mouseHook = IntPtr.Zero; }
            }
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException(nameof(WordMouseCaptureMonitor));
        }

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime);

        private delegate IntPtr LowLevelMouseDelegate(int code, IntPtr message, IntPtr data);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
            uint idThread, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelMouseDelegate callback, IntPtr module, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code,
            IntPtr message, IntPtr data);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window,
            out uint processId);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);
    }
}
