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

        private readonly Control dispatcher;
        private readonly WinEventDelegate callback;
        private readonly Timer captureEndTimer;
        private IntPtr hook;
        private bool disposed;

        internal WordMouseCaptureMonitor(Control dispatcher)
        {
            this.dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            callback = OnWinEvent;
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

        internal bool IsRunning => hook != IntPtr.Zero && !disposed;

        internal void Start()
        {
            ThrowIfDisposed();
            if (hook != IntPtr.Zero) return;

            // A VSTO add-in is loaded into WINWORD.EXE, so this scoping includes
            // every Word UI thread but excludes mouse capture in other programs.
            hook = SetWinEventHook(EventSystemCaptureStart, EventSystemCaptureEnd,
                IntPtr.Zero, callback, unchecked((uint)Process.GetCurrentProcess().Id), 0,
                WineventOutOfContext);
            if (hook == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    "Word could not register the native mouse-capture monitor.");
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
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException(nameof(WordMouseCaptureMonitor));
        }

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
            uint idThread, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);
    }
}
