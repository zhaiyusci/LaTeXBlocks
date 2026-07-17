using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace LaTeXBlocks.Word
{
    // Mirrors the StemTeX GUI lifecycle: one global profile, one renderer, and one
    // dedicated FIFO worker thread for create/render/destroy.
    internal sealed class StemTeXBackend : IDisposable
    {
        internal const string DefaultProfile = "xits_cjk";
        private readonly object gate = new object();
        private readonly Queue<Action> queue = new Queue<Action>();
        private readonly AutoResetEvent wake = new AutoResetEvent(false);
        private readonly Thread worker;
        private readonly string[] profiles;
        private readonly string stemTeXHome;
        private StemTeXRenderer renderer;
        private string rendererProfile;
        private string selectedProfile;
        private long generation;
        private long latestRequestId;
        private bool stopping;
        private string status = "not-started";
        private const int WorkerRestartingError = 5;
        private const int WorkerBusyError = 6;

        internal StemTeXBackend()
        {
            profiles = StemTeXRenderer.DiscoverProfiles();
            if (profiles.Length == 0) throw new InvalidOperationException("StemTeX has no usable profiles.");
            using (var probe = new StemTeXRenderer(DefaultAvailableProfile)) stemTeXHome = probe.StemTeXHome;
            worker = new Thread(WorkerLoop) { IsBackground = true, Name = "LaTeX Blocks StemTeX worker" };
            worker.Start();
        }

        internal string[] Profiles => (string[])profiles.Clone();
        internal string StemTeXHome => stemTeXHome;
        internal string Status { get { lock (gate) return status; } }

        internal string DefaultAvailableProfile
        {
            get
            {
                foreach (var profile in profiles)
                    if (string.Equals(profile, DefaultProfile, StringComparison.OrdinalIgnoreCase)) return profile;
                return profiles[0];
            }
        }

        internal void SwitchProfile(string profile)
        {
            var canonical = CanonicalProfile(profile);
            long requestedGeneration;
            lock (gate)
            {
                ThrowIfStopping();
                if (string.Equals(selectedProfile, canonical, StringComparison.OrdinalIgnoreCase) && generation != 0) return;
                selectedProfile = canonical;
                requestedGeneration = ++generation;
                ++latestRequestId;
                status = "warming:" + canonical;
                queue.Enqueue(() => InitializeRenderer(canonical, requestedGeneration));
            }
            wake.Set();
        }

        internal void WarmUp(string profile)
        {
            SwitchProfile(profile);
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            Post(() =>
            {
                try
                {
                    if (renderer == null || !string.Equals(rendererProfile, CanonicalProfile(profile), StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("StemTeX renderer did not initialize for profile " + profile +
                            ". Backend status: " + Status);
                    completion.SetResult(true);
                }
                catch (Exception exception) { completion.SetException(exception); }
            });
            completion.Task.GetAwaiter().GetResult();
        }

        internal Task<StemTeXSvgResult> RenderLatestAsync(string profile, string source, double widthPt, bool autoWidth,
            double fontSizePt = 10)
        {
            var canonical = CanonicalProfile(profile);
            SwitchProfile(canonical);
            var completion = new TaskCompletionSource<StemTeXSvgResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            long requestGeneration;
            long requestId;
            lock (gate)
            {
                ThrowIfStopping();
                requestGeneration = generation;
                requestId = ++latestRequestId;
                status = "queued:" + requestId;
                queue.Enqueue(() => ExecuteRender(canonical, source, widthPt, autoWidth, fontSizePt,
                    requestGeneration, requestId, completion));
            }
            wake.Set();
            return completion.Task;
        }

        internal StemTeXSvgResult RenderSvg(string profile, string source, double widthPt, bool autoWidth,
            double fontSizePt = 10)
        {
            return RenderLatestAsync(profile, source, widthPt, autoWidth, fontSizePt).GetAwaiter().GetResult();
        }

        private void InitializeRenderer(string profile, long requestedGeneration)
        {
            lock (gate) if (requestedGeneration != generation || stopping) return;
            StemTeXRenderer replacement = null;
            try
            {
                replacement = new StemTeXRenderer(profile);
                replacement.WarmUp();
                lock (gate)
                {
                    if (requestedGeneration != generation || stopping) return;
                    renderer?.Dispose();
                    renderer = replacement;
                    replacement = null;
                    rendererProfile = profile;
                    status = "ready:" + profile;
                }
            }
            catch (Exception exception)
            {
                lock (gate) if (requestedGeneration == generation) status = "failed:" + exception.Message;
            }
            finally { replacement?.Dispose(); }
        }

        private void ExecuteRender(string profile, string source, double widthPt, bool autoWidth, double fontSizePt,
            long requestGeneration, long requestId, TaskCompletionSource<StemTeXSvgResult> completion)
        {
            lock (gate)
            {
                if (stopping || requestGeneration != generation || requestId != latestRequestId)
                {
                    completion.SetCanceled();
                    return;
                }
                status = "rendering:" + requestId;
            }
            try
            {
                if (renderer == null || !string.Equals(rendererProfile, profile, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("StemTeX renderer is not ready for profile " + profile + ".");
                var result = RenderWithTransientRetry(source, widthPt, autoWidth, fontSizePt,
                    requestGeneration, requestId);
                lock (gate)
                {
                    if (stopping || requestGeneration != generation || requestId != latestRequestId)
                    {
                        completion.SetCanceled();
                        return;
                    }
                    status = "ready:" + profile;
                }
                completion.SetResult(result);
            }
            catch (Exception exception)
            {
                lock (gate) if (requestGeneration == generation && requestId == latestRequestId) status = "failed:" + exception.Message;
                completion.SetException(exception);
            }
        }

        private StemTeXSvgResult RenderWithTransientRetry(string source, double widthPt, bool autoWidth,
            double fontSizePt, long requestGeneration, long requestId)
        {
            try
            {
                return renderer.RenderSvg(source, widthPt, autoWidth, fontSizePt);
            }
            catch (StemTeXException exception) when (
                exception.ErrorCode == WorkerRestartingError || exception.ErrorCode == WorkerBusyError)
            {
                lock (gate)
                    if (stopping || requestGeneration != generation || requestId != latestRequestId)
                        throw new TaskCanceledException("StemTeX request was superseded.");
                // The hot worker can briefly report restarting/busy while recovering.
                // Retry this same latest request once; TeX snippet errors are never retried.
                Thread.Sleep(100);
                lock (gate)
                    if (stopping || requestGeneration != generation || requestId != latestRequestId)
                        throw new TaskCanceledException("StemTeX request was superseded.");
                return renderer.RenderSvg(source, widthPt, autoWidth, fontSizePt);
            }
        }

        private void Post(Action action)
        {
            lock (gate) { ThrowIfStopping(); queue.Enqueue(action); }
            wake.Set();
        }

        private void WorkerLoop()
        {
            while (true)
            {
                Action action = null;
                lock (gate)
                {
                    if (queue.Count > 0) action = queue.Dequeue();
                    else if (stopping) break;
                }
                if (action != null) action();
                else wake.WaitOne();
            }
            renderer?.Dispose();
            renderer = null;
        }

        private string CanonicalProfile(string profile)
        {
            foreach (var candidate in profiles)
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase)) return candidate;
            throw new ArgumentException("Unknown StemTeX profile: " + profile, nameof(profile));
        }

        private void ThrowIfStopping()
        {
            if (stopping) throw new ObjectDisposedException(nameof(StemTeXBackend));
        }

        public void Dispose()
        {
            StemTeXRenderer activeRenderer;
            lock (gate)
            {
                if (stopping) return;
                stopping = true;
                ++generation;
                ++latestRequestId;
                queue.Clear();
                status = "stopping";
                activeRenderer = renderer;
            }
            try { activeRenderer?.CancelCurrent(); } catch { }
            // Native cancellation normally recovers a worker for future requests.
            // During host shutdown there will be no future request, so terminate this
            // Word process's owned worker tree instead of waiting for recovery/rebuild.
            CancelOwnedWorkerTrees();
            wake.Set();
            // Office invokes add-in Shutdown on Word's UI thread. Shutdown is only a
            // cancellation signal: the background worker owns renderer destruction and
            // the StemTeX lifetime pipe owns the helper process. In particular, never
            // Join here—even a bounded wait makes closing Word visibly sluggish.
        }

        internal bool WaitForStopForTest(int millisecondsTimeout)
        {
            return worker.Join(millisecondsTimeout);
        }

        private void CancelOwnedWorkerTrees()
        {
            // stemtex_renderer_create does not publish its renderer pointer until the
            // warm-up process has completed. During that narrow window CancelCurrent
            // cannot address it, so terminate only the worker-host process tree created
            // by this Word process and this StemTeX installation. This is the process
            // equivalent of the GUI invalidating an in-flight renderer generation.
            var expectedHost = Path.GetFullPath(Path.Combine(stemTeXHome, "runtime", "bin", "windows",
                "stemtex-worker-host.exe"));
            var snapshot = ProcessTreeSnapshot.Capture();
            var ownerPid = Process.GetCurrentProcess().Id;
            foreach (var entry in snapshot)
            {
                if (entry.ParentProcessId != ownerPid ||
                    !string.Equals(entry.ExecutableName, "stemtex-worker-host.exe", StringComparison.OrdinalIgnoreCase))
                    continue;
                try
                {
                    using (var host = Process.GetProcessById(entry.ProcessId))
                    {
                        if (!string.Equals(Path.GetFullPath(host.MainModule.FileName), expectedHost,
                            StringComparison.OrdinalIgnoreCase)) continue;
                    }
                    ProcessTreeSnapshot.KillTree(entry.ProcessId, snapshot);
                }
                catch { }
            }
        }

        private sealed class ProcessTreeSnapshot
        {
            private const uint SnapshotProcesses = 0x00000002;
            private static readonly IntPtr InvalidHandle = new IntPtr(-1);

            internal int ProcessId;
            internal int ParentProcessId;
            internal string ExecutableName;

            internal static List<ProcessTreeSnapshot> Capture()
            {
                var result = new List<ProcessTreeSnapshot>();
                var handle = CreateToolhelp32Snapshot(SnapshotProcesses, 0);
                if (handle == InvalidHandle) return result;
                try
                {
                    var entry = new ProcessEntry32 { Size = (uint)Marshal.SizeOf(typeof(ProcessEntry32)) };
                    if (!Process32First(handle, ref entry)) return result;
                    do
                    {
                        result.Add(new ProcessTreeSnapshot
                        {
                            ProcessId = unchecked((int)entry.ProcessId),
                            ParentProcessId = unchecked((int)entry.ParentProcessId),
                            ExecutableName = entry.ExecutableName
                        });
                    } while (Process32Next(handle, ref entry));
                    return result;
                }
                finally { CloseHandle(handle); }
            }

            internal static void KillTree(int rootProcessId, List<ProcessTreeSnapshot> snapshot)
            {
                foreach (var child in snapshot)
                    if (child.ParentProcessId == rootProcessId) KillTree(child.ProcessId, snapshot);
                try { using (var process = Process.GetProcessById(rootProcessId)) process.Kill(); } catch { }
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            private struct ProcessEntry32
            {
                internal uint Size;
                internal uint Usage;
                internal uint ProcessId;
                internal IntPtr DefaultHeapId;
                internal uint ModuleId;
                internal uint Threads;
                internal uint ParentProcessId;
                internal int BasePriority;
                internal uint Flags;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] internal string ExecutableName;
            }

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);
            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool Process32First(IntPtr snapshot, ref ProcessEntry32 entry);
            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool Process32Next(IntPtr snapshot, ref ProcessEntry32 entry);
            [DllImport("kernel32.dll")]
            private static extern bool CloseHandle(IntPtr handle);
        }
    }
}
