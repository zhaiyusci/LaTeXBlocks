using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    // Mirrors the StemTeX GUI lifecycle: one global profile, one renderer, and one
    // dedicated FIFO worker thread for create/render/destroy.
    internal sealed class StemTeXBackend : IDisposable
    {
        internal const string DefaultProfile = "xits_cjk";
        private readonly object gate = new object();
        private readonly Queue<BackendWorkItem> queue = new Queue<BackendWorkItem>();
        private readonly AutoResetEvent wake = new AutoResetEvent(false);
        private readonly Thread worker;
        private readonly string[] profiles;
        private readonly string stemTeXHome;
        private StemTeXRenderer renderer;
        private string rendererProfile;
        private string selectedProfile;
        private long generation;
        private long nextRequestId;
        private long latestRequestId;
        private bool stopping;
        private string status = "not-started";
        private BackendWorkItem activeWorkItem;
        private const int WorkerRestartingError = 5;
        private const int WorkerBusyError = 6;
        private const int ShutdownReaperTimeoutMs = 100000;

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
                queue.Enqueue(new BackendWorkItem(
                    () => InitializeRenderer(canonical, requestedGeneration), null));
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
                    completion.TrySetResult(true);
                }
                catch (Exception exception) { completion.TrySetException(exception); }
            }, () => completion.TrySetCanceled());
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
                requestId = ++nextRequestId;
                latestRequestId = requestId;
                status = "queued:" + requestId;
                queue.Enqueue(new BackendWorkItem(
                    () => ExecuteRender(canonical, source, widthPt, autoWidth, fontSizePt,
                        requestGeneration, requestId, true, completion),
                    () => completion.TrySetCanceled()));
            }
            wake.Set();
            return completion.Task;
        }

        // Document mutations are durable work, unlike live previews. They retain FIFO
        // order and are canceled only by a profile replacement or host shutdown; a
        // later preview must never silently discard a width/font change already
        // committed by the user.
        internal Task<StemTeXSvgResult> RenderQueuedAsync(string profile, string source,
            double widthPt, bool autoWidth, double fontSizePt = 10)
        {
            var canonical = CanonicalProfile(profile);
            SwitchProfile(canonical);
            var completion = new TaskCompletionSource<StemTeXSvgResult>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            long requestGeneration;
            long requestId;
            lock (gate)
            {
                ThrowIfStopping();
                requestGeneration = generation;
                requestId = ++nextRequestId;
                status = "queued:" + requestId;
                queue.Enqueue(new BackendWorkItem(
                    () => ExecuteRender(canonical, source, widthPt, autoWidth, fontSizePt,
                        requestGeneration, requestId, false, completion),
                    () => completion.TrySetCanceled()));
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
            StemTeXRenderer retired = null;
            try
            {
                replacement = new StemTeXRenderer(profile);
                replacement.WarmUp();
                lock (gate)
                {
                    if (requestedGeneration != generation || stopping) return;
                    retired = renderer;
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
            finally
            {
                DisposeRendererUnlessHostIsStopping(retired);
                DisposeRendererUnlessHostIsStopping(replacement);
            }
        }

        private void DisposeRendererUnlessHostIsStopping(StemTeXRenderer candidate)
        {
            if (candidate == null) return;
            bool abandon;
            lock (gate) abandon = stopping;
            // Native destroy belongs to the backend worker, but it must never run while
            // holding gate: Office shutdown needs that lock only long enough to publish
            // managed cancellation. Once shutdown starts, the reaper and process exit
            // own native cleanup instead.
            if (!abandon)
            {
                try { candidate.Dispose(); }
                catch (Exception exception)
                {
                    lock (gate) if (!stopping) status = "failed:" + exception.Message;
                }
            }
        }

        private void ExecuteRender(string profile, string source, double widthPt, bool autoWidth,
            double fontSizePt, long requestGeneration, long requestId, bool latestOnly,
            TaskCompletionSource<StemTeXSvgResult> completion)
        {
            lock (gate)
            {
                if (stopping || requestGeneration != generation ||
                    (latestOnly && requestId != latestRequestId))
                {
                    completion.TrySetCanceled();
                    return;
                }
                status = "rendering:" + requestId;
            }
            try
            {
                if (renderer == null || !string.Equals(rendererProfile, profile, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("StemTeX renderer is not ready for profile " + profile + ".");
                var result = RenderWithTransientRetry(source, widthPt, autoWidth, fontSizePt,
                    requestGeneration, requestId, latestOnly);
                lock (gate)
                {
                    if (stopping || requestGeneration != generation ||
                        (latestOnly && requestId != latestRequestId))
                    {
                        completion.TrySetCanceled();
                        return;
                    }
                    status = "ready:" + profile;
                }
                completion.TrySetResult(result);
            }
            catch (Exception exception)
            {
                lock (gate)
                    if (requestGeneration == generation &&
                        (!latestOnly || requestId == latestRequestId))
                        status = "failed:" + exception.Message;
                completion.TrySetException(exception);
            }
        }

        private StemTeXSvgResult RenderWithTransientRetry(string source, double widthPt, bool autoWidth,
            double fontSizePt, long requestGeneration, long requestId, bool latestOnly)
        {
            try
            {
                return renderer.RenderSvg(source, widthPt, autoWidth, fontSizePt);
            }
            catch (StemTeXException exception) when (
                exception.ErrorCode == WorkerRestartingError || exception.ErrorCode == WorkerBusyError)
            {
                lock (gate)
                    if (stopping || requestGeneration != generation ||
                        (latestOnly && requestId != latestRequestId))
                        throw new TaskCanceledException("StemTeX request was superseded.");
                // The hot worker can briefly report restarting/busy while recovering.
                // Retry this same latest request once; TeX snippet errors are never retried.
                Thread.Sleep(100);
                lock (gate)
                    if (stopping || requestGeneration != generation ||
                        (latestOnly && requestId != latestRequestId))
                        throw new TaskCanceledException("StemTeX request was superseded.");
                return renderer.RenderSvg(source, widthPt, autoWidth, fontSizePt);
            }
        }

        private void Post(Action action, Action cancel = null)
        {
            lock (gate) { ThrowIfStopping(); queue.Enqueue(new BackendWorkItem(action, cancel)); }
            wake.Set();
        }

        private void WorkerLoop()
        {
            while (true)
            {
                BackendWorkItem item = null;
                lock (gate)
                {
                    if (queue.Count > 0)
                    {
                        item = queue.Dequeue();
                        activeWorkItem = item;
                    }
                    else if (stopping) break;
                }
                if (item != null)
                {
                    try { item.Execute(); }
                    catch (Exception exception)
                    {
                        lock (gate) if (!stopping) status = "failed:" + exception.Message;
                        item.Cancel();
                    }
                    finally
                    {
                        lock (gate)
                            if (ReferenceEquals(activeWorkItem, item)) activeWorkItem = null;
                    }
                }
                else wake.WaitOne();
            }
            // StemTeXBackend is process-lifetime infrastructure. During Word host
            // shutdown, native destroy can block in child-process waits while VSTO is
            // still on Word's UI shutdown path. Owned helpers are terminated by the
            // background reaper; the renderer/DLL are intentionally abandoned for the
            // OS to reclaim when WINWORD exits.
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
            BackendWorkItem[] abandoned;
            BackendWorkItem active;
            lock (gate)
            {
                if (stopping) return;
                stopping = true;
                ++generation;
                ++latestRequestId;
                abandoned = queue.ToArray();
                queue.Clear();
                status = "stopping";
                active = activeWorkItem;
            }
            foreach (var item in abandoned) item.Cancel();
            active?.Cancel();
            wake.Set();
            StartShutdownReaper();
            // Office invokes add-in Shutdown on Word's UI thread. This method performs
            // no native cancellation, process enumeration, destroy, or Join. A native
            // cancel can wait up to the worker termination timeout; even a bounded wait
            // here makes closing Word visibly sluggish.
        }

        internal bool WaitForStopForTest(int millisecondsTimeout)
        {
            return worker.Join(millisecondsTimeout);
        }

        internal bool WorkerHasActiveItemForTest
        {
            get { lock (gate) return activeWorkItem != null; }
        }

        internal bool HasOwnedWorkerHostForTest
        {
            get
            {
                var snapshot = ProcessTreeSnapshot.Capture();
                return FindOwnedWorkerHosts(snapshot).Count > 0;
            }
        }

        private void StartShutdownReaper()
        {
            var reaper = new Thread(() =>
            {
                var timer = Stopwatch.StartNew();
                do
                {
                    TerminateOwnedWorkerTrees();
                    if (worker.Join(100)) break;
                    Thread.Sleep(50);
                } while (timer.ElapsedMilliseconds < ShutdownReaperTimeoutMs);
                TerminateOwnedWorkerTrees();
            })
            {
                IsBackground = true,
                Name = "LaTeX Blocks StemTeX shutdown reaper"
            };
            try { reaper.Start(); } catch { }
        }

        private void TerminateOwnedWorkerTrees()
        {
            // stemtex_renderer_create does not publish its renderer pointer until the
            // warm-up process has completed. Repeatedly terminate only the worker-host
            // process tree created by this Word process and this StemTeX installation;
            // this also catches a host born after the initial shutdown snapshot.
            var snapshot = ProcessTreeSnapshot.Capture();
            foreach (var entry in FindOwnedWorkerHosts(snapshot))
            {
                try
                {
                    ProcessTreeSnapshot.KillTree(entry.ProcessId, snapshot);
                }
                catch { }
            }
        }

        private List<ProcessTreeSnapshot> FindOwnedWorkerHosts(List<ProcessTreeSnapshot> snapshot)
        {
            var matches = new List<ProcessTreeSnapshot>();
            var expectedHost = Path.GetFullPath(Path.Combine(stemTeXHome, "runtime", "bin", "windows",
                "stemtex-worker-host.exe"));
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
                    matches.Add(entry);
                }
                catch { }
            }
            return matches;
        }

        private sealed class BackendWorkItem
        {
            private readonly Action execute;
            private readonly Action cancel;

            internal BackendWorkItem(Action execute, Action cancel)
            {
                this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
                this.cancel = cancel;
            }

            internal void Execute() { execute(); }
            internal void Cancel()
            {
                try { cancel?.Invoke(); } catch { }
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
