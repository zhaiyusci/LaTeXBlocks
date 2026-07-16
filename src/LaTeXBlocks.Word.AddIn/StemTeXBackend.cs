using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace LaTeXBlocks.Word
{
    // Mirrors the StemTeX GUI lifecycle: one global profile, one renderer, and one
    // dedicated FIFO worker thread for create/render/destroy.
    internal sealed class StemTeXBackend : IDisposable
    {
        internal const string DefaultProfile = "unicodemath_cjk";
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
                        throw new InvalidOperationException("StemTeX renderer did not initialize for profile " + profile + ".");
                    completion.SetResult(true);
                }
                catch (Exception exception) { completion.SetException(exception); }
            });
            completion.Task.GetAwaiter().GetResult();
        }

        internal Task<StemTeXSvgResult> RenderLatestAsync(string profile, string source, double widthPt, bool autoWidth)
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
                queue.Enqueue(() => ExecuteRender(canonical, source, widthPt, autoWidth, requestGeneration, requestId, completion));
            }
            wake.Set();
            return completion.Task;
        }

        internal StemTeXSvgResult RenderSvg(string profile, string source, double widthPt, bool autoWidth)
        {
            return RenderLatestAsync(profile, source, widthPt, autoWidth).GetAwaiter().GetResult();
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

        private void ExecuteRender(string profile, string source, double widthPt, bool autoWidth,
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
                var result = renderer.RenderSvg(source, widthPt, autoWidth);
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
            lock (gate)
            {
                if (stopping) return;
                stopping = true;
                ++generation;
                ++latestRequestId;
                queue.Clear();
            }
            wake.Set();
            worker.Join();
            wake.Dispose();
        }
    }
}
