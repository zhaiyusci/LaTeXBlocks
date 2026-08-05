using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using LaTeXBlocks.PowerPoint;

namespace LaTeXBlocks.RenderHost
{
    /// <summary>
    /// Owns the StemTeX managed/native lifetime for exactly one render-host process.
    /// It intentionally knows nothing about Office, COM, selections, or documents.
    /// </summary>
    internal sealed class RenderHostEngine : IDisposable
    {
        internal const int MaximumRawSvgBytes = 5 * 1024 * 1024;
        private const int MaximumSourceCharacters = 250 * 1024;
        private const int MaximumProfileCharacters = 128;
        private const int MaximumSummaryCharacters = 64 * 1024;
        private const int MaximumOutcomeMessageCharacters = 16 * 1024;
        private const string ProtocolName = "LaTeXBlocks.RenderHost/1";

        private readonly object gate = new object();
        private readonly StemTeXBackend backend;
        private bool stopping;
        private bool disposed;

        internal RenderHostEngine()
        {
            // StemTeXBackend owns a dedicated FIFO worker. Its SwitchProfile method
            // starts native warm-up asynchronously, matching the StemTeX GUI model.
            backend = new StemTeXBackend();
        }

        internal bool IsStopping
        {
            get { lock (gate) return stopping; }
        }

        internal async Task<HostCommandOutcome> ExecuteAsync(RenderHostRequest request, string pipeName)
        {
            string command;
            var validation = ValidateEnvelope(request, out command);
            if (validation != null)
            {
                return new HostCommandOutcome(validation, false);
            }

            if (string.Equals(command, "ping", StringComparison.OrdinalIgnoreCase))
            {
                return new HostCommandOutcome(Success(request.Id, "pong", pipeName), false);
            }

            if (string.Equals(command, "shutdown", StringComparison.OrdinalIgnoreCase))
            {
                lock (gate)
                {
                    if (stopping)
                    {
                        return new HostCommandOutcome(Error(request.Id, "host_stopping",
                            "The render host is already shutting down."), false);
                    }
                    stopping = true;
                }
                return new HostCommandOutcome(Success(request.Id, "shutting_down", pipeName), true);
            }

            if (IsStopping)
            {
                return new HostCommandOutcome(Error(request.Id, "host_stopping",
                    "The render host is shutting down and no longer accepts work."), false);
            }

            if (string.Equals(command, "switchProfile", StringComparison.OrdinalIgnoreCase))
            {
                return SwitchProfile(request, pipeName);
            }

            if (string.Equals(command, "cancelLatestPreview", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    backend.CancelLatestPreview();
                    return new HostCommandOutcome(Success(request.Id, "latest_preview_cancelled", pipeName), false);
                }
                catch (Exception exception)
                {
                    return new HostCommandOutcome(ErrorForException(request.Id, exception), false);
                }
            }

            if (string.Equals(command, "renderLatest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(command, "renderQueued", StringComparison.OrdinalIgnoreCase))
            {
                return await RenderAsync(request, pipeName,
                    string.Equals(command, "renderLatest", StringComparison.OrdinalIgnoreCase)).ConfigureAwait(false);
            }

            return new HostCommandOutcome(Error(request.Id, "unsupported_command",
                "Supported commands are ping, switchProfile, renderLatest, renderQueued, " +
                "cancelLatestPreview, and shutdown."), false);
        }

        /// <summary>
        /// Called only after the shutdown response has been sent. No Office thread ever
        /// calls this: if native cleanup needs time, it belongs solely to this process.
        /// </summary>
        internal void CompleteShutdown()
        {
            lock (gate)
            {
                if (disposed)
                {
                    return;
                }
                disposed = true;
                stopping = true;
            }

            try { backend.Dispose(); }
            catch { }
        }

        public void Dispose()
        {
            lock (gate)
            {
                if (disposed)
                {
                    return;
                }
                stopping = true;
            }
            CompleteShutdown();
        }

        private HostCommandOutcome SwitchProfile(RenderHostRequest request, string pipeName)
        {
            try
            {
                var profile = ResolveProfile(request.Profile, true);
                backend.SwitchProfile(profile);
                var response = Success(request.Id, "profile_switch_queued", pipeName);
                response.Result.Profile = profile;
                response.Result.Status = backend.Status;
                response.Result.Profiles = backend.Profiles;
                return new HostCommandOutcome(response, false);
            }
            catch (Exception exception)
            {
                return new HostCommandOutcome(ErrorForException(request.Id, exception), false);
            }
        }

        private async Task<HostCommandOutcome> RenderAsync(RenderHostRequest request, string pipeName, bool latestOnly)
        {
            string profile;
            string source;
            double widthPt;
            bool autoWidth;
            double fontSizePt;
            RenderHostResponse error;
            if (!TryValidateRenderRequest(request, out profile, out source, out widthPt, out autoWidth,
                out fontSizePt, out error))
            {
                return new HostCommandOutcome(error, false);
            }

            try
            {
                StemTeXSvgResult render;
                if (latestOnly)
                {
                    render = await backend.RenderLatestAsync(profile, source, widthPt, autoWidth, fontSizePt)
                        .ConfigureAwait(false);
                }
                else
                {
                    render = await backend.RenderQueuedAsync(profile, source, widthPt, autoWidth, fontSizePt)
                        .ConfigureAwait(false);
                }

                if (render == null || render.Bytes == null)
                {
                    return new HostCommandOutcome(Error(request.Id, "render_failed",
                        "StemTeX returned no SVG output."), false);
                }
                if (render.Bytes.Length > MaximumRawSvgBytes)
                {
                    return new HostCommandOutcome(Error(request.Id, "render_too_large",
                        "The SVG output exceeds the host's 5 MiB raw-output limit."), false);
                }

                var response = Success(request.Id, latestOnly ? "render_latest" : "render_queued", pipeName);
                response.Result.Profile = profile;
                response.Result.Status = backend.Status;
                response.Result.SvgBase64 = Convert.ToBase64String(render.Bytes);
                response.Result.SummaryJson = LimitText(render.SummaryJson, MaximumSummaryCharacters);
                response.Result.OutcomeCode = render.OutcomeCode;
                response.Result.IssueFlags = render.IssueFlags;
                response.Result.OutcomeMessage = LimitText(render.OutcomeMessage, MaximumOutcomeMessageCharacters);
                response.Result.DepthPt = render.DepthPt;
                return new HostCommandOutcome(response, false);
            }
            catch (OperationCanceledException)
            {
                return new HostCommandOutcome(Error(request.Id, "cancelled",
                    "The render request was superseded or cancelled."), false);
            }
            catch (Exception exception)
            {
                return new HostCommandOutcome(ErrorForException(request.Id, exception), false);
            }
        }

        private RenderHostResponse ValidateEnvelope(RenderHostRequest request, out string command)
        {
            command = null;
            if (request == null)
            {
                return Error(null, "invalid_request", "The request is required.");
            }
            if (request.Id != null && request.Id.Length > 256)
            {
                return Error(null, "invalid_request", "The request id exceeds the 256-character limit.");
            }
            if (request.Version != PipeProtocol.Version)
            {
                return Error(request.Id, "unsupported_protocol",
                    "This host supports protocol version 1 only.");
            }
            command = (request.Command ?? string.Empty).Trim();
            if (command.Length == 0)
            {
                return Error(request.Id, "invalid_request", "A command is required.");
            }
            if (command.Length > 64)
            {
                return Error(request.Id, "invalid_request", "The command exceeds the 64-character limit.");
            }
            return null;
        }

        private bool TryValidateRenderRequest(RenderHostRequest request, out string profile, out string source,
            out double widthPt, out bool autoWidth, out double fontSizePt, out RenderHostResponse error)
        {
            profile = null;
            source = request.Source;
            widthPt = request.WidthPt ?? 360.0;
            autoWidth = request.AutoWidth ?? false;
            fontSizePt = request.FontSizePt ?? 10.0;
            error = null;

            try { profile = ResolveProfile(request.Profile, false); }
            catch (Exception exception)
            {
                error = ErrorForException(request.Id, exception);
                return false;
            }

            if (string.IsNullOrWhiteSpace(source))
            {
                error = Error(request.Id, "invalid_request", "LaTeX source cannot be empty.");
                return false;
            }
            if (source.Length > MaximumSourceCharacters)
            {
                error = Error(request.Id, "request_too_large",
                    "LaTeX source exceeds the 250,000-character host limit.");
                return false;
            }
            if (!IsFinite(widthPt) || widthPt <= 0 || widthPt > 50000)
            {
                error = Error(request.Id, "invalid_request",
                    "widthPt must be a finite value greater than zero and no more than 50,000.");
                return false;
            }
            if (!IsFinite(fontSizePt) || fontSizePt < 1 || fontSizePt > 200)
            {
                error = Error(request.Id, "invalid_request",
                    "fontSizePt must be a finite value from 1 through 200.");
                return false;
            }
            return true;
        }

        private string ResolveProfile(string profile, bool required)
        {
            if (profile == null || string.IsNullOrWhiteSpace(profile))
            {
                if (required)
                {
                    throw new ArgumentException("A profile is required.", nameof(profile));
                }
                return backend.DefaultAvailableProfile;
            }
            profile = profile.Trim();
            if (profile.Length > MaximumProfileCharacters)
            {
                throw new ArgumentException("The profile exceeds the 128-character limit.", nameof(profile));
            }
            foreach (var candidate in backend.Profiles)
            {
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }
            throw new ArgumentException("Unknown StemTeX profile: " + profile, nameof(profile));
        }

        private RenderHostResponse Success(string id, string kind, string pipeName)
        {
            return new RenderHostResponse
            {
                Version = PipeProtocol.Version,
                Id = id,
                Ok = true,
                Result = new RenderHostResult
                {
                    Kind = kind,
                    Protocol = ProtocolName,
                    ProcessId = Process.GetCurrentProcess().Id,
                    PipeName = pipeName,
                    Status = backend.Status
                }
            };
        }

        private static RenderHostResponse Error(string id, string code, string message)
        {
            return new RenderHostResponse
            {
                Version = PipeProtocol.Version,
                Id = id,
                Ok = false,
                Error = new RenderHostError
                {
                    Code = code,
                    Message = LimitText(message, 2048)
                }
            };
        }

        private static RenderHostResponse ErrorForException(string id, Exception exception)
        {
            if (exception is ArgumentException || exception is ArgumentOutOfRangeException)
            {
                return Error(id, "invalid_request", exception.Message);
            }
            if (exception is ObjectDisposedException)
            {
                return Error(id, "host_stopping", "The render host is shutting down.");
            }
            if (exception is StemTeXException)
            {
                return Error(id, "render_failed", exception.Message);
            }
            return Error(id, "render_failed", exception == null ? "Unknown render failure." : exception.Message);
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string LimitText(string value, int maximumCharacters)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maximumCharacters)
            {
                return value;
            }
            return value.Substring(0, maximumCharacters) + "…";
        }
    }

    internal sealed class HostCommandOutcome
    {
        internal HostCommandOutcome(RenderHostResponse response, bool shutdownAfterResponse)
        {
            Response = response ?? throw new ArgumentNullException(nameof(response));
            ShutdownAfterResponse = shutdownAfterResponse;
        }

        internal RenderHostResponse Response { get; }
        internal bool ShutdownAfterResponse { get; }
    }
}
