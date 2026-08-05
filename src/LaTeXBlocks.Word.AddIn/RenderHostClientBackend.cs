using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    /// <summary>
    /// Office-side proxy for one disposable RenderHost process.  It deliberately
    /// never loads stemtex-renderer.dll: all potentially unbounded native work
    /// (create, warm-up, rendering, and destroy) belongs to the child process.
    /// Consequently VSTO may unload this AppDomain without waiting for XeTeX.
    /// </summary>
    internal sealed class RenderHostClientBackend : IStemTeXBackend
    {
        private const string PipePrefix = "LaTeXBlocks.RenderHost.";
        private const int ProtocolVersion = 1;
        private const int MaximumRequestFrameBytes = 1024 * 1024;
        private const int MaximumResponseFrameBytes = 8 * 1024 * 1024;
        private const int PipeConnectTimeoutMs = 5000;
        private const int WarmUpTimeoutMs = 90000;
        private const uint JobObjectExtendedLimitInformationClass = 9;
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;

        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
        private static readonly DataContractJsonSerializer RequestSerializer =
            new DataContractJsonSerializer(typeof(RenderHostClientRequest));
        private static readonly DataContractJsonSerializer ResponseSerializer =
            new DataContractJsonSerializer(typeof(RenderHostClientResponse));

        private readonly object gate = new object();
        private readonly string[] profiles;
        private readonly string stemTeXHome;
        private readonly string pipeName;
        private readonly Process hostProcess;
        private IntPtr hostJob;
        private string selectedProfile;
        private string status = "starting";
        private bool disposed;

        internal RenderHostClientBackend()
        {
            profiles = StemTeXRenderer.DiscoverProfiles();
            if (profiles.Length == 0)
                throw new InvalidOperationException("StemTeX has no usable profiles.");
            using (var probe = new StemTeXRenderer(DefaultAvailableProfile))
                stemTeXHome = probe.StemTeXHome;

            pipeName = PipePrefix + Guid.NewGuid().ToString("N");
            hostProcess = StartHost(ResolveHostExecutable(), pipeName.Substring(PipePrefix.Length));
            hostJob = TryCreateKillOnCloseJob(hostProcess);
        }

        public string[] Profiles { get { return (string[])profiles.Clone(); } }
        public string StemTeXHome { get { return stemTeXHome; } }
        public string Status { get { lock (gate) return status; } }

        public string DefaultAvailableProfile
        {
            get
            {
                foreach (var profile in profiles)
                    if (string.Equals(profile, StemTeXBackend.DefaultProfile, StringComparison.OrdinalIgnoreCase))
                        return profile;
                return profiles[0];
            }
        }

        public void SwitchProfile(string profile)
        {
            var canonical = CanonicalProfile(profile);
            lock (gate)
            {
                ThrowIfDisposed();
                selectedProfile = canonical;
                status = "warming:" + canonical;
            }

            // Starting warm-up must not synchronously involve an Office UI thread.
            // The following render will await the same backend queue if it arrives
            // before the warm-up finishes.
            Task.Run(async () =>
            {
                try
                {
                    var response = await SendAsync(new RenderHostClientRequest
                    {
                        Version = ProtocolVersion,
                        Id = NewRequestId(),
                        Command = "switchProfile",
                        Profile = canonical
                    }).ConfigureAwait(false);
                    ApplySuccessfulStatus(response);
                }
                catch (Exception exception)
                {
                    SetFailedStatus(exception);
                }
            });
        }

        public void WarmUp(string profile)
        {
            var canonical = CanonicalProfile(profile);
            lock (gate)
            {
                ThrowIfDisposed();
                selectedProfile = canonical;
                status = "warming:" + canonical;
            }

            var switched = Send(new RenderHostClientRequest
            {
                Version = ProtocolVersion,
                Id = NewRequestId(),
                Command = "switchProfile",
                Profile = canonical
            });
            ApplySuccessfulStatus(switched);

            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < WarmUpTimeoutMs)
            {
                var snapshot = Send(new RenderHostClientRequest
                {
                    Version = ProtocolVersion,
                    Id = NewRequestId(),
                    Command = "ping"
                });
                ApplySuccessfulStatus(snapshot);
                var current = Status;
                if (current.StartsWith("ready:", StringComparison.OrdinalIgnoreCase)) return;
                if (current.StartsWith("failed:", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("StemTeX renderer warm-up failed: " + current);
                Thread.Sleep(50);
            }
            throw new TimeoutException("StemTeX renderer warm-up did not complete within " +
                WarmUpTimeoutMs + " ms.");
        }

        public Task<StemTeXSvgResult> RenderLatestAsync(string profile, string source, double widthPt,
            bool autoWidth, double fontSizePt = 10)
        {
            return RenderAsync("renderLatest", profile, source, widthPt, autoWidth, fontSizePt);
        }

        public Task<StemTeXSvgResult> RenderQueuedAsync(string profile, string source, double widthPt,
            bool autoWidth, double fontSizePt = 10)
        {
            return RenderAsync("renderQueued", profile, source, widthPt, autoWidth, fontSizePt);
        }

        public StemTeXSvgResult RenderSvg(string profile, string source, double widthPt, bool autoWidth,
            double fontSizePt = 10)
        {
            return RenderLatestAsync(profile, source, widthPt, autoWidth, fontSizePt).GetAwaiter().GetResult();
        }

        public void CancelLatestPreview()
        {
            lock (gate)
            {
                if (disposed) return;
            }

            // A distinct pipe connection reaches the host even while a previous
            // render request is awaiting its SVG response.
            Task.Run(async () =>
            {
                try
                {
                    var response = await SendAsync(new RenderHostClientRequest
                    {
                        Version = ProtocolVersion,
                        Id = NewRequestId(),
                        Command = "cancelLatestPreview"
                    }).ConfigureAwait(false);
                    ApplySuccessfulStatus(response);
                }
                catch
                {
                    // The editor is already closing or the host has been terminated.
                    // A cancelled preview must never surface as an Office UI failure.
                }
            });
        }

        public void Dispose()
        {
            Process process;
            IntPtr job;
            lock (gate)
            {
                if (disposed) return;
                disposed = true;
                status = "stopped";
                process = hostProcess;
                job = hostJob;
                hostJob = IntPtr.Zero;
            }

            // Closing the job kills the broker and every renderer helper it spawned.
            // Do not send a graceful shutdown and never wait: the whole point of this
            // boundary is that a native create/render cannot prolong Office teardown.
            if (job != IntPtr.Zero)
            {
                try { CloseHandle(job); }
                catch { }
            }
            else
            {
                try
                {
                    if (process != null && !process.HasExited) process.Kill();
                }
                catch { }
            }
            try { process?.Dispose(); }
            catch { }
        }

        private Task<StemTeXSvgResult> RenderAsync(string command, string profile, string source, double widthPt,
            bool autoWidth, double fontSizePt)
        {
            var canonical = CanonicalProfile(profile);
            lock (gate)
            {
                ThrowIfDisposed();
                selectedProfile = canonical;
                status = "queued";
            }
            return RenderResponseAsync(new RenderHostClientRequest
            {
                Version = ProtocolVersion,
                Id = NewRequestId(),
                Command = command,
                Profile = canonical,
                Source = source,
                WidthPt = widthPt,
                AutoWidth = autoWidth,
                FontSizePt = fontSizePt
            });
        }

        private async Task<StemTeXSvgResult> RenderResponseAsync(RenderHostClientRequest request)
        {
            try
            {
                var response = await SendAsync(request).ConfigureAwait(false);
                ApplySuccessfulStatus(response);
                var result = response.Result;
                if (result == null || string.IsNullOrWhiteSpace(result.SvgBase64))
                    throw new InvalidOperationException("The render host returned no SVG output.");
                byte[] bytes;
                try { bytes = Convert.FromBase64String(result.SvgBase64); }
                catch (FormatException exception)
                {
                    throw new InvalidDataException("The render host returned malformed SVG data.", exception);
                }
                return new StemTeXSvgResult(bytes, result.SummaryJson, result.OutcomeCode, result.IssueFlags,
                    result.OutcomeMessage, result.DepthPt);
            }
            catch (RenderHostRequestException exception) when (
                string.Equals(exception.Code, "cancelled", StringComparison.OrdinalIgnoreCase))
            {
                // Preserve the existing latest-only editor contract: an obsolete
                // preview is a cancellation, not a user-visible renderer failure.
                throw new TaskCanceledException("The render request was superseded.", exception);
            }
            catch (Exception exception)
            {
                SetFailedStatus(exception);
                throw;
            }
        }

        private Task<RenderHostClientResponse> SendAsync(RenderHostClientRequest request)
        {
            return Task.Run(() => Send(request));
        }

        private RenderHostClientResponse Send(RenderHostClientRequest request)
        {
            lock (gate) ThrowIfDisposed();
            try
            {
                using (var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut,
                    PipeOptions.None, TokenImpersonationLevel.Identification))
                {
                    pipe.Connect(PipeConnectTimeoutMs);
                    WriteRequest(pipe, request);
                    var response = ReadResponse(pipe);
                    if (response == null)
                        throw new EndOfStreamException("The render host closed the pipe without a response.");
                    if (response.Version != ProtocolVersion)
                        throw new InvalidOperationException("The render host returned an incompatible protocol version.");
                    if (!string.Equals(response.Id, request.Id, StringComparison.Ordinal))
                        throw new InvalidOperationException("The render host returned a response for a different request.");
                    if (!response.Ok)
                        throw new RenderHostRequestException(response.Error == null ? null : response.Error.Code,
                            FormatHostError(response.Error));
                    return response;
                }
            }
            catch (TimeoutException)
            {
                ThrowIfHostExited();
                throw new TimeoutException("Timed out waiting for the LaTeX Blocks render host.");
            }
            catch (IOException exception)
            {
                ThrowIfHostExited();
                throw new IOException("The LaTeX Blocks render host connection failed.", exception);
            }
        }

        private static void WriteRequest(Stream stream, RenderHostClientRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            using (var output = new MemoryStream())
            {
                RequestSerializer.WriteObject(output, request);
                WriteFrame(stream, Utf8.GetString(output.ToArray()), MaximumRequestFrameBytes, "request");
            }
        }

        private static RenderHostClientResponse ReadResponse(Stream stream)
        {
            var json = ReadFrame(stream, MaximumResponseFrameBytes, "response");
            if (json == null) return null;
            try
            {
                using (var input = new MemoryStream(Utf8.GetBytes(json), false))
                    return (RenderHostClientResponse)ResponseSerializer.ReadObject(input);
            }
            catch (SerializationException exception)
            {
                throw new InvalidDataException("The render host returned invalid protocol JSON.", exception);
            }
        }

        private static string ReadFrame(Stream stream, int maximumBytes, string kind)
        {
            var firstByte = stream.ReadByte();
            if (firstByte < 0) return null;
            var header = new byte[4];
            header[0] = (byte)firstByte;
            ReadExactly(stream, header, 1, 3);
            var length = (uint)(header[0] | (header[1] << 8) | (header[2] << 16) | (header[3] << 24));
            if (length > maximumBytes)
                throw new InvalidDataException("The render-host " + kind + " frame exceeds its size limit.");
            var payload = new byte[(int)length];
            ReadExactly(stream, payload, 0, payload.Length);
            try { return Utf8.GetString(payload); }
            catch (DecoderFallbackException exception)
            {
                throw new InvalidDataException("The render-host " + kind + " frame is not valid UTF-8.", exception);
            }
        }

        private static void WriteFrame(Stream stream, string json, int maximumBytes, string kind)
        {
            var payload = Utf8.GetBytes(json ?? string.Empty);
            if (payload.Length > maximumBytes)
                throw new InvalidDataException("The render-host " + kind + " frame exceeds its size limit.");
            var length = payload.Length;
            var header = new[]
            {
                (byte)(length & 0xff),
                (byte)((length >> 8) & 0xff),
                (byte)((length >> 16) & 0xff),
                (byte)((length >> 24) & 0xff)
            };
            stream.Write(header, 0, header.Length);
            stream.Write(payload, 0, payload.Length);
            stream.Flush();
        }

        private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count)
        {
            while (count > 0)
            {
                var read = stream.Read(buffer, offset, count);
                if (read <= 0) throw new EndOfStreamException("The pipe closed before its frame was complete.");
                offset += read;
                count -= read;
            }
        }

        private static string ResolveHostExecutable()
        {
            var assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var applicationDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var candidates = new[]
            {
                Path.Combine(assemblyDirectory ?? string.Empty, "LaTeXBlocks.RenderHost.host"),
                Path.Combine(applicationDirectory ?? string.Empty, "LaTeXBlocks.RenderHost.host")
            };
            foreach (var candidate in candidates)
                if (File.Exists(candidate)) return candidate;
            throw new FileNotFoundException(
                "LaTeX Blocks RenderHost was not found beside the installed add-in.", candidates[0]);
        }

        private static Process StartHost(string executable, string nonce)
        {
            var info = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = "--pipe-nonce " + nonce,
                WorkingDirectory = Path.GetDirectoryName(executable),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            var process = Process.Start(info);
            if (process == null)
                throw new InvalidOperationException("LaTeX Blocks RenderHost did not start.");
            return process;
        }

        private string CanonicalProfile(string profile)
        {
            if (string.IsNullOrWhiteSpace(profile)) return DefaultAvailableProfile;
            foreach (var candidate in profiles)
                if (string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase)) return candidate;
            throw new ArgumentException("Unknown StemTeX profile: " + profile, nameof(profile));
        }

        private void ApplySuccessfulStatus(RenderHostClientResponse response)
        {
            if (response == null || response.Result == null) return;
            lock (gate)
            {
                if (disposed) return;
                if (!string.IsNullOrWhiteSpace(response.Result.Profile))
                    selectedProfile = response.Result.Profile;
                if (!string.IsNullOrWhiteSpace(response.Result.Status))
                    status = response.Result.Status;
            }
        }

        private void SetFailedStatus(Exception exception)
        {
            lock (gate)
            {
                if (!disposed)
                    status = "failed:" + (exception == null ? "unknown error" : exception.Message);
            }
        }

        private void ThrowIfHostExited()
        {
            try
            {
                if (hostProcess != null && hostProcess.HasExited)
                    throw new InvalidOperationException("LaTeX Blocks RenderHost exited with code " +
                        hostProcess.ExitCode + ".");
            }
            catch (InvalidOperationException) { throw; }
            catch { }
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException(nameof(RenderHostClientBackend));
        }

        private static string NewRequestId() { return Guid.NewGuid().ToString("N"); }

        private static string FormatHostError(RenderHostClientError error)
        {
            if (error == null) return "The render host rejected a request without an error message.";
            return string.IsNullOrWhiteSpace(error.Code) ? error.Message : error.Code + ": " + error.Message;
        }

        private static IntPtr TryCreateKillOnCloseJob(Process process)
        {
            if (process == null) return IntPtr.Zero;
            IntPtr job = IntPtr.Zero;
            try
            {
                job = CreateJobObject(IntPtr.Zero, null);
                if (job == IntPtr.Zero) return IntPtr.Zero;
                var information = new JobObjectExtendedLimitInformation();
                information.BasicLimitInformation.LimitFlags = JobObjectLimitKillOnJobClose;
                var size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformation));
                if (!SetInformationJobObject(job, JobObjectExtendedLimitInformationClass, ref information, (uint)size) ||
                    !AssignProcessToJobObject(job, process.Handle))
                {
                    CloseHandle(job);
                    return IntPtr.Zero;
                }
                return job;
            }
            catch
            {
                if (job != IntPtr.Zero)
                    try { CloseHandle(job); } catch { }
                return IntPtr.Zero;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicLimitInformation
        {
            internal long PerProcessUserTimeLimit;
            internal long PerJobUserTimeLimit;
            internal uint LimitFlags;
            internal UIntPtr MinimumWorkingSetSize;
            internal UIntPtr MaximumWorkingSetSize;
            internal uint ActiveProcessLimit;
            internal IntPtr Affinity;
            internal uint PriorityClass;
            internal uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            internal ulong ReadOperationCount;
            internal ulong WriteOperationCount;
            internal ulong OtherOperationCount;
            internal ulong ReadTransferCount;
            internal ulong WriteTransferCount;
            internal ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectExtendedLimitInformation
        {
            internal JobObjectBasicLimitInformation BasicLimitInformation;
            internal IoCounters IoInfo;
            internal UIntPtr ProcessMemoryLimit;
            internal UIntPtr JobMemoryLimit;
            internal UIntPtr PeakProcessMemoryUsed;
            internal UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr jobAttributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetInformationJobObject(IntPtr job, uint informationClass,
            ref JobObjectExtendedLimitInformation information, uint informationLength);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);
    }

    [DataContract]
    internal sealed class RenderHostClientRequest
    {
        [DataMember(Name = "version", EmitDefaultValue = false)]
        internal int Version { get; set; }
        [DataMember(Name = "id", EmitDefaultValue = false)]
        internal string Id { get; set; }
        [DataMember(Name = "command", EmitDefaultValue = false)]
        internal string Command { get; set; }
        [DataMember(Name = "profile", EmitDefaultValue = false)]
        internal string Profile { get; set; }
        [DataMember(Name = "source", EmitDefaultValue = false)]
        internal string Source { get; set; }
        [DataMember(Name = "widthPt", EmitDefaultValue = false)]
        internal double? WidthPt { get; set; }
        [DataMember(Name = "autoWidth", EmitDefaultValue = false)]
        internal bool? AutoWidth { get; set; }
        [DataMember(Name = "fontSizePt", EmitDefaultValue = false)]
        internal double? FontSizePt { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostClientResponse
    {
        [DataMember(Name = "version")]
        internal int Version { get; set; }
        [DataMember(Name = "id", EmitDefaultValue = false)]
        internal string Id { get; set; }
        [DataMember(Name = "ok")]
        internal bool Ok { get; set; }
        [DataMember(Name = "result", EmitDefaultValue = false)]
        internal RenderHostClientResult Result { get; set; }
        [DataMember(Name = "error", EmitDefaultValue = false)]
        internal RenderHostClientError Error { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostClientResult
    {
        [DataMember(Name = "kind", EmitDefaultValue = false)]
        internal string Kind { get; set; }
        [DataMember(Name = "protocol", EmitDefaultValue = false)]
        internal string Protocol { get; set; }
        [DataMember(Name = "processId", EmitDefaultValue = false)]
        internal int ProcessId { get; set; }
        [DataMember(Name = "pipeName", EmitDefaultValue = false)]
        internal string PipeName { get; set; }
        [DataMember(Name = "status", EmitDefaultValue = false)]
        internal string Status { get; set; }
        [DataMember(Name = "profile", EmitDefaultValue = false)]
        internal string Profile { get; set; }
        [DataMember(Name = "profiles", EmitDefaultValue = false)]
        internal string[] Profiles { get; set; }
        [DataMember(Name = "svgBase64", EmitDefaultValue = false)]
        internal string SvgBase64 { get; set; }
        [DataMember(Name = "summaryJson", EmitDefaultValue = false)]
        internal string SummaryJson { get; set; }
        [DataMember(Name = "outcomeCode", EmitDefaultValue = false)]
        internal int OutcomeCode { get; set; }
        [DataMember(Name = "issueFlags", EmitDefaultValue = false)]
        internal int IssueFlags { get; set; }
        [DataMember(Name = "outcomeMessage", EmitDefaultValue = false)]
        internal string OutcomeMessage { get; set; }
        [DataMember(Name = "depthPt", EmitDefaultValue = false)]
        internal double DepthPt { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostClientError
    {
        [DataMember(Name = "code", EmitDefaultValue = false)]
        internal string Code { get; set; }
        [DataMember(Name = "message", EmitDefaultValue = false)]
        internal string Message { get; set; }
    }

    internal sealed class RenderHostRequestException : Exception
    {
        internal RenderHostRequestException(string code, string message) : base(message)
        {
            Code = code;
        }

        internal string Code { get; }
    }
}
