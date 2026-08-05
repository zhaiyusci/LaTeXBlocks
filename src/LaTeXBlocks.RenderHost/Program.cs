using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;

namespace LaTeXBlocks.RenderHost
{
    /// <summary>
    /// Standalone, current-user-only named-pipe front end for RenderHostEngine.
    /// The server deliberately has no Office, COM, VSTO, or windowing dependency.
    /// </summary>
    internal static class Program
    {
        private const string PipePrefix = "LaTeXBlocks.RenderHost.";
        private const int MaximumServerInstances = 16;
        private const uint JobObjectExtendedLimitInformationClass = 9;
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;
        // Deliberately kept open for the process lifetime. When this broker is
        // terminated by the Office-side job (or direct Kill fallback), Windows
        // closes this handle as part of process teardown and kills any StemTeX
        // worker helpers inherited into the job as well.
        private static IntPtr lifetimeChildJob;

        private static int Main(string[] args)
        {
            try
            {
                string nonce;
                if (IsHelp(args))
                {
                    WriteUsage();
                    return 0;
                }

                if (!TryGetNonce(args, out nonce))
                {
                    Console.Error.WriteLine("Missing or invalid --pipe-nonce argument.");
                    WriteUsage();
                    return 64;
                }

                var pipeName = PipePrefix + nonce;
                TryAttachProcessLifetimeJob();
                Console.WriteLine("LaTeX Blocks render host listening on {0}.", pipeName);
                return Serve(pipeName);
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("LaTeX Blocks render host failed: {0}", exception.Message);
                return 1;
            }
        }

        private static int Serve(string pipeName)
        {
            using (var shutdownSignal = new ManualResetEvent(false))
            using (var engine = new RenderHostEngine())
            {
                try
                {
                    // Keep one asynchronous listener pending while every accepted
                    // connection is independently read and dispatched. A render can
                    // therefore await StemTeX while a second pipe request, or a
                    // pipelined cancelLatestPreview request on the same pipe, arrives.
                    while (!shutdownSignal.WaitOne(0) && !engine.IsStopping)
                    {
                        NamedPipeServerStream listener = null;
                        try
                        {
                            listener = CreateRestrictedServer(pipeName);
                            var wait = listener.BeginWaitForConnection(null, null);
                            var signalled = WaitHandle.WaitAny(new[]
                            {
                                shutdownSignal,
                                wait.AsyncWaitHandle
                            });
                            if (signalled == 0 || engine.IsStopping)
                            {
                                listener.Dispose();
                                break;
                            }

                            listener.EndWaitForConnection(wait);
                            var accepted = listener;
                            listener = null;
                            Task.Run(() => ServeConnection(accepted, engine, pipeName, shutdownSignal));
                        }
                        catch (ObjectDisposedException)
                        {
                            // A listener can be disposed by a shutdown race. If it
                            // was an unrelated transient pipe failure, discard only
                            // this listener and keep the broker available.
                            if (!shutdownSignal.WaitOne(0) && !engine.IsStopping)
                            {
                                Console.Error.WriteLine("Render-host listener was disposed; retrying.");
                            }
                        }
                        catch (IOException exception)
                        {
                            // A malformed or disconnected client must never take
                            // down the renderer that other Office sessions use.
                            if (!shutdownSignal.WaitOne(0) && !engine.IsStopping)
                            {
                                Console.Error.WriteLine("Render-host pipe listener failed: {0}", exception.Message);
                            }
                        }
                        finally
                        {
                            if (listener != null)
                            {
                                listener.Dispose();
                            }
                        }
                    }
                    return 0;
                }
                finally
                {
                    // No native cleanup or cancellation ever runs in Office. This
                    // process owns it; its background worker is allowed to disappear
                    // with the process once the shutdown response has been sent.
                    engine.CompleteShutdown();
                }
            }
        }

        private static void ServeConnection(NamedPipeServerStream pipe, RenderHostEngine engine,
            string pipeName, EventWaitHandle shutdownSignal)
        {
            using (pipe)
            {
                var writeGate = new object();
                try
                {
                    while (!shutdownSignal.WaitOne(0))
                    {
                        RenderHostRequest request;
                        try
                        {
                            request = PipeProtocol.ReadRequest(pipe);
                        }
                        catch (ProtocolException exception)
                        {
                            TryWrite(pipe, writeGate, Error(null, exception.Code, exception.Message));
                            // A framing error can leave unread bytes in the pipe. Close
                            // just this connection; the host remains available.
                            return;
                        }
                        catch (EndOfStreamException)
                        {
                            return;
                        }

                        if (request == null)
                        {
                            return;
                        }

                        // Register each command on this reader, in wire order, but
                        // never await its eventual render response here. Calling the
                        // async method directly is intentional: its synchronous prefix
                        // publishes RenderLatestAsync to StemTeX before the next frame
                        // is read, so a following cancelLatestPreview cannot race ahead
                        // of the preview it is supposed to cancel.
                        var execution = engine.ExecuteAsync(request, pipeName);
                        ObserveRequestAsync(execution, request.Id, pipe, writeGate, engine, shutdownSignal);
                    }
                }
                catch (IOException)
                {
                    // Client disconnection is not a host failure.
                }
                catch (ObjectDisposedException)
                {
                    // The host is stopping.
                }
            }
        }

        private static async void ObserveRequestAsync(Task<HostCommandOutcome> execution, string requestId,
            NamedPipeServerStream pipe, object writeGate, RenderHostEngine engine, EventWaitHandle shutdownSignal)
        {
            HostCommandOutcome outcome;
            try
            {
                outcome = await execution.ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                outcome = new HostCommandOutcome(Error(requestId, "internal_error", LimitMessage(exception)), false);
            }

            var wrote = TryWrite(pipe, writeGate, outcome.Response);
            if (wrote && outcome.ShutdownAfterResponse)
            {
                engine.CompleteShutdown();
                shutdownSignal.Set();
            }
        }

        private static bool TryWrite(NamedPipeServerStream pipe, object writeGate, RenderHostResponse response)
        {
            try
            {
                lock (writeGate)
                {
                    PipeProtocol.WriteResponse(pipe, response);
                }
                return true;
            }
            catch (ProtocolException exception)
            {
                // The normal render response was too large. Send a compact response
                // where possible; an already closed client simply loses the error.
                try
                {
                    lock (writeGate)
                    {
                        PipeProtocol.WriteResponse(pipe, Error(response.Id, exception.Code, exception.Message));
                    }
                }
                catch { }
                return false;
            }
            catch (IOException)
            {
                return false;
            }
            catch (ObjectDisposedException)
            {
                return false;
            }
        }

        private static NamedPipeServerStream CreateRestrictedServer(string pipeName)
        {
            var identity = WindowsIdentity.GetCurrent();
            if (identity == null || identity.User == null)
            {
                throw new InvalidOperationException("The render host could not determine its Windows user identity.");
            }

            var security = new PipeSecurity();
            security.SetAccessRuleProtection(true, false);
            security.AddAccessRule(new PipeAccessRule(
                identity.User,
                PipeAccessRights.FullControl,
                AccessControlType.Allow));

            return new NamedPipeServerStream(
                pipeName,
                PipeDirection.InOut,
                MaximumServerInstances,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous,
                4096,
                4096,
                security);
        }

        private static void TryAttachProcessLifetimeJob()
        {
            if (lifetimeChildJob != IntPtr.Zero)
            {
                return;
            }

            IntPtr job = IntPtr.Zero;
            try
            {
                job = CreateJobObject(IntPtr.Zero, null);
                if (job == IntPtr.Zero)
                {
                    return;
                }
                var information = new JobObjectExtendedLimitInformation();
                information.BasicLimitInformation.LimitFlags = JobObjectLimitKillOnJobClose;
                var size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformation));
                if (!SetInformationJobObject(job, JobObjectExtendedLimitInformationClass, ref information,
                    (uint)size))
                {
                    CloseHandle(job);
                    return;
                }

                using (var current = Process.GetCurrentProcess())
                {
                    if (!AssignProcessToJobObject(job, current.Handle))
                    {
                        // The broker may already inherit an Office-controlled job.
                        // In that normal case the Office-side kill-on-close job is
                        // still authoritative; do not turn a best-effort guard into
                        // a launch failure.
                        CloseHandle(job);
                        return;
                    }
                }
                lifetimeChildJob = job;
            }
            catch
            {
                if (job != IntPtr.Zero)
                {
                    try { CloseHandle(job); }
                    catch { }
                }
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
                    Message = LimitMessage(message)
                }
            };
        }

        private static string LimitMessage(Exception exception)
        {
            return LimitMessage(exception == null ? null : exception.Message);
        }

        private static string LimitMessage(string message)
        {
            if (string.IsNullOrEmpty(message) || message.Length <= 2048)
            {
                return message ?? "The render-host request failed.";
            }
            return message.Substring(0, 2048) + "…";
        }

        private static bool TryGetNonce(string[] args, out string nonce)
        {
            nonce = null;
            if (args == null)
            {
                return false;
            }

            for (var index = 0; index + 1 < args.Length; index++)
            {
                if (!string.Equals(args[index], "--pipe-nonce", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var candidate = args[index + 1] ?? string.Empty;
                if (candidate.Length < 16 || candidate.Length > 128)
                {
                    return false;
                }

                foreach (var character in candidate)
                {
                    if (!IsAsciiLetterOrDigit(character) && character != '-' && character != '_')
                    {
                        return false;
                    }
                }

                nonce = candidate;
                return true;
            }

            return false;
        }

        private static bool IsAsciiLetterOrDigit(char character)
        {
            return (character >= 'A' && character <= 'Z')
                || (character >= 'a' && character <= 'z')
                || (character >= '0' && character <= '9');
        }

        private static bool IsHelp(string[] args)
        {
            return args != null
                && args.Length == 1
                && (string.Equals(args[0], "--help", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(args[0], "-h", StringComparison.OrdinalIgnoreCase));
        }

        private static void WriteUsage()
        {
            Console.WriteLine("Usage: LaTeXBlocks.RenderHost.host --pipe-nonce <unique-token>");
            Console.WriteLine("Pipe name: " + PipePrefix + "<unique-token>");
        }
    }
}
