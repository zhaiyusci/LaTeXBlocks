using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    internal sealed class StemTeXRenderer : IDisposable
    {
        private const int SvgOutputFormat = 1;
        // StemTeX 0.12 is the first bundled runtime for this host. Keep the
        // native contract explicit so an old STEMTEX_HOME cannot silently
        // substitute the 0.11 runtime for the packaged binary.
        private static readonly Version RequiredRuntimeVersion = new Version(0, 12, 0);
        private readonly object gate = new object();
        // RenderSvg keeps gate for the native render so create/destroy remain strictly
        // serialized. Cancellation must not wait for that lock, but it still needs a
        // small lifetime handshake so Destroy cannot free a renderer while an already
        // issued cancel call is using its native pointer.
        private readonly object cancellationGate = new object();
        private readonly string stemTeXHome;
        private readonly string runtimeRoot;
        private readonly string profileRoot;
        private IntPtr library;
        private IntPtr renderer;
        private NativeApi api;
        private bool disposed;
        private bool nativeRenderInProgress;
        private int activeCancelCalls;
        private long nativeRenderGeneration;
        private long cancellationRetryGeneration;
        private int nativeCancelAttempts;

        internal StemTeXRenderer(string profile = "unicodemath_cjk")
        {
            stemTeXHome = ResolveStemTeXHome(profile);
            runtimeRoot = Path.Combine(stemTeXHome, "runtime");
            profileRoot = Path.Combine(stemTeXHome, "gui", "profiles", profile);
        }

        internal string StemTeXHome => stemTeXHome;
        internal string ProfileRoot => profileRoot;
        internal bool NativeRenderInProgressForTest
        {
            get { lock (cancellationGate) return nativeRenderInProgress; }
        }
        internal int NativeCancelAttemptsForTest
        {
            get { lock (cancellationGate) return nativeCancelAttempts; }
        }

        internal void WarmUp()
        {
            lock (gate)
            {
                ThrowIfDisposed();
                EnsureRenderer();
            }
        }

        internal static string[] DiscoverProfiles()
        {
            var root = Path.Combine(ResolveStemTeXHome(null), "gui", "profiles");
            var profiles = new List<string>();
            foreach (var directory in Directory.GetDirectories(root))
            {
                if (File.Exists(Path.Combine(directory, "preamble.tex"))) profiles.Add(Path.GetFileName(directory));
            }
            profiles.Sort(StringComparer.OrdinalIgnoreCase);
            return profiles.ToArray();
        }

        internal StemTeXSvgResult RenderSvg(string source, double widthPt, bool autoWidth = false,
            double fontSizePt = 10, Func<bool> cancelBeforeNativeRender = null)
        {
            if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("LaTeX source cannot be empty.", nameof(source));
            if (widthPt <= 0) throw new ArgumentOutOfRangeException(nameof(widthPt));
            if (!(fontSizePt >= 1) || fontSizePt > 200) throw new ArgumentOutOfRangeException(nameof(fontSizePt));

            lock (gate)
            {
                ThrowIfDisposed();
                EnsureRenderer();
                var markedSource = AddMeasurementMarkers(source, autoWidth);
                var renderWidth = autoWidth ? 2000.0 : widthPt;
                using (var sourceUtf8 = new NativeUtf8(markedSource))
                {
                    var bytes = new StemTeXOutputBytes();
                    var result = new StemTeXRenderOutputResult();
                    IntPtr error;
                    int errorCode;
                    int ok;
                    lock (cancellationGate)
                    {
                        // Pair this check with CancelCurrent's native-render state.
                        // If a newer preview wins before we enter the native call, it
                        // returns here; if it wins after this point, CancelCurrent sees
                        // nativeRenderInProgress and interrupts the active XeTeX job.
                        if (cancelBeforeNativeRender != null && cancelBeforeNativeRender())
                            throw new OperationCanceledException("StemTeX request was superseded.");
                        ++nativeRenderGeneration;
                        nativeRenderInProgress = true;
                    }
                    try
                    {
                        ok = api.RenderOutputBytesWithFontSize(renderer, sourceUtf8.Pointer, renderWidth, fontSizePt,
                            SvgOutputFormat, out bytes, out result, out errorCode, out error);
                    }
                    finally
                    {
                        lock (cancellationGate)
                        {
                            nativeRenderInProgress = false;
                            Monitor.PulseAll(cancellationGate);
                        }
                    }
                    try
                    {
                        if (ok == 0)
                        {
                            throw new StemTeXException(errorCode, ReadUtf8(error) ?? "StemTeX render failed.");
                        }

                        var size = checked((int)bytes.Size.ToUInt64());
                        var managedBytes = new byte[size];
                        if (size > 0) Marshal.Copy(bytes.Data, managedBytes, 0, size);
                        double depthPt;
                        managedBytes = ProcessMeasurementMarkers(managedBytes, autoWidth, out depthPt);
                        return new StemTeXSvgResult(
                            managedBytes,
                            ReadUtf8(result.SummaryJson),
                            result.OutcomeCode,
                            result.IssueFlags,
                            ReadUtf8(result.OutcomeMessage),
                            depthPt);
                    }
                    finally
                    {
                        if (error != IntPtr.Zero) api.FreeString(error);
                        api.FreeOutputBytes(ref bytes);
                        api.FreeOutputResult(ref result);
                    }
                }
            }
        }

        private void EnsureRenderer()
        {
            if (renderer != IntPtr.Zero) return;
            try
            {
                var dll = Path.Combine(runtimeRoot, "bin", "sdk", "stemtex-renderer.dll");
                library = NativeMethods.LoadLibrary(dll);
                if (library == IntPtr.Zero)
                    throw new InvalidOperationException("Unable to load StemTeX renderer: " + dll + " (Win32 " + Marshal.GetLastWin32Error() + ").");
                api = new NativeApi(library);

                using (var home = new NativeUtf8(stemTeXHome))
                using (var runtime = new NativeUtf8(runtimeRoot))
                using (var profile = new NativeUtf8(profileRoot))
                {
                    var config = new StemTeXConfig
                    {
                        RepoRoot = home.Pointer,
                        RuntimeRoot = runtime.Pointer,
                        ProfileRoot = profile.Pointer,
                        RequestTimeoutMs = 90000,
                        XdvipdfmxTimeoutMs = 90000,
                        DefaultWidthPt = 360,
                        SpareWorkerCount = 0,
                        AutoRestart = 1,
                        DeleteIntermediates = 1
                    };
                    IntPtr error;
                    int errorCode;
                    renderer = api.Create(ref config, out errorCode, out error);
                    try
                    {
                        if (renderer == IntPtr.Zero)
                            throw new StemTeXException(errorCode, ReadUtf8(error) ?? "StemTeX initialization failed.");
                    }
                    finally
                    {
                        if (error != IntPtr.Zero) api.FreeString(error);
                    }
                }
            }
            catch
            {
                // A failed create/entry-point load must leave this instance retryable.
                // Without this cleanup a later retry overwrote library and leaked the
                // old module handle, which also made diagnostics misleading.
                try { ReleaseNativeResources(); }
                catch { }
                throw;
            }
        }

        private static string ResolveStemTeXHome(string profile)
        {
            var candidates = new List<string>();
            var configured = Environment.GetEnvironmentVariable("STEMTEX_HOME");
            if (!string.IsNullOrWhiteSpace(configured))
            {
                var configuredHome = ValidateStemTeXHome(configured, profile);
                if (configuredHome != null) return configuredHome;
                throw new DirectoryNotFoundException("STEMTEX_HOME does not contain a usable StemTeX runtime: " + configured);
            }
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\LaTeXBlocks"))
                {
                    var installedHome = key?.GetValue("StemTeXHome") as string;
                    if (!string.IsNullOrWhiteSpace(installedHome)) candidates.Add(installedHome);
                }
            }
            catch { }
            var systemRoot = Path.GetPathRoot(Environment.SystemDirectory);
            if (!string.IsNullOrWhiteSpace(systemRoot)) candidates.Add(Path.Combine(systemRoot, "StemTeX"));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Scholia", "StemTeX"));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "StemTeX"));
            // Installed runtimes precede development stages so an equal-version dev
            // tree cannot silently override the self-contained package. A strictly
            // newer development build can still win; STEMTEX_HOME remains the explicit
            // override for deterministic development and diagnostics.
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "dist", "stemtex-installer", "StemTeX"));
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "build", "stemtex-check-stage"));

            string bestHome = null;
            Version bestVersion = null;
            foreach (var candidate in candidates)
            {
                var home = ValidateStemTeXHome(candidate, profile);
                if (home == null) continue;
                Version version;
                var versionFile = Path.Combine(home, "runtime", "VERSION");
                var versionText = File.Exists(versionFile) ? File.ReadAllText(versionFile).Trim() : "0.0.0";
                if (!Version.TryParse(versionText, out version)) version = new Version(0, 0);
                if (bestHome == null || version.CompareTo(bestVersion) > 0) { bestHome = home; bestVersion = version; }
            }
            if (bestHome != null) return bestHome;
            throw new DirectoryNotFoundException(
                "StemTeX with SVG support was not found. Install StemTeX or set STEMTEX_HOME to its installation root.");
        }

        private static string ValidateStemTeXHome(string candidate, string profile)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return null;
            var home = Path.GetFullPath(candidate);
            Version runtimeVersion;
            var versionFile = Path.Combine(home, "runtime", "VERSION");
            if (!File.Exists(versionFile) ||
                !Version.TryParse(File.ReadAllText(versionFile).Trim(), out runtimeVersion) ||
                runtimeVersion.CompareTo(RequiredRuntimeVersion) < 0) return null;
            var profilesRoot = Path.Combine(home, "gui", "profiles");
            var hasProfile = profile == null
                ? Directory.Exists(profilesRoot) && Directory.GetFiles(profilesRoot, "preamble.tex", SearchOption.AllDirectories).Length > 0
                : File.Exists(Path.Combine(profilesRoot, profile, "preamble.tex"));
            return File.Exists(Path.Combine(home, "runtime", "bin", "sdk", "stemtex-renderer.dll")) &&
                   File.Exists(Path.Combine(home, "runtime", "bin", "windows", "dvisvgmdaemon.dll")) && hasProfile
                ? home : null;
        }

        public void Dispose()
        {
            lock (gate)
            {
                if (disposed) return;
                lock (cancellationGate)
                {
                    disposed = true;
                    while (activeCancelCalls != 0) Monitor.Wait(cancellationGate);
                }
                ReleaseNativeResources();
            }
        }

        // Reserves the current native render for one asynchronous cancellation loop.
        // Calls made while the render is still crossing into the C++ renderer are
        // intentionally coalesced; the loop below retries after active_slot is set.
        internal long BeginCancelCurrent()
        {
            lock (cancellationGate)
            {
                if (disposed || !nativeRenderInProgress || renderer == IntPtr.Zero || api == null) return 0;
                if (cancellationRetryGeneration == nativeRenderGeneration) return 0;
                cancellationRetryGeneration = nativeRenderGeneration;
                return nativeRenderGeneration;
            }
        }

        internal bool IsNativeRenderInProgress(long expectedGeneration)
        {
            lock (cancellationGate)
                return !disposed && nativeRenderInProgress && nativeRenderGeneration == expectedGeneration;
        }

        internal void ReleaseCancellationReservation(long expectedGeneration)
        {
            lock (cancellationGate)
                if (cancellationRetryGeneration == expectedGeneration) cancellationRetryGeneration = 0;
        }

        internal void CancelCurrent()
        {
            var generation = BeginCancelCurrent();
            if (generation != 0) CancelCurrent(generation);
        }

        internal void CancelCurrent(long expectedGeneration)
        {
            // Cancellation is intentionally concurrent with RenderSvg. The native API
            // uses its own control mutex and stops the active XeTeX child; taking gate
            // here would wait behind the render and defeat cancellation during shutdown.
            IntPtr currentRenderer;
            NativeApi currentApi;
            lock (cancellationGate)
            {
                if (disposed || !nativeRenderInProgress || nativeRenderGeneration != expectedGeneration ||
                    renderer == IntPtr.Zero || api == null) return;
                currentRenderer = renderer;
                currentApi = api;
                ++activeCancelCalls;
            }
            try
            {
                IntPtr error = IntPtr.Zero;
                int errorCode;
                try
                {
                    lock (cancellationGate) ++nativeCancelAttempts;
                    currentApi.CancelCurrent(currentRenderer, out errorCode, out error);
                }
                finally
                {
                    if (error != IntPtr.Zero)
                    {
                        try { currentApi.FreeString(error); }
                        catch { }
                    }
                }
            }
            catch
            {
                // The request is already obsolete. A best-effort native cancellation
                // must never surface as a new UI failure or break renderer shutdown.
            }
            finally
            {
                lock (cancellationGate)
                {
                    --activeCancelCalls;
                    Monitor.PulseAll(cancellationGate);
                }
            }
        }

        private void ReleaseNativeResources()
        {
            try
            {
                if (renderer != IntPtr.Zero && api != null) api.Destroy(renderer);
            }
            finally
            {
                renderer = IntPtr.Zero;
                api = null;
                if (library != IntPtr.Zero) NativeMethods.FreeLibrary(library);
                library = IntPtr.Zero;
            }
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException(nameof(StemTeXRenderer));
        }

        private static string ReadUtf8(IntPtr pointer)
        {
            if (pointer == IntPtr.Zero) return null;
            var length = 0;
            while (Marshal.ReadByte(pointer, length) != 0) length++;
            var bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            return Encoding.UTF8.GetString(bytes);
        }

        private static string AddMeasurementMarkers(string source, bool autoWidth)
        {
            // A full display environment owns a line-width box and therefore has no
            // natural inline width. Source newlines alone are harmless inside the hbox;
            // they let users format a single formula legibly without changing its box.
            var uncommented = RemoveTeXCommentsForDetection(source);
            if (uncommented.IndexOf("\\[", StringComparison.Ordinal) >= 0 ||
                uncommented.IndexOf("$$", StringComparison.Ordinal) >= 0 ||
                Regex.IsMatch(uncommented,
                    "\\\\begin\\s*\\{\\s*(?:displaymath|equation\\*?|align\\*?|gather\\*?|multline\\*?|flalign\\*?|minipage)\\s*\\}",
                    RegexOptions.CultureInvariant))
            {
                if (autoWidth) throw new ArgumentException(
                    "Auto width does not accept a full display or page-width environment. Use Fixed width for that content.",
                    nameof(source));
                return source;
            }

            if (autoWidth)
            {
                // The wrapper's line breaks are source formatting, not content. A bare
                // newline after "{" or before "}" becomes interword glue inside an
                // hbox, which used to add one hidden TeX space to each side of every
                // auto-width formula. Comment both wrapper newlines instead. Appending
                // '%' is also safe when the source already ends in a TeX comment, while
                // avoiding \unskip means explicit trailing \kern/\hspace/control-space
                // nodes remain part of the user's requested box.
                var boxedSource = source.TrimEnd('\r', '\n');
                return "\\begingroup\n\\setbox255=\\hbox{%\n" + boxedSource + "%\n}%\n" +
                       "\\leavevmode\\special{dvisvgm:raw <g id='latexblocks-start' data-x='{?x}' data-y='{?y}'/>}" +
                       "\\special{dvisvgm:bbox new latexblocksink}" +
                       "\\box255" +
                       "\\special{dvisvgm:raw <g id='latexblocks-ink' data-viewbox='{?bbox latexblocksink}'/>}" +
                       "\\special{dvisvgm:raw <g id='latexblocks-end' data-x='{?x}'/>}\\endgroup";
            }

            // dvisvgm expands {?y} to the current TeX baseline without adding visible geometry.
            return "\\begingroup\n\\leavevmode\\special{dvisvgm:raw <g id='latexblocks-baseline' data-y='{?y}'/>}\n" +
                   source + "\n\\endgroup";
        }

        internal static string RemoveTeXCommentsForDetection(string source)
        {
            if (string.IsNullOrEmpty(source)) return source ?? string.Empty;
            var result = new StringBuilder(source.Length);
            for (var index = 0; index < source.Length; index++)
            {
                if (source[index] == '%')
                {
                    var backslashes = 0;
                    for (var previous = index - 1; previous >= 0 && source[previous] == '\\'; previous--)
                        backslashes++;
                    if (backslashes % 2 == 0)
                    {
                        while (index + 1 < source.Length && source[index + 1] != '\r' &&
                               source[index + 1] != '\n') index++;
                        continue;
                    }
                }
                result.Append(source[index]);
            }
            return result.ToString();
        }

        private static byte[] ProcessMeasurementMarkers(byte[] svgBytes, bool autoWidth, out double depthPt)
        {
            depthPt = 0;
            var svg = Encoding.UTF8.GetString(svgBytes);
            var markerId = autoWidth ? "latexblocks-start" : "latexblocks-baseline";
            var marker = FindMarker(svg, markerId);
            if (!marker.Success) return svgBytes;
            var markerY = Regex.Match(marker.Value, "\\bdata-y=['\"](?<y>[-+0-9.eE]+)['\"]", RegexOptions.CultureInvariant);

            var viewBox = Regex.Match(svg,
                "\\bviewBox=['\"](?<x>[-+0-9.eE]+)\\s+(?<y>[-+0-9.eE]+)\\s+(?<w>[-+0-9.eE]+)\\s+(?<h>[-+0-9.eE]+)['\"]",
                RegexOptions.CultureInvariant);
            if (!viewBox.Success) throw new InvalidDataException("StemTeX SVG has a baseline marker but no numeric viewBox.");

            var baselineY = double.Parse(markerY.Groups["y"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var top = double.Parse(viewBox.Groups["y"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var height = double.Parse(viewBox.Groups["h"].Value, System.Globalization.CultureInfo.InvariantCulture);
            depthPt = Math.Max(0, top + height - baselineY);
            svg = svg.Remove(marker.Index, marker.Length);

            if (autoWidth)
            {
                var endMarker = FindMarker(svg, "latexblocks-end");
                if (!endMarker.Success) throw new InvalidDataException("StemTeX SVG is missing its auto-width end marker.");
                var inkMarker = FindMarker(svg, "latexblocks-ink");
                if (!inkMarker.Success) throw new InvalidDataException("StemTeX SVG is missing its auto-width ink marker.");
                var inkBox = Regex.Match(inkMarker.Value,
                    "\\bdata-viewbox=['\"](?<x>[-+0-9.eE]+)\\s+(?<y>[-+0-9.eE]+)\\s+" +
                    "(?<w>[-+0-9.eE]+)\\s+(?<h>[-+0-9.eE]+)['\"]",
                    RegexOptions.CultureInvariant);
                if (!inkBox.Success) throw new InvalidDataException("StemTeX SVG has no numeric auto-width ink bounds.");
                var startX = ReadMarkerCoordinate(marker.Value, "data-x");
                var endX = ReadMarkerCoordinate(endMarker.Value, "data-x");
                var naturalWidth = endX - startX;
                if (!(naturalWidth > 0) || naturalWidth > 2000) throw new InvalidDataException("StemTeX returned an invalid natural formula width.");
                var inkLeft = double.Parse(inkBox.Groups["x"].Value, System.Globalization.CultureInfo.InvariantCulture);
                var inkWidth = double.Parse(inkBox.Groups["w"].Value, System.Globalization.CultureInfo.InvariantCulture);
                var inkRight = inkLeft + inkWidth;
                if (!(inkWidth >= 0) || double.IsInfinity(inkWidth) || double.IsNaN(inkWidth))
                    throw new InvalidDataException("StemTeX returned invalid auto-width ink bounds.");
                svg = svg.Remove(endMarker.Index, endMarker.Length);
                svg = svg.Remove(inkMarker.Index, inkMarker.Length);

                // The renderer deliberately uses --bbox=papersize, so the page viewBox
                // cannot reveal content bounds. The named dvisvgm bbox marker above was
                // sampled immediately after drawing the TeX box and contains the exact
                // glyph/rule ink bounds. Crop to their union with the logical TeX box.
                // This removes preview.sty's generic 1pt horizontal page border while
                // retaining genuine accents, operators, rules, and glyph overhangs.
                // Do not add a renderer-side safety margin: an inline formula is a
                // faithful TeX box, and its only horizontal space is either part of
                // the user's TeX source or required by real ink outside that box.
                var croppedX = Math.Min(startX, inkLeft);
                var croppedRight = Math.Max(endX, inkRight);
                var croppedWidth = croppedRight - croppedX;
                var number = System.Globalization.CultureInfo.InvariantCulture;
                var newViewBox = "viewBox='" + croppedX.ToString("0.######", number) + " " +
                                 top.ToString("0.######", number) + " " + croppedWidth.ToString("0.######", number) + " " +
                                 height.ToString("0.######", number) + "'";
                svg = Regex.Replace(svg, "\\bviewBox=['\"][^'\"]+['\"]", newViewBox, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
                svg = Regex.Replace(svg, "(<svg\\b[^>]*?)\\bwidth=['\"][^'\"]+['\"]", "$1width='" +
                    croppedWidth.ToString("0.######", number) + "pt'", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            }
            return Encoding.UTF8.GetBytes(svg);
        }

        private static Match FindMarker(string svg, string id)
        {
            return Regex.Match(svg, "<g(?=[^>]*\\bid=['\"]" + Regex.Escape(id) + "['\"])[^>]*/>", RegexOptions.CultureInvariant);
        }

        private static double ReadMarkerCoordinate(string marker, string attribute)
        {
            var match = Regex.Match(marker, "\\b" + Regex.Escape(attribute) + "=['\"](?<v>[-+0-9.eE]+)['\"]", RegexOptions.CultureInvariant);
            if (!match.Success) throw new InvalidDataException("StemTeX SVG marker is missing " + attribute + ".");
            return double.Parse(match.Groups["v"].Value, System.Globalization.CultureInfo.InvariantCulture);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct StemTeXConfig
        {
            public IntPtr RepoRoot;
            public IntPtr RuntimeRoot;
            public IntPtr TexmfRoot;
            public IntPtr ProfileRoot;
            public IntPtr StateRoot;
            public IntPtr RendersRoot;
            public int RequestTimeoutMs;
            public int XdvipdfmxTimeoutMs;
            public double MinWidthPt;
            public double MaxWidthPt;
            public double DefaultWidthPt;
            public int SpareWorkerCount;
            public int AutoRestart;
            public int DeleteIntermediates;
            public IntPtr WorkerTemplate;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct StemTeXOutputBytes
        {
            public IntPtr Data;
            public UIntPtr Size;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct StemTeXRenderOutputResult
        {
            public IntPtr RequestId;
            public IntPtr OutputPath;
            public IntPtr OutputFormat;
            public IntPtr SummaryJson;
            public int OutcomeCode;
            public int IssueFlags;
            public IntPtr OutcomeMessage;
        }

        private sealed class NativeUtf8 : IDisposable
        {
            internal NativeUtf8(string value)
            {
                var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
                Pointer = Marshal.AllocHGlobal(bytes.Length + 1);
                Marshal.Copy(bytes, 0, Pointer, bytes.Length);
                Marshal.WriteByte(Pointer, bytes.Length, 0);
            }
            internal IntPtr Pointer { get; }
            public void Dispose() { Marshal.FreeHGlobal(Pointer); }
        }

        private sealed class NativeApi
        {
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate IntPtr CreateDelegate(ref StemTeXConfig config, out int errorCode, out IntPtr error);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate int RenderOutputBytesWithFontSizeDelegate(IntPtr renderer, IntPtr source, double widthPt,
                double fontSizePt, int format, out StemTeXOutputBytes bytes, out StemTeXRenderOutputResult result,
                out int errorCode, out IntPtr error);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeOutputBytesDelegate(ref StemTeXOutputBytes bytes);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeOutputResultDelegate(ref StemTeXRenderOutputResult result);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeStringDelegate(IntPtr value);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate int CancelCurrentDelegate(IntPtr renderer, out int errorCode, out IntPtr error);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void DestroyDelegate(IntPtr renderer);

            internal NativeApi(IntPtr module)
            {
                Create = Load<CreateDelegate>(module, "stemtex_renderer_create");
                RenderOutputBytesWithFontSize = Load<RenderOutputBytesWithFontSizeDelegate>(module,
                    "stemtex_renderer_render_output_bytes_with_font_size");
                FreeOutputBytes = Load<FreeOutputBytesDelegate>(module, "stemtex_renderer_free_output_bytes");
                FreeOutputResult = Load<FreeOutputResultDelegate>(module, "stemtex_renderer_free_output_result");
                FreeString = Load<FreeStringDelegate>(module, "stemtex_renderer_free_string");
                CancelCurrent = Load<CancelCurrentDelegate>(module, "stemtex_renderer_cancel_current");
                Destroy = Load<DestroyDelegate>(module, "stemtex_renderer_destroy");
            }

            internal CreateDelegate Create { get; }
            internal RenderOutputBytesWithFontSizeDelegate RenderOutputBytesWithFontSize { get; }
            internal FreeOutputBytesDelegate FreeOutputBytes { get; }
            internal FreeOutputResultDelegate FreeOutputResult { get; }
            internal FreeStringDelegate FreeString { get; }
            internal CancelCurrentDelegate CancelCurrent { get; }
            internal DestroyDelegate Destroy { get; }

            private static T Load<T>(IntPtr module, string name) where T : class
            {
                var address = NativeMethods.GetProcAddress(module, name);
                if (address == IntPtr.Zero) throw new EntryPointNotFoundException(name);
                return (T)(object)Marshal.GetDelegateForFunctionPointer(address, typeof(T));
            }
        }

        private static class NativeMethods
        {
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            internal static extern IntPtr LoadLibrary(string path);
            [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)]
            internal static extern IntPtr GetProcAddress(IntPtr module, string name);
            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool FreeLibrary(IntPtr module);
        }
    }

    internal sealed class StemTeXSvgResult
    {
        internal StemTeXSvgResult(byte[] bytes, string summaryJson, int outcomeCode, int issueFlags, string outcomeMessage, double depthPt)
        {
            Bytes = bytes; SummaryJson = summaryJson; OutcomeCode = outcomeCode;
            IssueFlags = issueFlags; OutcomeMessage = outcomeMessage; DepthPt = depthPt;
        }
        internal byte[] Bytes { get; }
        internal string SummaryJson { get; }
        internal int OutcomeCode { get; }
        internal int IssueFlags { get; }
        internal string OutcomeMessage { get; }
        internal double DepthPt { get; }
    }

    internal sealed class StemTeXException : Exception
    {
        internal StemTeXException(int errorCode, string message) : base(message) { ErrorCode = errorCode; }
        internal int ErrorCode { get; }
    }
}
