using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace LaTeXBlocks.Word
{
    internal sealed class StemTeXRenderer : IDisposable
    {
        private const int SvgOutputFormat = 1;
        private readonly object gate = new object();
        private readonly string stemTeXHome;
        private readonly string runtimeRoot;
        private readonly string profileRoot;
        private IntPtr library;
        private IntPtr renderer;
        private NativeApi api;
        private bool disposed;

        internal StemTeXRenderer(string profile = "unicodemath_cjk")
        {
            stemTeXHome = ResolveStemTeXHome(profile);
            runtimeRoot = Path.Combine(stemTeXHome, "runtime");
            profileRoot = Path.Combine(stemTeXHome, "gui", "profiles", profile);
        }

        internal string StemTeXHome => stemTeXHome;
        internal string ProfileRoot => profileRoot;

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

        internal StemTeXSvgResult RenderSvg(string source, double widthPt, bool autoWidth = false)
        {
            if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("LaTeX source cannot be empty.", nameof(source));
            if (widthPt <= 0) throw new ArgumentOutOfRangeException(nameof(widthPt));

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
                    var ok = api.RenderOutputBytes(renderer, sourceUtf8.Pointer, renderWidth, SvgOutputFormat,
                        out bytes, out result, out errorCode, out error);
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

        private static string ResolveStemTeXHome(string profile)
        {
            var candidates = new List<string>();
            var configured = Environment.GetEnvironmentVariable("STEMTEX_HOME");
            if (!string.IsNullOrWhiteSpace(configured)) candidates.Add(configured);
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Scholia", "StemTeX"));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "StemTeX"));
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "dist", "stemtex-installer", "StemTeX"));
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "build", "stemtex-check-stage"));

            foreach (var candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                var home = Path.GetFullPath(candidate);
                var profilesRoot = Path.Combine(home, "gui", "profiles");
                var hasProfile = profile == null
                    ? Directory.Exists(profilesRoot) && Directory.GetFiles(profilesRoot, "preamble.tex", SearchOption.AllDirectories).Length > 0
                    : File.Exists(Path.Combine(profilesRoot, profile, "preamble.tex"));
                if (File.Exists(Path.Combine(home, "runtime", "bin", "sdk", "stemtex-renderer.dll")) &&
                    File.Exists(Path.Combine(home, "runtime", "bin", "windows", "dvisvgmdaemon.dll")) && hasProfile)
                    return home;
            }
            throw new DirectoryNotFoundException(
                "StemTeX with SVG support was not found. Install StemTeX 0.10 or set STEMTEX_HOME to its installation root.");
        }

        public void Dispose()
        {
            lock (gate)
            {
                if (disposed) return;
                disposed = true;
                if (renderer != IntPtr.Zero && api != null) api.Destroy(renderer);
                renderer = IntPtr.Zero;
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
            // A block/display has no single baseline to align to surrounding Word text.
            if (source.IndexOf('\n') >= 0 || source.IndexOf('\r') >= 0 ||
                source.IndexOf("\\[", StringComparison.Ordinal) >= 0 ||
                source.IndexOf("$$", StringComparison.Ordinal) >= 0 ||
                source.IndexOf("\\begin{", StringComparison.Ordinal) >= 0)
            {
                if (autoWidth) throw new ArgumentException("Auto width supports one-line inline LaTeX only. Use Fixed width for display or multiline content.", nameof(source));
                return source;
            }

            if (autoWidth)
            {
                return "\\begingroup\\setbox255=\\hbox{" + source + "}" +
                       "\\leavevmode\\special{dvisvgm:raw <g id='latexblocks-start' data-x='{?x}' data-y='{?y}'/>}" +
                       "\\box255\\special{dvisvgm:raw <g id='latexblocks-end' data-x='{?x}'/>}\\endgroup";
            }

            // dvisvgm expands {?y} to the current TeX baseline without adding visible geometry.
            return "\\leavevmode\\special{dvisvgm:raw <g id='latexblocks-baseline' data-y='{?y}'/>}" + source;
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
                var startX = ReadMarkerCoordinate(marker.Value, "data-x");
                var endX = ReadMarkerCoordinate(endMarker.Value, "data-x");
                var naturalWidth = endX - startX;
                if (!(naturalWidth > 0) || naturalWidth > 2000) throw new InvalidDataException("StemTeX returned an invalid natural formula width.");
                svg = svg.Remove(endMarker.Index, endMarker.Length);

                const double borderPt = 1.0; // Matches the profile's existing PreviewBorder.
                var croppedX = startX - borderPt;
                var croppedWidth = naturalWidth + 2 * borderPt;
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
            internal delegate int RenderOutputBytesDelegate(IntPtr renderer, IntPtr source, double widthPt, int format,
                out StemTeXOutputBytes bytes, out StemTeXRenderOutputResult result, out int errorCode, out IntPtr error);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeOutputBytesDelegate(ref StemTeXOutputBytes bytes);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeOutputResultDelegate(ref StemTeXRenderOutputResult result);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void FreeStringDelegate(IntPtr value);
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            internal delegate void DestroyDelegate(IntPtr renderer);

            internal NativeApi(IntPtr module)
            {
                Create = Load<CreateDelegate>(module, "stemtex_renderer_create");
                RenderOutputBytes = Load<RenderOutputBytesDelegate>(module, "stemtex_renderer_render_output_bytes");
                FreeOutputBytes = Load<FreeOutputBytesDelegate>(module, "stemtex_renderer_free_output_bytes");
                FreeOutputResult = Load<FreeOutputResultDelegate>(module, "stemtex_renderer_free_output_result");
                FreeString = Load<FreeStringDelegate>(module, "stemtex_renderer_free_string");
                Destroy = Load<DestroyDelegate>(module, "stemtex_renderer_destroy");
            }

            internal CreateDelegate Create { get; }
            internal RenderOutputBytesDelegate RenderOutputBytes { get; }
            internal FreeOutputBytesDelegate FreeOutputBytes { get; }
            internal FreeOutputResultDelegate FreeOutputResult { get; }
            internal FreeStringDelegate FreeString { get; }
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
