using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace LaTeXBlocks.Word
{
    internal static class ColorSyncDiagnostics
    {
        private static readonly object Gate = new object();
        private static readonly bool Enabled = ResolveEnabled();
        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LaTeXBlocks");
        internal static readonly string LogPath = Path.Combine(LogDirectory,
            "color-sync.log");

        internal static void StartSession(string assemblyVersion)
        {
            if (!Enabled) return;
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogDirectory);
                    File.WriteAllText(LogPath, Format("session", "start version=" +
                        assemblyVersion), new UTF8Encoding(false));
                }
            }
            catch { }
        }

        internal static void Write(string stage, string message)
        {
            if (!Enabled) return;
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogDirectory);
                    File.AppendAllText(LogPath, Format(stage, message),
                        new UTF8Encoding(false));
                }
            }
            catch { }
        }

        internal static void Exception(string stage, Exception exception)
        {
            Write(stage, "exception=" + (exception == null ? "<null>" :
                exception.GetBaseException().GetType().FullName + ": " +
                exception.GetBaseException().Message));
        }

        private static string Format(string stage, string message)
        {
            return DateTime.Now.ToString("O", CultureInfo.InvariantCulture) +
                " pid=" + Process.GetCurrentProcess().Id +
                " tid=" + Thread.CurrentThread.ManagedThreadId +
                " stage=" + stage + " " + (message ?? string.Empty) +
                Environment.NewLine;
        }

        private static bool ResolveEnabled()
        {
            if (string.Equals(Environment.GetEnvironmentVariable(
                    "LATEXBLOCKS_COLOR_SYNC_TRACE"), "1",
                    StringComparison.Ordinal))
                return true;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                           @"Software\LaTeXBlocks\Word"))
                    return Convert.ToInt32(key?.GetValue(
                        "ColorSyncDiagnostics", 0)) != 0;
            }
            catch { return false; }
        }
    }
}
