using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    internal static class StemTeXUserTools
    {
        internal static bool TryChooseTexLiveRoot(IWin32Window owner, out string selectedRoot)
        {
            selectedRoot = null;
            var environmentRoot = Environment.GetEnvironmentVariable(
                StemTeXInstallation.TexLiveEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(environmentRoot))
                throw new InvalidOperationException("The TeX Live directory is controlled by " +
                    StemTeXInstallation.TexLiveEnvironmentVariable + ". Remove that environment " +
                    "variable before changing the saved LaTeXBlocks setting.");

            string current = null;
            try { current = StemTeXInstallation.ResolveTexLiveRoot(); }
            catch (DirectoryNotFoundException) { }
            using (var dialog = new FolderBrowserDialog
            {
                Description = "Select the system TeX Live root containing texmf-dist.",
                SelectedPath = current ?? string.Empty,
                ShowNewFolderButton = false
            })
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;
                if (!StemTeXInstallation.TryNormalizeTexLiveRoot(dialog.SelectedPath, out selectedRoot))
                    throw new DirectoryNotFoundException(
                        "The selected directory is not a TeX Live root containing texmf-dist and " +
                        "texmf-dist\\web2c: " + dialog.SelectedPath);
                return true;
            }
        }

        internal static Process StartProfileCreator(EventHandler exited)
        {
            var stemTeXHome = StemTeXInstallation.ResolveStemTeXHome();
            var texLiveRoot = StemTeXInstallation.ResolveTexLiveRoot();
            var executable = StemTeXInstallation.ResolveProfileCreatorExecutable(stemTeXHome);
            StemTeXInstallation.EnsureUserProfilesRoot();
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = executable,
                    Arguments = "--runtime " + Quote(Path.Combine(stemTeXHome, "runtime")) +
                                " --texmf " + Quote(texLiveRoot) +
                                " --profiles " + Quote(StemTeXInstallation.UserProfilesRoot),
                    WorkingDirectory = Path.GetDirectoryName(executable),
                    UseShellExecute = false
                },
                EnableRaisingEvents = true
            };
            if (exited != null) process.Exited += exited;
            try
            {
                if (!process.Start())
                    throw new InvalidOperationException("StemTeX Profile Creator did not start.");
                return process;
            }
            catch
            {
                process.Dispose();
                throw;
            }
        }

        private static string Quote(string value)
        {
            if (value == null) return "\"\"";
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }
    }
}
