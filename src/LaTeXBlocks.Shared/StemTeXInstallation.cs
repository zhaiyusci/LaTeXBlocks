using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

#if POWERPOINT
namespace LaTeXBlocks.PowerPoint
#else
namespace LaTeXBlocks.Word
#endif
{
    /// <summary>
    /// Resolves the immutable StemTeX engine distribution separately from the
    /// system TeX Live package/font tree and the user's generated profiles.
    /// </summary>
    internal static class StemTeXInstallation
    {
        internal const string HomeEnvironmentVariable = "STEMTEX_HOME";
        internal const string TexLiveEnvironmentVariable = "STEMTEX_TEXMF_ROOT";
        internal const string SettingsKey = @"Software\LaTeXBlocks";
        internal const string TexLiveRootValue = "TeXLiveRoot";
        internal static readonly Version RequiredRuntimeVersion = new Version(0, 15, 0);

        internal static string ApplicationDataRoot => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LaTeXBlocks");

        internal static string UserProfilesRoot => Path.Combine(ApplicationDataRoot, "profiles");

        internal static string LegacyUserProfilesRoot => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "StemTeX", "profiles");

        internal static string ResolveStemTeXHome()
        {
            var configured = Environment.GetEnvironmentVariable(HomeEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(configured))
            {
                var configuredHome = ValidateStemTeXHome(configured);
                if (configuredHome != null) return configuredHome;
                throw new DirectoryNotFoundException(HomeEnvironmentVariable +
                    " does not contain StemTeX " + RequiredRuntimeVersion +
                    " or newer with SVG support: " + configured);
            }

            var candidates = new List<string>();
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey))
                {
                    var installedHome = key?.GetValue("StemTeXHome") as string;
                    if (!string.IsNullOrWhiteSpace(installedHome)) candidates.Add(installedHome);
                }
            }
            catch { }

            var systemRoot = Path.GetPathRoot(Environment.SystemDirectory);
            if (!string.IsNullOrWhiteSpace(systemRoot))
                candidates.Add(Path.Combine(systemRoot, "StemTeX"));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "Scholia", "StemTeX"));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "StemTeX"));

            // Installed runtimes precede development stages at the same version.
            // A strictly newer read-only development stage may still win.
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            AddPackagedDevelopmentStages(Path.Combine(documents, "xetex", "stemtex", "dist",
                "package-stage"), candidates);
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "staging"));
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "dist",
                "stemtex-installer", "StemTeX"));
            candidates.Add(Path.Combine(documents, "xetex", "stemtex", "build",
                "stemtex-check-stage"));

            string bestHome = null;
            Version bestVersion = null;
            foreach (var candidate in candidates)
            {
                var home = ValidateStemTeXHome(candidate);
                if (home == null) continue;
                var version = ReadRuntimeVersion(home);
                if (bestHome == null || version.CompareTo(bestVersion) > 0)
                {
                    bestHome = home;
                    bestVersion = version;
                }
            }
            if (bestHome != null) return bestHome;
            throw new DirectoryNotFoundException(
                "StemTeX 0.13 or newer with SVG support was not found. Install LaTeXBlocks " +
                "or set " + HomeEnvironmentVariable + " to a compatible StemTeX distribution.");
        }

        internal static string[] DiscoverProfiles()
        {
            var home = ResolveStemTeXHome();
            EnsureUserProfilesRoot();
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // User profiles take precedence when their directory name matches an
            // installed profile. This mirrors the creator's writable collection and
            // lets a user intentionally replace the bundled default without editing it.
            AddProfiles(UserProfilesRoot, result, seen);
            // Only used if a legacy profile could not be copied during migration.
            AddProfiles(LegacyUserProfilesRoot, result, seen);
            AddProfiles(Path.Combine(home, "gui", "profiles"), result, seen);
            result.Sort(StringComparer.OrdinalIgnoreCase);
            return result.ToArray();
        }

        internal static string ResolveProfileRoot(string stemTeXHome, string profile)
        {
            if (string.IsNullOrWhiteSpace(profile))
                throw new ArgumentException("A StemTeX profile is required.", nameof(profile));
            EnsureUserProfilesRoot();
            foreach (var root in new[]
            {
                UserProfilesRoot,
                LegacyUserProfilesRoot,
                Path.Combine(stemTeXHome, "gui", "profiles")
            })
            {
                if (!Directory.Exists(root)) continue;
                foreach (var directory in Directory.GetDirectories(root))
                {
                    if (!string.Equals(Path.GetFileName(directory), profile,
                        StringComparison.OrdinalIgnoreCase)) continue;
                    if (IsUsableProfile(directory)) return Path.GetFullPath(directory);
                }
            }
            throw new DirectoryNotFoundException("StemTeX profile was not found: " + profile);
        }

        internal static void EnsureUserProfilesRoot()
        {
            MigrateUserProfiles(LegacyUserProfilesRoot, UserProfilesRoot);
        }

        internal static void MigrateUserProfiles(string legacyRoot, string targetRoot)
        {
            if (string.IsNullOrWhiteSpace(targetRoot))
                throw new ArgumentException("A target profile directory is required.",
                    nameof(targetRoot));
            Directory.CreateDirectory(targetRoot);
            if (string.IsNullOrWhiteSpace(legacyRoot) || !Directory.Exists(legacyRoot) ||
                string.Equals(Path.GetFullPath(legacyRoot).TrimEnd(Path.DirectorySeparatorChar),
                    Path.GetFullPath(targetRoot).TrimEnd(Path.DirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase)) return;

            foreach (var source in Directory.GetDirectories(legacyRoot))
            {
                if (!IsUsableProfile(source)) continue;
                var destination = Path.Combine(targetRoot, Path.GetFileName(source));
                if (Directory.Exists(destination)) continue;
                var staging = Path.Combine(targetRoot, ".migrate-" + Guid.NewGuid().ToString("N"));
                try
                {
                    CopyDirectory(source, staging);
                    if (!IsUsableProfile(staging)) continue;
                    try { Directory.Move(staging, destination); }
                    catch (IOException)
                    {
                        if (!Directory.Exists(destination)) throw;
                    }
                }
                catch (UnauthorizedAccessException) { }
                catch (IOException) { }
                finally
                {
                    try { if (Directory.Exists(staging)) Directory.Delete(staging, true); }
                    catch { }
                }
            }
        }

        internal static string ResolveTexLiveRoot()
        {
            var environmentRoot = Environment.GetEnvironmentVariable(TexLiveEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(environmentRoot))
            {
                string normalized;
                if (TryNormalizeTexLiveRoot(environmentRoot, out normalized)) return normalized;
                throw InvalidTexLiveRoot(TexLiveEnvironmentVariable, environmentRoot);
            }

            var configuredRoot = ReadConfiguredTexLiveRoot();
            if (!string.IsNullOrWhiteSpace(configuredRoot))
            {
                string normalized;
                if (TryNormalizeTexLiveRoot(configuredRoot, out normalized)) return normalized;
                throw InvalidTexLiveRoot("the saved LaTeXBlocks setting", configuredRoot);
            }

            foreach (var candidate in DiscoverSystemTexLiveRoots())
            {
                string normalized;
                if (TryNormalizeTexLiveRoot(candidate, out normalized)) return normalized;
            }
            throw new DirectoryNotFoundException(
                "A system TeX Live tree was not found. Use the LaTeXBlocks TeX Live command " +
                "to select a directory containing texmf-dist and texmf-dist\\web2c.");
        }

        internal static string ReadConfiguredTexLiveRoot()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey))
                    return key?.GetValue(TexLiveRootValue) as string;
            }
            catch { return null; }
        }

        internal static void WriteConfiguredTexLiveRoot(string value)
        {
            string normalized = null;
            if (!string.IsNullOrWhiteSpace(value) && !TryNormalizeTexLiveRoot(value, out normalized))
                throw InvalidTexLiveRoot("the selected directory", value);
            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
            {
                if (normalized == null) key.DeleteValue(TexLiveRootValue, false);
                else key.SetValue(TexLiveRootValue, normalized, RegistryValueKind.String);
            }
        }

        internal static bool TryNormalizeTexLiveRoot(string value, out string normalized)
        {
            normalized = null;
            if (string.IsNullOrWhiteSpace(value)) return false;
            try
            {
                var root = Path.GetFullPath(Environment.ExpandEnvironmentVariables(value.Trim()));
                if (string.Equals(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar)), "texmf-dist", StringComparison.OrdinalIgnoreCase))
                    root = Directory.GetParent(root)?.FullName;
                else if (string.Equals(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar,
                             Path.AltDirectorySeparatorChar)), "web2c", StringComparison.OrdinalIgnoreCase) &&
                         string.Equals(Directory.GetParent(root)?.Name, "texmf-dist",
                             StringComparison.OrdinalIgnoreCase))
                    root = Directory.GetParent(Directory.GetParent(root).FullName)?.FullName;
                if (string.IsNullOrWhiteSpace(root) ||
                    !Directory.Exists(Path.Combine(root, "texmf-dist")) ||
                    !Directory.Exists(Path.Combine(root, "texmf-dist", "web2c"))) return false;
                normalized = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar);
                return true;
            }
            catch (Exception exception) when (exception is ArgumentException ||
                                              exception is NotSupportedException ||
                                              exception is PathTooLongException ||
                                              exception is IOException ||
                                              exception is UnauthorizedAccessException)
            {
                return false;
            }
        }

        internal static string ResolveProfileCreatorExecutable(string stemTeXHome)
        {
            var executable = Path.Combine(stemTeXHome, "gui", "stemtex-profile-creator.exe");
            var library = Path.Combine(stemTeXHome, "runtime", "bin", "sdk", "stemtex-profile.dll");
            if (!File.Exists(executable) || !File.Exists(library))
                throw new FileNotFoundException(
                    "The bundled StemTeX Profile Creator is unavailable. Reinstall LaTeXBlocks.", executable);
            return executable;
        }

        private static void AddProfiles(string root, List<string> profiles, HashSet<string> seen)
        {
            if (!Directory.Exists(root)) return;
            foreach (var directory in Directory.GetDirectories(root))
            {
                var name = Path.GetFileName(directory);
                if (IsUsableProfile(directory) && seen.Add(name)) profiles.Add(name);
            }
        }

        private static bool IsUsableProfile(string directory)
        {
            return File.Exists(Path.Combine(directory, "preamble.tex")) &&
                   File.Exists(Path.Combine(directory, "warmup.tex"));
        }

        private static void CopyDirectory(string source, string destination)
        {
            var sourceInfo = new DirectoryInfo(source);
            if ((sourceInfo.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Profile directories cannot be reparse points: " + source);
            Directory.CreateDirectory(destination);
            foreach (var file in sourceInfo.GetFiles())
                file.CopyTo(Path.Combine(destination, file.Name), false);
            foreach (var directory in sourceInfo.GetDirectories())
            {
                if ((directory.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                CopyDirectory(directory.FullName, Path.Combine(destination, directory.Name));
            }
        }

        private static string ValidateStemTeXHome(string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return null;
            try
            {
                var home = Path.GetFullPath(Environment.ExpandEnvironmentVariables(candidate.Trim()));
                var version = ReadRuntimeVersion(home);
                if (version == null || version.CompareTo(RequiredRuntimeVersion) < 0) return null;
                return File.Exists(Path.Combine(home, "runtime", "bin", "sdk", "stemtex-renderer.dll")) &&
                       File.Exists(Path.Combine(home, "runtime", "bin", "windows", "dvisvgmdaemon.dll"))
                    ? home : null;
            }
            catch { return null; }
        }

        private static Version ReadRuntimeVersion(string home)
        {
            try
            {
                var path = Path.Combine(home, "runtime", "VERSION");
                Version version;
                return File.Exists(path) && Version.TryParse(File.ReadAllText(path).Trim(), out version)
                    ? version : null;
            }
            catch { return null; }
        }

        private static IEnumerable<string> DiscoverSystemTexLiveRoots()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();
            var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var entry in path.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                try
                {
                    var bin = entry.Trim().Trim('"');
                    if (!File.Exists(Path.Combine(bin, "kpsewhich.exe"))) continue;
                    AddCandidate(Path.GetFullPath(Path.Combine(bin, "..", "..")), result, seen);
                }
                catch { }
            }

            var systemDrive = Path.GetPathRoot(Environment.SystemDirectory);
            if (!string.IsNullOrWhiteSpace(systemDrive))
                AddVersionedTexLiveRoots(Path.Combine(systemDrive, "texlive"), result, seen);
            AddVersionedTexLiveRoots(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "texlive"), result, seen);
            return result;
        }

        private static void AddVersionedTexLiveRoots(string parent, List<string> result,
            HashSet<string> seen)
        {
            if (!Directory.Exists(parent)) return;
            var directories = new List<string>(Directory.GetDirectories(parent));
            directories.Sort((left, right) => StringComparer.OrdinalIgnoreCase.Compare(right, left));
            foreach (var directory in directories) AddCandidate(directory, result, seen);
        }

        private static void AddCandidate(string candidate, List<string> result, HashSet<string> seen)
        {
            if (!string.IsNullOrWhiteSpace(candidate) && seen.Add(candidate)) result.Add(candidate);
        }

        private static void AddPackagedDevelopmentStages(string parent, List<string> candidates)
        {
            try
            {
                if (!Directory.Exists(parent)) return;
                var directories = new List<string>(Directory.GetDirectories(parent, "StemTeX-*"));
                directories.Sort((left, right) =>
                    StringComparer.OrdinalIgnoreCase.Compare(right, left));
                candidates.AddRange(directories);
            }
            catch { }
        }

        private static DirectoryNotFoundException InvalidTexLiveRoot(string source, string value)
        {
            return new DirectoryNotFoundException("The TeX Live root from " + source +
                " is not usable: " + value + ". Select a TeX Live directory containing " +
                "texmf-dist and texmf-dist\\web2c.");
        }
    }
}
