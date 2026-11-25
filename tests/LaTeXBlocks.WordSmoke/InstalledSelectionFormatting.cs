using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using LaTeXBlocks.Word;
using Office = Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.WordSmoke
{
    internal static class InstalledSelectionFormatting
    {
        private const float FontSize = 24;
        private const float Tint = 0.8f;
        private const float Shade = -0.4f;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        internal static void Run(string directory)
        {
            var prior = Process.GetProcessesByName("WINWORD").Select(p => p.Id).ToArray();
            WordInterop.Application word = null;
            WordInterop.Document document = null;
            StemTeXBackend backend = null;
            var ownsWord = false;
            try
            {
                word = new WordInterop.Application();
                Require(!word.Visible && word.Documents.Count == 0, "A separate hidden Word instance is required.");
                document = word.Documents.Add();
                GetWindowThreadProcessId(new IntPtr(document.Windows[1].Hwnd), out var pid);
                // Other tests can start Word at the same time. The document HWND
                // identifies the process returned by our own COM activation.
                Require(pid != 0 && !prior.Contains((int)pid), "The new document belongs to an existing Word process.");
                ownsWord = true;
                word.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone;
                var options = Environment.GetCommandLineArgs();
                var workspaceOnly = options.Contains("--workspace-color-reconcile") ||
                    options.Contains("--adjacent-drawing-runs");
                string activeProfile = null;
                if (!workspaceOnly)
                {
                    var addIn = word.COMAddIns.Item("LaTeXBlocks.Word.AddIn");
                    Require(addIn.Connect, "The registered Word add-in did not connect at startup.");
                    dynamic diagnostics = addIn.Object;
                    activeProfile = diagnostics.CurrentProfile;
                    string loaded = diagnostics.AssemblyLocation;
                    using (var stream = File.OpenRead(loaded))
                    using (var algorithm = SHA256.Create())
                        Console.WriteLine("Word PID {0}, loaded {1}, SHA256 {2}, backend {3}", pid,
                            loaded, BitConverter.ToString(algorithm.ComputeHash(stream)).Replace("-", ""),
                            (string)diagnostics.BackendStartup);
                }

                // The test program uses its own renderer only to seed the fixture.
                // All subsequent changes come from native Word selection commands.
                document.Content.Text = "First xx tail\rSecond yy tail\rOutside selection.";
                document.Content.Font.Size = 12;
                backend = new StemTeXBackend();
                if (string.IsNullOrEmpty(activeProfile)) activeProfile = backend.DefaultAvailableProfile;
                var service = new LaTeXBlockService(word, backend);
                foreach (var item in new[] { ("xx", "$a^2$"), ("yy", "$b_2$") })
                {
                    var start = document.Content.Text.IndexOf(item.Item1, StringComparison.Ordinal);
                    document.Range(start, start + 2).Select();
                    var render = service.RenderCommittedAsync(item.Item2, 360,
                        LaTeXBlockLayoutMode.Auto, activeProfile, 12)
                        .GetAwaiter().GetResult();
                    service.InsertRendered(item.Item2, 360, LaTeXBlockLayoutMode.Auto, render);
                }
                LeaveSelection(document);
                var initial = new Dictionary<Guid, float>();
                var expectedFrames = new Dictionary<(Guid, float), (double Width, double Height)>();
                for (var index = 1; index <= document.InlineShapes.Count; index++)
                {
                    var shape = document.InlineShapes[index];
                    Require(LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) &&
                        metadata.FontSizePt == 12, "The fixture did not start with 12 pt formulas.");
                    initial.Add(metadata.Id, shape.Height);
                    if (!workspaceOnly && !options.Contains("--color-only"))
                    {
                        foreach (var size in new[] { 14f, FontSize })
                        {
                            var expected = service.RenderCommittedAsync(source, metadata.WidthPt, metadata.Mode,
                                activeProfile, size, renderKind: metadata.Kind).GetAwaiter().GetResult();
                            expectedFrames.Add((metadata.Id, size), (LaTeXBlockService.ReadSvgWidthPt(expected.SvgBytes),
                                LaTeXBlockService.ReadSvgHeightPt(expected.SvgBytes)));
                        }
                    }
                }
                Require(initial.Count == 2, "The fixture needs two distinct inline formulas.");
                Require(LaTeXBlockMetadata.TryParse(document.InlineShapes[1].AlternativeText,
                    out var firstFormula), "The first formula has no metadata.");
                var seed = Path.Combine(directory, "selection-fixture.docx");
                document.SaveAs2(seed, WordInterop.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                var themedSeed = Path.Combine(directory, "selection-theme-fixture.docx");
                foreach (WordInterop.InlineShape shape in document.InlineShapes)
                {
                    var format = shape.Range.Font.TextColor;
                    format.ObjectThemeColor = WordInterop.WdThemeColorIndex.wdThemeColorAccent6;
                    format.TintAndShade = Tint;
                    LaTeXBlockService.ApplySvgForegroundThemeFill(shape,
                        WordInterop.WdThemeColorIndex.wdThemeColorAccent6, Tint);
                }
                document.SaveAs2(themedSeed, WordInterop.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                if (!options.Contains("--exact-size-only") && !options.Contains("--held-selection-only") &&
                    !options.Contains("--held-rapid-only"))
                {
                    var adjacent = word.Documents.Add();
                    try { CheckAdjacentDrawingRuns(adjacent, document.InlineShapes[1], directory); }
                    finally { adjacent.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges); }
                }
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                document = null;
                if (options.Contains("--adjacent-drawing-runs")) return;
                var errors = new List<string>();
                var stages = options.Contains("--color-only")
                    ? new[] { "red", "theme", "red-leave", "theme-leave", "mixed" }
                    : new[] { "grow", "size", "red", "theme", "rapid", "size-leave", "red-leave", "theme-leave", "mixed" };
                if (options.Contains("--workspace-color-reconcile")) stages = new[] { "retained" };
                if (options.Contains("--exact-size-only")) stages = new[] { "size", "size-leave" };
                if (options.Contains("--held-selection-only")) stages = new[] { "grow", "size", "red", "theme", "rapid" };
                if (options.Contains("--held-rapid-only")) stages = new[] { "rapid" };
                foreach (var scenario in stages)
                {
                    // Start every command from a separate document. A Grow Font
                    // request can still be pending when another command is issued.
                    var leavesSelection = scenario.EndsWith("-leave", StringComparison.Ordinal);
                    var stage = leavesSelection ? scenario.Substring(0, scenario.Length - "-leave".Length) : scenario;
                    var path = Path.Combine(directory, "selection-" + scenario + ".docx");
                    var mixed = stage == "mixed" || stage == "retained";
                    File.Copy(mixed ? themedSeed : seed, path);
                    document = word.Documents.Open(path, AddToRecentFiles: false, Visible: false);
                    var lastShape = document.InlineShapes[2];
                    document.Range(0, lastShape.Range.End + 2).Select();
                    Pump(250);
                    var selection = word.Selection;
                    Require(selection.InlineShapes.Count == 2 && selection.Text.Contains("First"),
                        "The selection must span ordinary text and both formulas.");
                    var selectedStart = selection.Start;
                    var selectedEnd = selection.End;
                    var requestedSize = FontSize;
                    var retainsSelection = !leavesSelection && stage != "mixed";
                    var commit = stage == "retained" ? CaptureWorkspaceColorCommit(word, service) : null;
                    if (stage == "grow")
                    {
                        word.CommandBars.ExecuteMso("FontSizeIncreaseWord");
                        requestedSize = document.Range(0, 1).Font.Size;
                        Require(requestedSize == 14, "The native Grow Font command did not change 12 pt text to 14 pt.");
                    }
                    else if (stage == "size") selection.Font.Size = FontSize;
                    else if (stage == "red") selection.Font.Color = WordInterop.WdColor.wdColorRed;
                    else if (stage == "theme")
                    {
                        selection.Font.TextColor.ObjectThemeColor = WordInterop.WdThemeColorIndex.wdThemeColorAccent6;
                        selection.Font.TextColor.TintAndShade = Tint;
                    }
                    else if (stage == "rapid")
                    {
                        // Accept a color while a size render may still be pending,
                        // then change size twice without moving the mixed selection.
                        selection.Font.Size = FontSize;
                        selection.Font.TextColor.ObjectThemeColor = WordInterop.WdThemeColorIndex.wdThemeColorAccent6;
                        selection.Font.TextColor.TintAndShade = Tint;
                        Require(selection.Font.TextColor.ObjectThemeColor == WordInterop.WdThemeColorIndex.wdThemeColorAccent6 &&
                            Math.Abs(selection.Font.TextColor.TintAndShade - Tint) < 0.01,
                            "Word did not accept the new theme tint during the size change.");
                        Pump(180);
                        selection.Font.Size = 18;
                        Pump(120);
                        requestedSize = 14;
                        selection.Font.Size = requestedSize;
                    }
                    else
                    {
                        // Both drawings already have the same theme tint. Change just
                        // its variant on the first and reset the second while keeping
                        // the same selection. Each drawing must recover its own run.
                        document.InlineShapes[1].Range.Font.TextColor.TintAndShade = Shade;
                        document.InlineShapes[2].Range.Font.Color = WordInterop.WdColor.wdColorAutomatic;
                        if (commit != null)
                        {
                            Require(Math.Abs(document.InlineShapes[1].Fill.ForeColor.Brightness - Tint) < 0.01 &&
                                document.InlineShapes[2].Fill.ForeColor.Type == Office.MsoColorType.msoColorTypeScheme,
                                "A host event colored the fixture before the workspace callback ran.");
                            commit();
                        }
                    }
                    if (!retainsSelection) LeaveSelection(document);
                    var timeout = Stopwatch.StartNew();
                    var stable = new Stopwatch();
                    string mismatch;
                    while (true)
                    {
                        // Read only. Saving, switching views, running another Word
                        // command or leaving this selection can itself trigger an update.
                        if (retainsSelection && (word.Selection.Start != selectedStart ||
                            word.Selection.End != selectedEnd))
                        {
                            mismatch = "Word moved the original mixed selection";
                            break;
                        }
                        mismatch = LiveMismatch(document, initial, expectedFrames, stage, requestedSize, firstFormula.Id);
                        if (mismatch == null)
                        {
                            if (stage != "rapid") break;
                            if (!stable.IsRunning) stable.Start();
                            if (stable.ElapsedMilliseconds >= 750) break;
                        }
                        else stable.Reset();
                        if (timeout.Elapsed >= TimeSpan.FromSeconds(20))
                        {
                            if (mismatch == null) mismatch = "the final size and color did not stay stable";
                            break;
                        }
                        Pump(100);
                    }
                    if (mismatch != null)
                    {
                        var failure = scenario + " in Word: " + mismatch;
                        Console.WriteLine(failure);
                        if (scenario == "rapid")
                        {
                            foreach (WordInterop.InlineShape shape in document.InlineShapes)
                            {
                                var fill = shape.Fill.ForeColor;
                                LaTeXBlockMetadata.TryParse(shape.AlternativeText, out var meta);
                                Console.WriteLine($"Formula {meta?.Id}, run {shape.Range.Font.Size:g} pt, " +
                                    $"metadata {meta?.FontSizePt:g} pt, frame {shape.Width:g}×{shape.Height:g}, " +
                                    $"fill {fill.Type}, RGB {fill.RGB:X6}, brightness {fill.Brightness:g}");
                            }
                        }
                        errors.Add(failure);
                        document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                        document = null;
                        continue;
                    }
                    if (retainsSelection)
                    {
                        Require(word.Selection.Start == selectedStart && word.Selection.End == selectedEnd &&
                            word.Selection.InlineShapes.Count == 2,
                            scenario + " did not retain the mixed selection.");
                    }
                    else Require(word.Selection.Start == word.Selection.End,
                        "An asynchronous formatting update reselected the old range.");
                    timeout.Stop();
                    document.Save();
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    document = null;
                    Console.WriteLine("{0} passed live in {1:F2} s. {2}", scenario, timeout.Elapsed.TotalSeconds, path);
                    try { CheckPackage(path, initial, stage, requestedSize, firstFormula.Id); }
                    catch (Exception ex) { errors.Add(stage + " on disk: " + ex.Message); }
                }
                Require(errors.Count == 0, string.Join(Environment.NewLine, errors));
            }
            finally
            {
                backend?.Dispose();
                try { document?.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges); }
                finally
                {
                    if (ownsWord && word.Documents.Count == 0)
                        word.Quit(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                }
            }
        }

        private static void CheckAdjacentDrawingRuns(WordInterop.Document document,
            WordInterop.InlineShape source, string directory)
        {
            document.Content.Text = "AB";
            foreach (var marker in new[] { 'A', 'B' })
            {
                var position = document.Content.Text.IndexOf(marker);
                document.Range(position, position + 1).FormattedText = source.Range.FormattedText;
            }
            var first = document.InlineShapes[1];
            var second = document.InlineShapes[2];
            Require(document.InlineShapes.Count == 2 && first.Range.End == second.Range.Start,
                $"Expected two adjacent drawings, found {document.InlineShapes.Count} " +
                $"at {first.Range.Start}–{first.Range.End} and {second.Range.Start}–{second.Range.End}.");
            Require(LaTeXBlockMetadata.TryParse(first.AlternativeText, out var original) &&
                LaTeXBlockMetadata.TryParse(second.AlternativeText, out var copied) && original.Id == copied.Id,
                "Word did not retain the formula identity in its formatted copy.");

            var leftTheme = WordInterop.WdThemeColorIndex.wdThemeColorAccent2;
            var rightTheme = WordInterop.WdThemeColorIndex.wdThemeColorAccent6;
            first.Range.Font.TextColor.ObjectThemeColor = leftTheme;
            first.Range.Font.TextColor.TintAndShade = 0.6f;
            second.Range.Font.TextColor.ObjectThemeColor = rightTheme;
            second.Range.Font.TextColor.TintAndShade = Shade;
            AssertDrawingRun(first, leftTheme, 0.6f);
            AssertDrawingRun(second, rightTheme, Shade);
            document.SaveAs2(Path.Combine(directory, "selection-adjacent-copies.docx"),
                WordInterop.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);

            // Word exposes ordinary pictures and formula pictures through the same
            // InlineShapes collection. Expand each target's run across a picture
            // without our metadata, once on the left and once on the right.
            var formulaDescription = first.AlternativeText;
            first.AlternativeText = "Ordinary Office SVG";
            Require(!LaTeXBlockService.TryReadContract(first, out _, out _),
                "The left control drawing still has a formula contract.");
            AssertDrawingRun(second, rightTheme, Shade);
            document.SaveAs2(Path.Combine(directory, "selection-adjacent-ordinary-first.docx"),
                WordInterop.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            first.AlternativeText = formulaDescription;
            second.AlternativeText = "Ordinary Office SVG";
            Require(!LaTeXBlockService.TryReadContract(second, out _, out _),
                "The right control drawing still has a formula contract.");
            AssertDrawingRun(first, leftTheme, 0.6f);
            Console.WriteLine("Adjacent duplicate and ordinary drawing runs matched their own themes and variants.");
        }

        private static void AssertDrawingRun(WordInterop.InlineShape shape,
            WordInterop.WdThemeColorIndex theme, float variant)
        {
            Require(LaTeXBlockService.NativeTextColorDescriptor.TryCaptureDrawingRun(shape, out var color) &&
                color.Kind == LaTeXBlockService.NativeTextColorKind.Theme && color.ThemeColor == theme &&
                Math.Abs(color.TintAndShade - variant) < 0.01,
                $"Drawing at {shape.Range.Start} expected {theme} {variant}, got " +
                $"{color.Kind} {color.ThemeColor} {color.TintAndShade}.");
        }

        // The Ribbon palette cannot be shown in a hidden Word process. Enter its
        // actual capture/commit callbacks from the workspace assembly while real
        // native ranges and fills remain live and the original selection is intact.
        private static Action CaptureWorkspaceColorCommit(WordInterop.Application word, LaTeXBlockService service)
        {
            const BindingFlags members = BindingFlags.Instance | BindingFlags.NonPublic;
            var assembly = typeof(LaTeXBlockService).Assembly;
            var type = assembly.GetType("LaTeXBlocks.Word.ThisAddIn", true);
            var owner = FormatterServices.GetUninitializedObject(type);
            type.GetField("Application", members).SetValue(owner, word);
            type.GetField("blocks", members).SetValue(owner, service);
            var captured = type.GetMethod("CaptureFontColorInteraction", members).Invoke(owner, new object[] { 1L });
            Require(captured != null, "The workspace color callback could not capture the mixed selection.");
            Console.WriteLine("Workspace palette callback: " + assembly.Location);
            return () => type.GetMethod("ReconcileFontColorInteraction", members).Invoke(owner, new[] { captured });
        }

        private static void LeaveSelection(WordInterop.Document doc)
        {
            var end = doc.Content.End - 1;
            doc.Range(end, end).Select();
        }

        private static void Pump(int milliseconds)
        {
            var timer = Stopwatch.StartNew();
            do { Application.DoEvents(); Thread.Sleep(15); } while (timer.ElapsedMilliseconds < milliseconds);
        }

        private static string LiveMismatch(WordInterop.Document doc, Dictionary<Guid, float> initial,
            Dictionary<(Guid, float), (double Width, double Height)> expectedFrames,
            string stage, float requestedSize, Guid firstFormula)
        {
            try { return ReadLiveMismatch(doc, initial, expectedFrames, stage, requestedSize, firstFormula); }
            catch (COMException error) when (error.HResult == unchecked((int)0x800A16C1))
            {
                // A rendered SVG may replace its drawing between two COM reads.
                // The caller will look up fresh shapes on the next poll.
                return "Word is replacing a formula drawing";
            }
        }

        private static string ReadLiveMismatch(WordInterop.Document doc, Dictionary<Guid, float> initial,
            Dictionary<(Guid, float), (double Width, double Height)> expectedFrames,
            string stage, float requestedSize, Guid firstFormula)
        {
            if (doc.InlineShapes.Count != initial.Count) return "formula count changed";
            for (var i = 1; i <= doc.InlineShapes.Count; i++)
            {
                var shape = doc.InlineShapes[i];
                if (!LaTeXBlockMetadata.TryParse(shape.AlternativeText, out var metadata) ||
                    !initial.TryGetValue(metadata.Id, out var originalHeight)) return "formula identity changed";
                var run = shape.Range.Font;
                if (stage == "size" || stage == "grow" || stage == "rapid")
                {
                    var expected = expectedFrames[(metadata.Id, requestedSize)];
                    if (Math.Abs(run.Size - requestedSize) > 0.01 ||
                        Math.Abs(metadata.FontSizePt - requestedSize) > 0.01 ||
                        Math.Abs(shape.Width - expected.Width) > 0.15 || Math.Abs(shape.Height - expected.Height) > 0.15)
                        return $"formula {i}, run {run.Size} pt, metadata {metadata.FontSizePt} pt, " +
                            $"frame {shape.Width:F2}×{shape.Height:F2} pt, expected {expected.Width:F2}×{expected.Height:F2} pt " +
                            $"(original height {originalHeight:F2} pt)";
                    if (stage != "rapid") continue;
                }
                var fill = shape.Fill;
                if (fill.Visible != Office.MsoTriState.msoTrue ||
                    fill.Type != Office.MsoFillType.msoFillSolid) return "formula " + i + " has no solid fill";
                var foreground = fill.ForeColor;
                if (stage == "red")
                {
                    if (run.Color != WordInterop.WdColor.wdColorRed ||
                        foreground.Type != Office.MsoColorType.msoColorTypeRGB ||
                        (int)foreground.RGB != (int)WordInterop.WdColor.wdColorRed)
                        return $"formula {i}, run {run.Color}, SVG fill {foreground.Type}/{foreground.RGB:X6}, expected red";
                }
                else if ((stage == "mixed" || stage == "retained") && metadata.Id != firstFormula)
                {
                    if (run.Color != WordInterop.WdColor.wdColorAutomatic ||
                        foreground.Type != Office.MsoColorType.msoColorTypeRGB || foreground.RGB != 0)
                        return "the second formula did not switch to automatic text and a black SVG fill";
                }
                else
                {
                    var variant = stage == "theme" || stage == "rapid" ? Tint : Shade;
                    if (foreground.Type != Office.MsoColorType.msoColorTypeScheme ||
                         run.TextColor.ObjectThemeColor != WordInterop.WdThemeColorIndex.wdThemeColorAccent6 ||
                         Math.Abs(run.TextColor.TintAndShade - variant) > 0.01 ||
                         (int)foreground.ObjectThemeColor != (int)Office.MsoThemeColorIndex.msoThemeColorAccent6 ||
                         Math.Abs(foreground.Brightness - variant) > 0.01)
                        return $"formula {i}, run {run.TextColor.ObjectThemeColor}/{run.TextColor.TintAndShade:F2}, " +
                            $"SVG fill {foreground.Type}/{foreground.Brightness:F2}, expected Accent 6/{variant:F2}";
                }
            }
            return null;
        }

        private static void CheckPackage(string path, Dictionary<Guid, float> initial,
            string stage, float requestedSize, Guid firstFormula)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = Package.Open(stream, FileMode.Open, FileAccess.Read))
            {
                var main = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
                var xml = LoadXml(main);
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                ns.AddNamespace("asvg", "http://schemas.microsoft.com/office/drawing/2016/SVG/main");
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                var seen = new HashSet<Guid>();
                foreach (XmlElement drawing in xml.SelectNodes("//wp:inline", ns))
                {
                    var alt = drawing.SelectSingleNode("wp:docPr/@descr", ns)?.Value;
                    Require(LaTeXBlockMetadata.TryParse(alt, out var meta) && initial.ContainsKey(meta.Id)
                        && seen.Add(meta.Id), "Missing or duplicated formula identity in the saved document.");
                    var run = drawing.SelectSingleNode("ancestor::w:r[1]/w:rPr", ns);
                    var embed = drawing.SelectSingleNode(".//asvg:svgBlip/@r:embed", ns)?.Value;
                    Require(embed != null, "Saved formula has no SVG relationship.");
                    var target = PackUriHelper.ResolvePartUri(main.Uri, main.GetRelationship(embed).TargetUri);
                    var svg = LoadXml(package.GetPart(target));
                    if (stage == "size" || stage == "grow" || stage == "rapid")
                    {
                        var cy = drawing.SelectSingleNode("wp:extent/@cy", ns)?.Value;
                        var hostHeight = double.Parse(cy, CultureInfo.InvariantCulture) / 12700;
                        var svgHeight = LaTeXBlockService.ReadSvgHeightPt(Encoding.UTF8.GetBytes(svg.OuterXml));
                        var halfPoints = (requestedSize * 2).ToString("0", CultureInfo.InvariantCulture);
                        Require(meta.FontSizePt == requestedSize &&
                            run?.SelectSingleNode("w:sz/@w:val", ns)?.Value == halfPoints &&
                            hostHeight > initial[meta.Id] + 0.5 && Math.Abs(hostHeight - svgHeight) < 0.1,
                            "Saved drawing, embedded SVG, run size or metadata did not reflect the new font size.");
                        if (stage != "rapid") continue;
                    }
                    if (stage == "red")
                    {
                        Require(run?.SelectSingleNode("w:color/@w:val", ns)?.Value == "FF0000" &&
                            HasSavedPaint(svg, "#FF0000"),
                            "The red Word run and embedded red SVG paint did not both persist.");
                    }
                    else if ((stage == "mixed" || stage == "retained") && meta.Id != firstFormula)
                    {
                        var color = run?.SelectSingleNode("w:color", ns);
                        var value = color?.SelectSingleNode("@w:val", ns)?.Value;
                        Require((value == null || value == "auto") &&
                            color?.SelectSingleNode("@w:themeColor", ns) == null &&
                            HasSavedPaint(svg, "#000000") &&
                            !svg.SelectNodes("//*[@class]").Cast<XmlElement>().Any(e =>
                                e.GetAttribute("class").Contains("MsftOfcThm_")),
                            "The automatic drawing retained its previous theme paint in the saved document.");
                    }
                    else
                    {
                        var shade = stage != "theme" && stage != "rapid";
                        Require(run?.SelectSingleNode("w:color/@w:themeColor", ns)?.Value == "accent6" &&
                            run.SelectSingleNode(shade ? "w:color/@w:themeShade" : "w:color/@w:themeTint", ns)?.Value ==
                                (shade ? "99" : "33") &&
                            svg.SelectNodes("//*[@class]").Cast<XmlElement>().Any(e =>
                                e.GetAttribute("class").IndexOf(shade ? "Accent6_lumMod_60_Fill" :
                                    "Accent6_lumMod_20_lumOff_80_Fill",
                                    StringComparison.OrdinalIgnoreCase) >= 0),
                            "The saved Word run and embedded SVG did not both retain the Accent 6 variant.");
                    }
                }
                Require(seen.Count == initial.Count, "One of the formulas is missing from the saved package.");
            }
        }

        private static XmlDocument LoadXml(PackagePart part)
        {
            var xml = new XmlDocument { XmlResolver = null };
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read)) xml.Load(stream);
            return xml;
        }

        private static bool HasSavedPaint(XmlDocument svg, string expected)
        {
            var styles = string.Join("\n", svg.SelectNodes("//*[local-name()='style']")
                .Cast<XmlElement>().Select(e => e.InnerText));
            foreach (XmlElement element in svg.SelectNodes("//*[@class or @fill]"))
            {
                var paint = element.GetAttribute("fill");
                foreach (var name in element.GetAttribute("class").Split(
                    new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    // Word can retain an earlier fill attribute underneath a new
                    // Office responsive-colour class. Its CSS rule takes precedence.
                    var rule = Regex.Match(styles, @"\." + Regex.Escape(name) +
                        @"\s*\{[^{}]*?\bfill\s*:\s*(#[0-9a-f]{6})\b", RegexOptions.IgnoreCase);
                    if (rule.Success) paint = rule.Groups[1].Value;
                }
                if (string.Equals(paint, expected, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static void Require(bool valid, string message)
        {
            if (!valid) throw new InvalidOperationException(message);
        }
    }
}
