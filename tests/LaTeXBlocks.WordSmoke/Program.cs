using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using LaTeXBlocks.Word;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.WordSmoke
{
    internal static class Program
    {
        private const string StartupShutdownProbeChild = "LATEXBLOCKS_STARTUP_SHUTDOWN_PROBE_CHILD";
        private const string WordJoiner = "\u2060";

        [STAThread]
        private static int Main()
        {
            const string source = "$C_{ij}$";
            const string updatedSource = "$E=mc^2$";
            WordInterop.Application word = null;
            WordInterop.Document document = null;
            StemTeXBackend renderer = null;
            var ownsWord = false;
            var artifactDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "artifacts");
            var documentPath = Path.Combine(artifactDirectory, "LaTeXBlocks-Smoke.docx");
            var numberedDocumentPath = Path.Combine(artifactDirectory, "LaTeXBlocks-Numbered-Smoke.docx");
            var spacingDocumentPath = Path.Combine(artifactDirectory, "LaTeXBlocks-Inline-Spacing-Smoke.docx");
            var textColorDocumentPath = Path.Combine(artifactDirectory, "LaTeXBlocks-Text-Color-Smoke.docx");
            var floatingBlockDocumentPath = Path.Combine(artifactDirectory, "LaTeXBlocks-Floating-Block-Smoke.docx");
            try
            {
                Directory.CreateDirectory(artifactDirectory);
                if (string.Equals(Environment.GetEnvironmentVariable(StartupShutdownProbeChild), "1",
                    StringComparison.Ordinal))
                {
                    Console.WriteLine("StemTeX: testing immediate shutdown during renderer initialization...");
                    RunStartupShutdownProbe();
                    return 0;
                }
                Console.WriteLine("StemTeX: testing immediate shutdown during renderer initialization...");
                RunStartupShutdownProbeInIsolatedHost();
                renderer = new StemTeXBackend();
                var profile = renderer.DefaultAvailableProfile;
                var alternateProfile = profile;
                foreach (var candidate in renderer.Profiles)
                    if (!string.Equals(candidate, profile, StringComparison.OrdinalIgnoreCase)) { alternateProfile = candidate; break; }
                var cjkProfile = profile;
                if (profile.IndexOf("cjk", StringComparison.OrdinalIgnoreCase) < 0)
                    foreach (var candidate in renderer.Profiles)
                        if (candidate.IndexOf("cjk", StringComparison.OrdinalIgnoreCase) >= 0) { cjkProfile = candidate; break; }
                RunRenderHostClientSmoke(profile);
                Console.WriteLine("StemTeX: warming the default profile...");
                renderer.WarmUp(profile);
                Assert(WordSelectionLaTeXExporter.EscapeText("10% & x_1 # {a} ~ ^ \\") ==
                    "10\\% \\& x\\_1 \\# \\{a\\} \\textasciitilde{} \\textasciicircum{} \\textbackslash{}",
                    "Word text was not escaped safely for LaTeX export.");
                if (string.Equals(Environment.GetEnvironmentVariable("LATEXBLOCKS_SMOKE_SHUTDOWN_ONLY"), "1",
                    StringComparison.Ordinal))
                {
                    RunShutdownProbe(renderer, profile);
                    Console.WriteLine("StemTeX shutdown smoke test passed.");
                    return 0;
                }
                Console.WriteLine("StemTeX: testing latest-only preview scheduling...");
                var staleOne = renderer.RenderLatestAsync(profile, "$a_1$", 360, true);
                var staleTwo = renderer.RenderLatestAsync(profile, "$a_2$", 360, true);
                var latest = renderer.RenderLatestAsync(profile, "$a_3$", 360, true);
                var latestResult = latest.GetAwaiter().GetResult();
                Assert(latestResult.Bytes.Length > 0, "The latest queued render did not complete.");
                Assert(LaTeXBlockEditorForm.BuildPreviewHtml(latestResult.Bytes).IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0,
                    "The editor preview document does not contain the latest SVG.");
                Assert(staleOne.IsCanceled && staleTwo.IsCanceled,
                    "The single-worker scheduler did not discard superseded queued renders.");
                Console.WriteLine("StemTeX: testing cancellation of an active latest-only preview...");
                var activePreview = renderer.RenderLatestAsync(profile, "\\loop\\iftrue\\repeat", 360, true, 11);
                WaitFor(() => renderer.NativeRenderInProgressForTest, 5000,
                    "The active-preview cancellation probe did not enter native rendering.");
                var cancellationTimer = Stopwatch.StartNew();
                var replacementPreview = renderer.RenderLatestAsync(profile, "$a_4$", 360, true, 11);
                WaitFor(() => activePreview.IsCompleted, 2000,
                    "A superseded active preview did not cancel promptly. Native cancel attempts=" +
                    renderer.NativeCancelAttemptsForTest + ".");
                try
                {
                    activePreview.GetAwaiter().GetResult();
                    throw new InvalidOperationException("The active superseded preview completed instead of canceling.");
                }
                catch (System.Threading.Tasks.TaskCanceledException) { }
                Assert(cancellationTimer.ElapsedMilliseconds < 2000,
                    "Active latest-only preview cancellation took " + cancellationTimer.ElapsedMilliseconds + " ms.");
                Assert(renderer.NativeCancelAttemptsForTest > 0,
                    "The superseded active preview was not sent to StemTeX cancellation.");
                var cancellationMilliseconds = cancellationTimer.ElapsedMilliseconds;
                // The native renderer may need to rebuild a primary worker after killing
                // the canceled one. Its recovery latency is distinct from cancellation;
                // the important UI guarantee above is that the obsolete task finishes
                // promptly and the editor can accept a newer request immediately.
                WaitFor(() => replacementPreview.IsCompleted, 30000,
                    "The preview submitted after native cancellation did not recover.");
                Assert(replacementPreview.GetAwaiter().GetResult().Bytes.Length > 0,
                    "The preview submitted after native cancellation did not complete.");
                Console.WriteLine("StemTeX: active preview canceled in " + cancellationMilliseconds +
                    " ms; replacement recovered in " + cancellationTimer.ElapsedMilliseconds + " ms.");
                Console.WriteLine("StemTeX: testing fixed-width and inline auto-width rendering...");
                var svg = renderer.RenderSvg(profile, source, 360, false);
                var prefix = Encoding.UTF8.GetString(svg.Bytes, 0, Math.Min(svg.Bytes.Length, 512));
                Assert(prefix.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0, "StemTeX did not return SVG bytes.");
                Assert(Encoding.UTF8.GetString(svg.Bytes).IndexOf("latexblocks-baseline", StringComparison.Ordinal) < 0,
                    "The temporary baseline marker leaked into the embedded SVG.");
                Assert(svg.DepthPt > 0, "StemTeX inline baseline marker did not produce a positive TeX depth.");
                var autoSvg = renderer.RenderSvg(profile, source, 360, true);
                var autoSvg11 = renderer.RenderSvg(profile, source, 360, true, 11);
                var terminalNewlineAutoSvg = renderer.RenderSvg(profile, source + "\n", 360, true);
                var coloredAutoSvg = renderer.RenderSvg(profile,
                    LaTeXBlockService.ApplyTextColor(source, 0x00ff0000, true), 360, true);
                var coloredTerminalNewlineSvg = renderer.RenderSvg(profile,
                    LaTeXBlockService.ApplyTextColor(source + "\n", 0x00ff0000, true), 360, true);
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(autoSvg.Bytes) -
                                LaTeXBlockService.ReadSvgWidthPt(coloredAutoSvg.Bytes)) < 0.01,
                    "The TeX color wrapper changed the auto-width formula's logical box.");
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(autoSvg.Bytes) -
                                LaTeXBlockService.ReadSvgWidthPt(terminalNewlineAutoSvg.Bytes)) < 0.01,
                    "A terminal inline-source newline changed the automatic TeX box.");
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(autoSvg.Bytes) -
                                LaTeXBlockService.ReadSvgWidthPt(coloredTerminalNewlineSvg.Bytes)) < 0.01,
                    "The TeX color wrapper turned a terminal inline-source newline into horizontal space.");
                Console.WriteLine("StemTeX: color wrapper preserves auto-width geometry.");
                // A deliberately edge-trimmed reference must have the same width. If
                // the measurement wrapper contributes either of its own line-break
                // spaces, the ordinary render is wider by a font interword space.
                var edgeTrimReference = renderer.RenderSvg(profile, "\\unskip" + source + "%", 360, true);
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(autoSvg.Bytes) -
                                LaTeXBlockService.ReadSvgWidthPt(edgeTrimReference.Bytes)) < 0.01,
                    "The auto-width TeX wrapper added horizontal interword glue around the formula.");
                var explicitTrailingSpace = renderer.RenderSvg(profile,
                    source + "\\hspace{2pt}% trailing source comment", 360, true);
                var autoWidthPt = LaTeXBlockService.ReadSvgWidthPt(autoSvg.Bytes);
                var explicitWidthPt = LaTeXBlockService.ReadSvgWidthPt(explicitTrailingSpace.Bytes);
                Assert(Math.Abs(explicitWidthPt - autoWidthPt - 2) < 0.01,
                    "Suppressing wrapper glue removed an explicit trailing TeX space " +
                    "(base=" + autoWidthPt + "pt, explicit=" + explicitWidthPt + "pt).");
                var logicalBox = renderer.RenderSvg(profile,
                    "\\hbox to 10pt{\\vrule width.1pt height1pt depth0pt\\hfil}", 360, true);
                var logicalBoxWidthPt = LaTeXBlockService.ReadSvgWidthPt(logicalBox.Bytes);
                const double texPointToWordPoint = 72.0 / 72.27;
                Assert(Math.Abs(logicalBoxWidthPt - (10 * texPointToWordPoint)) < 0.01,
                    "Auto-width SVG added horizontal padding around the logical TeX box " +
                    "(width=" + logicalBoxWidthPt + "pt).");
                var overhangingInk = renderer.RenderSvg(profile,
                    "\\hbox to 10pt{\\kern-2pt\\vrule width1pt height1pt depth0pt\\hfil}", 360, true);
                var overhangingWidthPt = LaTeXBlockService.ReadSvgWidthPt(overhangingInk.Bytes);
                Assert(Math.Abs(overhangingWidthPt - (12 * texPointToWordPoint)) < 0.01,
                    "Removing horizontal SVG padding clipped a real TeX ink overhang " +
                    "(width=" + overhangingWidthPt + "pt).");
                var rhoGhSvg = renderer.RenderSvg(profile, "$\\rho gh$", 360, true, 11);
                Assert(rhoGhSvg.Bytes.Length > 0 && rhoGhSvg.DepthPt > 0,
                    "A valid rho-gh inline formula did not render.");
                var autoText = Encoding.UTF8.GetString(autoSvg.Bytes);
                Assert(autoText.IndexOf("latexblocks-start", StringComparison.Ordinal) < 0 &&
                    autoText.IndexOf("latexblocks-end", StringComparison.Ordinal) < 0 &&
                    autoText.IndexOf("latexblocks-ink", StringComparison.Ordinal) < 0,
                    "Auto-width measurement markers leaked into the embedded SVG.");
                Assert(Convert.ToBase64String(autoSvg11.Bytes) != Convert.ToBase64String(autoSvg.Bytes),
                    "Changing the requested TeX font size did not produce a new SVG render.");
                Assert(Regex.IsMatch(autoSvg11.SummaryJson ?? string.Empty,
                    "\\\"fontSizePt\\\"\\s*:\\s*11(?:\\.0+)?(?:[,}])"),
                    "StemTeX did not report the native per-request 11 pt font size.");
                var mixedCjkSvg = renderer.RenderSvg(cjkProfile, "Einstein's $E=mc^2$ 这样的", 360, true);
                Console.WriteLine("StemTeX: CJK-profile render completed.");
                var mixedCjkText = Encoding.UTF8.GetString(mixedCjkSvg.Bytes);
                Assert(mixedCjkSvg.DepthPt > 0, "Mixed CJK/Western inline TeX lost its TeX baseline depth.");
                Assert(mixedCjkText.IndexOf("latexblocks-start", StringComparison.Ordinal) < 0 &&
                    mixedCjkText.IndexOf("latexblocks-end", StringComparison.Ordinal) < 0,
                    "Mixed CJK/Western rendering leaked its baseline measurement markers.");
                Console.WriteLine("StemTeX: confirming repeated warm-up is idempotent...");
                renderer.WarmUp(profile);
                Assert(LaTeXBlockService.IsSupportedInlineShapeType(WordInterop.WdInlineShapeType.wdInlineShapePicture),
                    "Picture objects are not eligible for the LaTeX Block contract.");
                Assert(LaTeXBlockService.IsSupportedInlineShapeType((WordInterop.WdInlineShapeType)17),
                    "Word SVG objects are not eligible for the LaTeX Block contract.");
                Assert(!LaTeXBlockService.IsSupportedInlineShapeType(
                        WordInterop.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject),
                    "Embedded OLE objects such as MathType must never be probed as LaTeX Blocks.");
                var legacyTitle = LaTeXBlockMetadata.Prefix + "id=" + Guid.NewGuid().ToString("D") +
                    ";width=360;depth=0;mode=fixed;size=10";
                Assert(LaTeXBlockMetadata.TryParse(legacyTitle, out var legacyMetadata) &&
                    legacyMetadata.Role == LaTeXBlockRole.Content,
                    "Metadata written before the role field no longer defaults to ordinary content.");
                var framedMetadata = new LaTeXBlockMetadata(Guid.NewGuid(), 180, 2.5,
                    LaTeXBlockLayoutMode.Fixed, 14, LaTeXBlockRole.Content,
                    182.75, 48.25);
                Assert(LaTeXBlockMetadata.TryParse(framedMetadata.ToString(), out var reparsedFrameMetadata) &&
                       Math.Abs(reparsedFrameMetadata.FrameWidthPt - 182.75) < 0.001 &&
                       Math.Abs(reparsedFrameMetadata.FrameHeightPt - 48.25) < 0.001,
                    "Floating frame geometry did not round-trip through Word metadata.");
                var styledBlockStyle = new LaTeXBlockStyle(1.45, 8.5,
                    LaTeXBlockVerticalAlignment.Middle,
                    System.Drawing.Color.FromArgb(0x12, 0x34, 0x56), true,
                    System.Drawing.Color.FromArgb(0xf0, 0xee, 0xd0), 1.25,
                    System.Drawing.Color.FromArgb(0x65, 0x43, 0x21));
                var styledMetadata = LaTeXBlockMetadata.Create(180, 2.5,
                    LaTeXBlockLayoutMode.Fixed, 14, LaTeXBlockRole.Content,
                    styledBlockStyle).WithFrameSize(182.75, 48.25);
                Assert(LaTeXBlockMetadata.TryParse(styledMetadata.ToString(),
                        out var reparsedStyledMetadata) &&
                       reparsedStyledMetadata.HasExplicitStyle &&
                       reparsedStyledMetadata.Style.Equals(styledBlockStyle),
                    "A fixed Block style did not round-trip through Word Title metadata.");
                var autoMetadataWithStyle = LaTeXBlockMetadata.Create(180, 2.5,
                    LaTeXBlockLayoutMode.Auto, 14, LaTeXBlockRole.Content,
                    styledBlockStyle);
                Assert(!autoMetadataWithStyle.HasExplicitStyle,
                    "An Auto formula retained Fixed-Block style metadata after a layout-mode change.");
                var explicitDefaultStyle = LaTeXBlockStyle.Default;
                var explicitDefaultSource = explicitDefaultStyle.WrapSource("\\[E=mc^2\\]",
                    14, true);
                Assert(explicitDefaultSource.IndexOf("\\renewcommand{\\baselinestretch}{1.2}",
                           StringComparison.Ordinal) >= 0 &&
                       explicitDefaultSource.IndexOf("\\noindent", StringComparison.Ordinal) < 0 &&
                       explicitDefaultSource.IndexOf("\\strut", StringComparison.Ordinal) < 0 &&
                       explicitDefaultSource.IndexOf("\\color{latexblocksforeground}",
                           StringComparison.Ordinal) < 0,
                    "A standalone Word display Block gained paragraph material outside its TeX display.");
                var styledSource = styledBlockStyle.WrapSource("\\[E=mc^2\\]", 14);
                Assert(styledSource.IndexOf("\\colorbox", StringComparison.Ordinal) < 0 &&
                       styledSource.IndexOf("\\fbox", StringComparison.Ordinal) < 0 &&
                       styledSource.IndexOf("\\color{latexblocksforeground}",
                           StringComparison.Ordinal) < 0,
                    "The Word display Block style moved paragraph decoration into TeX.");
                var styledSvgBytes = LaTeXBlockSvgFrame.Decorate(svg.Bytes, styledBlockStyle,
                    213.25, 89.75);
                var styledSvgText = Encoding.UTF8.GetString(styledSvgBytes);
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(styledSvgBytes) - 213.25) < 0.001 &&
                       Math.Abs(LaTeXBlockService.ReadSvgHeightPt(styledSvgBytes) - 89.75) < 0.001 &&
                       styledSvgText.IndexOf("data-latexblocks-frame='1'", StringComparison.Ordinal) >= 0 &&
                       styledSvgText.IndexOf("data-latexblocks-border='1'", StringComparison.Ordinal) >= 0 &&
                       styledSvgText.IndexOf("#F0EED0", StringComparison.Ordinal) >= 0 &&
                       styledSvgText.IndexOf("#654321", StringComparison.Ordinal) >= 0 &&
                       styledSvgText.IndexOf("fill='#123456'", StringComparison.OrdinalIgnoreCase) >= 0,
                    "The shared styled SVG frame did not preserve Word's exact frame or paint all shell layers.");
                var framedSvgBytes = LaTeXBlockService.FrameSvg(svg.Bytes, 213.25, 89.75);
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(framedSvgBytes) - 213.25) < 0.001 &&
                       Math.Abs(LaTeXBlockService.ReadSvgHeightPt(framedSvgBytes) - 89.75) < 0.001 &&
                       Encoding.UTF8.GetString(framedSvgBytes).IndexOf("overflow='hidden'",
                           StringComparison.Ordinal) >= 0,
                    "A reframed SVG did not retain the requested exact clipping viewport.");
                var oversizedFrameSvgBytes = LaTeXBlockService.FrameSvg(svg.Bytes, 2500.25, 61.5);
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(oversizedFrameSvgBytes) -
                           2500.25) < 0.001 &&
                       Math.Abs(LaTeXBlockService.ReadSvgHeightPt(oversizedFrameSvgBytes) -
                           61.5) < 0.001 &&
                       Math.Abs(LaTeXBlockService.ClampFloatingFrameExtent(2500.25) -
                           2500.25) < 0.001,
                    "A Word-owned frame wider than the editor policy was silently clipped.");
                Assert(!LaTeXBlockService.HasNativeFrameGeometryChanged(213.25, 89.75,
                           213.25, 89.75) &&
                       LaTeXBlockService.HasNativeFrameGeometryChanged(213.25, 89.75,
                           213.31, 89.75),
                    "Move/rotation geometry was not distinguished from a native frame resize.");
                var firstRapidResizeWidth = LaTeXBlockService.ComposeNativeFrameLayoutWidth(180,
                    182.75, 213.25);
                var secondRapidResizeWidth = LaTeXBlockService.ComposeNativeFrameLayoutWidth(
                    firstRapidResizeWidth, 213.25, 244.5);
                Assert(Math.Abs(secondRapidResizeWidth - (180 + 244.5 - 182.75)) < 0.001,
                    "Rapid native frame resizes did not compose from the pending TeX width.");
                Assert(Math.Abs(LaTeXBlockWidthPolicy.ResolveDefaultFixedWidth() - 360) < 0.001 &&
                    Math.Abs(LaTeXBlockWidthPolicy.WidthStepPt - 0.5) < 0.001 &&
                    LaTeXBlockWidthPolicy.IsValidWidth(30) &&
                    LaTeXBlockWidthPolicy.IsValidWidth(2000) &&
                    !LaTeXBlockWidthPolicy.IsValidWidth(29.9) &&
                    !LaTeXBlockWidthPolicy.IsValidWidth(2000.1),
                    "The fixed-block width policy does not support native floating frames.");
                Assert(LaTeXBlockWidthPolicy.TryParseWidth("360.5", out var parsedWidthPt) &&
                    Math.Abs(parsedWidthPt - 360.5) < 0.001 &&
                    !LaTeXBlockWidthPolicy.TryParseWidth("Natural", out _),
                    "The Ribbon width field does not parse precise point values safely.");
                var ribbonXml = LaTeXBlocksRibbon.BuildCustomUi();
                Assert(ribbonXml.IndexOf(LaTeXBlocksRibbon.WidthControlId,
                           StringComparison.Ordinal) >= 0 &&
                       ribbonXml.IndexOf("getEnabled=\"GetWidthEnabled\"",
                           StringComparison.Ordinal) >= 0,
                    "The Word Ribbon does not expose the selection-aware width control.");
                Assert(ribbonXml.IndexOf("id=\"LaTeXBlocks.Edit\"", StringComparison.Ordinal) >= 0 &&
                       ribbonXml.IndexOf("imageMso=\"" + LaTeXBlocksRibbon.EditBlockImageMso + "\"",
                           StringComparison.Ordinal) >= 0,
                    "The Word Edit Block command does not expose its edit icon.");
                Assert(ribbonXml.IndexOf("id=\"" + LaTeXBlocksRibbon.ReflowFrameControlId + "\"",
                           StringComparison.Ordinal) >= 0 &&
                       ribbonXml.IndexOf("onAction=\"OnReflowFrame\"", StringComparison.Ordinal) >= 0,
                    "The Word Ribbon does not expose the floating-frame reflow command.");
                Assert(ribbonXml.IndexOf("id=\"LaTeXBlocks.InsertEquationReference\"",
                           StringComparison.Ordinal) >= 0 &&
                       ribbonXml.IndexOf("onAction=\"OnInsertEquationReference\"",
                           StringComparison.Ordinal) >= 0,
                    "The Word Ribbon does not expose the equation-reference command.");
                Assert(!LaTeXBlockService.ShouldRefreshForHostFontSizeChange(11, 11, 10),
                    "Selecting and leaving an unchanged formula would spuriously rerender it at the host character size.");
                Assert(LaTeXBlockService.ShouldRefreshForHostFontSizeChange(11, 12, 11),
                    "An actual host font-size change is no longer detected when the selection is left.");
                Assert(!LaTeXBlockService.ShouldRefreshForHostFontSizeChange(11, 12, 12),
                    "A formula already rendered at the new host font size would be rendered a second time.");
                Assert(LaTeXBlockService.PrepareDisplayMathSource("\\[E=mc^2\\]") ==
                    "\\(\n\\displaystyle\nE=mc^2\n\\)",
                    "The numbered-equation render wrapper did not preserve the formula body.");
                const string commentedNumberedSource =
                    "\\begin {align}\r\nE&=mc^2 % exact source comment\r\n\\end {align}";
                var canonicalCommentedNumberedSource =
                    LaTeXBlockService.NormalizeSourceText(commentedNumberedSource);
                var preparedCommentedDisplay = LaTeXBlockService.PrepareDisplayMathSource(commentedNumberedSource);
                Assert(preparedCommentedDisplay.IndexOf("\\begin{aligned}\n", StringComparison.Ordinal) >= 0 &&
                    Regex.IsMatch(preparedCommentedDisplay,
                        "% exact source comment\\r?\\n\\\\end\\{aligned\\}", RegexOptions.CultureInvariant),
                    "The display wrapper did not preserve a TeX comment boundary while reducing align.");
                Assert(StemTeXRenderer.RemoveTeXCommentsForDetection(
                        "$a % commented \\[ is not a display\r\n+b$")
                        .IndexOf("\\[", StringComparison.Ordinal) < 0,
                    "Auto-width display detection still reads TeX comments as source.");
                var texTagRejected = false;
                try { LaTeXBlockService.PrepareDisplayMathSource("\\[E=mc^2 \\tag{A}\\]"); }
                catch (ArgumentException) { texTagRejected = true; }
                Assert(texTagRejected, "A TeX-side tag was allowed to compete with Word-owned numbering.");
                Console.WriteLine("StemTeX: testing natural-width display-style rendering...");
                var displaySvg = renderer.RenderSvg(profile,
                    LaTeXBlockService.PrepareDisplayMathSource("\\[\\sum_{i=1}^n \\frac{1}{i}\\]"),
                    360, true, 10);
                Assert(LaTeXBlockService.ReadSvgWidthPt(displaySvg.Bytes) < 100,
                    "Display-style math retained StemTeX's fixed-width page canvas.");
                var commentedDisplaySvg = renderer.RenderSvg(profile, preparedCommentedDisplay, 360, true, 10);
                Assert(commentedDisplaySvg.Bytes.Length > 0,
                    "A legal trailing TeX comment swallowed the generated display or measurement boundary.");

                Console.WriteLine("Word: starting an isolated application instance...");
                var existingWordProcesses = new HashSet<int>();
                foreach (var process in Process.GetProcessesByName("WINWORD"))
                {
                    try { existingWordProcesses.Add(process.Id); }
                    finally { process.Dispose(); }
                }
                word = new WordInterop.Application();
                foreach (var process in Process.GetProcessesByName("WINWORD"))
                {
                    try
                    {
                        if (!existingWordProcesses.Contains(process.Id)) ownsWord = true;
                    }
                    finally { process.Dispose(); }
                }
                if (!ownsWord)
                    throw new InvalidOperationException("The smoke test attached to an existing visible Word instance; no document was changed.");
                word.Visible = false;
                word.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone;
                document = word.Documents.Add();
                document.Range(0, 0).Text = "Einstein's theory ";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 11;
                document.Range(0, 1).Font.Size = 13;
                document.Range(1, 2).Font.Size = 17;
                document.Range(1, 1).Select();
                Assert(Math.Abs((double)word.Selection.Range.Font.Size - 17) < 0.001 &&
                    Math.Abs((double)word.Selection.Font.Size - 13) < 0.001 &&
                    Math.Abs(LaTeXBlockService.ResolveFontSize(word.Selection,
                        LaTeXBlockLayoutMode.Auto, 10) - 13) < 0.001,
                    "A run-boundary insertion used the right-hand character instead of Word's typing size.");
                word.Selection.Font.Size = 36;
                Assert(Math.Abs((double)word.Selection.Range.Font.Size - 17) < 0.001 &&
                    Math.Abs((double)word.Selection.Font.Size - 36) < 0.001 &&
                    Math.Abs(LaTeXBlockService.ResolveFontSize(word.Selection,
                        LaTeXBlockLayoutMode.Auto, 10) - 36) < 0.001,
                    "An explicit caret-only Word font size was not used for inline rendering.");
                document.Range(0, 2).Select();
                Assert(Math.Abs(LaTeXBlockService.ResolveFontSize(word.Selection,
                    LaTeXBlockLayoutMode.Auto, 10) - 13) < 0.001,
                    "A mixed-size replacement selection silently fell back to 10 pt.");
                document.Content.Font.Size = 11;
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var service = new LaTeXBlockService(word, renderer);
                var page = document.Sections[1].PageSetup;
                var expectedTextAreaWidth = (double)page.PageWidth - page.LeftMargin - page.RightMargin;
                var textAreaWidth = service.ResolveTextAreaWidth(word.Selection.Range);
                Assert(Math.Abs(textAreaWidth - expectedTextAreaWidth) < 0.01,
                    "The fixed-block editor did not resolve the current Word text area width.");
                var textColumns = page.TextColumns;
                textColumns.SetCount(2);
                // Word quantizes each stored TextColumn.Width independently.  The
                // page-width formula can therefore differ from the authoritative
                // column object by a few hundredths of a point (197.025 vs 197.000
                // in the default two-column fixture).
                var expectedColumnWidth = (double)textColumns[1].Width;
                Assert(Math.Abs(service.ResolveTextAreaWidth(word.Selection.Range) - expectedColumnWidth) < 0.01,
                    "The fixed-block editor did not resolve the current Word column width.");
                textColumns.SetCount(1);
                Console.WriteLine("Word: testing natural and exact point-width editors...");
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var editorFontSize = LaTeXBlockService.ResolveFontSize(word.Selection,
                    LaTeXBlockLayoutMode.Auto, 10);
                using (var editor = new LaTeXBlockEditorForm(service, "$x_1$", 360,
                    LaTeXBlockLayoutMode.Auto, profile, selected => { }, false,
                    editorFontSize))
                {
                    editor.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    editor.Location = new System.Drawing.Point(100, 100);
                    editor.Show();
                    Assert(editor.WidthIsNatural && Math.Abs(editor.WidthPt - 360) < 0.001,
                        "The auto-width editor exposed a fixed text-area width.");
                    WaitFor(() => editor.PreviewIsCurrent, 10000, "The editor did not produce its initial live preview.");
                    Assert(Math.Abs(editor.CurrentRender.FontSizePt - 11) < 0.001,
                        "The insert editor did not preview at Word's current insertion font size.");
                    var firstPreview = Convert.ToBase64String(editor.CurrentRender.SvgBytes);
                    editor.SetSourceForTest("$x_2+y_2$");
                    WaitFor(() => editor.PreviewIsCurrent &&
                        Convert.ToBase64String(editor.CurrentRender.SvgBytes) != firstPreview,
                        10000, "Changing editor text did not replace the live SVG preview.");
                    editor.SetSourceForTest("$\\rho gh$");
                    WaitFor(() => editor.PreviewIsCurrent &&
                        Encoding.UTF8.GetString(editor.CurrentRender.SvgBytes).IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0,
                        10000, "The live editor failed to preview a valid rho-gh formula.");
                    for (var frame = 0; frame < 10; frame++) { System.Windows.Forms.Application.DoEvents(); Thread.Sleep(50); }
                    using (var screenshot = new System.Drawing.Bitmap(editor.Width, editor.Height))
                    using (var graphics = System.Drawing.Graphics.FromImage(screenshot))
                    {
                        graphics.CopyFromScreen(editor.Location, System.Drawing.Point.Empty, editor.Size);
                        screenshot.Save(Path.Combine(artifactDirectory, "editor-live-preview.png"));
                    }
                    editor.Close();
                }
                const double legacyFixedWidthPt = 181.234;
                using (var editor = new LaTeXBlockEditorForm(service, "\\[x_1+x_2\\]",
                    legacyFixedWidthPt, LaTeXBlockLayoutMode.Fixed, profile,
                    selected => { }, true, 10))
                {
                    Assert(!editor.WidthIsNatural &&
                        Math.Abs(editor.WidthPt - legacyFixedWidthPt) < 0.0001,
                        "Opening a legacy fixed-width block changed its absolute metadata width.");
                    editor.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    editor.Location = new System.Drawing.Point(100, 100);
                    editor.Show();
                    WaitFor(() => editor.PreviewIsCurrent, 10000,
                        "The fixed-width editor did not produce its initial live preview.");
                    const double requestedFixedWidthPt = 234.5;
                    editor.SetWidthPtForTest(requestedFixedWidthPt);
                    WaitFor(() => editor.PreviewIsCurrent &&
                        Math.Abs(editor.WidthPt - requestedFixedWidthPt) < 0.01,
                        10000, "Changing the exact point width did not update the fixed block.");
                    var renderedPointWidth =
                        LaTeXBlockService.ReadSvgWidthPt(editor.CurrentRender.SvgBytes);
                    // Opening the style editor upgrades a legacy Fixed Block into
                    // an explicitly styled viewport. Its requested width is now
                    // the exact Word/SVG outer frame; TeX receives only the inner
                    // measure (after any SVG shell inset) and is never scaled.
                    var expectedPointSvgWidth = requestedFixedWidthPt;
                    Console.WriteLine("Fixed width editor: requested=" +
                        requestedFixedWidthPt.ToString("0.###") + "pt, SVG=" +
                        renderedPointWidth.ToString("0.###") + "pt");
                    Assert(Math.Abs(renderedPointWidth - expectedPointSvgWidth) < 0.02,
                        "The styled Block editor did not preserve the exact Word outer-frame width.");
                    editor.Close();
                }
                Console.WriteLine("Word: exact point-width editor passed.");
                // Run the independent Font.Color path before the longer inline-spacing
                // matrix, so a host-specific spacing tolerance cannot mask a color
                // geometry regression.
                RunTextColorSmoke(word, service, profile, textColorDocumentPath);
                RunInlineSpacingSmoke(word, service, profile, spacingDocumentPath);
                RunFloatingBlockSmoke(word, service, profile, floatingBlockDocumentPath);
                const string fractionSource = "$\\frac{1}{2}E=mc^2$";
                var fractionStart = document.Content.End - 1;
                document.Range(fractionStart, fractionStart).Text = "What?";
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                word.Selection.Font.Name = "Times New Roman";
                word.Selection.Font.Size = 36;
                word.Selection.Font.Position = 0;
                word.Selection.NoProofing = 0;
                var fractionRender = service.RenderPreview(fractionSource, 360,
                    LaTeXBlockLayoutMode.Auto, profile, 36);
                var fractionShape = service.InsertRendered(fractionSource, 360,
                    LaTeXBlockLayoutMode.Auto, fractionRender);
                var fractionDepth = (int)Math.Round(fractionRender.DepthPt,
                    MidpointRounding.AwayFromZero);
                Assert(fractionDepth > 1 && fractionShape.Range.Font.Position == -fractionDepth,
                    "U+2060 boundary insertion discarded the large inline formula's baseline position.");
                Assert(Math.Abs((double)fractionShape.Range.Font.Size - 36) < 0.001,
                    "U+2060 boundary insertion discarded the large inline formula's host font size.");
                Assert(document.Range(fractionStart, fractionShape.Range.Start - 1).Text == "What?" &&
                       document.Range(fractionShape.Range.Start - 1, fractionShape.Range.Start).Text == WordJoiner,
                    "Restoring the drawing run format duplicated adjacent running text.");
                AssertInlineWordJoinerBoundary(fractionShape, 2, "Large inline formula");
                var fractionFollowingStart = word.Selection.Start;
                word.Selection.TypeText("abc");
                var fractionFollowing = document.Range(fractionFollowingStart, fractionFollowingStart + 3);
                Assert(fractionFollowing.Text == "abc" && fractionFollowing.Font.Position == 0 &&
                    fractionFollowing.NoProofing == 0 &&
                    document.Range(fractionShape.Range.End, fractionShape.Range.End + 1).Text == WordJoiner,
                    "Text after a large inline formula inherited its compensated picture position.");
                document.Range(fractionStart, document.Content.End - 1).Delete();
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                word.Selection.Font.Name = "Times New Roman";
                word.Selection.Font.Size = 11;
                word.Selection.Font.Position = 0;
                word.Selection.NoProofing = 0;

                var reusableInlineRender = new LaTeXBlockRender(null, autoSvg11.Bytes, autoSvg11.DepthPt, 11);
                var existingTextStart = document.Content.End - 1;
                document.Range(existingTextStart, existingTextStart).Text = "and we.";
                document.Range(existingTextStart, existingTextStart).Select();
                var beforeExistingText = service.InsertRendered(source, 360, LaTeXBlockLayoutMode.Auto,
                    reusableInlineRender);
                var insertedRunningStart = word.Selection.Start;
                word.Selection.TypeText(" think ");
                var insertedRunning = document.Range(insertedRunningStart, insertedRunningStart + 7);
                var untouchedFollowing = document.Range(insertedRunningStart + 7, insertedRunningStart + 14);
                AssertInlineWordJoinerBoundary(beforeExistingText, 2,
                    "Formula before existing running text");
                Assert(document.InlineShapes.Count == 1 && beforeExistingText.Range.End + 1 == insertedRunningStart &&
                    insertedRunning.Text == " think " && insertedRunning.Font.Position == 0 &&
                    insertedRunning.NoProofing == 0 && untouchedFollowing.Text == "and we." &&
                    untouchedFollowing.Font.Position == 0 && untouchedFollowing.NoProofing == 0,
                    "Inserting before existing running text leaked the picture baseline or changed the following run.");
                document.Range(existingTextStart, document.Content.End - 1).Delete();
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                word.Selection.Font.Position = 0;
                word.Selection.NoProofing = 0;

                var raisedStart = document.Content.End - 1;
                document.Range(raisedStart, raisedStart).Select();
                word.Selection.Font.Position = 2;
                word.Selection.NoProofing = -1;
                var raisedShape = service.InsertRendered(source, 360, LaTeXBlockLayoutMode.Auto,
                    reusableInlineRender);
                var raisedDepth = (int)Math.Round(autoSvg11.DepthPt, MidpointRounding.AwayFromZero);
                var raisedTextStart = word.Selection.Start;
                word.Selection.TypeText("x");
                var raisedText = document.Range(raisedTextStart, raisedTextStart + 1);
                Assert(raisedShape.Range.Font.Position == 2 - raisedDepth && raisedText.Font.Position == 0 &&
                    raisedText.NoProofing == 0 &&
                    document.Range(raisedShape.Range.End, raisedShape.Range.End + 1).Text == WordJoiner,
                    "Inline baseline compensation was not relative to the deliberate host position, or its " +
                    "picture-only formatting leaked into the paragraph insertion format.");
                document.Range(raisedStart, document.Content.End - 1).Delete();
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                word.Selection.Font.Position = 0;
                word.Selection.NoProofing = 0;

                var inserted = service.InsertBlock(source, 360, LaTeXBlockLayoutMode.Auto, profile);
                Assert(inserted.AlternativeText == source, "Alternative Text is not the exact TeX source.");
                Assert(LaTeXBlockMetadata.TryParse(inserted.Title, out var firstMetadata), "Title metadata is invalid.");
                Assert(firstMetadata.DepthPt > 0, "TeX depth was not stored in block metadata.");
                Assert(Math.Abs(firstMetadata.DepthPt - autoSvg11.DepthPt) < 0.01 &&
                    Math.Abs(firstMetadata.FontSizePt - 11) < 0.001,
                    "The inline formula was not rerendered at Word's 11 pt insertion font.");
                var compensatedSvg = Encoding.UTF8.GetString(LaTeXBlockService.ApplyFractionalBaselineCompensation(
                    autoSvg11.Bytes, autoSvg11.DepthPt, 1));
                Assert(compensatedSvg.IndexOf("data-latexblocks-baseline-residual", StringComparison.Ordinal) >= 0,
                    "The sub-point TeX depth residual was not encoded in the inline SVG viewport.");
                Assert(firstMetadata.Mode == LaTeXBlockLayoutMode.Auto, "Auto-width mode was not stored in metadata.");
                Assert(inserted.Width < 100, "Auto-width formula retained the fixed typesetting canvas width.");
                Assert(Math.Abs((double)inserted.Range.Font.Size - firstMetadata.FontSizePt) < 0.001,
                    "The inserted formula's Word run size does not match its TeX design size.");
                Assert(inserted.Range.Font.Position == -(int)Math.Round(firstMetadata.DepthPt, MidpointRounding.AwayFromZero),
                    "Word baseline compensation does not equal the rounded TeX depth.");
                AssertInlineWordJoinerBoundary(inserted, 2, "Inserted inline formula");
                Assert(word.Selection.Start == word.Selection.End && word.Selection.Start == inserted.Range.End + 1 &&
                    word.Selection.Font.Position == 0 && word.Selection.NoProofing == 0,
                    "Inline insertion left the picture selected or leaked its character formatting into the caret.");
                var runningTextStart = word.Selection.Start;
                word.Selection.TypeText(" running");
                var runningText = document.Range(runningTextStart, runningTextStart + 8);
                Assert(document.InlineShapes.Count == 1 && runningText.Text == " running" &&
                    runningText.Font.Position == 0 && runningText.NoProofing == 0 &&
                    document.Range(inserted.Range.End, inserted.Range.End + 1).Text == WordJoiner,
                    "Text typed after an inline formula inherited the picture run's baseline or no-proof formatting.");
                var inlineExport = WordSelectionLaTeXExporter.Export(document.Range(
                    inserted.Range.Start - 1, runningText.End));
                Assert(inlineExport == source + " running" &&
                    inlineExport.IndexOf(WordJoiner, StringComparison.Ordinal) < 0,
                    "Selected Word text did not export the inline Block as its exact LaTeX source.");
                document.Range(0, 0).Select();
                var hostSizeBeforeSelection = (double)inserted.Range.Font.Size;
                inserted.Range.Select();
                var hostSizeWhileSelected = (double)inserted.Range.Font.Size;
                document.Range(0, 0).Select();
                var hostSizeAfterSelection = (double)inserted.Range.Font.Size;
                Assert(Math.Abs(hostSizeBeforeSelection - hostSizeWhileSelected) < 0.001 &&
                    Math.Abs(hostSizeBeforeSelection - hostSizeAfterSelection) < 0.001,
                    "Selecting and leaving an InlineShape changed Word's host character font size.");
                var stableId = firstMetadata.Id;
                var end = document.Range(document.Content.End - 1, document.Content.End - 1);
                end.Select();
                const string fixedBlockSource = "\\[x^2\\]";
                var fixedBlock = service.InsertBlock(fixedBlockSource, 180, LaTeXBlockLayoutMode.Fixed, alternateProfile);
                Assert(LaTeXBlockMetadata.TryParse(fixedBlock.Title, out var fixedMetadata) &&
                    fixedMetadata.Mode == LaTeXBlockLayoutMode.Fixed, "Fixed-width block mode was not persisted.");
                Assert(fixedBlock.Width > 150, "Fixed-width block lost its requested canvas width.");
                Assert(fixedBlock.LockAspectRatio == Microsoft.Office.Core.MsoTriState.msoFalse,
                    "A fixed Content Block did not expose independent native frame dimensions.");
                Assert(document.Range(fixedBlock.Range.Start - 1, fixedBlock.Range.Start).Text != WordJoiner &&
                       document.Range(fixedBlock.Range.End, fixedBlock.Range.End + 1).Text != WordJoiner,
                    "A fixed-width block unexpectedly received inline U+2060 boundaries.");
                Assert(document.InlineShapes.Count == 2, "The two insertion modes did not produce two InlineShapes.");

                // An In Line with Text Block is still a Block frame, not an ordinary
                // proportional picture. Simulate Word's native independent resize,
                // then replace it with an exact framed SVG as the mouse-up path does.
                const double requestedInlineFrameWidthPt = 226.25;
                const double requestedInlineFrameHeightPt = 61.5;
                fixedBlock.Width = (float)requestedInlineFrameWidthPt;
                fixedBlock.Height = (float)requestedInlineFrameHeightPt;
                var expectedInlineLayoutWidth = fixedMetadata.WidthPt +
                    requestedInlineFrameWidthPt - fixedMetadata.FrameWidthPt;
                expectedInlineLayoutWidth = Math.Max(LaTeXBlockWidthPolicy.MinimumWidthPt,
                    Math.Min(LaTeXBlockWidthPolicy.MaximumWidthPt, expectedInlineLayoutWidth));
                var inlineReflowRawRender = service.RenderPreview(fixedBlockSource,
                    expectedInlineLayoutWidth, LaTeXBlockLayoutMode.Fixed, alternateProfile,
                    fixedMetadata.FontSizePt);
                var inlineReflowFrameRender = service.FrameFloatingRender(inlineReflowRawRender,
                    requestedInlineFrameWidthPt, requestedInlineFrameHeightPt);
                fixedBlock = service.UpdateRendered(fixedBlock, fixedBlockSource,
                    expectedInlineLayoutWidth, LaTeXBlockLayoutMode.Fixed,
                    inlineReflowFrameRender, false);
                Assert(LaTeXBlockMetadata.TryParse(fixedBlock.Title,
                        out var reflowedInlineMetadata) &&
                       Math.Abs(fixedBlock.Width - requestedInlineFrameWidthPt) < 0.05 &&
                       Math.Abs(fixedBlock.Height - requestedInlineFrameHeightPt) < 0.05 &&
                       Math.Abs(reflowedInlineMetadata.WidthPt - expectedInlineLayoutWidth) < 0.01 &&
                       Math.Abs(reflowedInlineMetadata.FrameWidthPt - requestedInlineFrameWidthPt) < 0.01 &&
                       Math.Abs(reflowedInlineMetadata.FrameHeightPt - requestedInlineFrameHeightPt) < 0.01 &&
                       fixedBlock.LockAspectRatio == Microsoft.Office.Core.MsoTriState.msoFalse,
                    "An inline fixed Block resize did not persist an exact unscaled SVG frame.");
                document.SaveAs2(documentPath, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(documentPath, ReadOnly: false);
                Assert(document.InlineShapes.Count == 2, "The SVG objects did not survive save and reopen.");
                var reopened = document.InlineShapes[1];
                Assert(reopened.AlternativeText == source, "Exact TeX source did not survive save and reopen.");
                Assert(LaTeXBlockMetadata.TryParse(reopened.Title, out var reopenedMetadata) && reopenedMetadata.Id == stableId,
                    "Block identity did not survive save and reopen.");
                Assert(reopened.Range.Font.Position == -(int)Math.Round(reopenedMetadata.DepthPt, MidpointRounding.AwayFromZero),
                    "Baseline compensation did not survive save and reopen.");
                AssertInlineWordJoinerBoundary(reopened, 2, "Reopened inline formula");

                var updated = service.UpdateBlock(reopened, updatedSource, 420, LaTeXBlockLayoutMode.Auto, profile, 14);
                Assert(document.InlineShapes.Count == 2, "Update changed the document's SVG count.");
                Assert(updated.AlternativeText == updatedSource, "Update did not replace the authoritative TeX source.");
                Assert(LaTeXBlockMetadata.TryParse(updated.Title, out var updatedMetadata) && updatedMetadata.Id == stableId,
                    "Update changed the block identity.");
                Assert(Math.Abs(updatedMetadata.WidthPt - 420) < 0.001, "Update did not persist the new typesetting width.");
                Assert(Math.Abs(updatedMetadata.FontSizePt - 14) < 0.001,
                    "Update did not rerender and persist the requested TeX font size.");
                Assert(Math.Abs((double)updated.Range.Font.Size - updatedMetadata.FontSizePt) < 0.001,
                    "Updating an inline formula discarded its Word host font size.");
                Assert(updated.Range.Font.Position ==
                    -(int)Math.Round(updatedMetadata.DepthPt, MidpointRounding.AwayFromZero),
                    "Updating an inline formula discarded its baseline position.");
                AssertInlineWordJoinerBoundary(updated, 2, "Updated inline formula");
                Assert(word.Selection.Start == word.Selection.End && word.Selection.Start == updated.Range.End + 1 &&
                    word.Selection.Font.Position == 0 && word.Selection.NoProofing == 0,
                    "Updating an inline formula left its compensated picture run selected.");
                var updatedRunningStart = word.Selection.Start;
                word.Selection.TypeText(" updated");
                var updatedRunning = document.Range(updatedRunningStart, updatedRunningStart + 8);
                Assert(updatedRunning.Text == " updated" && updatedRunning.Font.Position == 0 &&
                    updatedRunning.NoProofing == 0 &&
                    document.Range(updated.Range.End, updated.Range.End + 1).Text == WordJoiner,
                    "Text typed after updating a formula inherited the picture run's formatting.");
                document.Save();
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(documentPath, ReadOnly: true);
                Assert(document.InlineShapes.Count == 2 && document.InlineShapes[1].AlternativeText == updatedSource,
                    "Updated block did not survive the second reopen.");
                var reopenedUpdated = document.InlineShapes[1];
                Assert(LaTeXBlockMetadata.TryParse(reopenedUpdated.Title, out var reopenedUpdatedMetadata) &&
                    Math.Abs((double)reopenedUpdated.Range.Font.Size - reopenedUpdatedMetadata.FontSizePt) < 0.001 &&
                    reopenedUpdated.Range.Font.Position ==
                    -(int)Math.Round(reopenedUpdatedMetadata.DepthPt, MidpointRounding.AwayFromZero),
                    "The updated host font size or baseline position did not survive the second reopen.");
                AssertInlineWordJoinerBoundary(reopenedUpdated, 2,
                    "Second-reopened inline formula");
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Add();
                document.Range(0, 0).Text = "Alpha beta";
                document.Range(2, 4).Select();
                var expandedTargetRejected = false;
                try { LaTeXBlockService.ValidateNumberedEquationTarget(word.Selection.Range); }
                catch (InvalidOperationException) { expandedTargetRejected = true; }
                Assert(expandedTargetRejected && document.Tables.Count == 0 &&
                    document.Content.Text.IndexOf("Alpha beta", StringComparison.Ordinal) >= 0,
                    "A numbered equation was allowed to replace an expanded selection.");
                document.Content.Text = "Alpha\tbeta";
                document.Range(5, 5).Select();
                var ordinaryTabRejected = false;
                try { LaTeXBlockService.ValidateNumberedEquationTarget(word.Selection.Range); }
                catch (InvalidOperationException) { ordinaryTabRejected = true; }
                Assert(ordinaryTabRejected,
                    "Numbered-equation tab stops were allowed to overwrite an ordinary tab layout.");
                document.Content.Text = "Alpha beta";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 11;
                document.Paragraphs[1].Format.LeftIndent = 24;
                document.Paragraphs[1].Format.RightIndent = 18;
                document.Paragraphs[1].Format.FirstLineIndent = 12;
                document.Paragraphs[1].Format.SpaceBefore = 3;
                document.Paragraphs[1].Format.SpaceAfter = 5;
                document.Paragraphs[1].Format.LineSpacingRule = WordInterop.WdLineSpacing.wdLineSpaceExactly;
                document.Range(5, 5).Select();
                var exactSpacingRejected = false;
                try { LaTeXBlockService.ValidateNumberedEquationTarget(word.Selection.Range); }
                catch (InvalidOperationException) { exactSpacingRejected = true; }
                Assert(exactSpacingRejected,
                    "A display equation was allowed into an Exact-line-spacing paragraph where it cannot expand.");
                document.Paragraphs[1].Format.LineSpacingRule = WordInterop.WdLineSpacing.wdLineSpaceSingle;
                document.Paragraphs[1].Range.ParagraphFormat.TabStops.Add(90,
                    WordInterop.WdTabAlignment.wdAlignTabLeft, WordInterop.WdTabLeader.wdTabLeaderSpaces);
                document.Range(5, 5).Select();
                var customStopsRejected = false;
                try { LaTeXBlockService.ValidateNumberedEquationTarget(word.Selection.Range); }
                catch (InvalidOperationException) { customStopsRejected = true; }
                Assert(customStopsRejected,
                    "Numbered-equation insertion was allowed to erase an existing custom tab layout.");
                document.Paragraphs[1].Range.ParagraphFormat.TabStops.ClearAll();
                document.Range(5, 5).Select();
                LaTeXBlockService.ValidateNumberedEquationTarget(word.Selection.Range);
                var numberedSource = "\\[E=mc^2\\]";
                const double numberedWidth = 360;
                Console.WriteLine("Numbered equation: rendering natural-width display SVG...");
                var numberedRender = service.RenderPreview(numberedSource, numberedWidth, LaTeXBlockLayoutMode.Auto,
                    profile, 11, true);
                Console.WriteLine("Numbered equation: inserting same-paragraph scaffold...");
                var firstNumbered = service.InsertNumberedRendered(numberedSource, numberedWidth,
                    LaTeXBlockLayoutMode.Auto, numberedRender);
                Console.WriteLine("Numbered equation: first insertion returned.");
                Assert(LaTeXBlockMetadata.TryParse(firstNumbered.Title, out var firstNumberedMetadata) &&
                    firstNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation &&
                    firstNumberedMetadata.Mode == LaTeXBlockLayoutMode.Auto,
                    "The natural-width numbered-equation contract was not stored in SVG metadata.");
                Assert(firstNumbered.AlternativeText == numberedSource,
                    "The display-style render wrapper leaked into Alternative Text.");
                var expectedNumberedCaret = document.Fields[1].Result.End + 2;
                if (expectedNumberedCaret < document.Content.End &&
                    document.Range(expectedNumberedCaret, expectedNumberedCaret + 1).Text == "\v")
                    expectedNumberedCaret++;
                Assert(word.Selection.Start == word.Selection.End && word.Selection.Start == expectedNumberedCaret &&
                    word.Selection.Font.Position == 0 && word.Selection.NoProofing == 0,
                    "Numbered-equation insertion left the picture selected or the caret before its line break.");
                Assert(firstNumbered.Width < 100,
                    "The numbered equation retained a fixed-width display canvas.");
                Assert(document.Tables.Count == 0 && document.InlineShapes.Count == 1 &&
                    document.Paragraphs.Count == 1 && document.Fields.Count == 1,
                    "The first numbered equation did not remain in one table-free Word paragraph.");
                Assert(HasCaptionLabel(word, LaTeXBlockService.EquationCategoryIdentifier),
                    "Word did not register the LaTeXBlockEq caption category for numbered equations.");
                Assert(Regex.IsMatch(document.Fields[1].Code.Text ?? string.Empty,
                    "^\\s*SEQ\\s+" + LaTeXBlockService.EquationSequenceIdentifier + "(?:\\s|$)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
                    "A new numbered equation did not use the LaTeXBlockEq SEQ category.");
                var firstParagraphText = document.Paragraphs[1].Range.Text ?? string.Empty;
                Assert(firstParagraphText.StartsWith("Alpha\v\t", StringComparison.Ordinal) &&
                    firstParagraphText.IndexOf("\t(1)\v beta", StringComparison.Ordinal) >= 0,
                    "The numbered equation does not use the expected manual-break/tab scaffold.");
                var numberedExport = WordSelectionLaTeXExporter.Export(document.Content);
                Assert(numberedExport == "Alpha\n" + numberedSource + "\n beta",
                    "A numbered equation exported its Word tab, parentheses, or SEQ-field scaffold.");
                AssertEquationTabStops(document.Paragraphs[1]);
                var insertedParagraphFormat = document.Paragraphs[1].Range.ParagraphFormat;
                Assert(Math.Abs(insertedParagraphFormat.LeftIndent - 24) < 0.01 &&
                    Math.Abs(insertedParagraphFormat.RightIndent - 18) < 0.01 &&
                    Math.Abs(insertedParagraphFormat.FirstLineIndent - 12) < 0.01 &&
                    Math.Abs(insertedParagraphFormat.SpaceBefore - 3) < 0.01 &&
                    Math.Abs(insertedParagraphFormat.SpaceAfter - 5) < 0.01 &&
                    insertedParagraphFormat.LineSpacingRule == WordInterop.WdLineSpacing.wdLineSpaceSingle,
                    "SVG normalization or editing did not preserve the paragraph's direct formatting.");
                Assert(document.Bookmarks.Exists(LaTeXBlockService.EquationBookmarkName(firstNumberedMetadata.Id)),
                    "The first equation number has no stable bookmark.");
                Assert(document.Bookmarks[LaTeXBlockService.EquationBookmarkName(firstNumberedMetadata.Id)].Range.Text == "1",
                    "The equation bookmark does not identify the SEQ field result.");
                var insideScaffoldRejected = false;
                try
                {
                    var insideScaffold = document.Range(firstNumbered.Range.End, firstNumbered.Range.End);
                    LaTeXBlockService.ValidateNumberedEquationTarget(insideScaffold);
                }
                catch (InvalidOperationException) { insideScaffoldRejected = true; }
                Assert(insideScaffoldRejected,
                    "An insertion point inside an existing equation scaffold was accepted.");

                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                Console.WriteLine("Numbered equation: inserting second equation...");
                var secondNumberedRender = service.RenderPreview(commentedNumberedSource, numberedWidth,
                    LaTeXBlockLayoutMode.Auto, profile, 11, true);
                var secondNumbered = service.InsertNumberedRendered(commentedNumberedSource, numberedWidth,
                    LaTeXBlockLayoutMode.Auto, secondNumberedRender);
                Assert(LaTeXBlockMetadata.TryParse(secondNumbered.Title, out var secondNumberedMetadata) &&
                    secondNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation &&
                    secondNumbered.AlternativeText == canonicalCommentedNumberedSource,
                    "The second equation did not receive numbered metadata.");
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var thirdNumbered = service.InsertNumberedRendered(numberedSource, numberedWidth,
                    LaTeXBlockLayoutMode.Auto, numberedRender);
                Assert(LaTeXBlockMetadata.TryParse(thirdNumbered.Title, out var thirdNumberedMetadata) &&
                    thirdNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation,
                    "The third equation did not receive numbered metadata.");
                Assert(service.UpdateEquationNumbers(document) == 3,
                    "Word did not find all three LaTeX equation sequence fields.");
                Assert(EquationNumberText(document.Fields[1]) == "1" &&
                    EquationNumberText(document.Fields[2]) == "2" &&
                    EquationNumberText(document.Fields[3]) == "3",
                    "Word SEQ fields did not number equations in document order.");
                Assert(document.Paragraphs.Count == 1 && document.Tables.Count == 0,
                    "A second numbered equation introduced a paragraph or table boundary.");

                var oversizedSvg = Regex.Replace(Encoding.UTF8.GetString(numberedRender.SvgBytes),
                    "(<svg\\b[^>]*?\\bwidth=['\"])[^'\"]+(['\"])", "${1}1000pt${2}",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var oversizedRender = new LaTeXBlockRender(numberedRender.SvgPath,
                    Encoding.UTF8.GetBytes(oversizedSvg), numberedRender.DepthPt, numberedRender.FontSizePt);
                var oversizedUpdateRejected = false;
                try
                {
                    service.UpdateRendered(firstNumbered, numberedSource, numberedWidth, LaTeXBlockLayoutMode.Auto,
                        oversizedRender, false);
                }
                catch (InvalidOperationException) { oversizedUpdateRejected = true; }
                Assert(oversizedUpdateRejected && document.Tables.Count == 0 && document.InlineShapes.Count == 3 &&
                    firstNumbered.AlternativeText == numberedSource && EquationNumberText(document.Fields[1]) == "1",
                    "An oversized edit damaged or replaced the previous numbered equation.");

                var updatedNumbered = service.UpdateBlock(firstNumbered, "\\[E=h\\nu\\]", numberedWidth,
                    LaTeXBlockLayoutMode.Auto, profile, 11);
                Assert(LaTeXBlockMetadata.TryParse(updatedNumbered.Title, out var updatedNumberedMetadata) &&
                    updatedNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation,
                    "Editing a numbered equation discarded its numbered role.");
                Assert(document.Tables.Count == 0 && EquationNumberText(document.Fields[1]) == "1" &&
                    document.Bookmarks.Exists(LaTeXBlockService.EquationBookmarkName(updatedNumberedMetadata.Id)),
                    "Editing the formula disturbed its Word-native field or bookmark.");
                var expectedUpdatedNumberedCaret = document.Fields[1].Result.End + 2;
                if (expectedUpdatedNumberedCaret < document.Content.End &&
                    document.Range(expectedUpdatedNumberedCaret, expectedUpdatedNumberedCaret + 1).Text == "\v")
                    expectedUpdatedNumberedCaret++;
                Assert(word.Selection.Start == word.Selection.End &&
                    word.Selection.Start == expectedUpdatedNumberedCaret,
                    "Editing a numbered equation left the caret inside its tab/number scaffold.");

                Console.WriteLine("Numbered equation: saving tab-stop document...");
                document.SaveAs2(numberedDocumentPath, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(numberedDocumentPath, ReadOnly: false);
                Assert(document.Tables.Count == 0 && document.InlineShapes.Count == 3 &&
                    document.Paragraphs.Count == 1,
                    "The same-paragraph numbered equations did not survive save and reopen.");
                Assert(EquationNumberText(document.Fields[1]) == "1" &&
                    EquationNumberText(document.Fields[2]) == "2" &&
                    EquationNumberText(document.Fields[3]) == "3",
                    "Equation number results did not survive save and reopen.");
                var reopenedNumbered = document.InlineShapes[1];
                Assert(LaTeXBlockMetadata.TryParse(reopenedNumbered.Title, out var reopenedNumberedMetadata) &&
                    reopenedNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation,
                    "The numbered role did not survive save and reopen.");
                AssertEquationTabStops(document.Paragraphs[1]);
                var reopenedParagraphFormat = document.Paragraphs[1].Range.ParagraphFormat;
                Assert(Math.Abs(reopenedParagraphFormat.LeftIndent - 24) < 0.01 &&
                    Math.Abs(reopenedParagraphFormat.RightIndent - 18) < 0.01 &&
                    Math.Abs(reopenedParagraphFormat.FirstLineIndent - 12) < 0.01 &&
                    Math.Abs(reopenedParagraphFormat.SpaceBefore - 3) < 0.01 &&
                    Math.Abs(reopenedParagraphFormat.SpaceAfter - 5) < 0.01 &&
                    reopenedParagraphFormat.LineSpacingRule == WordInterop.WdLineSpacing.wdLineSpaceSingle,
                    "Paragraph formatting did not survive formula editing and DOCX reopen.");
                var firstBookmarkName = LaTeXBlockService.EquationBookmarkName(firstNumberedMetadata.Id);
                var secondBookmarkName = LaTeXBlockService.EquationBookmarkName(secondNumberedMetadata.Id);
                var thirdBookmarkName = LaTeXBlockService.EquationBookmarkName(thirdNumberedMetadata.Id);
                Assert(document.Bookmarks.Exists(firstBookmarkName) && document.Bookmarks.Exists(secondBookmarkName) &&
                    document.Bookmarks.Exists(thirdBookmarkName),
                    "Equation bookmarks did not survive save and reopen.");
                Assert(document.InlineShapes[2].AlternativeText == canonicalCommentedNumberedSource,
                    "Word lost the canonical multiline/commented TeX source on save and reopen.");

                var referenceTargets = service.GetEquationReferenceTargets(document);
                Assert(referenceTargets.Count == 3 &&
                       referenceTargets[0].Id == updatedNumberedMetadata.Id &&
                       referenceTargets[1].Id == secondNumberedMetadata.Id &&
                       referenceTargets[2].Id == thirdNumberedMetadata.Id &&
                       referenceTargets[1].BookmarkName == secondBookmarkName &&
                       referenceTargets[1].Number == "2" &&
                       referenceTargets[1].Source == canonicalCommentedNumberedSource,
                    "The bookmark-backed equation-reference picker did not enumerate the three equations in document order.");
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var referenceField = service.InsertEquationReference(referenceTargets[1]);
                Assert(LaTeXBlockService.IsEquationReferenceField(referenceField) &&
                       (referenceField.Code.Text ?? string.Empty).IndexOf(secondBookmarkName,
                           StringComparison.OrdinalIgnoreCase) >= 0 &&
                       (referenceField.Code.Text ?? string.Empty).IndexOf("\\h",
                           StringComparison.OrdinalIgnoreCase) >= 0 &&
                       (document.Content.Text ?? string.Empty).EndsWith("(2)\r", StringComparison.Ordinal) &&
                       word.Selection.Start == document.Content.End - 1,
                    "The equation-reference command did not insert one native hyperlink REF field in parentheses.");

                // The reference is a native document field, not a cached plugin
                // string: it must persist together with its bookmark target.
                document.Save();
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = word.Documents.Open(numberedDocumentPath, ReadOnly: false);
                Assert(document.InlineShapes.Count == 3 && document.Fields.Count == 4,
                    "The equation reference did not survive save and reopen.");
                reopenedNumbered = document.InlineShapes[1];
                reopenedParagraphFormat = document.Paragraphs[1].Range.ParagraphFormat;
                referenceField = FindEquationReferenceField(document, secondBookmarkName);
                Assert((document.Content.Text ?? string.Empty).EndsWith("(2)\r", StringComparison.Ordinal),
                    "The persisted equation reference did not retain its target number.");
                var reopenedPage = document.Paragraphs[1].Range.Sections[1].PageSetup;
                var reopenedColumnWidth = (double)reopenedPage.PageWidth - reopenedPage.LeftMargin - reopenedPage.RightMargin;
                var reopenedColumns = reopenedPage.TextColumns;
                if (reopenedColumns.Count > 1)
                    reopenedColumnWidth = (reopenedColumnWidth - reopenedColumns.Spacing * (reopenedColumns.Count - 1)) /
                                          reopenedColumns.Count;
                var oldUsableWidth = reopenedColumnWidth - 24 - 18;
                reopenedParagraphFormat.TabStops.ClearAll();
                reopenedParagraphFormat.TabStops.Add((float)(24 + oldUsableWidth / 2),
                    WordInterop.WdTabAlignment.wdAlignTabCenter, WordInterop.WdTabLeader.wdTabLeaderSpaces);
                reopenedParagraphFormat.TabStops.Add((float)(24 + oldUsableWidth),
                    WordInterop.WdTabAlignment.wdAlignTabRight, WordInterop.WdTabLeader.wdTabLeaderSpaces);
                reopenedParagraphFormat.LeftIndent = 0;
                reopenedParagraphFormat.RightIndent = 0;
                Assert(service.UpdateEquationNumbers(document) == 3,
                    "Updating equation numbers did not visit every numbered equation.");
                AssertEquationTabStops(document.Paragraphs[1]);
                Assert(Math.Abs(reopenedParagraphFormat.LeftIndent) < 0.01 &&
                       Math.Abs(reopenedParagraphFormat.RightIndent) < 0.01,
                    "Migrating stale equation tab stops changed the paragraph indents.");
                LaTeXBlockService.NumberedEquationLineRange(reopenedNumbered).Delete();
                Assert(service.UpdateEquationNumbers(document) == 2 && EquationNumberText(document.Fields[1]) == "1" &&
                    EquationNumberText(document.Fields[2]) == "2",
                    "The remaining equation did not renumber after deleting an earlier equation.");
                Assert(!document.Bookmarks.Exists(firstBookmarkName) && document.Bookmarks.Exists(secondBookmarkName) &&
                    document.Bookmarks.Exists(thirdBookmarkName),
                    "Deleting one equation removed the wrong bookmark or left a stale bookmark behind.");
                Assert(document.Bookmarks[secondBookmarkName].Range.Text == "1",
                    "The surviving equation bookmark did not follow its updated SEQ result.");
                Assert(document.Bookmarks[thirdBookmarkName].Range.Text == "2",
                    "The third equation bookmark did not follow its updated SEQ result.");
                Assert((document.Content.Text ?? string.Empty).EndsWith("(1)\r", StringComparison.Ordinal),
                    "The equation REF field did not follow its target after renumbering.");
                var adjacentEquation = document.InlineShapes[1];
                LaTeXBlockService.NumberedEquationLineRange(adjacentEquation).Delete();
                Assert(service.UpdateEquationNumbers(document) == 1 &&
                    EquationNumberText(document.Fields[1]) == "1" &&
                    !document.Bookmarks.Exists(secondBookmarkName) &&
                    document.Bookmarks.Exists(thirdBookmarkName) &&
                    document.Bookmarks[thirdBookmarkName].Range.Text == "1" &&
                    (document.Paragraphs[1].Range.Text ?? string.Empty).IndexOf("\v\t", StringComparison.Ordinal) >= 0,
                    "Deleting an equation adjacent to another display removed their shared visual-line boundary.");

                // Documents written before the public LaTeXBlockEq category used a
                // private LaTeXEquation sequence. A regular Update Numbers command
                // must move such fields into the one current category without
                // changing their visible numbering result.
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = word.Documents.Add();
                var legacyField = document.Fields.Add(document.Range(0, 0),
                    WordInterop.WdFieldType.wdFieldSequence,
                    "LaTeXEquation \\* ARABIC", false);
                Assert(legacyField.Update(), "Word could not create the legacy equation SEQ field.");
                Assert(service.UpdateEquationNumbers(document) == 1,
                    "Update Numbers did not find the legacy equation sequence.");
                Assert(Regex.IsMatch(legacyField.Code.Text ?? string.Empty,
                    "^\\s*SEQ\\s+" + LaTeXBlockService.EquationSequenceIdentifier + "(?:\\s|$)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) &&
                    EquationNumberText(legacyField) == "1",
                    "Update Numbers did not migrate a legacy equation into LaTeXBlockEq.");

                RunStyledBlockSmoke(word, service, profile);
                RunShutdownProbe(renderer, profile);

                Console.WriteLine("LaTeX Blocks smoke test passed.");
                Console.WriteLine("StemTeX: " + renderer.StemTeXHome);
                Console.WriteLine("Verified: SVG insertion, metadata, update, Word-native equation numbering and references, and DOCX persistence.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
            finally
            {
                if (document != null) { document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges); Release(document); }
                if (word != null)
                {
                    if (ownsWord) word.Quit(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    Release(word);
                }
                renderer?.Dispose();
            }
        }

        private static void RunRenderHostClientSmoke(string profile)
        {
            Console.WriteLine("RenderHost: testing isolated SVG rendering...");
            var preexistingHostProcesses = CaptureProcessIds("LaTeXBlocks.RenderHost.host");
            var preexistingWorkerProcesses = CaptureProcessIds("stemtex-worker-host");
            using (var remote = new RenderHostClientBackend())
            {
                Assert(remote.Profiles.Length > 0 &&
                       Array.IndexOf(remote.Profiles, profile) >= 0,
                    "The isolated RenderHost client did not discover the installed profile set.");
                remote.SwitchProfile(profile);
                var result = remote.RenderQueuedAsync(profile, "$E=mc^2$", 180, true, 11)
                    .GetAwaiter().GetResult();
                Assert(result != null && result.Bytes != null && result.Bytes.Length > 0 &&
                       Encoding.UTF8.GetString(result.Bytes).IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0,
                    "The isolated RenderHost did not return SVG output.");
                Assert(result.DepthPt > 0,
                    "The isolated RenderHost lost the TeX baseline-depth result.");
                Assert(remote.Status.StartsWith("ready:", StringComparison.OrdinalIgnoreCase),
                    "The isolated RenderHost did not report a ready renderer after a completed request.");
            }
            WaitFor(() => NoNewProcesses("LaTeXBlocks.RenderHost.host", preexistingHostProcesses) &&
                          NoNewProcesses("stemtex-worker-host", preexistingWorkerProcesses), 5000,
                "Disposing the RenderHost client left a renderer broker or XeTeX worker process behind.");
            Console.WriteLine("RenderHost: isolated SVG rendering and parent-owned disposal passed.");
        }

        private static HashSet<int> CaptureProcessIds(string processName)
        {
            var result = new HashSet<int>();
            foreach (var process in Process.GetProcessesByName(processName))
            {
                try { result.Add(process.Id); }
                finally { process.Dispose(); }
            }
            return result;
        }

        private static bool NoNewProcesses(string processName, HashSet<int> preexisting)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                try
                {
                    if (!preexisting.Contains(process.Id)) return false;
                }
                finally { process.Dispose(); }
            }
            return true;
        }

        private static void RunStyledBlockSmoke(WordInterop.Application word,
            LaTeXBlockService service, string profile)
        {
            WordInterop.Document document = null;
            WordInterop.Document previousDocument = null;
            var previousSelectionStart = 0;
            var previousSelectionEnd = 0;
            var previousFontSize = 10.0;
            var previousFontPosition = 0.0;
            var previousNoProofing = 0;
            var previousColor = LaTeXBlockService.AutomaticTextColor;
            try
            {
                previousDocument = word.ActiveDocument;
                previousSelectionStart = word.Selection.Start;
                previousSelectionEnd = word.Selection.End;
                previousFontSize = (double)word.Selection.Font.Size;
                previousFontPosition = (double)word.Selection.Font.Position;
                previousNoProofing = word.Selection.NoProofing;
                previousColor = (int)word.Selection.Font.Color;
                document = word.Documents.Add();
                document.Range(0, 0).Text = "Styled ";
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var style = new LaTeXBlockStyle(1.35, 6,
                    LaTeXBlockVerticalAlignment.Bottom,
                    System.Drawing.Color.FromArgb(0x00, 0x55, 0xaa), true,
                    System.Drawing.Color.FromArgb(0xff, 0xfa, 0xe8), 1.5,
                    System.Drawing.Color.FromArgb(0x33, 0x22, 0x11));
                var render = service.RenderPreview("\\[E=mc^2\\]", 240,
                    LaTeXBlockLayoutMode.Fixed, profile, 14, false,
                    LaTeXBlockService.ToWordColor(style.TextColor), style);
                Assert(Encoding.UTF8.GetString(render.SvgBytes).IndexOf(
                        "data-latexblocks-border='1'", StringComparison.Ordinal) >= 0 &&
                       render.ContentSvgBytes != render.SvgBytes,
                    "The styled fixed Block preview did not retain raw TeX content and one SVG shell.");
                var shape = service.InsertRendered("\\[E=mc^2\\]", 240,
                    LaTeXBlockLayoutMode.Fixed, render, style);
                Assert(LaTeXBlockService.TryReadContract(shape, out var metadata, out var source) &&
                       source == "\\[E=mc^2\\]" && metadata.HasExplicitStyle &&
                       metadata.Style.Equals(style),
                    "Inserting a styled Word Block lost its TeX source or persistent style.");
                const double resizedWidth = 278.25;
                const double resizedHeight = 76.5;
                var resizedRaw = service.RenderPreview(source, resizedWidth,
                    LaTeXBlockLayoutMode.Fixed, profile, 14, false,
                    LaTeXBlockService.ToWordColor(style.TextColor), style);
                var resized = service.FrameFloatingRender(resizedRaw, resizedWidth,
                    resizedHeight, style);
                shape = service.UpdateRendered(shape, source, resizedWidth,
                    LaTeXBlockLayoutMode.Fixed, resized, false, style);
                var decoratedText = Encoding.UTF8.GetString(resized.SvgBytes);
                Assert(LaTeXBlockService.TryReadContract(shape, out metadata, out source) &&
                       metadata.Style.Equals(style) &&
                       Math.Abs(shape.Width - resizedWidth) < 0.05 &&
                       Math.Abs(shape.Height - resizedHeight) < 0.05 &&
                       CountOccurrences(decoratedText, "data-latexblocks-frame='1'") == 1 &&
                       CountOccurrences(decoratedText, "data-latexblocks-border='1'") == 1,
                    "A Word fixed-Block resize did not repaint exactly one persistent SVG style shell.");
                // The async native-resize completion deliberately does not pass a
                // separate style argument: it must recover durable Title data from
                // the current object and retain exactly one SVG shell.
                shape = service.UpdateRendered(shape, source, resizedWidth,
                    LaTeXBlockLayoutMode.Fixed, resized, false);
                var recoveredText = Encoding.UTF8.GetString(resized.SvgBytes);
                Assert(LaTeXBlockService.TryReadContract(shape, out metadata, out source) &&
                       metadata.HasExplicitStyle && metadata.Style.Equals(style) &&
                       CountOccurrences(recoveredText, "data-latexblocks-frame='1'") == 1 &&
                       CountOccurrences(recoveredText, "data-latexblocks-border='1'") == 1,
                    "A Word reflow completion without a transient style argument lost or doubled the persistent SVG shell.");
                Console.WriteLine("Word: styled fixed Block persistence and exact frame reflow passed.");
            }
            finally
            {
                if (document != null)
                {
                    try { document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges); }
                    catch { }
                    Release(document);
                }
                if (previousDocument != null)
                {
                    try
                    {
                        previousDocument.Range(previousSelectionStart, previousSelectionEnd).Select();
                        word.Selection.Font.Size = (float)previousFontSize;
                        word.Selection.Font.Position = (int)Math.Round(previousFontPosition);
                        word.Selection.NoProofing = previousNoProofing;
                        word.Selection.Font.Color = (WordInterop.WdColor)previousColor;
                    }
                    catch { }
                }
            }
        }

        private static void RunTextColorSmoke(WordInterop.Application word,
            LaTeXBlockService service, string profile, string documentPath)
        {
            const string inlineSource = "$E=mc^2$";
            const string displaySource = "\\[E=mc^2\\]";
            const int wordRed = 0x0000ff;   // WdColor is BGR, so this is RGB FF0000.
            const int wordBlue = 0x00ff0000; // WdColor is BGR, so this is RGB 0000FF.
            WordInterop.Document document = null;
            try
            {
                Assert(LaTeXBlockService.ApplyTextColor(inlineSource, wordRed)
                           .IndexOf("\\color[HTML]{FF0000}", StringComparison.Ordinal) >= 0 &&
                       LaTeXBlockService.ApplyTextColor(inlineSource, wordBlue)
                           .IndexOf("\\color[HTML]{0000FF}", StringComparison.Ordinal) >= 0 &&
                       LaTeXBlockService.ApplyTextColor(inlineSource,
                           LaTeXBlockService.AutomaticTextColor) == inlineSource,
                    "Word BGR text colors were not converted into the intended TeX RGB colors.");

                Console.WriteLine("Word: testing native text color for inline and display formulas...");
                if (File.Exists(documentPath)) File.Delete(documentPath);
                document = word.Documents.Add();
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 14;
                document.Range(0, 0).Select();
                word.Selection.Font.Color = (WordInterop.WdColor)wordRed;
                Assert(LaTeXBlockService.TextColorsEqual(
                        LaTeXBlockService.ResolveTextColor(word.Selection), wordRed),
                    "A collapsed Word selection did not expose its native text color.");

                var automaticRender = service.RenderPreview(inlineSource, 360,
                    LaTeXBlockLayoutMode.Auto, profile, 14);
                var inline = service.InsertBlock(inlineSource, 360,
                    LaTeXBlockLayoutMode.Auto, profile);
                Assert(inline.AlternativeText == inlineSource &&
                       LaTeXBlockService.TextColorsEqual((int)inline.Range.Font.Color, wordRed),
                    "An inline formula did not preserve its raw TeX source and native Word text color.");
                var inlineRedSvg = service.RenderPreview(inlineSource, 360,
                    LaTeXBlockLayoutMode.Auto, profile, 14, false, wordRed);
                Assert(Convert.ToBase64String(automaticRender.SvgBytes) !=
                       Convert.ToBase64String(inlineRedSvg.SvgBytes),
                    "The SVG generated for an explicit Word text color is identical to automatic text color.");
                Assert(Math.Abs(LaTeXBlockService.ReadSvgWidthPt(automaticRender.SvgBytes) -
                                LaTeXBlockService.ReadSvgWidthPt(inlineRedSvg.SvgBytes)) < 0.01,
                    "Applying Word Font.Color changed the inline formula's TeX box width.");

                // A native color command changes the drawing run first. The event-driven
                // refresh path then renders the SVG from that same authoritative value.
                inline.Range.Font.Color = (WordInterop.WdColor)wordBlue;
                var recoloredInline = service.UpdateBlock(inline, inlineSource, 360,
                    LaTeXBlockLayoutMode.Auto, profile, 14, false);
                Assert(recoloredInline.AlternativeText == inlineSource &&
                       LaTeXBlockService.TextColorsEqual((int)recoloredInline.Range.Font.Color,
                           wordBlue),
                    "Updating an inline formula lost the user-selected Word text color.");
                document.SaveAs2(documentPath, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(documentPath, ReadOnly: false);
                var reopenedInline = document.InlineShapes[1];
                Assert(reopenedInline.AlternativeText == inlineSource &&
                       LaTeXBlockService.TextColorsEqual((int)reopenedInline.Range.Font.Color,
                           wordBlue),
                    "The recolored inline formula did not preserve its Word text color after save and reopen.");
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Add();
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 14;
                document.Range(0, 0).Select();
                word.Selection.Font.Color = (WordInterop.WdColor)wordRed;
                var display = service.InsertBlock(displaySource, 300,
                    LaTeXBlockLayoutMode.Fixed, profile);
                Assert(display.AlternativeText == displaySource &&
                       LaTeXBlockService.TextColorsEqual((int)display.Range.Font.Color, wordRed),
                    "A fixed-width display formula did not inherit Word's text color.");
                display.Range.Font.Color = (WordInterop.WdColor)wordBlue;
                var recoloredDisplay = service.UpdateBlock(display, displaySource, 300,
                    LaTeXBlockLayoutMode.Fixed, profile, 14, false);
                Assert(LaTeXBlockService.TextColorsEqual((int)recoloredDisplay.Range.Font.Color,
                           wordBlue),
                    "Updating a fixed-width display formula lost the Word text color.");
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Add();
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 14;
                document.Range(0, 0).Select();
                word.Selection.Font.Color = (WordInterop.WdColor)wordRed;
                var numbered = service.InsertNumberedBlock(displaySource, 360,
                    LaTeXBlockLayoutMode.Auto, profile);
                Assert(numbered.AlternativeText == displaySource &&
                       LaTeXBlockService.TextColorsEqual((int)numbered.Range.Font.Color, wordRed),
                    "A numbered display formula did not inherit Word's text color.");
                numbered.Range.Font.Color = (WordInterop.WdColor)wordBlue;
                var recoloredNumbered = service.UpdateBlock(numbered, displaySource, 360,
                    LaTeXBlockLayoutMode.Auto, profile, 14, false);
                Assert(recoloredNumbered.AlternativeText == displaySource &&
                       LaTeXBlockService.TextColorsEqual((int)recoloredNumbered.Range.Font.Color,
                           wordBlue),
                    "Updating a numbered display formula lost the Word text color.");
                Console.WriteLine("Word: text color insertion, refresh, and persistence passed.");
            }
            finally
            {
                if (document != null)
                {
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    Release(document);
                }
            }
        }

        private static void RunFloatingBlockSmoke(WordInterop.Application word, LaTeXBlockService service,
            string profile, string documentPath)
        {
            const string source = "\\[E=mc^2\\]";
            const string updatedSource = "\\[x^2+y^2\\]";
            WordInterop.Document document = null;
            try
            {
                if (File.Exists(documentPath)) File.Delete(documentPath);
                document = word.Documents.Add();
                document.Range(0, 0).Text = "A floating block anchor.";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 14;
                document.Range(4, 4).Select();
                var initialRender = service.RenderPreview(source, 180,
                    LaTeXBlockLayoutMode.Fixed, profile, 14);
                var inline = service.InsertRendered(source, 180,
                    LaTeXBlockLayoutMode.Fixed, initialRender);
                Assert(LaTeXBlockMetadata.TryParse(inline.Title, out var beforeMetadata),
                    "The fixed block did not receive its metadata before becoming floating.");

                // A fixed Block starts as an InlineShape, and Word still lets a
                // user resize that picture. Reframe it before converting it to a
                // floating Shape: Block reflow must be independent of its current
                // text-layout participation. Fixed blocks deliberately have no
                // U+2060 word-joiner scaffold.
                inline.Width = 206.5f;
                inline.Height = 61.75f;
                var requestedInlineFrameWidthPt = (double)inline.Width;
                var requestedInlineFrameHeightPt = (double)inline.Height;
                var expectedInlineLayoutWidth = beforeMetadata.WidthPt +
                    requestedInlineFrameWidthPt - beforeMetadata.FrameWidthPt;
                expectedInlineLayoutWidth = Math.Max(LaTeXBlockWidthPolicy.MinimumWidthPt,
                    Math.Min(LaTeXBlockWidthPolicy.MaximumWidthPt, expectedInlineLayoutWidth));
                var inlineReflowRawRender = service.RenderPreview(source, expectedInlineLayoutWidth,
                    LaTeXBlockLayoutMode.Fixed, profile, 14);
                var inlineReflowFrameRender = service.FrameFloatingRender(inlineReflowRawRender,
                    requestedInlineFrameWidthPt, requestedInlineFrameHeightPt);
                inline = service.UpdateRendered(inline, source, expectedInlineLayoutWidth,
                    LaTeXBlockLayoutMode.Fixed, inlineReflowFrameRender, false);
                Assert(LaTeXBlockMetadata.TryParse(inline.Title, out var inlineReflowMetadata) &&
                       Math.Abs(inline.Width - requestedInlineFrameWidthPt) < 0.05 &&
                       Math.Abs(inline.Height - requestedInlineFrameHeightPt) < 0.05 &&
                       Math.Abs(inlineReflowMetadata.WidthPt - expectedInlineLayoutWidth) < 0.01 &&
                       Math.Abs(inlineReflowMetadata.FrameWidthPt - requestedInlineFrameWidthPt) < 0.01 &&
                       Math.Abs(inlineReflowMetadata.FrameHeightPt - requestedInlineFrameHeightPt) < 0.01,
                    "A manually resized inline fixed Block did not retain its exact reframed SVG extent.");
                var precedingInlineCharacter = document.Range(inline.Range.Start - 1,
                    inline.Range.Start).Text;
                var followingInlineCharacter = document.Range(inline.Range.End,
                    inline.Range.End + 1).Text;
                Assert(precedingInlineCharacter != "\u2060" && followingInlineCharacter != "\u2060",
                    "A fixed inline Block accidentally acquired auto-formula word-joiner boundaries.");

                var floating = inline.ConvertToShape();
                floating.RelativeHorizontalPosition = (WordInterop.WdRelativeHorizontalPosition)1;
                floating.RelativeVerticalPosition = (WordInterop.WdRelativeVerticalPosition)1;
                floating.WrapFormat.Type = (WordInterop.WdWrapType)3; // In Front of Text.
                floating.Left = 101.25f;
                floating.Top = 83.5f;
                floating.WrapFormat.Side = (WordInterop.WdWrapSideType)0;
                floating.WrapFormat.DistanceLeft = 2.5f;
                floating.WrapFormat.DistanceRight = 3.5f;
                floating.WrapFormat.DistanceTop = 4.5f;
                floating.WrapFormat.DistanceBottom = 5.5f;
                floating.WrapFormat.AllowOverlap = -1;
                floating.Rotation = 7.5f;
                var expectedRotation = floating.Rotation;
                floating.Select();

                Assert(word.Selection.ShapeRange.Count == 1 && word.Selection.InlineShapes.Count == 0 &&
                       document.Shapes.Count == 1 && document.InlineShapes.Count == 0,
                    "Changing Wrap Text did not produce Word's expected floating Shape selection.");
                Assert(service.TryGetSelectedFloatingBlock(out var selected, out var floatingMetadata) &&
                       floatingMetadata.Id == beforeMetadata.Id && selected.AlternativeText == source,
                    "A selected floating LaTeX Block was not recognized by its metadata contract.");

                var updateRender = service.RenderPreview(updatedSource, 180,
                    LaTeXBlockLayoutMode.Fixed, profile, 14);
                var preservedFrameRender = service.FrameFloatingRender(updateRender,
                    selected.Width, selected.Height);
                var updated = service.UpdateFloatingRendered(selected, updatedSource, 180,
                    LaTeXBlockLayoutMode.Fixed, preservedFrameRender);
                LaTeXBlockMetadata updatedMetadata;
                string updatedText;
                var hasUpdatedContract = LaTeXBlockService.TryReadContract(updated,
                    out updatedMetadata, out updatedText);
                Assert(document.Shapes.Count == 1 && document.InlineShapes.Count == 0 &&
                       hasUpdatedContract &&
                       updatedMetadata.Id == beforeMetadata.Id && updatedText == updatedSource,
                    "Updating a floating block lost its Shape contract or changed it back to inline.");
                Assert(service.TryGetSelectedFloatingBlock(out var reselected, out var reselectedMetadata) &&
                       reselectedMetadata.Id == beforeMetadata.Id && reselected.AlternativeText == updatedSource,
                    "Updating a floating block did not leave its replacement selected for another edit.");
                Console.WriteLine("Floating block placement: wrap=" + (int)updated.WrapFormat.Type +
                    ", rel=" + (int)updated.RelativeHorizontalPosition + "/" +
                    (int)updated.RelativeVerticalPosition + ", left/top=" +
                    updated.Left.ToString("0.##") + "/" + updated.Top.ToString("0.##") +
                    ", distances=" + updated.WrapFormat.DistanceLeft.ToString("0.##") + "/" +
                    updated.WrapFormat.DistanceRight.ToString("0.##") + "/" +
                    updated.WrapFormat.DistanceTop.ToString("0.##") + "/" +
                    updated.WrapFormat.DistanceBottom.ToString("0.##") + ", rotation=" +
                    updated.Rotation.ToString("0.##"));
                Assert((int)updated.WrapFormat.Type == 3 &&
                       (int)updated.RelativeHorizontalPosition == 1 &&
                       (int)updated.RelativeVerticalPosition == 1 &&
                       Math.Abs(updated.Left - 101.25f) < 0.05 &&
                       Math.Abs(updated.Top - 83.5f) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceLeft - 2.5f) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceRight - 3.5f) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceTop - 4.5f) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceBottom - 5.5f) < 0.05 &&
                       Math.Abs(updated.Rotation - expectedRotation) < 0.05,
                    "Updating a floating block changed its Word wrapping, position, margins, or rotation.");

                // A Block is a host-independent object: it must follow the same
                // SVG replacement path when it participates in Word text wrapping,
                // not only when it is In Front of Text.  Switch the verified block
                // to Square wrapping, then perform another floating replacement and
                // retain the actual Word values (Word is allowed to normalize its
                // position when the wrap mode changes).
                updated.WrapFormat.Type = (WordInterop.WdWrapType)0; // Square.
                updated.WrapFormat.Side = (WordInterop.WdWrapSideType)0; // Both sides.
                updated.WrapFormat.DistanceLeft = 6.25f;
                updated.WrapFormat.DistanceRight = 7.25f;
                updated.WrapFormat.DistanceTop = 8.25f;
                updated.WrapFormat.DistanceBottom = 9.25f;
                updated.WrapFormat.AllowOverlap = 0;
                var expectedSquareSide = (int)updated.WrapFormat.Side;
                var expectedSquareLeft = updated.Left;
                var expectedSquareTop = updated.Top;
                var expectedSquareDistanceLeft = updated.WrapFormat.DistanceLeft;
                var expectedSquareDistanceRight = updated.WrapFormat.DistanceRight;
                var expectedSquareDistanceTop = updated.WrapFormat.DistanceTop;
                var expectedSquareDistanceBottom = updated.WrapFormat.DistanceBottom;
                var expectedSquareAllowOverlap = updated.WrapFormat.AllowOverlap;
                var squareFrameRender = service.FrameFloatingRender(updateRender,
                    updated.Width, updated.Height);
                updated = service.UpdateFloatingRendered(updated, updatedSource, 180,
                    LaTeXBlockLayoutMode.Fixed, squareFrameRender, false);
                Assert((int)updated.WrapFormat.Type == 0 &&
                       (int)updated.WrapFormat.Side == expectedSquareSide &&
                       (int)updated.RelativeHorizontalPosition == 1 &&
                       (int)updated.RelativeVerticalPosition == 1 &&
                       Math.Abs(updated.Left - expectedSquareLeft) < 0.05 &&
                       Math.Abs(updated.Top - expectedSquareTop) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceLeft - expectedSquareDistanceLeft) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceRight - expectedSquareDistanceRight) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceTop - expectedSquareDistanceTop) < 0.05 &&
                       Math.Abs(updated.WrapFormat.DistanceBottom - expectedSquareDistanceBottom) < 0.05 &&
                       updated.WrapFormat.AllowOverlap == expectedSquareAllowOverlap &&
                       Math.Abs(updated.Rotation - expectedRotation) < 0.05,
                    "Updating a Square-wrapped floating block changed its Word wrap layout.");
                Assert(LaTeXBlockService.TryReadContract(updated, out updatedMetadata,
                    out updatedText) && updatedText == updatedSource,
                    "A Square-wrapped floating block lost its LaTeX Block contract.");

                // A fixed Content Block has the same outer-frame semantics under
                // every floating Word wrap mode.  Its relationship to surrounding
                // text is not a criterion for re-rendering: update each native
                // floating variant through the same Shape -> InlineShape -> Shape
                // replacement path and retain the post-Word values exactly.
                foreach (var wrapType in new[]
                {
                    (WordInterop.WdWrapType)0, // Square
                    (WordInterop.WdWrapType)1, // Tight
                    (WordInterop.WdWrapType)2, // Through
                    (WordInterop.WdWrapType)4, // Top and Bottom
                    (WordInterop.WdWrapType)5, // Behind Text
                    (WordInterop.WdWrapType)3  // In Front of Text
                })
                {
                    updated.WrapFormat.Type = wrapType;
                    updated.WrapFormat.Side = (WordInterop.WdWrapSideType)0;
                    updated.WrapFormat.DistanceLeft = 3.25f;
                    updated.WrapFormat.DistanceRight = 4.25f;
                    updated.WrapFormat.DistanceTop = 5.25f;
                    updated.WrapFormat.DistanceBottom = 6.25f;
                    var expectedWrapType = (int)updated.WrapFormat.Type;
                    var expectedWrapSide = (int)updated.WrapFormat.Side;
                    var expectedWrapLeft = updated.Left;
                    var expectedWrapTop = updated.Top;
                    var expectedWrapDistanceLeft = updated.WrapFormat.DistanceLeft;
                    var expectedWrapDistanceRight = updated.WrapFormat.DistanceRight;
                    var expectedWrapDistanceTop = updated.WrapFormat.DistanceTop;
                    var expectedWrapDistanceBottom = updated.WrapFormat.DistanceBottom;
                    var wrapFrameRender = service.FrameFloatingRender(updateRender,
                        updated.Width, updated.Height);
                    updated = service.UpdateFloatingRendered(updated, updatedSource, 180,
                        LaTeXBlockLayoutMode.Fixed, wrapFrameRender, false);
                    Assert((int)updated.WrapFormat.Type == expectedWrapType &&
                           (int)updated.WrapFormat.Side == expectedWrapSide &&
                           Math.Abs(updated.Left - expectedWrapLeft) < 0.05 &&
                           Math.Abs(updated.Top - expectedWrapTop) < 0.05 &&
                           Math.Abs(updated.WrapFormat.DistanceLeft - expectedWrapDistanceLeft) < 0.05 &&
                           Math.Abs(updated.WrapFormat.DistanceRight - expectedWrapDistanceRight) < 0.05 &&
                           Math.Abs(updated.WrapFormat.DistanceTop - expectedWrapDistanceTop) < 0.05 &&
                           Math.Abs(updated.WrapFormat.DistanceBottom - expectedWrapDistanceBottom) < 0.05 &&
                           Math.Abs(updated.Rotation - expectedRotation) < 0.05,
                        "Updating a floating fixed Block changed its layout under wrap mode " +
                        expectedWrapType + ".");
                    Assert(LaTeXBlockService.TryReadContract(updated, out updatedMetadata,
                        out updatedText) && updatedText == updatedSource,
                        "A floating fixed Block lost its contract under wrap mode " + expectedWrapType + ".");
                }

                // Persist the next assertion under Square wrapping, whose ordinary
                // rectangular text-flow behavior makes a good representative DOCX
                // round-trip case after all floating variants above.
                updated.WrapFormat.Type = (WordInterop.WdWrapType)0;
                updated.WrapFormat.Side = (WordInterop.WdWrapSideType)0;
                updated.WrapFormat.DistanceLeft = expectedSquareDistanceLeft;
                updated.WrapFormat.DistanceRight = expectedSquareDistanceRight;
                updated.WrapFormat.DistanceTop = expectedSquareDistanceTop;
                updated.WrapFormat.DistanceBottom = expectedSquareDistanceBottom;
                expectedSquareSide = (int)updated.WrapFormat.Side;
                expectedSquareLeft = updated.Left;
                expectedSquareTop = updated.Top;
                expectedSquareDistanceLeft = updated.WrapFormat.DistanceLeft;
                expectedSquareDistanceRight = updated.WrapFormat.DistanceRight;
                expectedSquareDistanceTop = updated.WrapFormat.DistanceTop;
                expectedSquareDistanceBottom = updated.WrapFormat.DistanceBottom;
                expectedSquareAllowOverlap = updated.WrapFormat.AllowOverlap;

                // Simulate Word's native image resize, then replace it with an SVG
                // whose own physical root is the user-selected outer frame.  The
                // final Shape must retain that frame without a persisted xfrm scale.
                const double requestedFrameWidthPt = 246.25;
                const double requestedFrameHeightPt = 72.5;
                var expectedLayoutWidth = updatedMetadata.WidthPt + requestedFrameWidthPt -
                    updatedMetadata.FrameWidthPt;
                expectedLayoutWidth = Math.Max(LaTeXBlockWidthPolicy.MinimumWidthPt,
                    Math.Min(LaTeXBlockWidthPolicy.MaximumWidthPt, expectedLayoutWidth));
                var reflowRawRender = service.RenderPreview(updatedSource, expectedLayoutWidth,
                    LaTeXBlockLayoutMode.Fixed, profile, 14);
                var reflowFrameRender = service.FrameFloatingRender(reflowRawRender,
                    requestedFrameWidthPt, requestedFrameHeightPt);
                var reflowed = service.UpdateFloatingRendered(updated, updatedSource,
                    expectedLayoutWidth, LaTeXBlockLayoutMode.Fixed, reflowFrameRender, false);
                Assert(LaTeXBlockService.TryReadContract(reflowed, out var reflowedMetadata,
                        out var reflowedSource) && reflowedSource == updatedSource &&
                       Math.Abs(reflowed.Width - requestedFrameWidthPt) < 0.05 &&
                       Math.Abs(reflowed.Height - requestedFrameHeightPt) < 0.05 &&
                       Math.Abs(reflowedMetadata.WidthPt - expectedLayoutWidth) < 0.01 &&
                       Math.Abs(reflowedMetadata.FrameWidthPt - requestedFrameWidthPt) < 0.01 &&
                       Math.Abs(reflowedMetadata.FrameHeightPt - requestedFrameHeightPt) < 0.01,
                    "A floating block resize did not persist an exact unscaled SVG frame.");
                updated = reflowed;
                updatedMetadata = reflowedMetadata;

                var plainRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                var plainInline = plainRange.InlineShapes.AddPicture(initialRender.SvgPath,
                    LinkToFile: false, SaveWithDocument: true, Range: plainRange);
                var plainFloating = plainInline.ConvertToShape();
                plainFloating.Select();
                Assert(!service.TryGetSelectedFloatingBlock(out _, out _),
                    "An ordinary floating SVG picture was mistaken for a LaTeX Block.");
                plainFloating.Delete();

                document.SaveAs2(documentPath, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = word.Documents.Open(documentPath, ReadOnly: false);
                var reopened = document.Shapes[1];
                Assert(document.Shapes.Count == 1 && document.InlineShapes.Count == 0 &&
                       LaTeXBlockService.TryReadContract(reopened, out var reopenedMetadata, out var reopenedText) &&
                       reopenedMetadata.Id == beforeMetadata.Id && reopenedText == updatedSource &&
                       Math.Abs(reopenedMetadata.FrameWidthPt - requestedFrameWidthPt) < 0.01 &&
                       Math.Abs(reopenedMetadata.FrameHeightPt - requestedFrameHeightPt) < 0.01 &&
                       Math.Abs(reopened.Width - requestedFrameWidthPt) < 0.05 &&
                       Math.Abs(reopened.Height - requestedFrameHeightPt) < 0.05 &&
                       (int)reopened.WrapFormat.Type == 0 &&
                       (int)reopened.WrapFormat.Side == expectedSquareSide &&
                       Math.Abs(reopened.Left - expectedSquareLeft) < 0.05 &&
                       Math.Abs(reopened.Top - expectedSquareTop) < 0.05 &&
                       Math.Abs(reopened.WrapFormat.DistanceLeft - expectedSquareDistanceLeft) < 0.05 &&
                       Math.Abs(reopened.WrapFormat.DistanceRight - expectedSquareDistanceRight) < 0.05 &&
                       Math.Abs(reopened.WrapFormat.DistanceTop - expectedSquareDistanceTop) < 0.05 &&
                       Math.Abs(reopened.WrapFormat.DistanceBottom - expectedSquareDistanceBottom) < 0.05 &&
                       reopened.WrapFormat.AllowOverlap == expectedSquareAllowOverlap,
                    "A floating block did not persist its editable contract and position after DOCX reopen.");
                Console.WriteLine("Word: floating fixed block update and persistence passed.");
            }
            finally
            {
                if (document != null)
                {
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    Release(document);
                }
            }
        }

        private static void RunInlineSpacingSmoke(WordInterop.Application word, LaTeXBlockService service,
            string profile, string documentPath)
        {
            WordInterop.Document document = null;
            try
            {
                var render = service.RenderPreview("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, profile, 16);
                if (File.Exists(documentPath)) File.Delete(documentPath);
                document = word.Documents.Add();
                document.Range(0, 0).Text = "a b\rA xx B\rC,D";
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 16;

                // Replace "xx" in "A xx B" so both surrounding U+0020 characters can
                // remain ordinary Word spaces outside the U+2060 boundaries.
                document.Range(6, 8).Select();
                var shape = service.InsertRendered("$E=mc^2$", 360, LaTeXBlockLayoutMode.Auto, render);
                AssertExactSvgDrawingExtents(shape, render.SvgBytes,
                    "Initial inline formula");
                AssertInlineWordJoinerBoundary(shape, 2, "Initial inline formula");
                var leftSpace = document.Range(shape.Range.Start - 2, shape.Range.Start - 1);
                var rightSpace = document.Range(shape.Range.End + 1, shape.Range.End + 2);
                Assert(leftSpace.Text == " " && rightSpace.Text == " " &&
                       (double)leftSpace.Font.Scaling == 100 && (double)rightSpace.Font.Scaling == 100 &&
                       Math.Abs((double)leftSpace.Font.Spacing) < 0.001 &&
                       Math.Abs((double)rightSpace.Font.Spacing) < 0.001,
                    "U+2060 boundaries changed the adjacent ordinary spaces.");
                var svgWidth = LaTeXBlockService.ReadSvgWidthPt(render.SvgBytes);
                Assert(Math.Abs((double)shape.Width - svgWidth) < 0.35,
                    "The Word Joiner boundaries changed the SVG image's physical width.");

                var updatedRender = service.RenderPreview("$x^2$", 360,
                    LaTeXBlockLayoutMode.Auto, profile, 16);
                shape = service.UpdateRendered(shape, "$x^2$", 360,
                    LaTeXBlockLayoutMode.Auto, updatedRender);
                AssertExactSvgDrawingExtents(shape, updatedRender.SvgBytes,
                    "Updated inline formula");
                Assert(shape.AlternativeText == "$x^2$",
                    "Updating the inline formula lost the authoritative TeX source.");
                AssertInlineWordJoinerBoundary(shape, 2, "Updated inline formula");
                Assert(word.Selection.Start == shape.Range.End + 1,
                    "Update left the insertion caret before the trailing U+2060.");
                var typedAfterUpdateStart = word.Selection.Start;
                word.Selection.TypeText(" after");
                var typedAfterUpdate = document.Range(typedAfterUpdateStart, typedAfterUpdateStart + 6);
                Assert(typedAfterUpdate.Text == " after" &&
                       document.Range(shape.Range.End, shape.Range.End + 1).Text == WordJoiner,
                    "Typing after an inline formula split its trailing U+2060 boundary.");

                var punctuationInsertion = (document.Content.Text ?? string.Empty)
                    .IndexOf("C,D", StringComparison.Ordinal) + 1;
                Assert(punctuationInsertion > 0, "The no-space test text was not found.");
                document.Range(punctuationInsertion, punctuationInsertion).Select();
                var punctuationShape = service.InsertRendered("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                AssertInlineWordJoinerBoundary(punctuationShape, 4,
                    "No-space inline formula");
                Assert(document.Range(punctuationShape.Range.Start - 2,
                           punctuationShape.Range.Start - 1).Text == "C" &&
                       document.Range(punctuationShape.Range.End + 1,
                           punctuationShape.Range.End + 2).Text == ",",
                    "The no-space U+2060 boundaries altered the surrounding punctuation.");

                document.SaveAs2(documentPath, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(documentPath, ReadOnly: false);
                Assert(document.InlineShapes.Count == 2,
                    "The inline-spacing SVGs did not survive save and reopen.");
                var reopened = document.InlineShapes[1];
                AssertInlineWordJoinerBoundary(reopened, 4, "Reopened inline formula");
                AssertExactSvgDrawingExtents(reopened, updatedRender.SvgBytes,
                    "Reopened inline formula");
                RunInlineWordJoinerReuseSmoke(word, service, profile, render);
            }
            finally
            {
                if (document != null)
                {
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    Release(document);
                }
            }
        }

        private static void RunInlineWordJoinerReuseSmoke(WordInterop.Application word,
            LaTeXBlockService service, string profile, LaTeXBlockRender render)
        {
            WordInterop.Document document = null;
            try
            {
                document = word.Documents.Add();
                document.Content.Font.Name = "Times New Roman";
                document.Content.Font.Size = 16;

                // An existing pair is the normal update/migration case. The plugin must
                // reuse both characters rather than turning it into four joiners.
                document.Content.Text = "A " + WordJoiner + "xx" + WordJoiner + " B";
                document.Range(3, 5).Select();
                var shape = service.InsertRendered("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                AssertInlineWordJoinerBoundary(shape, 2, "Existing two-sided U+2060 boundary");
                for (var iteration = 0; iteration < 3; iteration++)
                {
                    shape = service.UpdateRendered(shape, "$x^2$", 360,
                        LaTeXBlockLayoutMode.Auto, render, false);
                    AssertInlineWordJoinerBoundary(shape, 2,
                        "Repeated U+2060-boundary update " + iteration);
                }

                // Each side is idempotent independently, so an interrupted or manually
                // edited boundary is repaired without duplicating the side that exists.
                document.Content.Text = "A " + WordJoiner + "xx B";
                document.Range(3, 5).Select();
                shape = service.InsertRendered("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                AssertInlineWordJoinerBoundary(shape, 2, "Existing left U+2060 boundary");

                document.Content.Text = "A xx" + WordJoiner + " B";
                document.Range(2, 4).Select();
                shape = service.InsertRendered("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                AssertInlineWordJoinerBoundary(shape, 2, "Existing right U+2060 boundary");

                // Two adjacent formulas share the joiner between them. This exercises
                // the normal typing-caret path as well as the no-duplicate invariant.
                document.Content.Text = "A B";
                document.Range(2, 2).Select();
                var first = service.InsertRendered("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                var second = service.InsertRendered("$x^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render);
                AssertInlineWordJoinerBoundary(first, 3, "First adjacent inline formula");
                AssertInlineWordJoinerBoundary(second, 3, "Second adjacent inline formula");
                first = service.UpdateRendered(first, "$C_{ij}$", 360,
                    LaTeXBlockLayoutMode.Auto, render, false);
                AssertInlineWordJoinerBoundary(first, 3, "Updated first adjacent inline formula");
                AssertInlineWordJoinerBoundary(second, 3, "Second adjacent inline formula after neighbor update");

                // Editing an auto formula into a fixed block removes its own boundary
                // characters. The shared middle joiner remains only while it is needed
                // by the neighboring auto formula.
                var fixedRender = service.RenderPreview("$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Fixed, profile, 16);
                first = service.UpdateRendered(first, "$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Fixed, fixedRender, false);
                Assert(document.Range(first.Range.Start - 1, first.Range.Start).Text != WordJoiner &&
                       document.Range(first.Range.End, first.Range.End + 1).Text == WordJoiner &&
                       CountOccurrences(document.Content.Text ?? string.Empty, WordJoiner) == 2,
                    "Changing the first auto formula to Fixed did not release only its unshared U+2060 boundary.");
                AssertInlineWordJoinerBoundary(second, 2,
                    "Neighboring auto formula after the first formula became Fixed");

                second = service.UpdateRendered(second, "$E=mc^2$", 360,
                    LaTeXBlockLayoutMode.Fixed, fixedRender, false);
                Assert(CountOccurrences(document.Content.Text ?? string.Empty, WordJoiner) == 0 &&
                       document.Range(second.Range.Start - 1, second.Range.Start).Text != WordJoiner &&
                       document.Range(second.Range.End, second.Range.End + 1).Text != WordJoiner,
                    "Changing the last adjacent auto formula to Fixed left U+2060 boundaries behind.");

                // The recovery-safe update path prepares a Fixed -> Auto transition
                // before deleting the old drawing. Verify it still creates exactly the
                // same two ownership boundaries as a fresh inline formula.
                second = service.UpdateRendered(second, "$x^2$", 360,
                    LaTeXBlockLayoutMode.Auto, render, false);
                AssertInlineWordJoinerBoundary(second, 2,
                    "Changing a fixed formula back to Auto");
                Assert(CountOccurrences(document.Content.Text ?? string.Empty, WordJoiner) == 2,
                    "Changing a fixed formula back to Auto did not create exactly two U+2060 boundaries.");
            }
            finally
            {
                if (document != null)
                {
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    Release(document);
                }
            }
        }

        private static void AssertInlineWordJoinerBoundary(WordInterop.InlineShape shape,
            int expectedJoinerCount, string context)
        {
            var document = shape.Range.Document;
            var left = document.Range(shape.Range.Start - 1, shape.Range.Start);
            var right = document.Range(shape.Range.End, shape.Range.End + 1);
            Assert(left.Text == WordJoiner && right.Text == WordJoiner,
                context + " is not directly bounded by one U+2060 on each side.");
            AssertZeroEffectExtent(shape, context);

            var content = document.Content.Text ?? string.Empty;
            Assert(CountOccurrences(content, WordJoiner) == expectedJoinerCount &&
                   content.IndexOf(WordJoiner + WordJoiner, StringComparison.Ordinal) < 0,
                context + " duplicated a U+2060 boundary character.");
        }

        private static void AssertZeroEffectExtent(WordInterop.InlineShape shape, string context)
        {
            var effect = Regex.Match(shape.Range.WordOpenXML,
                "<wp:effectExtent\\b(?=[^>]*\\bl=\"(?<left>-?[0-9]+)\")" +
                "(?=[^>]*\\bt=\"(?<top>-?[0-9]+)\")" +
                "(?=[^>]*\\br=\"(?<right>-?[0-9]+)\")" +
                "(?=[^>]*\\bb=\"(?<bottom>-?[0-9]+)\")[^>]*/>",
                RegexOptions.CultureInvariant);
            Assert(effect.Success && effect.Groups["left"].Value == "0" &&
                   effect.Groups["top"].Value == "0" && effect.Groups["right"].Value == "0" &&
                   effect.Groups["bottom"].Value == "0",
                context + " retained a non-zero wp:effectExtent.");
        }

        private static int CountOccurrences(string text, string value)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value)) return 0;
            var count = 0;
            for (var index = 0; index <= text.Length - value.Length;)
            {
                var next = text.IndexOf(value, index, StringComparison.Ordinal);
                if (next < 0) break;
                count++;
                index = next + value.Length;
            }
            return count;
        }

        private static void AssertExactSvgDrawingExtents(WordInterop.InlineShape shape,
            byte[] svgBytes, string context)
        {
            const double emusPerPoint = 12700.0;
            var expectedWidthEmu = checked((long)Math.Round(
                LaTeXBlockService.ReadSvgWidthPt(svgBytes) * emusPerPoint,
                MidpointRounding.AwayFromZero));
            var expectedHeightEmu = checked((long)Math.Round(
                LaTeXBlockService.ReadSvgHeightPt(svgBytes) * emusPerPoint,
                MidpointRounding.AwayFromZero));
            var flatOpc = shape.Range.WordOpenXML;

            var inlineExtent = Regex.Match(flatOpc,
                "<wp:extent\\b(?=[^>]*\\bcx=\"(?<cx>[-+0-9]+)\")" +
                "(?=[^>]*\\bcy=\"(?<cy>[-+0-9]+)\")[^>]*/>",
                RegexOptions.CultureInvariant);
            Assert(inlineExtent.Success,
                context + " has no readable wp:extent in its inline DrawingML.");

            // Deliberately scope a:ext to pic:spPr/a:xfrm. The SVG blip extension
            // list also contains an a:ext element, but that element has a uri and is
            // not the picture's physical transform extent.
            var pictureProperties = Regex.Match(flatOpc,
                "<pic:spPr\\b[^>]*>.*?</pic:spPr>",
                RegexOptions.CultureInvariant | RegexOptions.Singleline);
            Assert(pictureProperties.Success,
                context + " has no pic:spPr drawing properties.");
            var transform = Regex.Match(pictureProperties.Value,
                "<a:xfrm\\b[^>]*>.*?</a:xfrm>",
                RegexOptions.CultureInvariant | RegexOptions.Singleline);
            Assert(transform.Success,
                context + " has no pic:spPr/a:xfrm picture transform.");
            var transformExtent = Regex.Match(transform.Value,
                "<a:ext\\b(?=[^>]*\\bcx=\"(?<cx>[-+0-9]+)\")" +
                "(?=[^>]*\\bcy=\"(?<cy>[-+0-9]+)\")[^>]*/>",
                RegexOptions.CultureInvariant);
            Assert(transformExtent.Success,
                context + " has no numeric pic:spPr/a:xfrm/a:ext.");

            var inlineWidthEmu = long.Parse(inlineExtent.Groups["cx"].Value);
            var inlineHeightEmu = long.Parse(inlineExtent.Groups["cy"].Value);
            var transformWidthEmu = long.Parse(transformExtent.Groups["cx"].Value);
            var transformHeightEmu = long.Parse(transformExtent.Groups["cy"].Value);
            Console.WriteLine(context + " SVG extents: expected=" + expectedWidthEmu + "x" +
                expectedHeightEmu + ", wp=" + inlineWidthEmu + "x" + inlineHeightEmu +
                ", transform=" + transformWidthEmu + "x" + transformHeightEmu);
            Assert(inlineWidthEmu == expectedWidthEmu && inlineHeightEmu == expectedHeightEmu &&
                   transformWidthEmu == expectedWidthEmu && transformHeightEmu == expectedHeightEmu,
                context + " did not preserve the SVG's exact physical EMU dimensions " +
                "(expected " + expectedWidthEmu + "x" + expectedHeightEmu +
                ", wp:extent " + inlineWidthEmu + "x" + inlineHeightEmu +
                ", pic transform " + transformWidthEmu + "x" + transformHeightEmu + ").");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void RunShutdownProbe(StemTeXBackend renderer, string profile)
        {
            renderer.RenderLatestAsync(profile, "\\loop\\iftrue\\repeat", 360, true, 11);
            WaitFor(() => renderer.Status.StartsWith("rendering:", StringComparison.Ordinal),
                5000, "The shutdown probe did not enter a render.");
            var shutdownTimer = Stopwatch.StartNew();
            renderer.Dispose();
            var disposeMilliseconds = shutdownTimer.ElapsedMilliseconds;
            Assert(disposeMilliseconds < 250,
                "StemTeX shutdown blocked the Office UI thread for " + disposeMilliseconds + " ms.");
            Assert(renderer.WaitForStopForTest(2000),
                "StemTeX background worker did not actually stop after shutdown cancellation.");
            WaitFor(() => !renderer.HasOwnedWorkerHostForTest, 2000,
                "StemTeX left an owned worker-host process after render shutdown.");
            Console.WriteLine("StemTeX: active-render shutdown returned in " +
                disposeMilliseconds + " ms.");
        }

        private static void RunStartupShutdownProbe()
        {
            var startupBackend = new StemTeXBackend();
            startupBackend.SwitchProfile(startupBackend.DefaultAvailableProfile);
            WaitFor(() => startupBackend.WorkerHasActiveItemForTest &&
                    startupBackend.Status.StartsWith("warming:", StringComparison.Ordinal) &&
                    startupBackend.HasOwnedWorkerHostForTest,
                10000, "The startup-shutdown probe did not observe a live worker during renderer initialization.");
            var shutdownTimer = Stopwatch.StartNew();
            startupBackend.Dispose();
            var disposeMilliseconds = shutdownTimer.ElapsedMilliseconds;
            Assert(disposeMilliseconds < 250,
                "Shutdown during renderer initialization blocked for " + disposeMilliseconds + " ms.");
            WaitFor(() => !startupBackend.HasOwnedWorkerHostForTest, 2000,
                "StemTeX left an owned worker-host process after initialization shutdown.");
            Console.WriteLine("StemTeX: initialization shutdown returned in " +
                disposeMilliseconds + " ms.");
        }

        private static void RunStartupShutdownProbeInIsolatedHost()
        {
            // stemtex_renderer_create does not publish its renderer pointer until native
            // initialization completes. If shutdown wins that race, the backend deliberately
            // abandons its background initializer after returning immediately and terminating
            // its owned helper tree; Office process exit reclaims the blocked native call. Run
            // that process-lifetime contract in a child host so its long native timeout and
            // shutdown reaper cannot interfere with the main smoke-test backend.
            var executable = Process.GetCurrentProcess().MainModule.FileName;
            var startInfo = new ProcessStartInfo(executable)
            {
                UseShellExecute = false
            };
            startInfo.EnvironmentVariables[StartupShutdownProbeChild] = "1";
            using (var child = Process.Start(startInfo))
            {
                Assert(child != null, "The isolated renderer-initialization shutdown probe did not start.");
                if (!child.WaitForExit(15000))
                {
                    try { child.Kill(); } catch { }
                    throw new InvalidOperationException(
                        "The isolated renderer-initialization shutdown probe did not exit within 15 seconds.");
                }
                Assert(child.ExitCode == 0,
                    "The isolated renderer-initialization shutdown probe exited with code " + child.ExitCode + ".");
            }
        }

        private static void WaitFor(Func<bool> condition, int timeoutMs, string message)
        {
            var timer = Stopwatch.StartNew();
            while (timer.ElapsedMilliseconds < timeoutMs)
            {
                System.Windows.Forms.Application.DoEvents();
                if (condition()) return;
                Thread.Sleep(20);
            }
            throw new InvalidOperationException(message);
        }

        private static string EquationNumberText(WordInterop.Field field)
        {
            Assert(LaTeXBlockService.IsEquationSequenceField(field),
                "The numbered-equation line does not contain a LaTeX SEQ field.");
            Assert((field.Code.Text ?? string.Empty).IndexOf("\\* ARABIC", StringComparison.OrdinalIgnoreCase) >= 0,
                "The equation SEQ field does not request Arabic numbering.");
            return (field.Result.Text ?? string.Empty).Trim();
        }

        private static WordInterop.Field FindEquationReferenceField(WordInterop.Document document,
            string bookmarkName)
        {
            for (var index = 1; index <= document.Fields.Count; index++)
            {
                var field = document.Fields[index];
                if (LaTeXBlockService.IsEquationReferenceField(field) &&
                    (field.Code.Text ?? string.Empty).IndexOf(bookmarkName,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return field;
            }
            throw new InvalidOperationException("The expected equation REF field was not found.");
        }

        private static bool HasCaptionLabel(WordInterop.Application application, string name)
        {
            var labels = application.CaptionLabels;
            for (var index = 1; index <= labels.Count; index++)
            {
                object item = index;
                WordInterop.CaptionLabel label = null;
                try
                {
                    label = labels.get_Item(ref item);
                    if (string.Equals(label.Name, name, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                finally { Release(label); }
            }
            return false;
        }

        private static void AssertEquationTabStops(WordInterop.Paragraph paragraph)
        {
            var tabs = paragraph.Range.ParagraphFormat.TabStops;
            Assert(tabs.Count == 2, "The numbered-equation paragraph does not have exactly two custom tab stops.");
            WordInterop.TabStop center = null;
            WordInterop.TabStop right = null;
            for (var index = 1; index <= tabs.Count; index++)
            {
                var tab = tabs[index];
                if (tab.Alignment == WordInterop.WdTabAlignment.wdAlignTabCenter) center = tab;
                if (tab.Alignment == WordInterop.WdTabAlignment.wdAlignTabRight) right = tab;
            }
            Assert(center != null && right != null,
                "The numbered-equation paragraph lost its center or right tab stop.");

            var page = paragraph.Range.Sections[1].PageSetup;
            var columnWidth = (double)page.PageWidth - page.LeftMargin - page.RightMargin;
            var columns = page.TextColumns;
            if (columns.Count > 1)
                columnWidth = (columnWidth - columns.Spacing * (columns.Count - 1)) / columns.Count;
            Assert(Math.Abs(center.Position - columnWidth / 2) < 0.1 &&
                   Math.Abs(right.Position - columnWidth) < 0.1,
                "The equation tab stops are not centered and right-aligned within the text column.");
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
