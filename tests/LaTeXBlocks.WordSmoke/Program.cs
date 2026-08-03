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
                Console.WriteLine("StemTeX: warming the default profile...");
                renderer.WarmUp(profile);
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
                Assert(Math.Abs(LaTeXBlockWidthPolicy.ResolveDefaultFixedWidth() - 360) < 0.001 &&
                    Math.Abs(LaTeXBlockWidthPolicy.WidthStepPt - 0.5) < 0.001 &&
                    LaTeXBlockWidthPolicy.IsValidWidth(30) &&
                    LaTeXBlockWidthPolicy.IsValidWidth(450) &&
                    !LaTeXBlockWidthPolicy.IsValidWidth(29.9) &&
                    !LaTeXBlockWidthPolicy.IsValidWidth(450.1),
                    "The fixed-block width policy does not match StemTeX GUI's point range.");
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
                    var expectedPointSvgWidth = requestedFixedWidthPt * texPointToWordPoint + 2;
                    Console.WriteLine("Fixed width editor: requested=" +
                        requestedFixedWidthPt.ToString("0.###") + "pt, SVG=" +
                        renderedPointWidth.ToString("0.###") + "pt");
                    Assert(Math.Abs(renderedPointWidth - expectedPointSvgWidth) < 0.02,
                        "The exact point width was not passed to StemTeX unchanged.");
                    editor.Close();
                }
                Console.WriteLine("Word: exact point-width editor passed.");
                RunInlineSpacingSmoke(word, service, profile, spacingDocumentPath);
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
                var fixedBlock = service.InsertBlock("\\[x^2\\]", 180, LaTeXBlockLayoutMode.Fixed, alternateProfile);
                Assert(LaTeXBlockMetadata.TryParse(fixedBlock.Title, out var fixedMetadata) &&
                    fixedMetadata.Mode == LaTeXBlockLayoutMode.Fixed, "Fixed-width block mode was not persisted.");
                Assert(fixedBlock.Width > 150, "Fixed-width block lost its requested canvas width.");
                Assert(document.Range(fixedBlock.Range.Start - 1, fixedBlock.Range.Start).Text != WordJoiner &&
                       document.Range(fixedBlock.Range.End, fixedBlock.Range.End + 1).Text != WordJoiner,
                    "A fixed-width block unexpectedly received inline U+2060 boundaries.");
                Assert(document.InlineShapes.Count == 2, "The two insertion modes did not produce two InlineShapes.");
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
                var firstParagraphText = document.Paragraphs[1].Range.Text ?? string.Empty;
                Assert(firstParagraphText.StartsWith("Alpha\v\t", StringComparison.Ordinal) &&
                    firstParagraphText.IndexOf("\t(1)\v beta", StringComparison.Ordinal) >= 0,
                    "The numbered equation does not use the expected manual-break/tab scaffold.");
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
                var adjacentEquation = document.InlineShapes[1];
                LaTeXBlockService.NumberedEquationLineRange(adjacentEquation).Delete();
                Assert(service.UpdateEquationNumbers(document) == 1 &&
                    EquationNumberText(document.Fields[1]) == "1" &&
                    !document.Bookmarks.Exists(secondBookmarkName) &&
                    document.Bookmarks.Exists(thirdBookmarkName) &&
                    document.Bookmarks[thirdBookmarkName].Range.Text == "1" &&
                    (document.Paragraphs[1].Range.Text ?? string.Empty).IndexOf("\v\t", StringComparison.Ordinal) >= 0,
                    "Deleting an equation adjacent to another display removed their shared visual-line boundary.");

                RunShutdownProbe(renderer, profile);

                Console.WriteLine("LaTeX Blocks smoke test passed.");
                Console.WriteLine("StemTeX: " + renderer.StemTeXHome);
                Console.WriteLine("Verified: SVG insertion, metadata, update, Word-native equation numbering, and DOCX persistence.");
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
                document.Repaginate();
                var ordinarySpaceWidth = MeasureHorizontalAdvance(document.Range(1, 2));
                Assert(ordinarySpaceWidth > 3 && ordinarySpaceWidth < 5,
                    "The isolated Word document did not produce an ordinary 16 pt space.");

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
                var leftInlineWidth = MeasureHorizontalAdvance(leftSpace);
                var rightInlineWidth = MeasureHorizontalAdvance(rightSpace);
                Console.WriteLine("Inline U+2060 spacing: ordinary=" + ordinarySpaceWidth.ToString("0.###") +
                    ", actual=" + leftInlineWidth.ToString("0.###") + "/" +
                    rightInlineWidth.ToString("0.###"));
                Assert(Math.Abs(leftInlineWidth - ordinarySpaceWidth) < 0.16 &&
                       Math.Abs(rightInlineWidth - ordinarySpaceWidth) < 0.16,
                    "A Word Joiner boundary did not restore the Times New Roman word spaces.");
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

        private static double MeasureHorizontalAdvance(WordInterop.Range range)
        {
            var start = range.Duplicate;
            start.Collapse(WordInterop.WdCollapseDirection.wdCollapseStart);
            var end = range.Duplicate;
            end.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
            var startX = Convert.ToDouble(start.get_Information(
                WordInterop.WdInformation.wdHorizontalPositionRelativeToPage));
            var endX = Convert.ToDouble(end.get_Information(
                WordInterop.WdInformation.wdHorizontalPositionRelativeToPage));
            return endX - startX;
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
