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
            try
            {
                Directory.CreateDirectory(artifactDirectory);
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
                Console.WriteLine("StemTeX: testing fixed-width and inline auto-width rendering...");
                var svg = renderer.RenderSvg(profile, source, 360, false);
                var prefix = Encoding.UTF8.GetString(svg.Bytes, 0, Math.Min(svg.Bytes.Length, 512));
                Assert(prefix.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0, "StemTeX did not return SVG bytes.");
                Assert(Encoding.UTF8.GetString(svg.Bytes).IndexOf("latexblocks-baseline", StringComparison.Ordinal) < 0,
                    "The temporary baseline marker leaked into the embedded SVG.");
                Assert(svg.DepthPt > 0, "StemTeX inline baseline marker did not produce a positive TeX depth.");
                var autoSvg = renderer.RenderSvg(profile, source, 360, true);
                var autoSvg11 = renderer.RenderSvg(profile, source, 360, true, 11);
                var rhoGhSvg = renderer.RenderSvg(profile, "$\\rho gh$", 360, true, 11);
                Assert(rhoGhSvg.Bytes.Length > 0 && rhoGhSvg.DepthPt > 0,
                    "A valid rho-gh inline formula did not render.");
                var autoText = Encoding.UTF8.GetString(autoSvg.Bytes);
                Assert(autoText.IndexOf("latexblocks-start", StringComparison.Ordinal) < 0 &&
                    autoText.IndexOf("latexblocks-end", StringComparison.Ordinal) < 0,
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
                document.Range(document.Content.End - 1, document.Content.End - 1).Select();
                var service = new LaTeXBlockService(word, renderer);
                using (var editor = new LaTeXBlockEditorForm(service, "$x_1$", 360, LaTeXBlockLayoutMode.Auto,
                    profile, selected => { }, false))
                {
                    editor.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    editor.Location = new System.Drawing.Point(100, 100);
                    editor.Show();
                    WaitFor(() => editor.PreviewIsCurrent, 10000, "The editor did not produce its initial live preview.");
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
                Assert(inserted.Range.Font.Position == -(int)Math.Round(firstMetadata.DepthPt, MidpointRounding.AwayFromZero),
                    "Word baseline compensation does not equal the rounded TeX depth.");
                Assert(Regex.IsMatch(inserted.Range.WordOpenXML,
                    "<wp:effectExtent\\b(?=[^>]*\\bb=\"0\")[^>]*/>"),
                    "Word's host-only bottom effect extent was not removed from the inline SVG.");
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
                Assert(Regex.IsMatch(reopened.Range.WordOpenXML,
                    "<wp:effectExtent\\b(?=[^>]*\\bb=\"0\")[^>]*/>"),
                    "The zero bottom effect extent did not survive save and reopen.");

                var updated = service.UpdateBlock(reopened, updatedSource, 420, LaTeXBlockLayoutMode.Auto, profile, 14);
                Assert(document.InlineShapes.Count == 2, "Update changed the document's SVG count.");
                Assert(updated.AlternativeText == updatedSource, "Update did not replace the authoritative TeX source.");
                Assert(LaTeXBlockMetadata.TryParse(updated.Title, out var updatedMetadata) && updatedMetadata.Id == stableId,
                    "Update changed the block identity.");
                Assert(Math.Abs(updatedMetadata.WidthPt - 420) < 0.001, "Update did not persist the new typesetting width.");
                Assert(Math.Abs(updatedMetadata.FontSizePt - 14) < 0.001,
                    "Update did not rerender and persist the requested TeX font size.");
                document.Save();
                document.Close(WordInterop.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;

                document = word.Documents.Open(documentPath, ReadOnly: true);
                Assert(document.InlineShapes.Count == 2 && document.InlineShapes[1].AlternativeText == updatedSource,
                    "Updated block did not survive the second reopen.");
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
                    LaTeXBlockLayoutMode.Auto, profile, 11, false);
                Assert(LaTeXBlockMetadata.TryParse(updatedNumbered.Title, out var updatedNumberedMetadata) &&
                    updatedNumberedMetadata.Role == LaTeXBlockRole.NumberedEquation,
                    "Editing a numbered equation discarded its numbered role.");
                Assert(document.Tables.Count == 0 && EquationNumberText(document.Fields[1]) == "1" &&
                    document.Bookmarks.Exists(LaTeXBlockService.EquationBookmarkName(updatedNumberedMetadata.Id)),
                    "Editing the formula disturbed its Word-native field or bookmark.");

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
            Assert(shutdownTimer.ElapsedMilliseconds < 250,
                "StemTeX shutdown blocked the Office UI thread for " + shutdownTimer.ElapsedMilliseconds + " ms.");
            Assert(renderer.WaitForStopForTest(2000),
                "StemTeX background worker did not actually stop after shutdown cancellation.");
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
