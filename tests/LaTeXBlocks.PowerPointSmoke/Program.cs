using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using LaTeXBlocks.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace LaTeXBlocks.PowerPointSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            try
            {
                if (Process.GetProcessesByName("POWERPNT").Length != 0)
                    throw new InvalidOperationException(
                        "Close PowerPoint before running the PowerPoint smoke test.");
                Run();
                Console.WriteLine("LaTeX Blocks PowerPoint smoke test passed.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static void Run()
        {
            var ribbonXml = PowerPointRibbonContract.BuildCustomUi();
            var ribbonDocument = new System.Xml.XmlDocument();
            ribbonDocument.LoadXml(ribbonXml);
            Assert(ribbonXml.IndexOf(PowerPointRibbonContract.FontSizeControlId,
                       StringComparison.Ordinal) >= 0,
                "The PowerPoint Ribbon does not expose its TeX font-size control.");
            Assert(PowerPointRibbonContract.TryParseFontSize("27.5", out var parsedSize) &&
                   Math.Abs(parsedSize - 27.5) < 0.001 &&
                   !PowerPointRibbonContract.TryParseFontSize("0", out _) &&
                   !PowerPointRibbonContract.TryParseFontSize("not-a-size", out _),
                "The PowerPoint TeX font-size control did not validate point sizes.");
            Assert(ribbonXml.IndexOf(PowerPointRibbonContract.LayoutWidthControlId,
                       StringComparison.Ordinal) >= 0 &&
                   PowerPointRibbonContract.TryParseLayoutWidthPt("360.5",
                       out var parsedWidthPt) &&
                   Math.Abs(parsedWidthPt - 360.5) < 0.001 &&
                   !PowerPointRibbonContract.TryParseLayoutWidthPt("29.9", out _) &&
                   !PowerPointRibbonContract.TryParseLayoutWidthPt("450.1", out _) &&
                   !PowerPointRibbonContract.TryParseLayoutWidthPt("wide", out _) &&
                   Math.Abs(BlockLayoutWidthPolicy.DefaultPt - 360) < 0.001 &&
                   Math.Abs(BlockLayoutWidthPolicy.StepPt - 0.5) < 0.001,
                "The PowerPoint Ribbon does not expose StemTeX's point-width control.");

            StemTeXBackend backend = null;
            PowerPointInterop.Application application = null;
            PowerPointInterop.Presentation presentation = null;
            PowerPointInterop.Slide slide = null;
            PowerPointInterop.Slide secondSlide = null;
            PowerPointInterop.EApplication_AfterShapeSizeChangeEventHandler sizeChangeHandler = null;
            var sizeChangeEvents = 0;
            var artifacts = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "..", "..", "..", "artifacts"));
            Directory.CreateDirectory(artifacts);
            var documentPath = Path.Combine(artifacts, "latex-blocks-powerpoint-smoke.pptx");
            try
            {
                backend = new StemTeXBackend();
                var profile = backend.DefaultAvailableProfile;
                backend.SwitchProfile(profile);
                var durableFirst = backend.RenderQueuedAsync(profile, "$x_1$", 180,
                    false, 10);
                var stalePreview = backend.RenderLatestAsync(profile, "$x_2$", 180,
                    false, 10);
                var latestPreview = backend.RenderLatestAsync(profile, "$x_3$", 180,
                    false, 10);
                var durableSecond = backend.RenderQueuedAsync(profile, "$x_4$", 180,
                    false, 10);
                Assert(durableFirst.GetAwaiter().GetResult().Bytes.Length > 0 &&
                       durableSecond.GetAwaiter().GetResult().Bytes.Length > 0 &&
                       latestPreview.GetAwaiter().GetResult().Bytes.Length > 0,
                    "A live preview discarded durable document-format rendering work.");
                try
                {
                    stalePreview.GetAwaiter().GetResult();
                    throw new InvalidOperationException(
                        "The superseded live preview was not canceled.");
                }
                catch (System.Threading.Tasks.TaskCanceledException) { }
                application = new PowerPointInterop.Application();
                application.Visible = Office.MsoTriState.msoTrue;
                sizeChangeHandler = changedShape => sizeChangeEvents++;
                application.AfterShapeSizeChange += sizeChangeHandler;
                presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);
                slide = presentation.Slides.Add(1,
                    PowerPointInterop.PpSlideLayout.ppLayoutBlank);

                var service = new PowerPointBlockService(application, backend);
                var slideWidth = service.GetActiveSlideWidthPt();
                Assert(Math.Abs(slideWidth - presentation.PageSetup.SlideWidth) < 0.01,
                    "PowerPoint did not expose the active slide width in points.");
                var textBox = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    72, 54, 288, 72);
                textBox.TextFrame.TextRange.Text = "Ordinary PowerPoint text";
                textBox.TextFrame.TextRange.Font.Name = "Times New Roman";
                textBox.TextFrame.TextRange.Font.Size = 27.5f;
                textBox.TextFrame.TextRange.Select();

                var inheritedSize = PowerPointBlockService.ResolveFontSize(application, 18);
                Assert(Math.Abs(inheritedSize - 27.5) < 0.01,
                    "A PowerPoint text selection did not supply its ordinary font size.");
                Assert(Math.Abs(service.ResolveInitialWidth(360) - 360) < 0.01,
                    "A new PowerPoint block did not use StemTeX's standard initial width.");

                const string editorSource = "$E=mc^2$";
                using (var editor = new LaTeXBlockEditorForm(service, editorSource, 288,
                           inheritedSize, profile, backend.SwitchProfile, false))
                {
                    editor.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    editor.Location = new System.Drawing.Point(100, 100);
                    editor.Show();
                    WaitFor(() => editor.PreviewIsCurrent, 30000,
                        () => "The PowerPoint editor did not produce its live preview.");
                    var initialPreview = Convert.ToBase64String(editor.CurrentRender.SvgBytes);
                    editor.SetWidthPtForTest(234.5);
                    WaitFor(() => editor.PreviewIsCurrent &&
                        Math.Abs(editor.WidthPt - 234.5) < 0.01 &&
                        Convert.ToBase64String(editor.CurrentRender.SvgBytes) != initialPreview,
                        30000, () => "Changing the PowerPoint editor width did not reflow its live preview.");
                    var renderedBeforeClose = editor.CurrentRender;
                    var insertButton = editor.AcceptButton as System.Windows.Forms.Button;
                    Assert(insertButton != null && insertButton.Enabled,
                        "The PowerPoint editor did not enable Insert for its current preview.");
                    insertButton.PerformClick();
                    System.Windows.Forms.Application.DoEvents();
                    Assert(editor.DialogResult == System.Windows.Forms.DialogResult.OK &&
                           ReferenceEquals(editor.AcceptedRender, renderedBeforeClose) &&
                           editor.AcceptedSource == editorSource &&
                           Math.Abs(editor.AcceptedWidthPt - 234.5) < 0.01,
                        "Closing the PowerPoint editor invalidated its accepted preview.");
                }

                const string source = "PowerPoint block with $E=mc^2$.";
                var render = service.RenderPreviewAsync(source, 288, profile,
                    inheritedSize).GetAwaiter().GetResult();
                var smallerRender = service.RenderPreviewAsync(source, 288, profile, 10)
                    .GetAwaiter().GetResult();
                Assert(PowerPointBlockService.ReadSvgHeightPt(render.SvgBytes) >
                       PowerPointBlockService.ReadSvgHeightPt(smallerRender.SvgBytes) + 1 &&
                       Math.Abs(PowerPointBlockService.ReadSvgWidthPt(render.SvgBytes) -
                                PowerPointBlockService.ReadSvgWidthPt(smallerRender.SvgBytes)) < 0.02,
                    "PowerPoint ordinary-text size was recorded but not used as the real StemTeX design size.");
                var block = service.InsertRendered(source, 288, render);
                Assert(PowerPointBlockService.TryReadContract(block, out var metadata,
                           out var storedSource),
                    "The inserted PowerPoint shape is not recognized as a LaTeX Block.");
                Assert(storedSource == source && metadata.Role == LaTeXBlockRole.Content &&
                       metadata.Mode == LaTeXBlockLayoutMode.Fixed &&
                       Math.Abs(metadata.FontSizePt - 27.5) < 0.01,
                    "The inserted PowerPoint block lost its source, role, layout, or TeX size.");
                Assert(string.Equals(block.Tags[PowerPointBlockService.KindTag],
                           PowerPointBlockService.KindValue, StringComparison.OrdinalIgnoreCase),
                    "The inserted PowerPoint block lost its explicit identity tag.");
                Assert(Math.Abs(block.Width - PowerPointBlockService.ReadSvgWidthPt(
                           render.SvgBytes)) < 0.02 &&
                       Math.Abs(block.Height - PowerPointBlockService.ReadSvgHeightPt(
                           render.SvgBytes)) < 0.02,
                    "PowerPoint did not preserve the SVG's physical point dimensions.");
                Assert(block.LockAspectRatio == Office.MsoTriState.msoFalse &&
                       Math.Abs(PowerPointBlockService.ReadPositiveTag(block,
                            PowerPointBlockService.VisualScaleTag, 0) - 1) < 0.001,
                    "The inserted block did not expose separate layout-width and visual-scale gestures.");
                var duplicatedRange = block.Duplicate();
                var duplicatedBlock = duplicatedRange[1];
                Assert(PowerPointBlockService.TryReadContract(duplicatedBlock,
                           out var duplicatedMetadata, out _) &&
                       duplicatedMetadata.Id == metadata.Id &&
                       !PowerPointBlockService.GetShapeKey(duplicatedBlock).Equals(
                           PowerPointBlockService.GetShapeKey(block)),
                    "Copied blocks with duplicate semantic IDs do not have distinct physical locators.");
                duplicatedBlock.Delete();
                Release(duplicatedBlock);
                Release(duplicatedRange);

                // This section tests the resize classifier directly. An installed copy
                // of the add-in can be auto-loaded into this COM PowerPoint instance and
                // otherwise consume AfterShapeSizeChange first, restoring the old SVG
                // before the smoke test can inspect the drag geometry. Temporarily hide
                // only this test shape from the production event handler; the installed-
                // host integration test covers the complete asynchronous reflow path.
                block.Tags.Delete(PowerPointBlockService.KindTag);
                try
                {
                    var originalBlockWidth = block.Width;
                    var originalBlockLeft = block.Left;
                    var eventsBeforeResize = sizeChangeEvents;
                    block.Width = originalBlockWidth * 0.8f;
                    WaitFor(() => sizeChangeEvents > eventsBeforeResize, 2000,
                        () => "PowerPoint did not raise AfterShapeSizeChange after a block resize.");
                    var horizontalResize = PowerPointBlockService.ClassifyResize(block, metadata);
                    Assert(horizontalResize.Kind == PowerPointResizeKind.LayoutWidth &&
                           Math.Abs(horizontalResize.LayoutWidthPt - metadata.WidthPt * 0.8) < 0.05,
                        "A horizontal PowerPoint handle resize was not classified as TeX reflow.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                    Assert(Math.Abs(block.Width - originalBlockWidth) < 0.02 &&
                           Math.Abs(block.Left - originalBlockLeft) < 0.02,
                        "A failed horizontal reflow could not restore valid SVG geometry.");
                    block.Width *= 1.1f;
                    block.Height *= 1.1f;
                    var scaleResize = PowerPointBlockService.ClassifyResize(block, metadata);
                    Assert(scaleResize.Kind == PowerPointResizeKind.VisualScale &&
                           Math.Abs(scaleResize.VisualScale - 1.1) < 0.01,
                        "A proportional PowerPoint resize was not classified as visual scale.");
                    PowerPointBlockService.NormalizeVisualScale(block, scaleResize.VisualScale);
                    Assert(Math.Abs(block.Width / block.Height -
                               PowerPointBlockService.ReadSvgWidthPt(render.SvgBytes) /
                               PowerPointBlockService.ReadSvgHeightPt(render.SvgBytes)) < 0.01,
                        "Visual-scale normalization distorted the SVG.");
                    Assert(Math.Abs(PowerPointBlockService.LayoutWidthForVisibleWidth(block,
                               metadata, block.Width) - metadata.WidthPt) < 0.05,
                        "Slide-relative visible width was confused with the block's independent visual scale.");
                    PowerPointBlockService.NormalizeVisualScale(block, 1.0);
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }

                var ordinarySvg = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                    18, 18, 18, 18);
                ordinarySvg.AlternativeText = source;
                ordinarySvg.Title = metadata.ToString();
                Assert(!PowerPointBlockService.TryReadContract(ordinarySvg, out _, out _),
                    "An ordinary shape without the explicit LaTeX Block tag was misidentified.");

                secondSlide = presentation.Slides.Add(2,
                    PowerPointInterop.PpSlideLayout.ppLayoutBlank);
                application.ActiveWindow.View.GotoSlide(2);
                var resizedRender = service.RenderPreviewAsync(source, 288, profile, 19)
                    .GetAwaiter().GetResult();
                block = service.UpdateRendered(block, source, 288, resizedRender, false);
                Assert(PowerPointBlockService.TryReadContract(block, out var resizedMetadata,
                           out var resizedSource) && resizedMetadata.Id == metadata.Id &&
                       resizedSource == source && Math.Abs(resizedMetadata.FontSizePt - 19) < 0.01 &&
                       secondSlide.Shapes.Count == 0,
                    "A PowerPoint block did not rerender at its requested TeX font size on its owning slide.");
                application.ActiveWindow.View.GotoSlide(1);

                var back = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                    10, 10, 20, 20);
                back.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                var front = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                    400, 10, 20, 20);
                front.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                block.Rotation = 17.5f;
                PowerPointBlockService.NormalizeVisualScale(block, 1.15);
                var expectedLeft = block.Left;
                var expectedTop = block.Top;
                var expectedWidth = block.Width;
                var expectedScale = PowerPointBlockService.ReadPositiveTag(block,
                    PowerPointBlockService.VisualScaleTag, 1);
                var expectedRotation = block.Rotation;
                var expectedZ = block.ZOrderPosition;
                var originalId = metadata.Id;

                const string updatedSource = "Updated block: $\\int_0^1 x^2\\,dx=1/3$.";
                var updatedRender = service.RenderPreviewAsync(updatedSource, 288, profile, 19)
                    .GetAwaiter().GetResult();
                var updated = service.UpdateRendered(block, updatedSource, 288, updatedRender);
                Assert(PowerPointBlockService.TryReadContract(updated, out var updatedMetadata,
                           out var updatedStoredSource) &&
                       updatedMetadata.Id == originalId && updatedStoredSource == updatedSource &&
                       Math.Abs(updatedMetadata.FontSizePt - 19) < 0.01,
                    "PowerPoint edit did not preserve identity while replacing source and TeX size.");
                Console.WriteLine("PowerPoint replacement geometry: expected=" +
                    expectedLeft.ToString("0.###") + "/" + expectedTop.ToString("0.###") + "/" +
                    expectedWidth.ToString("0.###") + "/" + expectedRotation.ToString("0.###") +
                    "/z" + expectedZ + ", actual=" + updated.Left.ToString("0.###") + "/" +
                    updated.Top.ToString("0.###") + "/" + updated.Width.ToString("0.###") + "/" +
                    updated.Rotation.ToString("0.###") + "/z" + updated.ZOrderPosition);
                Assert(Math.Abs(updated.Left - expectedLeft) < 0.02 &&
                       Math.Abs(updated.Top - expectedTop) < 0.02 &&
                       Math.Abs(updated.Width - expectedWidth) < 0.03 &&
                       Math.Abs(updated.Rotation - expectedRotation) < 0.02 &&
                       updated.ZOrderPosition == expectedZ,
                    "PowerPoint edit did not preserve position, width scale, rotation, and z-order.");
                var expectedHeight = PowerPointBlockService.ReadSvgHeightPt(updatedRender.SvgBytes) *
                                     expectedScale;
                Assert(Math.Abs(updated.Height - expectedHeight) < 0.04,
                    "PowerPoint edit distorted the new SVG instead of preserving the user's scale.");

                if (File.Exists(documentPath)) File.Delete(documentPath);
                presentation.SaveAs(documentPath,
                    PowerPointInterop.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    Office.MsoTriState.msoFalse);
                Release(secondSlide); secondSlide = null;
                presentation.Close();
                Release(slide); slide = null;
                Release(presentation); presentation = null;

                presentation = application.Presentations.Open(documentPath,
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoFalse);
                slide = presentation.Slides[1];
                PowerPointInterop.Shape reopened = null;
                foreach (PowerPointInterop.Shape candidate in slide.Shapes)
                {
                    if (PowerPointBlockService.TryReadContract(candidate, out var candidateMetadata,
                            out _) && candidateMetadata.Id == originalId)
                    {
                        reopened = candidate;
                        break;
                    }
                }
                Assert(reopened != null &&
                       PowerPointBlockService.TryReadContract(reopened, out var reopenedMetadata,
                           out var reopenedSource) &&
                       reopenedMetadata.Id == originalId && reopenedSource == updatedSource &&
                       Math.Abs(reopenedMetadata.FontSizePt - 19) < 0.01 &&
                       Math.Abs(PowerPointBlockService.ReadPositiveTag(reopened,
                           PowerPointBlockService.VisualScaleTag, 0) - expectedScale) < 0.01,
                    "The PowerPoint LaTeX Block contract did not survive PPTX save/reopen.");
                Release(reopened);
            }
            finally
            {
                if (presentation != null)
                {
                    try { presentation.Close(); } catch { }
                    Release(presentation);
                }
                Release(slide);
                Release(secondSlide);
                if (application != null)
                {
                    if (sizeChangeHandler != null)
                        try { application.AfterShapeSizeChange -= sizeChangeHandler; } catch { }
                    try { application.Quit(); } catch { }
                    Release(application);
                }
                backend?.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void WaitFor(Func<bool> predicate, int timeoutMs,
            Func<string> message)
        {
            var stopwatch = Stopwatch.StartNew();
            while (!predicate() && stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(25);
            }
            Assert(predicate(), message());
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                try { Marshal.FinalReleaseComObject(value); } catch { }
        }
    }
}
