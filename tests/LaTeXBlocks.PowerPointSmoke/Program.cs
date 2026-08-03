using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
            Assert(ribbonXml.IndexOf("id=\"LaTeXBlocks.PowerPoint.Edit\"", StringComparison.Ordinal) >= 0 &&
                   ribbonXml.IndexOf("imageMso=\"" + PowerPointRibbonContract.EditBlockImageMso + "\"",
                       StringComparison.Ordinal) >= 0,
                "The PowerPoint Edit Block command does not expose its edit icon.");
            Assert(PowerPointRibbonContract.TryParseFontSize("27.5", out var parsedSize) &&
                   Math.Abs(parsedSize - 27.5) < 0.001 &&
                   !PowerPointRibbonContract.TryParseFontSize("0", out _) &&
                   !PowerPointRibbonContract.TryParseFontSize("not-a-size", out _),
                "The PowerPoint TeX font-size control did not validate point sizes.");
            Assert(LaTeXBlockEditorForm.IsTransientPreviewNavigationFailure(
                       new COMException("busy", unchecked((int)0x8001010A))) &&
                   LaTeXBlockEditorForm.IsTransientPreviewNavigationFailure(
                       new COMException("rejected", unchecked((int)0x80010001))) &&
                   !LaTeXBlockEditorForm.IsTransientPreviewNavigationFailure(
                       new COMException("other", unchecked((int)0x80004005))),
                "The PowerPoint preview did not classify OLE busy responses as retryable.");
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
                Assert(block.LockAspectRatio == Office.MsoTriState.msoFalse,
                    "The inserted PowerPoint block did not expose an editable host frame.");
                AssertHostFrameGeometry(block,
                    PowerPointBlockService.ReadSvgWidthPt(render.SvgBytes),
                    PowerPointBlockService.ReadSvgHeightPt(render.SvgBytes),
                    "The inserted PowerPoint block");
                // TeX owns content colour and leading. The SVG owns padding, fill,
                // border and vertical placement; PowerPoint still receives one
                // ordinary picture whose author-facing source remains unchanged.
                const string styledSource = "First styled line.\\par Second styled line with $E=mc^2$.";
                var styledStyle = new LaTeXBlockStyle(1.5, 8,
                    LaTeXBlockVerticalAlignment.Middle, Color.FromArgb(24, 55, 102),
                    true, Color.FromArgb(241, 245, 255), 1.25,
                    Color.FromArgb(51, 98, 162));
                Assert(LaTeXBlockStyle.ReadFromTag(styledStyle.ToString()).Equals(styledStyle) &&
                       LaTeXBlockStyle.ReadFromTag(null).Equals(LaTeXBlockStyle.Default) &&
                       LaTeXBlockStyle.ReadFromTag("LaTeXBlocksStyle/1;leading=not-a-number;" +
                           "padding=not-a-number;border=not-a-number")
                           .Equals(LaTeXBlockStyle.Default),
                    "PowerPoint TeX style tags did not round-trip or safely recover from malformed values.");
                var styledWrapper = styledStyle.WrapSource(styledSource, inheritedSize);
                Assert(styledWrapper.IndexOf("\\global\\PreviewBorder=0pt",
                           StringComparison.Ordinal) >= 0 &&
                       styledWrapper.IndexOf("\\renewcommand{\\baselinestretch}{1.5}",
                           StringComparison.Ordinal) >= 0 &&
                       styledWrapper.IndexOf("\\fbox", StringComparison.Ordinal) < 0 &&
                       styledWrapper.IndexOf("\\colorbox", StringComparison.Ordinal) < 0 &&
                       styledWrapper.IndexOf("\\vbox to ", StringComparison.Ordinal) < 0 &&
                       styledWrapper.IndexOf(styledSource, StringComparison.Ordinal) >= 0,
                    "The PowerPoint TeX wrapper still contains outer-box styling.");
                var styledRender = service.RenderPreviewAsync(styledSource, 288, profile,
                    inheritedSize, styledStyle, 126).GetAwaiter().GetResult();
                Assert(styledRender.SvgBytes.Length > 0,
                    "StemTeX could not render a PowerPoint TeX-styled block.");
                Assert(Math.Abs(PowerPointBlockService.ReadSvgWidthPt(styledRender.SvgBytes) -
                           288) < 0.05 &&
                       Math.Abs(PowerPointBlockService.ReadSvgHeightPt(styledRender.SvgBytes) -
                           126) < 0.05,
                    "A styled SVG block did not make its requested PowerPoint frame.");
                var styledSvg = Encoding.UTF8.GetString(styledRender.SvgBytes);
                Assert(styledSvg.IndexOf("#f1f5ff", StringComparison.OrdinalIgnoreCase) >= 0 &&
                       styledSvg.IndexOf("#183766", StringComparison.OrdinalIgnoreCase) >= 0 &&
                       styledSvg.IndexOf("#3362a2", StringComparison.OrdinalIgnoreCase) >= 0,
                    "The styled SVG did not contain its text, fill, and border colors.");
                Assert(styledSvg.IndexOf("data-latexblocks-frame='1'",
                           StringComparison.Ordinal) >= 0 &&
                       styledSvg.IndexOf("data-latexblocks-border='1'",
                           StringComparison.Ordinal) >= 0,
                    "The styled SVG did not contain its PowerPoint-composed frame.");
                AssertSvgRectanglePaintFitsViewport(styledRender.SvgBytes,
                    "A PowerPoint-composed SVG border was clipped by its viewport.");
                var topStyle = new LaTeXBlockStyle(1.5, 8,
                    LaTeXBlockVerticalAlignment.Top, styledStyle.TextColor, true,
                    styledStyle.BackgroundColor, styledStyle.BorderThicknessPt,
                    styledStyle.BorderColor);
                var bottomStyle = new LaTeXBlockStyle(1.5, 8,
                    LaTeXBlockVerticalAlignment.Bottom, styledStyle.TextColor, true,
                    styledStyle.BackgroundColor, styledStyle.BorderThicknessPt,
                    styledStyle.BorderColor);
                var topRender = service.RenderPreviewAsync(styledSource, 288, profile,
                    inheritedSize, topStyle, 150).GetAwaiter().GetResult();
                var bottomRender = service.RenderPreviewAsync(styledSource, 288, profile,
                    inheritedSize, bottomStyle, 150).GetAwaiter().GetResult();
                Assert(Math.Abs(PowerPointBlockService.ReadSvgHeightPt(topRender.SvgBytes) -
                           150) < 0.05 &&
                       Math.Abs(PowerPointBlockService.ReadSvgHeightPt(bottomRender.SvgBytes) -
                           150) < 0.05 &&
                       ReadSvgViewBoxY(topRender.SvgBytes) >
                           ReadSvgViewBoxY(bottomRender.SvgBytes) + 0.1 &&
                       !string.Equals(Convert.ToBase64String(topRender.SvgBytes),
                           Convert.ToBase64String(bottomRender.SvgBytes),
                           StringComparison.Ordinal),
                    "Top and bottom SVG vertical alignment did not produce distinct fixed-height boxes.");
                const string leadingSource = "A deliberately long ordinary sentence wraps over several lines " +
                    "inside this narrow TeX block so the selected line spacing is measurable.";
                var tightLeading = service.RenderPreviewAsync(leadingSource, 110, profile,
                    inheritedSize, new LaTeXBlockStyle(1.0), null).GetAwaiter().GetResult();
                var looseLeading = service.RenderPreviewAsync(leadingSource, 110, profile,
                    inheritedSize, new LaTeXBlockStyle(1.8), null).GetAwaiter().GetResult();
                Assert(PowerPointBlockService.ReadSvgHeightPt(looseLeading.SvgBytes) >
                       PowerPointBlockService.ReadSvgHeightPt(tightLeading.SvgBytes) + 8,
                    "The TeX line-spacing control did not change ordinary paragraph leading.");
                var constrainedStyle = new LaTeXBlockStyle(1.2, 6,
                    LaTeXBlockVerticalAlignment.Bottom, Color.Black, true,
                    Color.FromArgb(250, 250, 250), 0.75, Color.Black);
                var constrainedRender = service.RenderPreviewAsync(leadingSource, 110,
                    profile, inheritedSize, constrainedStyle, 24).GetAwaiter().GetResult();
                Assert(PowerPointBlockService.ReadSvgHeightPt(constrainedRender.SvgBytes) >
                       100,
                    "An SVG-styled block cropped content when its requested height was too small.");
                // Styled requests set PreviewBorder to zero at TeX shipout. A later
                // ordinary block must restore the profile's historical border rather
                // than inherit that temporary renderer state.
                var restoredDefaultRender = service.RenderPreviewAsync(source, 288, profile,
                    inheritedSize).GetAwaiter().GetResult();
                Assert(Math.Abs(PowerPointBlockService.ReadSvgWidthPt(
                           restoredDefaultRender.SvgBytes) -
                               PowerPointBlockService.ReadSvgWidthPt(render.SvgBytes)) < 0.02 &&
                       Math.Abs(PowerPointBlockService.ReadSvgHeightPt(
                           restoredDefaultRender.SvgBytes) -
                               PowerPointBlockService.ReadSvgHeightPt(render.SvgBytes)) < 0.02,
                    "A TeX-styled render leaked PreviewBorder state into an ordinary PowerPoint block.");
                var styledBlock = service.InsertRendered(styledSource, 288, styledRender,
                    styledStyle);
                Assert(PowerPointBlockService.TryReadContract(styledBlock, out _,
                           out var storedStyledSource) && storedStyledSource == styledSource &&
                       PowerPointBlockService.ReadStyle(styledBlock).Equals(styledStyle) &&
                       string.Equals(styledBlock.Tags[LaTeXBlockStyle.TagName],
                           styledStyle.ToString(), StringComparison.Ordinal),
                    "A TeX-styled PowerPoint block did not retain its raw source and style metadata.");
                var styledUpdatedRender = service.RenderPreviewAsync(styledSource, 288,
                    profile, inheritedSize, styledStyle, 150, 330).GetAwaiter().GetResult();
                styledBlock = service.UpdateRendered(styledBlock, styledSource, 288,
                    styledUpdatedRender, false, 150, 330, styledStyle);
                Assert(PowerPointBlockService.ReadStyle(styledBlock).Equals(styledStyle) &&
                       styledBlock.AlternativeText == styledSource,
                    "An SVG-styled PowerPoint block lost its style or raw source during a re-render.");
                AssertHostFrameGeometry(styledBlock, 330, 150,
                    "A styled PowerPoint host-frame update");
                var styledBlockId = Guid.Empty;
                if (PowerPointBlockService.TryReadContract(styledBlock,
                    out var styledMetadata, out _))
                    styledBlockId = styledMetadata.Id;
                Assert(styledBlockId != Guid.Empty,
                    "A TeX-styled PowerPoint block lost its identity before persistence.");
                Release(styledBlock);

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

                // Test all native handles through the one host-frame contract. An installed
                // copy of the add-in can be auto-loaded into this COM PowerPoint instance
                // and otherwise consume AfterShapeSizeChange first. Temporarily hide only
                // this test shape from that production handler while capturing a gesture;
                // each captured gesture is then committed through UpdateRendered below.
                block.Tags.Delete(PowerPointBlockService.KindTag);
                PowerPointFrameUpdate horizontalFrame = null;
                try
                {
                    var originalBlockWidth = block.Width;
                    var originalBlockLeft = block.Left;
                    var eventsBeforeResize = sizeChangeEvents;
                    var requestedHorizontalWidth = originalBlockWidth * 1.2f;
                    block.Width = requestedHorizontalWidth;
                    WaitFor(() => sizeChangeEvents > eventsBeforeResize, 2000,
                        () => "PowerPoint did not raise AfterShapeSizeChange after a block resize.");
                    horizontalFrame = PowerPointBlockService.CaptureFrameResize(block, metadata);
                    Assert(horizontalFrame.HasChange && horizontalFrame.WidthChanged &&
                           !horizontalFrame.HeightChanged &&
                           Math.Abs(horizontalFrame.FrameWidthPt - requestedHorizontalWidth) < 0.05,
                        "A horizontal PowerPoint handle resize was not captured as a host-frame width update.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                    Assert(Math.Abs(block.Width - originalBlockWidth) < 0.02 &&
                           Math.Abs(block.Left - originalBlockLeft) < 0.02,
                        "Capturing a horizontal host-frame update could not restore stored geometry.");
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }

                var horizontalRender = service.RenderPreviewAsync(source,
                    metadata.WidthPt, profile, metadata.FontSizePt)
                    .GetAwaiter().GetResult();
                var horizontalNaturalWidth = PowerPointBlockService.ReadSvgWidthPt(
                    horizontalRender.SvgBytes);
                var horizontalNaturalHeight = PowerPointBlockService.ReadSvgHeightPt(
                    horizontalRender.SvgBytes);
                block = service.UpdateRendered(block, source, metadata.WidthPt,
                    horizontalRender, false, horizontalFrame.FrameHeightPt,
                    horizontalFrame.FrameWidthPt);
                Assert(PowerPointBlockService.TryReadContract(block, out metadata, out _) &&
                       Math.Abs(metadata.WidthPt - 288) < 0.05,
                    "A native host-frame update changed the stored TeX layout width.");
                AssertHostFrameGeometry(block,
                    Math.Max(horizontalNaturalWidth, horizontalFrame.FrameWidthPt),
                    Math.Max(horizontalNaturalHeight, horizontalFrame.FrameHeightPt),
                    "A horizontal host-frame update");

                // A frame smaller than the unchanged TeX box is clamped rather
                // than being treated as an instruction to scale or crop the SVG.
                var requestedShortWidth = Math.Max(1, horizontalNaturalWidth / 2.0);
                PowerPointFrameUpdate shortWidthFrame = null;
                block.Tags.Delete(PowerPointBlockService.KindTag);
                try
                {
                    var originalBlockWidth = block.Width;
                    var originalBlockHeight = block.Height;
                    block.Width = (float)requestedShortWidth;
                    shortWidthFrame = PowerPointBlockService.CaptureFrameResize(block,
                        metadata);
                    Assert(shortWidthFrame.HasChange && shortWidthFrame.WidthChanged &&
                           !shortWidthFrame.HeightChanged &&
                           Math.Abs(shortWidthFrame.FrameWidthPt - requestedShortWidth) < 0.05 &&
                           shortWidthFrame.FrameWidthPt < originalBlockWidth - 1,
                        "A narrow horizontal handle resize was not captured as a host-frame request.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                    Assert(Math.Abs(block.Height - originalBlockHeight) < 0.02,
                        "Capturing a narrow horizontal host-frame update changed stored height.");
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }
                var shortWidthRender = service.RenderPreviewAsync(source, metadata.WidthPt,
                    profile, metadata.FontSizePt).GetAwaiter().GetResult();
                var shortWidthNatural = PowerPointBlockService.ReadSvgWidthPt(
                    shortWidthRender.SvgBytes);
                var shortWidthNaturalHeight = PowerPointBlockService.ReadSvgHeightPt(
                    shortWidthRender.SvgBytes);
                block = service.UpdateRendered(block, source, metadata.WidthPt, shortWidthRender,
                    false, shortWidthFrame.FrameHeightPt, shortWidthFrame.FrameWidthPt);
                AssertHostFrameGeometry(block, shortWidthNatural,
                    Math.Max(shortWidthNaturalHeight, shortWidthFrame.FrameHeightPt),
                    "A too-narrow host-frame update");
                Assert(block.Width >= shortWidthNatural - 0.05,
                    "A host frame narrower than the natural TeX SVG clipped the content.");

                var verticalRender = service.RenderPreviewAsync(source, metadata.WidthPt,
                    profile, metadata.FontSizePt).GetAwaiter().GetResult();
                var verticalNaturalWidth = PowerPointBlockService.ReadSvgWidthPt(verticalRender.SvgBytes);
                var verticalNaturalHeight = PowerPointBlockService.ReadSvgHeightPt(verticalRender.SvgBytes);
                var requestedVerticalHeight = Math.Max(verticalNaturalHeight + 42, block.Height + 42);
                PowerPointFrameUpdate verticalFrame = null;
                block.Tags.Delete(PowerPointBlockService.KindTag);
                try
                {
                    var originalBlockWidth = block.Width;
                    block.Height = (float)requestedVerticalHeight;
                    verticalFrame = PowerPointBlockService.CaptureFrameResize(block, metadata);
                    Assert(verticalFrame.HasChange && !verticalFrame.WidthChanged &&
                           verticalFrame.HeightChanged &&
                           Math.Abs(verticalFrame.FrameHeightPt - requestedVerticalHeight) < 0.05,
                        "A vertical PowerPoint handle resize was not captured as a host-frame height update.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                    Assert(Math.Abs(block.Width - originalBlockWidth) < 0.02,
                        "Capturing a vertical host-frame update changed the stored width.");
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }
                block = service.UpdateRendered(block, source, metadata.WidthPt, verticalRender,
                    false, verticalFrame.FrameHeightPt, verticalFrame.FrameWidthPt);
                AssertHostFrameGeometry(block,
                    Math.Max(verticalNaturalWidth, verticalFrame.FrameWidthPt),
                    Math.Max(verticalNaturalHeight, verticalFrame.FrameHeightPt),
                    "A vertical host-frame update");
                Assert(block.Height > verticalNaturalHeight + 30,
                    "A vertical host-frame update did not add an empty SVG frame around the natural TeX content.");

                var shortFrameHeight = Math.Max(1, verticalNaturalHeight / 2.0);
                PowerPointFrameUpdate shortFrame = null;
                block.Tags.Delete(PowerPointBlockService.KindTag);
                try
                {
                    block.Height = (float)shortFrameHeight;
                    shortFrame = PowerPointBlockService.CaptureFrameResize(block, metadata);
                    Assert(shortFrame.HasChange && !shortFrame.WidthChanged &&
                           shortFrame.HeightChanged &&
                           shortFrame.FrameHeightPt < verticalNaturalHeight - 1,
                        "A short vertical handle resize was not captured as a host-frame request.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }
                var shortRender = service.RenderPreviewAsync(source, metadata.WidthPt, profile,
                    metadata.FontSizePt).GetAwaiter().GetResult();
                var shortNaturalWidth = PowerPointBlockService.ReadSvgWidthPt(shortRender.SvgBytes);
                var shortNaturalHeight = PowerPointBlockService.ReadSvgHeightPt(shortRender.SvgBytes);
                block = service.UpdateRendered(block, source, metadata.WidthPt, shortRender,
                    false, shortFrame.FrameHeightPt, shortFrame.FrameWidthPt);
                AssertHostFrameGeometry(block,
                    Math.Max(shortNaturalWidth, shortFrame.FrameWidthPt), shortNaturalHeight,
                    "A too-short host-frame update");
                Assert(block.Height >= shortNaturalHeight - 0.05,
                    "A host frame shorter than the natural TeX SVG clipped the content.");

                PowerPointFrameUpdate cornerFrame = null;
                var requestedCornerWidth = block.Width * 1.15f;
                var requestedCornerHeight = block.Height + 48;
                block.Tags.Delete(PowerPointBlockService.KindTag);
                try
                {
                    block.Width = requestedCornerWidth;
                    block.Height = requestedCornerHeight;
                    cornerFrame = PowerPointBlockService.CaptureFrameResize(block, metadata);
                    Assert(cornerFrame.HasChange && cornerFrame.WidthChanged &&
                           cornerFrame.HeightChanged &&
                           Math.Abs(cornerFrame.FrameWidthPt - requestedCornerWidth) < 0.08 &&
                           Math.Abs(cornerFrame.FrameHeightPt - requestedCornerHeight) < 0.05,
                        "A corner PowerPoint handle resize was not captured as one host-frame update.");
                    PowerPointBlockService.RestoreStoredGeometry(block, metadata);
                }
                finally
                {
                    block.Tags.Add(PowerPointBlockService.KindTag,
                        PowerPointBlockService.KindValue);
                }
                var cornerRender = service.RenderPreviewAsync(source, metadata.WidthPt,
                    profile, metadata.FontSizePt).GetAwaiter().GetResult();
                var cornerNaturalWidth = PowerPointBlockService.ReadSvgWidthPt(cornerRender.SvgBytes);
                var cornerNaturalHeight = PowerPointBlockService.ReadSvgHeightPt(cornerRender.SvgBytes);
                block = service.UpdateRendered(block, source, metadata.WidthPt,
                    cornerRender, false, cornerFrame.FrameHeightPt,
                    cornerFrame.FrameWidthPt);
                Assert(PowerPointBlockService.TryReadContract(block, out metadata, out _) &&
                       Math.Abs(metadata.WidthPt - 288) < 0.05,
                    "A corner host-frame update changed the stored TeX layout width.");
                AssertHostFrameGeometry(block,
                    Math.Max(cornerNaturalWidth, cornerFrame.FrameWidthPt),
                    Math.Max(cornerNaturalHeight, cornerFrame.FrameHeightPt),
                    "A corner host-frame update");

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
                var expectedLeft = block.Left;
                var expectedTop = block.Top;
                var expectedRotation = block.Rotation;
                var expectedZ = block.ZOrderPosition;
                var originalId = metadata.Id;

                const string updatedSource = "Updated block: $\\int_0^1 x^2\\,dx=1/3$.";
                var updatedRender = service.RenderPreviewAsync(updatedSource, 288, profile, 19)
                    .GetAwaiter().GetResult();
                var expectedWidth = Math.Max(PowerPointBlockService.ReadSvgWidthPt(
                    updatedRender.SvgBytes), PowerPointBlockService.ReadFrameWidthPt(block));
                var expectedHeight = Math.Max(PowerPointBlockService.ReadSvgHeightPt(
                    updatedRender.SvgBytes), PowerPointBlockService.ReadFrameHeightPt(block));
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
                    "PowerPoint edit did not preserve position, host-frame width, rotation, and z-order.");
                Assert(Math.Abs(updated.Height - expectedHeight) < 0.04,
                    "PowerPoint edit did not preserve the new SVG's host-frame height.");
                AssertHostFrameGeometry(updated, expectedWidth, expectedHeight,
                    "A PowerPoint edit after host-frame resizes");

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
                PowerPointInterop.Shape reopenedStyled = null;
                foreach (PowerPointInterop.Shape candidate in slide.Shapes)
                {
                    if (PowerPointBlockService.TryReadContract(candidate, out var candidateMetadata,
                            out _))
                    {
                        if (candidateMetadata.Id == originalId)
                            reopened = candidate;
                        else if (candidateMetadata.Id == styledBlockId)
                            reopenedStyled = candidate;
                    }
                }
                Assert(reopened != null &&
                        PowerPointBlockService.TryReadContract(reopened, out var reopenedMetadata,
                            out var reopenedSource) &&
                        reopenedMetadata.Id == originalId && reopenedSource == updatedSource &&
                        Math.Abs(reopenedMetadata.FontSizePt - 19) < 0.01,
                    "The PowerPoint LaTeX Block contract did not survive PPTX save/reopen.");
                AssertHostFrameGeometry(reopened, expectedWidth, expectedHeight,
                    "A reopened PowerPoint host frame");
                Assert(reopenedStyled != null &&
                       PowerPointBlockService.TryReadContract(reopenedStyled,
                           out var reopenedStyledMetadata, out var reopenedStyledSource) &&
                       reopenedStyledMetadata.Id == styledBlockId &&
                       reopenedStyledSource == styledSource &&
                       PowerPointBlockService.ReadStyle(reopenedStyled).Equals(styledStyle) &&
                       string.Equals(reopenedStyled.Tags[LaTeXBlockStyle.TagName],
                           styledStyle.ToString(), StringComparison.Ordinal),
                    "A TeX-styled PowerPoint block did not preserve raw source and style after PPTX save/reopen.");
                Release(reopened);
                Release(reopenedStyled);
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

        private static void AssertHostFrameGeometry(PowerPointInterop.Shape shape,
            double expectedWidthPt, double expectedHeightPt, string context)
        {
            const double tolerancePt = 0.06;
            var taggedWidth = PowerPointBlockService.ReadPositiveTag(shape,
                PowerPointBlockService.SvgWidthTag, 0);
            var taggedHeight = PowerPointBlockService.ReadPositiveTag(shape,
                PowerPointBlockService.SvgHeightTag, 0);
            Assert(taggedWidth > 0 && taggedHeight > 0 &&
                   Math.Abs(shape.Width - taggedWidth) < tolerancePt &&
                   Math.Abs(shape.Height - taggedHeight) < tolerancePt,
                context + " did not keep its SVG width/height tags equal to its shape geometry.");
            Assert(Math.Abs(shape.Width - expectedWidthPt) < tolerancePt &&
                   Math.Abs(shape.Height - expectedHeightPt) < tolerancePt,
                context + " did not retain the expected renderer-root host-frame dimensions.");
            Assert(string.IsNullOrEmpty(ReadTag(shape, "LATEXBLOCKS_VISUAL_SCALE")),
                context + " retained obsolete visual-scale metadata.");
        }

        private static void AssertSvgRectanglePaintFitsViewport(byte[] svgBytes, string message)
        {
            if (svgBytes == null || svgBytes.Length == 0)
                throw new InvalidOperationException(message + " SVG output was empty.");

            var document = new System.Xml.XmlDocument();
            using (var stream = new MemoryStream(svgBytes)) document.Load(stream);
            var root = document.DocumentElement;
            var viewBoxParts = (root?.GetAttribute("viewBox") ?? string.Empty).Split(
                new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (viewBoxParts.Length != 4)
                throw new InvalidOperationException(message + " SVG root had no usable viewBox.");

            var viewLeft = ParseSvgNumber(viewBoxParts[0]);
            var viewTop = ParseSvgNumber(viewBoxParts[1]);
            var viewWidth = ParseSvgNumber(viewBoxParts[2]);
            var viewHeight = ParseSvgNumber(viewBoxParts[3]);
            var viewRight = viewLeft + viewWidth;
            var viewBottom = viewTop + viewHeight;
            const double tolerance = 0.05;
            var rectCount = 0;
            foreach (System.Xml.XmlNode node in document.GetElementsByTagName("rect"))
            {
                var element = node as System.Xml.XmlElement;
                if (element == null || !element.HasAttribute("x") || !element.HasAttribute("width") ||
                    !element.HasAttribute("y") || !element.HasAttribute("height"))
                    continue;
                var left = ParseSvgNumber(element.GetAttribute("x"));
                var right = left + ParseSvgNumber(element.GetAttribute("width"));
                var top = ParseSvgNumber(element.GetAttribute("y"));
                var bottom = top + ParseSvgNumber(element.GetAttribute("height"));
                ++rectCount;
                Assert(left >= viewLeft - tolerance && right <= viewRight + tolerance &&
                       top >= viewTop - tolerance && bottom <= viewBottom + tolerance,
                    message + " Paint extends outside viewBox: " + left.ToString("0.###", CultureInfo.InvariantCulture) +
                    ".." + right.ToString("0.###", CultureInfo.InvariantCulture) + ", " +
                    top.ToString("0.###", CultureInfo.InvariantCulture) + ".." +
                    bottom.ToString("0.###", CultureInfo.InvariantCulture) + " vs " +
                    viewLeft.ToString("0.###", CultureInfo.InvariantCulture) + ".." +
                    viewRight.ToString("0.###", CultureInfo.InvariantCulture) + ", " +
                    viewTop.ToString("0.###", CultureInfo.InvariantCulture) + ".." +
                    viewBottom.ToString("0.###", CultureInfo.InvariantCulture) + ".");
            }
            Assert(rectCount > 0, message + " No rectangle paint was emitted.");
        }

        private static double ParseSvgNumber(string value)
        {
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var result))
                throw new InvalidOperationException("SVG contained an invalid numeric attribute: " + value);
            return result;
        }

        private static double ReadSvgViewBoxY(byte[] svgBytes)
        {
            var document = new System.Xml.XmlDocument();
            using (var stream = new MemoryStream(svgBytes)) document.Load(stream);
            var parts = (document.DocumentElement?.GetAttribute("viewBox") ?? string.Empty)
                .Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
                throw new InvalidOperationException("SVG root had no usable viewBox.");
            return ParseSvgNumber(parts[1]);
        }

        private static string ReadTag(PowerPointInterop.Shape shape, string name)
        {
            try { return shape.Tags[name] ?? string.Empty; }
            catch (COMException) { return string.Empty; }
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
