using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordObjectComparison
{
    internal static class Program
    {
        private const float BodyPointSize = 20.0f;
        private const int DarkBlue = 0x794E1F;
        private const int BodyGray = 0x595959;
        private const int White = 0xFFFFFF;
        private const int PaleBlue = 0xFAF6F2;
        private const int BlueGuide = 0xB6752E;

        [STAThread]
        private static int Main(string[] args)
        {
            var root = args.Length == 0 || string.IsNullOrWhiteSpace(args[0])
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.GetFullPath(args[0]);
            Directory.CreateDirectory(root);

            var artifacts = Path.Combine(root, "artifacts");
            var assets = CreateArtwork(Path.Combine(artifacts, "assets"));
            var docx = Path.Combine(root, "Word-Object-Comparison.docx");
            var pdf = Path.Combine(root, "Word-Object-Comparison.pdf");

            Word.Application word = null;
            Word.Document document = null;
            var retainedRanges = new List<object>();
            try
            {
                Console.WriteLine("Starting Word COM.");
                word = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = word.Documents.Add();
                ConfigurePage(document);

                AddTitle(document);
                var records = AddCases(document, assets, retainedRanges);
                document.Repaginate();
                AddMeasurementTable(document, records, retainedRanges);
                AddInterpretation(document);
                document.Repaginate();

                Console.WriteLine("Saving DOCX.");
                document.SaveAs2(docx, Word.WdSaveFormat.wdFormatDocumentDefault);
                Console.WriteLine("Exporting Word PDF proof.");
                document.ExportAsFixedFormat(pdf, Word.WdExportFormat.wdExportFormatPDF);
                Console.WriteLine("Created: " + docx);
                Console.WriteLine("Word PDF proof: " + pdf);
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
            finally
            {
                for (var index = retainedRanges.Count - 1; index >= 0; index--)
                {
                    Release(retainedRanges[index]);
                }
                if (document != null)
                {
                    try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Release(document);
                }
                if (word != null)
                {
                    try { word.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Release(word);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void ConfigurePage(Word.Document document)
        {
            document.PageSetup.TopMargin = 48.0f;
            document.PageSetup.BottomMargin = 48.0f;
            document.PageSetup.LeftMargin = 58.0f;
            document.PageSetup.RightMargin = 58.0f;
        }

        private static void AddTitle(Word.Document document)
        {
            AddRunAndRelease(document, "Word inline-object comparison", "Aptos Display", 22, true, DarkBlue);
            AddParagraphBreak(document, 2.0f);
            AddRunAndRelease(document,
                "Actual Word COM sample: the same formula is represented as an ordinary text run, SVG, PNG, and EMF.",
                "Aptos", 10.5f, false, BodyGray);
            AddParagraphBreak(document, 10.0f);

            var label = AddRun(document, "Reading the test  ", "Aptos", 9.5f, true, DarkBlue);
            try
            {
                label.Shading.BackgroundPatternColor = (Word.WdColor)Rgb(226, 239, 218);
            }
            finally
            {
                Release(label);
            }
            AddRunAndRelease(document,
                "Yellow swatches are actual ordinary U+0020 spaces. Blue lines in the graphics mark their declared internal baselines; the ordinary-text row is the Word text-baseline reference. Gray frames show the InlineShape boundary.",
                "Aptos", 9.5f, false, 0x444444);
            AddParagraphBreak(document, 11.0f);
        }

        private static List<CaseRecord> AddCases(
            Word.Document document,
            Artwork assets,
            ICollection<object> retainedRanges)
        {
            var cases = new[]
            {
                new CaseDefinition("Ordinary text", "The formula is a genuine Word text run.", null),
                new CaseDefinition("SVG InlineShape", "Vector graphic payload in w:drawing.", assets.Svg),
                new CaseDefinition("PNG InlineShape", "Raster graphic payload in w:drawing.", assets.Png),
                new CaseDefinition("EMF InlineShape", "Metafile graphic payload in w:drawing.", assets.Emf),
            };
            var records = new List<CaseRecord>();

            foreach (var definition in cases)
            {
                AddRunAndRelease(document, definition.Name, "Aptos", 10.0f, true, DarkBlue);
                AddRunAndRelease(document, "  " + definition.Note, "Aptos", 9.0f, false, BodyGray);
                AddParagraphBreak(document, 1.0f);

                var lineStart = document.Content.End - 1;
                AddRunAndRelease(document, "left", "Times New Roman", BodyPointSize, false, 0);
                var leftSpaceStart = document.Content.End - 1;
                AddRunAndRelease(document, " ", "Times New Roman", BodyPointSize, false, 0);
                if (definition.Path == null)
                {
                    AddRunAndRelease(document, "E = mc²", "Cambria Math", BodyPointSize, false, 0);
                }
                else
                {
                    InsertPicture(document, definition.Path);
                }
                var rightSpaceStart = document.Content.End - 1;
                AddRunAndRelease(document, " ", "Times New Roman", BodyPointSize, false, 0);
                AddRunAndRelease(document, "right", "Times New Roman", BodyPointSize, false, 0);

                var lineEnd = document.Content.End - 1;
                // Range objects are live in Word: a run held while later
                // content is appended may grow. Re-create these one-character
                // ranges from their fixed document offsets before highlighting
                // or measuring them.
                var leftSpace = document.Range(leftSpaceStart, leftSpaceStart + 1);
                var rightSpace = document.Range(rightSpaceStart, rightSpaceStart + 1);
                leftSpace.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                rightSpace.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                retainedRanges.Add(leftSpace);
                retainedRanges.Add(rightSpace);
                var line = document.Range(lineStart, lineEnd);
                line.ParagraphFormat.SpaceAfter = 10.0f;
                line.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                retainedRanges.Add(line);
                records.Add(new CaseRecord(definition.Name, definition.Path == null, leftSpace, rightSpace));
                AddParagraphBreak(document, 4.0f);
            }
            return records;
        }

        private static void InsertPicture(Word.Document document, string path)
        {
            Word.Range insertion = null;
            Word.InlineShape picture = null;
            try
            {
                var start = document.Content.End - 1;
                insertion = document.Range(start, start);
                picture = document.InlineShapes.AddPicture(path, false, true, insertion);
                picture.Width = 106.0f;
                picture.Height = 31.8f;
                picture.Range.Font.Position = 0;
            }
            finally
            {
                Release(picture);
                Release(insertion);
            }
        }

        private static void AddMeasurementTable(
            Word.Document document,
            IList<CaseRecord> records,
            ICollection<object> retainedRanges)
        {
            AddRunAndRelease(document, "Observed Word layout", "Aptos Display", 13.0f, true, DarkBlue);
            AddParagraphBreak(document, 3.0f);
            AddRunAndRelease(document,
                "Measured from Word range endpoints after pagination. The values describe the highlighted literal U+0020 characters only; they do not include formula ink or image bounds.",
                "Aptos", 9.0f, false, BodyGray);
            AddParagraphBreak(document, 5.0f);

            Word.Range insertion = null;
            Word.Table table = null;
            try
            {
                var position = document.Content.End - 1;
                insertion = document.Range(position, position);
                table = document.Tables.Add(insertion, records.Count + 1, 4);
                table.AllowAutoFit = false;
                var headers = new[] { "representation", "left U+0020", "right U+0020", "Word object kind" };
                for (var column = 1; column <= 4; column++)
                {
                    SetCell(table.Cell(1, column), headers[column - 1], true, White, DarkBlue);
                    table.Columns[column].Width = new[] { 125.0f, 85.0f, 85.0f, 145.0f }[column - 1];
                }

                for (var index = 0; index < records.Count; index++)
                {
                    var record = records[index];
                    var row = index + 2;
                    var values = new[]
                    {
                        record.Name,
                        string.Format("{0:0.00} pt", GetRangeWidth(record.LeftSpace)),
                        string.Format("{0:0.00} pt", GetRangeWidth(record.RightSpace)),
                        record.IsText ? "w:t text run" : "wdInlineShapePicture",
                    };
                    for (var column = 1; column <= 4; column++)
                    {
                        SetCell(table.Cell(row, column), values[column - 1], false, 0, row % 2 == 0 ? PaleBlue : 0);
                    }
                }
            }
            finally
            {
                Release(table);
                Release(insertion);
            }
            AddParagraphBreak(document, 5.0f);
        }

        private static void AddInterpretation(Word.Document document)
        {
            AddRunAndRelease(document,
                "Interpretation. Only the first row is a character-shaped Word text run. The other three rows use different payload formats, but Word lays every one out as an InlineShape object.",
                "Aptos", 9.5f, false, 0x444444);
            AddParagraphBreak(document, 0.0f);
        }

        private static void SetCell(Word.Cell cell, string text, bool header, int fontColor, int fillColor)
        {
            Word.Range range = null;
            try
            {
                range = cell.Range;
                range.End -= 1;
                range.Text = text;
                range.Font.Name = "Aptos";
                range.Font.Size = 8.5f;
                range.Font.Bold = header ? 1 : 0;
                range.Font.Color = (Word.WdColor)fontColor;
                range.ParagraphFormat.SpaceAfter = 0.0f;
                if (fillColor != 0)
                {
                    cell.Shading.BackgroundPatternColor = (Word.WdColor)fillColor;
                }
            }
            finally
            {
                Release(range);
                Release(cell);
            }
        }

        private static Word.Range AddRun(
            Word.Document document,
            string text,
            string fontName,
            float fontSize,
            bool bold,
            int color,
            bool highlight = false,
            bool underline = false)
        {
            var start = document.Content.End - 1;
            Word.Range anchor = null;
            try
            {
                anchor = document.Range(start, start);
                anchor.Text = text;
            }
            finally
            {
                Release(anchor);
            }

            var result = document.Range(start, start + text.Length);
            result.Font.Name = fontName;
            result.Font.Size = fontSize;
            result.Font.Bold = bold ? 1 : 0;
            result.Font.Color = (Word.WdColor)color;
            result.Font.Underline = underline
                ? Word.WdUnderline.wdUnderlineSingle
                : Word.WdUnderline.wdUnderlineNone;
            if (highlight)
            {
                result.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            }
            if (underline)
            {
                result.Font.UnderlineColor = (Word.WdColor)BlueGuide;
            }
            return result;
        }

        private static void AddRunAndRelease(
            Word.Document document,
            string text,
            string fontName,
            float fontSize,
            bool bold,
            int color,
            bool highlight = false,
            bool underline = false)
        {
            var range = AddRun(document, text, fontName, fontSize, bold, color, highlight, underline);
            Release(range);
        }

        private static void AddParagraphBreak(Word.Document document, float spaceAfter)
        {
            Word.Range range = null;
            try
            {
                range = AddRun(document, "\r", "Aptos", 1.0f, false, 0);
                range.ParagraphFormat.SpaceBefore = 0.0f;
                range.ParagraphFormat.SpaceAfter = spaceAfter;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
            finally
            {
                Release(range);
            }
        }

        private static float GetRangeWidth(Word.Range range)
        {
            Word.Range start = null;
            Word.Range end = null;
            try
            {
                start = range.Duplicate;
                end = range.Duplicate;
                start.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                var x1 = Convert.ToSingle(start.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage));
                var x2 = Convert.ToSingle(end.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage));
                return x2 - x1;
            }
            finally
            {
                Release(start);
                Release(end);
            }
        }

        private static Artwork CreateArtwork(string directory)
        {
            Directory.CreateDirectory(directory);
            var png = Path.Combine(directory, "formula.png");
            var emf = Path.Combine(directory, "formula.emf");
            var svg = Path.Combine(directory, "formula.svg");
            const int width = 480;
            const int height = 144;
            const float scale = 3.0f;

            using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                bitmap.SetResolution(288, 288);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    DrawFormulaArtwork(graphics, width, height, scale);
                }
                bitmap.Save(png, ImageFormat.Png);
            }

            using (var referenceBitmap = new Bitmap(1, 1))
            using (var referenceGraphics = Graphics.FromImage(referenceBitmap))
            {
                var hdc = referenceGraphics.GetHdc();
                try
                {
                    using (var metafile = new Metafile(
                        emf,
                        hdc,
                        new Rectangle(0, 0, 160, 48),
                        MetafileFrameUnit.Point,
                        EmfType.EmfPlusDual))
                    using (var graphics = Graphics.FromImage(metafile))
                    {
                        // GDI records EMF coordinates in the reference DC's
                        // display units. The 3x artwork matches the 160pt
                        // frame's effective Word viewport; without it Word
                        // leaves most of the EMF canvas visibly empty.
                        DrawFormulaArtwork(graphics, width, height, scale, 27.6f);
                    }
                }
                finally
                {
                    referenceGraphics.ReleaseHdc(hdc);
                }
            }

            File.WriteAllText(svg,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160pt\" height=\"48pt\" viewBox=\"0 0 160 48\">\n" +
                "  <rect x=\"0.5\" y=\"0.5\" width=\"159\" height=\"47\" fill=\"none\" stroke=\"#B2BFCC\" stroke-width=\"0.7\"/>\n" +
                "  <text x=\"8\" y=\"32\" font-family=\"Cambria Math\" font-size=\"28\" fill=\"#19202A\">E = mc²</text>\n" +
                "  <line x1=\"1\" y1=\"37\" x2=\"159\" y2=\"37\" stroke=\"#2E75B6\" stroke-width=\"0.8\"/>\n" +
                "</svg>\n");
            return new Artwork(svg, png, emf);
        }

        private static void DrawFormulaArtwork(
            Graphics graphics,
            int width,
            int height,
            float scale,
            float declaredBaseline = 37.0f)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            graphics.Clear(Color.Transparent);
            using (var frame = new Pen(Color.FromArgb(178, 191, 204), 1.0f * scale))
            using (var baseline = new Pen(Color.FromArgb(46, 117, 182), 1.5f * scale))
            using (var ink = new SolidBrush(Color.FromArgb(25, 32, 42)))
            using (var font = new Font("Cambria Math", 28.0f * scale, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var format = StringFormat.GenericTypographic)
            {
                graphics.DrawRectangle(frame, 0.5f * scale, 0.5f * scale,
                    width - 1.0f * scale, height - 1.0f * scale);
                graphics.DrawString("E = mc²", font, ink, 8.0f * scale, 10.0f * scale, format);
                graphics.DrawLine(baseline,
                    1.0f * scale,
                    declaredBaseline * scale,
                    159.0f * scale,
                    declaredBaseline * scale);
            }
        }

        private static int Rgb(int red, int green, int blue)
        {
            return red + (green << 8) + (blue << 16);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        private sealed class Artwork
        {
            internal Artwork(string svg, string png, string emf)
            {
                Svg = svg;
                Png = png;
                Emf = emf;
            }

            internal string Svg { get; }
            internal string Png { get; }
            internal string Emf { get; }
        }

        private sealed class CaseDefinition
        {
            internal CaseDefinition(string name, string note, string path)
            {
                Name = name;
                Note = note;
                Path = path;
            }

            internal string Name { get; }
            internal string Note { get; }
            internal string Path { get; }
        }

        private sealed class CaseRecord
        {
            internal CaseRecord(string name, bool isText, Word.Range leftSpace, Word.Range rightSpace)
            {
                Name = name;
                IsText = isText;
                LeftSpace = leftSpace;
                RightSpace = rightSpace;
            }

            internal string Name { get; }
            internal bool IsText { get; }
            internal Word.Range LeftSpace { get; }
            internal Word.Range RightSpace { get; }
        }
    }
}
