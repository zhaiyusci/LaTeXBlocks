using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.FontGlyphSpike
{
    internal static class Program
    {
        private const string FormulaFamily = "TeX Formula Glyph Spike 2";
        // @ is a temporary ASCII carrier. The generated font maps it to the
        // formula outline so Word must choose the hAnsi/ASCII font slot. The
        // font also contains U+E000; testing that PUA carrier is a separate
        // compatibility question, not a premise of this layout spike.
        private const char FormulaCharacter = '@';
        private const float BodySize = 24.0f;

        private static readonly string Root = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", ".."));
        private static readonly string FontPath = Path.Combine(Root, "TeXFormulaGlyphSpike2.ttf");
        private static readonly string LiveDocument = Path.Combine(Root, "01-font-glyph-live.docx");
        private static readonly string EmbeddedDocument = Path.Combine(Root, "02-font-glyph-embedded.docx");
        private static readonly string ResultsPath = Path.Combine(Root, "results.txt");
        private static readonly string VerificationPdf = Path.Combine(Root, "02-font-glyph-embedded-reopen.pdf");
        private static readonly string RemovalVerificationPdf = Path.Combine(Root, "03-font-glyph-after-font-removal.pdf");
        private static readonly string PuaDocument = Path.Combine(Root, "03-pua-formula-glyph.docx");
        private static readonly string PuaVerificationPdf = Path.Combine(Root, "03-pua-formula-glyph.pdf");

        private static int Main(string[] arguments)
        {
            try
            {
                return Run(arguments);
            }
            catch (Exception exception)
            {
                File.WriteAllText(Path.Combine(Root, "error.txt"), exception.ToString());
                Console.Error.WriteLine("Unhandled " + exception.GetType().FullName);
                Console.Error.WriteLine("HRESULT=0x" + exception.HResult.ToString("X8", CultureInfo.InvariantCulture));
                Console.Error.WriteLine(exception.Message);
                return 1;
            }
        }

        private static int Run(string[] arguments)
        {
            if (arguments.Length == 1 && string.Equals(arguments[0], "verify-after-uninstall", StringComparison.Ordinal))
            {
                Console.WriteLine(ReopenDocumentAndExportPdf(EmbeddedDocument, RemovalVerificationPdf));
                return 0;
            }
            if (arguments.Length == 1 && string.Equals(arguments[0], "verify-pua", StringComparison.Ordinal))
            {
                Console.WriteLine(ReopenDocumentAndExportPdf(PuaDocument, PuaVerificationPdf));
                return 0;
            }

            if (!File.Exists(FontPath))
            {
                Console.Error.WriteLine("Build TeXFormulaGlyphSpike.ttf before running this probe.");
                return 2;
            }

            int registrations = AddFontResourceEx(FontPath, 0, IntPtr.Zero);
            if (registrations <= 0)
            {
                Console.Error.WriteLine("Windows rejected the experimental TrueType font.");
                return 3;
            }

            BroadcastFontChange();
            try
            {
                var liveResult = CreateDocument(LiveDocument, embedFont: false);
                var embeddedResult = CreateDocument(EmbeddedDocument, embedFont: true);
                File.WriteAllText(ResultsPath,
                    "Font glyph experiment (Word COM)\r\n" +
                    "Family: " + FormulaFamily + "\r\n" +
                    "Carrier: U+0040 (@), drawn as E = mc² by the custom font\r\n\r\n" +
                    "Live document\r\n" + liveResult + "\r\n\r\n" +
                    "Embedded document\r\n" + embeddedResult + "\r\n",
                    new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

                // Start a fresh Word process only after the temporary font is
                // unregistered.  The separate method does that below.
            }
            finally
            {
                RemoveFontResourceEx(FontPath, 0, IntPtr.Zero);
                BroadcastFontChange();
            }

            var reopenResult = ReopenDocumentAndExportPdf(EmbeddedDocument, VerificationPdf);
            File.AppendAllText(ResultsPath,
                "\r\nReopen after unregistering the font\r\n" + reopenResult + "\r\n",
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            Console.WriteLine(File.ReadAllText(ResultsPath));
            return 0;
        }

        private static string CreateDocument(string targetPath, bool embedFont)
        {
            Word.Application application = null;
            Word.Document document = null;
            string stage = "initializing";
            try
            {
                stage = "starting Word";
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                stage = "creating document";
                document = application.Documents.Add();
                stage = "setting font embedding flags";
                document.EmbedTrueTypeFonts = embedFont;
                // Word returns "command unavailable" when SaveSubsetFonts is
                // set while embedding is disabled.
                if (embedFont) document.SaveSubsetFonts = true;
                document.DoNotEmbedSystemFonts = false;
                stage = "setting page geometry";
                document.PageSetup.TopMargin = application.InchesToPoints(0.85f);
                document.PageSetup.BottomMargin = application.InchesToPoints(0.85f);
                document.PageSetup.LeftMargin = application.InchesToPoints(0.85f);
                document.PageSetup.RightMargin = application.InchesToPoints(0.85f);

                stage = "writing title";
                AddRun(document, "True glyph formula experiment", "Calibri", 16, bold: true);
                AddParagraph(document);
                AddRun(document,
                    "The second line contains an ASCII carrier from a temporary TrueType font; it is a literal text character, not an InlineShape.",
                    "Calibri", 10.5f);
                AddParagraph(document);
                AddParagraph(document);

                AddRun(document, "Reference formula written as ordinary text", "Calibri", 10.5f, bold: true);
                AddParagraph(document);
                AddRun(document, "E = mc²", "Times New Roman", BodySize);
                AddParagraph(document);
                AddParagraph(document);

                AddRun(document, "Ordinary character spacing control", "Calibri", 10.5f, bold: true);
                AddParagraph(document);
                AddRun(document, "What does", "Times New Roman", BodySize);
                RangeRecord ordinaryLine = AddInlineOrdinaryCharacter(document, "x");
                AddRun(document, "stand for?", "Times New Roman", BodySize);
                AddParagraph(document);
                AddParagraph(document);

                AddRun(document, "Formula glyph using an ASCII carrier (@)", "Calibri", 10.5f, bold: true);
                AddParagraph(document);
                AddRun(document, "What does", "Times New Roman", BodySize);
                RangeRecord formulaLine = AddInlineFormulaGlyph(document);
                AddRun(document, "stand for?", "Times New Roman", BodySize);
                document.Bookmarks.Add("TeXFormulaGlyph", formulaLine.Formula);
                AddParagraph(document);
                AddParagraph(document);

                AddRun(document, "Same glyph without surrounding U+0020 spaces", "Calibri", 10.5f, bold: true);
                AddParagraph(document);
                AddRun(document, "What does", "Times New Roman", BodySize);
                AddRun(document, FormulaCharacter.ToString(), FormulaFamily, BodySize);
                AddRun(document, "stand for?", "Times New Roman", BodySize);
                AddParagraph(document);
                AddParagraph(document);

                AddRun(document, "Measured immediately after Word pagination", "Calibri", 10.5f, bold: true);
                AddParagraph(document);
                stage = "paginating";
                document.Repaginate();
                application.ScreenRefresh();

                double ordinaryLeftSpacePt = HorizontalAdvance(ordinaryLine.LeftSpace);
                double ordinaryRightSpacePt = HorizontalAdvance(ordinaryLine.RightSpace);
                double leftSpacePt = HorizontalAdvance(formulaLine.LeftSpace);
                double formulaPt = HorizontalAdvance(formulaLine.Formula);
                double rightSpacePt = HorizontalAdvance(formulaLine.RightSpace);
                string formulaFontInSession = Convert.ToString(formulaLine.Formula.Font.Name, CultureInfo.InvariantCulture);
                AddRun(document,
                    string.Format(CultureInfo.InvariantCulture,
                        "Ordinary-character spaces: {0:0.00} / {1:0.00} pt. Formula-glyph spaces: {2:0.00} / {3:0.00} pt. Formula glyph advance: {4:0.00} pt at {5:0} pt.",
                        ordinaryLeftSpacePt, ordinaryRightSpacePt, leftSpacePt, rightSpacePt, formulaPt, BodySize),
                    "Calibri", 10.5f);
                AddParagraph(document);
                AddRun(document,
                    "Expected interpretation: both U+0020 values are ordinary Times New Roman word spaces. There is no drawing extent or effectExtent in this representation.",
                    "Calibri", 10.5f);
                AddParagraph(document);

                stage = "saving document";
                document.SaveAs2(targetPath, Word.WdSaveFormat.wdFormatXMLDocument);
                return string.Format(CultureInfo.InvariantCulture,
                    "embed={0}; left U+0020={1:0.00} pt; right U+0020={2:0.00} pt; glyph advance={3:0.00} pt",
                    embedFont, leftSpacePt, rightSpacePt, formulaPt) + string.Format(CultureInfo.InvariantCulture,
                    "; ordinary U+0020={0:0.00}/{1:0.00} pt; Word reports formula run font='{2}'",
                    ordinaryLeftSpacePt, ordinaryRightSpacePt, formulaFontInSession);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException("CreateDocument failed at " + stage, exception);
            }
            finally
            {
                if (document != null)
                {
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseCom(document);
                }
                if (application != null)
                {
                    application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseCom(application);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static string ReopenDocumentAndExportPdf(string inputDocument, string outputPdf)
        {
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                document = application.Documents.Open(inputDocument, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                document.Repaginate();
                application.ScreenRefresh();
                document.ExportAsFixedFormat(outputPdf, Word.WdExportFormat.wdExportFormatPDF);

                if (!document.Bookmarks.Exists("TeXFormulaGlyph"))
                {
                    return "could not find the formula glyph bookmark after reopen";
                }
                object bookmarkName = "TeXFormulaGlyph";
                var formula = document.Bookmarks.get_Item(ref bookmarkName).Range;

                string reportedFont = Convert.ToString(formula.Font.Name, CultureInfo.InvariantCulture);
                return "reopen PDF=" + outputPdf + "; Word reports formula run font='" + reportedFont + "'";
            }
            finally
            {
                if (document != null)
                {
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseCom(document);
                }
                if (application != null)
                {
                    application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseCom(application);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static RangeRecord AddInlineFormulaGlyph(Word.Document document)
        {
            return AddInlineOrdinaryCharacter(document, FormulaCharacter.ToString(), FormulaFamily);
        }

        private static RangeRecord AddInlineOrdinaryCharacter(Word.Document document, string character, string fontName = "Times New Roman")
        {
            RangeRecord record = new RangeRecord();
            record.LeftSpace = AddRun(document, " ", "Times New Roman", BodySize);
            record.Formula = AddRun(document, character, fontName, BodySize);
            record.RightSpace = AddRun(document, " ", "Times New Roman", BodySize);
            return record;
        }

        private static Word.Range AddRun(Word.Document document, string text, string fontName, float size, bool bold = false)
        {
            int start = document.Content.End - 1;
            var insertion = document.Range(start, start);
            insertion.InsertAfter(text);
            var written = document.Range(start, start + text.Length);
            written.Font.Name = fontName;
            written.Font.Size = size;
            written.Font.Bold = bold ? -1 : 0;
            ReleaseCom(insertion);
            return written;
        }

        private static void AddParagraph(Word.Document document)
        {
            Word.Range paragraphMark = AddRun(document, "\r", "Calibri", 10.5f);
            ReleaseCom(paragraphMark);
        }

        private static double HorizontalAdvance(Word.Range range)
        {
            Word.Range start = range.Duplicate;
            Word.Range end = range.Duplicate;
            try
            {
                start.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                double x0 = Convert.ToDouble(start.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage), CultureInfo.InvariantCulture);
                double x1 = Convert.ToDouble(end.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage), CultureInfo.InvariantCulture);
                return x1 - x0;
            }
            finally
            {
                ReleaseCom(start);
                ReleaseCom(end);
            }
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int AddFontResourceEx(string name, uint flags, IntPtr reserved);

        [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RemoveFontResourceEx(string name, uint flags, IntPtr reserved);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr windowHandle, uint message, UIntPtr wParam, string lParam,
            uint flags, uint timeoutMilliseconds, out UIntPtr result);

        private static void BroadcastFontChange()
        {
            UIntPtr ignored;
            SendMessageTimeout(new IntPtr(0xffff), 0x001d, UIntPtr.Zero, null, 0x0002, 1000, out ignored);
        }

        private sealed class RangeRecord
        {
            internal Word.Range LeftSpace;
            internal Word.Range Formula;
            internal Word.Range RightSpace;
        }
    }
}
