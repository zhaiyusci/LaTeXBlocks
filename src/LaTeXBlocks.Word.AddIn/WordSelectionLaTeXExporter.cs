using System;
using System.Collections.Generic;
using System.Text;
using WordInterop = Microsoft.Office.Interop.Word;

namespace LaTeXBlocks.Word
{
    internal static class WordSelectionLaTeXExporter
    {
        private const char InlineShapeCharacter = '\u0001';
        private const char WordJoiner = '\u2060';

        internal static string Export(WordInterop.Range selection)
        {
            if (selection == null) throw new ArgumentNullException(nameof(selection));
            if (selection.Start == selection.End)
                throw new InvalidOperationException("Select some Word text before copying it as LaTeX.");

            var tokens = ReadInlineBlockTokens(selection);
            var result = new StringBuilder();
            var position = selection.Start;
            while (position < selection.End)
            {
                if (tokens.TryGetValue(position, out var token))
                {
                    result.Append(token.Source);
                    position = Math.Max(position + 1, token.End);
                    continue;
                }

                var characterRange = selection.Document.Range(position, position + 1);
                var text = characterRange.Text ?? string.Empty;
                if (text.Length == 0)
                {
                    position++;
                    continue;
                }

                var character = text[0];
                // WORD JOINER belongs to the auto-width formula scaffold, never to
                // author text. An unrecognized drawing is intentionally omitted:
                // there is no lossless LaTeX representation to export for it.
                if (character != WordJoiner && character != InlineShapeCharacter)
                    AppendEscapedWordCharacter(result, character);
                position++;
            }
            return TrimDocumentEnd(result.ToString());
        }

        internal static string EscapeText(string text)
        {
            var result = new StringBuilder();
            foreach (var character in text ?? string.Empty)
                AppendEscapedWordCharacter(result, character);
            return result.ToString();
        }

        private static Dictionary<int, ExportToken> ReadInlineBlockTokens(WordInterop.Range selection)
        {
            var tokens = new Dictionary<int, ExportToken>();
            for (var index = 1; index <= selection.InlineShapes.Count; index++)
            {
                var shape = selection.InlineShapes[index];
                if (!LaTeXBlockService.TryReadContract(shape, out var metadata, out var source))
                    continue;
                var start = shape.Range.Start;
                var end = shape.Range.End;
                var kind = LaTeXBlockService.ResolveKind(metadata, source);
                if (kind == LaTeXBlockKind.NumberedMath)
                {
                    start = NumberedScaffoldStart(shape, selection.Start);
                    end = NumberedScaffoldEnd(shape, selection.End);
                }
                else if (kind == LaTeXBlockKind.DisplayMath)
                {
                    if (start > selection.Start &&
                        ReadCharacter(shape.Range.Document, start - 1) == '\v') start--;
                    if (end < selection.End &&
                        ReadCharacter(shape.Range.Document, end) == '\v') end++;
                }
                if (kind == LaTeXBlockKind.InlineMath)
                    source = "\\(" + LaTeXBlockService.NormalizeMathBody(source) + "\\)";
                else if (kind == LaTeXBlockKind.DisplayMath ||
                         kind == LaTeXBlockKind.NumberedMath)
                    source = "\n\\[" + LaTeXBlockService.NormalizeMathBody(source) +
                             "\\]\n";
                tokens[start] = new ExportToken(source, end);
            }
            return tokens;
        }

        private static int NumberedScaffoldStart(WordInterop.InlineShape shape, int selectionStart)
        {
            var start = shape.Range.Start;
            if (start > selectionStart && ReadCharacter(shape.Range.Document, start - 1) == '\t')
                start--;
            if (start > selectionStart && ReadCharacter(shape.Range.Document, start - 1) == '\v')
                start--;
            return start;
        }

        private static int NumberedScaffoldEnd(WordInterop.InlineShape shape, int selectionEnd)
        {
            var document = shape.Range.Document;
            var end = shape.Range.End;
            // The owned suffix is: TAB, '(', SEQ field result, ')', and an optional
            // manual line break. Stop defensively if a document has been edited out
            // of that contract instead of consuming unrelated author text.
            if (end >= selectionEnd || ReadCharacter(document, end) != '\t') return end;
            var cursor = end + 1;
            if (cursor >= selectionEnd || ReadCharacter(document, cursor) != '(') return end;
            cursor++;
            var paragraphEnd = Math.Min(selectionEnd, shape.Range.Paragraphs[1].Range.End);
            while (cursor < paragraphEnd && ReadCharacter(document, cursor) != ')') cursor++;
            if (cursor >= paragraphEnd) return end;
            cursor++;
            if (cursor < selectionEnd && ReadCharacter(document, cursor) == '\v') cursor++;
            return cursor;
        }

        private static char ReadCharacter(WordInterop.Document document, int position)
        {
            var text = document.Range(position, position + 1).Text;
            return string.IsNullOrEmpty(text) ? '\0' : text[0];
        }

        private static void AppendEscapedWordCharacter(StringBuilder result, char character)
        {
            switch (character)
            {
                case '\r':
                    result.Append("\n\n");
                    break;
                case '\v':
                    result.Append("\n");
                    break;
                case '\t':
                    result.Append("\t");
                    break;
                case '\\': result.Append("\\textbackslash{}"); break;
                case '#': case '$': case '%': case '&': case '_': case '{': case '}':
                    result.Append('\\').Append(character);
                    break;
                case '~': result.Append("\\textasciitilde{}"); break;
                case '^': result.Append("\\textasciicircum{}"); break;
                case '\a': case '\f':
                    // Word table-cell/end-of-row markers are structural, not text.
                    break;
                default:
                    result.Append(character);
                    break;
            }
        }

        private static string TrimDocumentEnd(string text)
        {
            return (text ?? string.Empty).TrimEnd('\n', '\r');
        }

        private sealed class ExportToken
        {
            internal ExportToken(string source, int end)
            {
                Source = LaTeXBlockService.NormalizeSourceText(source) ?? string.Empty;
                End = end;
            }
            internal string Source { get; }
            internal int End { get; }
        }
    }
}
