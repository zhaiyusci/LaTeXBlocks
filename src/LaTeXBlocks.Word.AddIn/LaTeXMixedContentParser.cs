using System;
using System.Collections.Generic;
using System.Text;

namespace LaTeXBlocks.Word
{
    internal enum LaTeXContentKind { Text, InlineMath, DisplayMath }
    internal enum LaTeXTextFontFamily { Inherited, Roman, SansSerif, Monospace }

    internal sealed class LaTeXContentSegment
    {
        internal LaTeXContentSegment(LaTeXContentKind kind, string source,
            bool bold = false, bool italic = false,
            LaTeXTextFontFamily fontFamily = LaTeXTextFontFamily.Inherited)
        {
            Kind = kind;
            Source = source ?? string.Empty;
            Bold = bold;
            Italic = italic;
            FontFamily = fontFamily;
        }
        internal LaTeXContentKind Kind { get; }
        internal string Source { get; }
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal LaTeXTextFontFamily FontFamily { get; }

        internal LaTeXContentSegment WithTextStyle(bool bold, bool italic,
            LaTeXTextFontFamily fontFamily)
        {
            if (Kind != LaTeXContentKind.Text) return this;
            return new LaTeXContentSegment(Kind, Source, Bold || bold, Italic || italic,
                FontFamily != LaTeXTextFontFamily.Inherited ? FontFamily : fontFamily);
        }
    }

    internal static class LaTeXMixedContentParser
    {
        private static readonly HashSet<string> DisplayMathEnvironments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "equation", "equation*", "align", "align*", "alignat", "alignat*",
                "gather", "gather*", "multline", "multline*", "flalign", "flalign*",
                "displaymath"
            };
        private static readonly HashSet<string> InlineMathEnvironments =
            new HashSet<string>(StringComparer.Ordinal) { "math" };

        internal static IReadOnlyList<LaTeXContentSegment> Parse(string source)
        {
            source = LaTeXBlockService.NormalizeSourceText(source) ?? string.Empty;
            var segments = new List<LaTeXContentSegment>();
            var text = new StringBuilder();
            var position = 0;
            while (position < source.Length)
            {
                if (source[position] == '%' && !IsEscaped(source, position))
                {
                    while (position < source.Length && source[position] != '\n') position++;
                    continue;
                }

                if (TryReadDelimitedMath(source, position, out var end, out var kind))
                {
                    FlushText(segments, text);
                    segments.Add(new LaTeXContentSegment(kind, source.Substring(position, end - position)));
                    position = end;
                    continue;
                }

                if (TryReadMathEnvironment(source, position, out end, out kind))
                {
                    FlushText(segments, text);
                    segments.Add(new LaTeXContentSegment(kind,
                        source.Substring(position, end - position)));
                    position = end;
                    continue;
                }

                if (TryReadStyledText(source, position, out var styled, out end))
                {
                    FlushText(segments, text);
                    foreach (var segment in styled) segments.Add(segment);
                    position = end;
                    continue;
                }

                if (TryReadTextEscape(source, position, out var decoded, out end))
                {
                    text.Append(decoded);
                    position = end;
                    continue;
                }

                text.Append(source[position++]);
            }
            FlushText(segments, text);
            return segments;
        }

        private static bool TryReadDelimitedMath(string source, int start, out int end,
            out LaTeXContentKind kind)
        {
            end = start;
            kind = LaTeXContentKind.Text;
            string open;
            string close;
            if (source[start] == '$' && !IsEscaped(source, start))
            {
                if (start + 1 < source.Length && source[start + 1] == '$')
                {
                    open = close = "$$";
                    kind = LaTeXContentKind.DisplayMath;
                }
                else
                {
                    open = close = "$";
                    kind = LaTeXContentKind.InlineMath;
                }
            }
            else if (StartsWith(source, start, "\\(") && !IsEscaped(source, start))
            {
                open = "\\("; close = "\\)"; kind = LaTeXContentKind.InlineMath;
            }
            else if (StartsWith(source, start, "\\[") && !IsEscaped(source, start))
            {
                open = "\\["; close = "\\]"; kind = LaTeXContentKind.DisplayMath;
            }
            else return false;

            var cursor = start + open.Length;
            while (cursor < source.Length)
            {
                if (source[cursor] == '%' && !IsEscaped(source, cursor))
                {
                    while (cursor < source.Length && source[cursor] != '\n') cursor++;
                    continue;
                }
                if (StartsWith(source, cursor, close) && !IsEscaped(source, cursor) &&
                    (close != "$" || cursor + 1 >= source.Length || source[cursor + 1] != '$'))
                {
                    end = cursor + close.Length;
                    return true;
                }
                cursor++;
            }
            throw new ArgumentException("An opened LaTeX math delimiter has no matching close delimiter.");
        }

        private static bool TryReadMathEnvironment(string source, int start, out int end,
            out LaTeXContentKind kind)
        {
            end = start;
            kind = LaTeXContentKind.Text;
            if (!TryReadEnvironmentCommand(source, start, "begin", out var environment, out var commandEnd) ||
                (!DisplayMathEnvironments.Contains(environment) &&
                 !InlineMathEnvironments.Contains(environment))) return false;
            kind = InlineMathEnvironments.Contains(environment)
                ? LaTeXContentKind.InlineMath : LaTeXContentKind.DisplayMath;

            var depth = 1;
            var cursor = commandEnd;
            while (cursor < source.Length)
            {
                if (source[cursor] == '%' && !IsEscaped(source, cursor))
                {
                    while (cursor < source.Length && source[cursor] != '\n') cursor++;
                    continue;
                }
                if (TryReadEnvironmentCommand(source, cursor, "begin", out var nested, out var nestedEnd) &&
                    string.Equals(nested, environment, StringComparison.Ordinal))
                {
                    depth++;
                    cursor = nestedEnd;
                    continue;
                }
                if (TryReadEnvironmentCommand(source, cursor, "end", out nested, out nestedEnd) &&
                    string.Equals(nested, environment, StringComparison.Ordinal))
                {
                    depth--;
                    cursor = nestedEnd;
                    if (depth == 0) { end = cursor; return true; }
                    continue;
                }
                cursor++;
            }
            throw new ArgumentException("The LaTeX environment '" + environment + "' has no matching \\end.");
        }

        private static bool TryReadEnvironmentCommand(string source, int start, string command,
            out string environment, out int end)
        {
            environment = null;
            end = start;
            var prefix = "\\" + command;
            if (!StartsWith(source, start, prefix) || IsEscaped(source, start)) return false;
            var cursor = start + prefix.Length;
            while (cursor < source.Length && char.IsWhiteSpace(source[cursor])) cursor++;
            if (cursor >= source.Length || source[cursor] != '{') return false;
            var close = source.IndexOf('}', cursor + 1);
            if (close < 0) return false;
            environment = source.Substring(cursor + 1, close - cursor - 1).Trim();
            end = close + 1;
            return environment.Length > 0;
        }

        private static bool TryReadTextEscape(string source, int start, out string decoded, out int end)
        {
            decoded = null;
            end = start;
            if (source[start] != '\\' || IsEscaped(source, start)) return false;
            if (start + 1 < source.Length)
            {
                var next = source[start + 1];
                if ("%&#_${}".IndexOf(next) >= 0)
                {
                    decoded = next.ToString(); end = start + 2; return true;
                }
                if (next == '\\')
                {
                    decoded = "\n"; end = start + 2; return true;
                }
            }
            foreach (var item in new[] {
                new[] { "\\textbackslash{}", "\\" },
                new[] { "\\textasciitilde{}", "~" },
                new[] { "\\textasciicircum{}", "^" } })
            {
                if (!StartsWith(source, start, item[0])) continue;
                decoded = item[1]; end = start + item[0].Length; return true;
            }
            return false;
        }

        private static bool TryReadStyledText(string source, int start,
            out IReadOnlyList<LaTeXContentSegment> styled, out int end)
        {
            styled = null;
            end = start;
            string command = null;
            bool bold = false;
            bool italic = false;
            var fontFamily = LaTeXTextFontFamily.Inherited;
            foreach (var candidate in new[] { "textit", "emph", "textbf", "textsf", "textrm", "texttt" })
            {
                var prefix = "\\" + candidate;
                if (!StartsWith(source, start, prefix) || IsEscaped(source, start)) continue;
                command = candidate;
                end = start + prefix.Length;
                break;
            }
            if (command == null) return false;
            while (end < source.Length && char.IsWhiteSpace(source[end])) end++;
            if (end >= source.Length || source[end] != '{') { end = start; return false; }
            var close = FindBalancedBrace(source, end);
            if (close < 0)
                throw new ArgumentException("The LaTeX text command \\" + command + " has no closing brace.");

            switch (command)
            {
                case "textit": case "emph": italic = true; break;
                case "textbf": bold = true; break;
                case "textsf": fontFamily = LaTeXTextFontFamily.SansSerif; break;
                case "textrm": fontFamily = LaTeXTextFontFamily.Roman; break;
                case "texttt": fontFamily = LaTeXTextFontFamily.Monospace; break;
            }
            var inner = Parse(source.Substring(end + 1, close - end - 1));
            var result = new List<LaTeXContentSegment>();
            foreach (var segment in inner)
                result.Add(segment.WithTextStyle(bold, italic, fontFamily));
            styled = result;
            end = close + 1;
            return true;
        }

        private static int FindBalancedBrace(string source, int open)
        {
            var depth = 0;
            for (var position = open; position < source.Length; position++)
            {
                if (source[position] == '%' && !IsEscaped(source, position))
                {
                    while (position < source.Length && source[position] != '\n') position++;
                    if (position >= source.Length) return -1;
                }
                if (source[position] == '{' && !IsEscaped(source, position)) depth++;
                else if (source[position] == '}' && !IsEscaped(source, position))
                {
                    depth--;
                    if (depth == 0) return position;
                }
            }
            return -1;
        }

        private static bool IsEscaped(string source, int position)
        {
            var slashes = 0;
            for (var index = position - 1; index >= 0 && source[index] == '\\'; index--) slashes++;
            return (slashes & 1) != 0;
        }

        private static bool StartsWith(string source, int start, string value)
        {
            return start >= 0 && start + value.Length <= source.Length &&
                string.CompareOrdinal(source, start, value, 0, value.Length) == 0;
        }

        private static void FlushText(List<LaTeXContentSegment> segments, StringBuilder text)
        {
            if (text.Length == 0) return;
            segments.Add(new LaTeXContentSegment(LaTeXContentKind.Text, text.ToString()));
            text.Clear();
        }
    }
}
