using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    public static class SafeDraftHtml
    {
        public static string Format(
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            return FormatContent(body, boldPhrases).Html;
        }

        public static SafeDraftContent FormatContent(
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            var text = TextBoundary.PlainText(
                body,
                TextBoundary.MaxAssistantCharacters);
            if (text.Length == 0)
            {
                throw new InvalidOperationException(
                    "A non-empty draft body is required.");
            }

            List<TextRange> markdownRanges;
            text = NormalizeMarkdownStrong(
                text,
                out markdownRanges);
            var ranges = FindBoldRanges(text, boldPhrases);
            ranges.AddRange(markdownRanges);
            ranges = MergeRanges(ranges);
            var html = new StringBuilder(text.Length + 128);
            html.Append(
                "<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt;\">");
            var position = 0;
            foreach (var range in ranges)
            {
                AppendEncoded(html, text.Substring(
                    position,
                    range.Start - position));
                html.Append("<strong>");
                AppendEncoded(html, text.Substring(
                    range.Start,
                    range.Length));
                html.Append("</strong>");
                position = range.Start + range.Length;
            }

            AppendEncoded(html, text.Substring(position));
            html.Append("</div>");
            return new SafeDraftContent(
                text,
                html.ToString());
        }

        private static string NormalizeMarkdownStrong(
            string text,
            out List<TextRange> boldRanges)
        {
            boldRanges = new List<TextRange>();
            var normalized = new StringBuilder(text.Length);
            var position = 0;
            var lineStart = true;
            while (position < text.Length)
            {
                if (lineStart &&
                    position + 1 < text.Length &&
                    text[position] == '*' &&
                    text[position + 1] == ' ')
                {
                    normalized.Append("- ");
                    position += 2;
                    lineStart = false;
                    continue;
                }

                if (position + 3 < text.Length &&
                    ((text[position] == '*' &&
                      text[position + 1] == '*') ||
                     (text[position] == '_' &&
                      text[position + 1] == '_')))
                {
                    var marker = text.Substring(position, 2);
                    var close = text.IndexOf(
                        marker,
                        position + 2,
                        StringComparison.Ordinal);
                    if (close > position + 2)
                    {
                        var content = text.Substring(
                            position + 2,
                            close - position - 2);
                        var start = normalized.Length;
                        normalized.Append(content);
                        boldRanges.Add(
                            new TextRange(start, content.Length));
                        lineStart = content.EndsWith(
                            "\n",
                            StringComparison.Ordinal) ||
                            content.EndsWith(
                                "\r",
                                StringComparison.Ordinal);
                        position = close + 2;
                        continue;
                    }
                }

                var value = text[position];
                normalized.Append(value);
                lineStart = value == '\n' || value == '\r';
                position++;
            }

            return normalized.ToString();
        }

        private static List<TextRange> FindBoldRanges(
            string text,
            IReadOnlyList<string> boldPhrases)
        {
            var candidates = new List<TextRange>();
            if (boldPhrases == null)
            {
                return candidates;
            }

            var phraseCount = Math.Min(12, boldPhrases.Count);
            for (var phraseIndex = 0;
                 phraseIndex < phraseCount;
                 phraseIndex++)
            {
                var phrase = TextBoundary.PlainText(
                    boldPhrases[phraseIndex],
                    200);
                if (phrase.Length == 0)
                {
                    continue;
                }

                var start = 0;
                while (start < text.Length)
                {
                    var found = text.IndexOf(
                        phrase,
                        start,
                        StringComparison.OrdinalIgnoreCase);
                    if (found < 0)
                    {
                        break;
                    }

                    candidates.Add(
                        new TextRange(found, phrase.Length));
                    start = found + phrase.Length;
                }
            }

            return MergeRanges(candidates);
        }

        private static List<TextRange> MergeRanges(
            List<TextRange> candidates)
        {
            candidates.Sort((left, right) =>
            {
                var startComparison = left.Start.CompareTo(right.Start);
                return startComparison != 0
                    ? startComparison
                    : right.Length.CompareTo(left.Length);
            });

            var accepted = new List<TextRange>();
            var end = -1;
            foreach (var candidate in candidates)
            {
                if (candidate.Start < end)
                {
                    continue;
                }

                accepted.Add(candidate);
                end = candidate.Start + candidate.Length;
            }

            return accepted;
        }

        private static void AppendEncoded(
            StringBuilder output,
            string value)
        {
            output.Append(
                WebUtility.HtmlEncode(value)
                    .Replace("\r\n", "<br>")
                    .Replace("\n", "<br>")
                    .Replace("\r", "<br>"));
        }

        private sealed class TextRange
        {
            public TextRange(int start, int length)
            {
                Start = start;
                Length = length;
            }

            public int Start { get; }

            public int Length { get; }
        }
    }

    public sealed class SafeDraftContent
    {
        internal SafeDraftContent(
            string plainText,
            string html)
        {
            PlainText = plainText ?? string.Empty;
            Html = html ?? string.Empty;
        }

        public string PlainText { get; }

        public string Html { get; }
    }
}
