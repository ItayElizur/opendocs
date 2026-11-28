using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // Finds a phrase in a Word paragraph's raw text the way a reader sees it. Range.Text carries control characters
    // that are not part of the visible words - footnote/endnote reference marks (0x02), field begin/separator/end
    // marks (0x13-0x15), inline-picture anchors (0x01) - and they sit between words, so a plain search for
    // "the mouse" misses "the<0x02> mouse". Clean() drops them and remembers where each kept character came from.
    // Read with field codes included (which keeps every character at its document position), a field also carries its
    // code text ("PAGE") before the result the reader sees ("3"); the code part is dropped too, so only what is
    // displayed can be matched.
    public sealed class CleanText
    {
        public string Text;
        public int[] Map; // Map[i] = index in the raw string of Text[i]
    }

    public static class TextAnchors
    {
        // keepBreaks: paragraph marks, line breaks and cell ends (\r, \v) become "\n" instead of being dropped, so a whole
        // body can be searched across paragraphs.
        public static CleanText Clean(string raw, bool keepBreaks = false)
        {
            raw = raw ?? "";
            var sb = new StringBuilder(raw.Length);
            var map = new List<int>(raw.Length);
            // one entry per open field: false = inside its code part (not displayed), true = inside its result
            var fields = new Stack<bool>();
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                if (c == '\u0013') { fields.Push(false); continue; }
                if (c == '\u0014') { if (fields.Count > 0) { fields.Pop(); fields.Push(true); } continue; }
                if (c == '\u0015') { if (fields.Count > 0) fields.Pop(); continue; }
                if (fields.Contains(false)) continue;
                if (keepBreaks && (c == '\r' || c == '\v')) { sb.Append('\n'); map.Add(i); continue; }
                if (c < ' ' && c != '\t') continue;
                sb.Append(c == ' ' ? ' ' : c);
                map.Add(i);
            }
            return new CleanText { Text = sb.ToString(), Map = map.ToArray() };
        }

        // Range.Text is only a usable map of document positions when it has one character per position. The one known
        // exception is the end-of-cell mark, which reads "\r\a" (two characters) but occupies a single position; every
        // character before it still lines up.
        public static bool PositionsAlign(string raw, int span)
        {
            raw = raw ?? "";
            return raw.Length == span || (raw.EndsWith("\r\a", StringComparison.Ordinal) && raw.Length == span + 1);
        }

        // The document position of each character of a whole story's Range.Text: one per character, except that an
        // end-of-cell or end-of-row mark reads "\r\a" but is one position, so both of its characters get the same one.
        public static int[] DocumentPositions(string raw)
        {
            raw = raw ?? "";
            var pos = new int[raw.Length];
            int shift = 0;
            for (int i = 0; i < raw.Length; i++)
            {
                if (raw[i] == '\a' && i > 0 && raw[i - 1] == '\r') shift++;
                pos[i] = i - shift;
            }
            return pos;
        }

        // Start indices (into text) of non-overlapping, case-insensitive matches, at most max of them.
        public static List<int> FindAll(string text, string needle, int max)
        {
            var found = new List<int>();
            if (string.IsNullOrEmpty(needle)) return found;
            needle = needle.Replace(' ', ' ');
            int from = 0, i;
            while (found.Count < max && from <= text.Length && (i = text.IndexOf(needle, from, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                found.Add(i);
                from = i + needle.Length;
            }
            return found;
        }

        // "...before[[match]]after..." with up to `context` characters each side; the match is marked so close-together
        // matches stay distinguishable.
        public static string Snippet(string text, int start, int length, int context)
        {
            int from = Math.Max(0, start - context);
            int to = Math.Min(text.Length, start + length + context);
            return "..." + text.Substring(from, start - from) + "[[" + text.Substring(start, length) + "]]"
                + text.Substring(start + length, to - start - length) + "...";
        }
    }
}
