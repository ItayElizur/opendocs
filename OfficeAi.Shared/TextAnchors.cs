using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // Finds a phrase in a Word paragraph's raw text the way a reader sees it. Range.Text carries control characters
    // that are not part of the visible words - footnote/endnote reference marks (0x02), field begin/separator/end
    // marks (0x13-0x15), inline-picture anchors (0x01) - and they sit between words, so a plain search for
    // "the mouse" misses "the<0x02> mouse". Clean() drops them and remembers where each kept character came from.
    public sealed class CleanText
    {
        public string Text;
        public int[] Map; // Map[i] = index in the raw string of Text[i]
    }

    public static class TextAnchors
    {
        public static CleanText Clean(string raw)
        {
            raw = raw ?? "";
            var sb = new StringBuilder(raw.Length);
            var map = new List<int>(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                if (c < ' ' && c != '\t') continue;
                sb.Append(c == ' ' ? ' ' : c);
                map.Add(i);
            }
            return new CleanText { Text = sb.ToString(), Map = map.ToArray() };
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
