using System;
using System.Text;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure string helpers shared by the Word/Excel/PowerPoint tool layers.
    /// Before adding anything new here, read the seam rules in
    /// TextUtil.cs.md - in short: no COM/Office-interop types, and never
    /// expose one as a generic type argument (carry int/string and cast at
    /// the app-side call site instead; see ColorUtil.HexToOle, ShapeTypes.ByName).
    /// </summary>
    public static class TextUtil
    {
        // Order matters: & must be escaped first, or a literal "<" would
        // become "&amp;lt;" instead of "&lt;".
        public static string HtmlEscape(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        public static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }

        // .NET Framework 4.8 has no String.Replace(string, string, StringComparison)
        // overload - hence the hand-rolled scan.
        public static string ReplaceAllOccurrences(string input, string find, string replace, StringComparison comparison)
        {
            if (comparison == StringComparison.Ordinal) return input.Replace(find, replace);
            var sb = new StringBuilder();
            int pos = 0;
            while (true)
            {
                int idx = input.IndexOf(find, pos, comparison);
                if (idx < 0) { sb.Append(input, pos, input.Length - pos); break; }
                sb.Append(input, pos, idx - pos);
                sb.Append(replace);
                pos = idx + find.Length;
            }
            return sb.ToString();
        }

        public static int CountOccurrences(string haystack, string needle, StringComparison comparison)
        {
            if (string.IsNullOrEmpty(needle)) return 0;
            int count = 0, pos = 0;
            while (true)
            {
                int idx = haystack.IndexOf(needle, pos, comparison);
                if (idx < 0) break;
                count++;
                pos = idx + needle.Length;
            }
            return count;
        }

        // PowerPoint's TextRange never auto-flips paragraph direction the way
        // Word's editor does with "detect language automatically" - direction
        // has to be decided per write from the text's own script mix.
        public static bool IsRtlMajority(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            int rtl = 0, ltr = 0;
            foreach (char c in text)
            {
                bool isRtlChar = (c >= '֐' && c <= '׿') // Hebrew
                    || (c >= '؀' && c <= 'ۿ') // Arabic
                    || (c >= 'ݐ' && c <= 'ݿ') // Arabic Supplement
                    || (c >= 'יִ' && c <= '﷿') // Hebrew/Arabic presentation forms A
                    || (c >= 'ﹰ' && c <= '﻿'); // Arabic presentation forms B
                if (isRtlChar) rtl++;
                else if (char.IsLetter(c)) ltr++;
            }
            return rtl > 0 && rtl >= ltr;
        }
    }
}
