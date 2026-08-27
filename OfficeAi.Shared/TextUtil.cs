using System;
using System.Text;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure string helpers shared by the Word/Excel/PowerPoint tool layers.
    /// Extracted here (Phase 0) so they are unit-testable at all - the
    /// *Tools.cs files live in VSTO projects whose private members no test
    /// project can reach.
    ///
    /// SEAM RULES for this assembly (OfficeAi.Shared) - read before adding
    /// anything new here:
    ///
    /// 1. Anything free of COM/Office-interop types goes here, and gets a
    ///    test. Anything touching Word.*/Excel.*/PowerPoint.* stays in its
    ///    app's own *Tools.cs - this assembly has no interop seam for actual
    ///    COM calls yet (that is a future phase, not this one). Without this
    ///    rule written down, the next helper gets added to a *Tools.cs out of
    ///    habit and this seam quietly stops growing.
    ///
    /// 2. NEVER expose an Office interop type as a generic type argument from
    ///    this assembly - e.g. Dictionary&lt;string, MsoAutoShapeType&gt; as a
    ///    public member. This assembly embeds the Office PIA
    ///    (EmbedInteropTypes=true, same as every app project), and an
    ///    embedded interop type used as a generic type argument cannot cross
    ///    an assembly boundary - confirmed via a spike while building
    ///    ShapeTypes.cs (Phase 0 Task 4): CS1769, "cannot be used across
    ///    assembly boundaries because it has a generic type argument that is
    ///    an embedded interop type." The type is fine used BARE (a plain
    ///    parameter or return value, not inside a generic), just never as a
    ///    generic argument.
    ///    Carry it as int/string instead, and cast at the app-side call site
    ///    - see ColorUtil.HexToOle and ShapeTypes.ByName, both int-valued for
    ///    exactly this reason. This costs a build cycle to rediscover if you
    ///    don't already know it; don't make the next person hit it head-on.
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
