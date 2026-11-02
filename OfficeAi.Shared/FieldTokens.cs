using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeAi.Shared
{
    // The small set of Word fields a header/footer can carry as plain-text tokens, e.g. "Page {PAGE} of {NUMPAGES}".
    // Pure string logic so it can be unit-tested; the COM side lives in WordTools.ObjectEdits.cs.
    public static class FieldTokens
    {
        private static readonly Regex Token = new Regex(@"\{(PAGE|NUMPAGES|DATE|TIME|FILENAME|AUTHOR|TITLE)\}", RegexOptions.Compiled);
        // A field as Range.Text shows it with field codes on: begin(0x13) code [separator(0x14) result] end(0x15).
        private static readonly Regex CodedField = new Regex("\u0013([^\u0013\u0014\u0015]*)(?:\u0014([^\u0013\u0015]*))?\u0015", RegexOptions.Compiled);

        // Field names written in the text, in order of appearance (a name repeats if its token does).
        public static List<string> Find(string text)
        {
            var names = new List<string>();
            foreach (Match m in Token.Matches(text ?? "")) names.Add(m.Groups[1].Value);
            return names;
        }

        // Turns Range.Text (read with field codes included) into text with {NAME} tokens. Fields that have no token
        // keep their displayed result.
        public static string FromFieldCodes(string rangeText)
        {
            return CodedField.Replace(rangeText ?? "", m =>
            {
                string code = m.Groups[1].Value.Trim();
                int space = code.IndexOf(' ');
                string name = (space < 0 ? code : code.Substring(0, space)).ToUpperInvariant();
                return Token.IsMatch("{" + name + "}") ? "{" + name + "}" : m.Groups[2].Value;
            });
        }
    }
}
