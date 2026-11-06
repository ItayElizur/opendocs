using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeAi.Shared
{
    // One piece of header/footer text: plain text, or (Field != null) a field to insert.
    public sealed class FieldSegment
    {
        public string Text;  // plain text; empty for a field
        public string Field; // PAGE, NUMPAGES, ... ; null for plain text
    }

    // The small set of Word fields a header/footer can carry as plain-text tokens, e.g. "Page {PAGE} of {NUMPAGES}".
    // Pure string logic so it can be unit-tested; the COM side lives in WordTools.ObjectEdits.cs.
    public static class FieldTokens
    {
        private static readonly Regex Token = new Regex(@"\{(PAGE|NUMPAGES|DATE|TIME|FILENAME|AUTHOR|TITLE)\}", RegexOptions.Compiled);
        // A field as Range.Text shows it with field codes on: begin(0x13) code [separator(0x14) result] end(0x15).
        private static readonly Regex CodedField = new Regex("\u0013([^\u0013\u0014\u0015]*)(?:\u0014([^\u0013\u0015]*))?\u0015", RegexOptions.Compiled);
        private static readonly Regex FieldName = new Regex(@"^\s*([A-Za-z]+)", RegexOptions.Compiled);

        // Field names written in the text, in order of appearance (a name repeats if its token does).
        public static List<string> Find(string text)
        {
            var names = new List<string>();
            foreach (Match m in Token.Matches(text ?? "")) names.Add(m.Groups[1].Value);
            return names;
        }

        // Cuts the text into plain parts and fields, in order, so a header can be built left to right: insert the text,
        // add the field, insert the next text. Nothing is searched for in the document afterwards.
        public static List<FieldSegment> Split(string text)
        {
            var parts = new List<FieldSegment>();
            text = text ?? "";
            int pos = 0;
            foreach (Match m in Token.Matches(text))
            {
                if (m.Index > pos) parts.Add(new FieldSegment { Text = text.Substring(pos, m.Index - pos) });
                parts.Add(new FieldSegment { Text = "", Field = m.Groups[1].Value });
                pos = m.Index + m.Length;
            }
            if (pos < text.Length) parts.Add(new FieldSegment { Text = text.Substring(pos) });
            return parts;
        }

        // Turns Range.Text (read with field codes included) into text with {NAME} tokens. Fields that have no token
        // keep their displayed result. The field name is the leading letters of the code, so "DATE \@ ..." and "DATE\@ ..."
        // both read as DATE. Nested fields are resolved from the inside out; any control character left over is dropped.
        public static string FromFieldCodes(string rangeText)
        {
            string text = rangeText ?? "";
            for (int pass = 0; pass < 5 && CodedField.IsMatch(text); pass++)
            {
                text = CodedField.Replace(text, m =>
                {
                    Match n = FieldName.Match(m.Groups[1].Value);
                    string name = n.Success ? n.Groups[1].Value.ToUpperInvariant() : "";
                    return Token.IsMatch("{" + name + "}") ? "{" + name + "}" : m.Groups[2].Value;
                });
            }
            return text.Replace("\u0013", "").Replace("\u0014", "").Replace("\u0015", "");
        }
    }
}
