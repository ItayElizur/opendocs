using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeAi.Shared.AttachmentText
{
    internal static class PlainTextReader
    {
        public static string ReadText(string path)
        {
            // detectEncodingFromByteOrderMarks: true - honors a UTF-8/UTF-16 BOM,
            // falls back to UTF-8 otherwise (the common case for these types).
            using (StreamReader reader = new StreamReader(path, Encoding.UTF8, true))
            {
                return reader.ReadToEnd();
            }
        }

        private static readonly Regex ScriptStyle = new Regex(
            @"<(script|style)\b[^>]*>.*?</\1>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private static readonly Regex Tags = new Regex(@"<[^>]+>", RegexOptions.Singleline);

        private static readonly Regex ExcessBlankLines = new Regex(@"(\s*\r?\n){3,}");

        public static string ReadHtml(string path)
        {
            string html = ReadText(path);
            html = ScriptStyle.Replace(html, " ");
            html = Tags.Replace(html, " ");
            html = WebUtility.HtmlDecode(html);
            html = ExcessBlankLines.Replace(html, "\n\n");
            return html.Trim();
        }
    }
}
