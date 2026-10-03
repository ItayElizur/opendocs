using System.IO;

namespace OfficeAi.Shared.AttachmentText
{
    // The single seam for turning a saved email attachment into plain text for
    // the model (TryExtract). Handles text-family types and OpenXML
    // (.docx/.xlsx/.pptx) only; .pdf and images return false. See
    // AttachmentTextExtractor.cs.md for why this is structured as a seam.
    public static class AttachmentTextExtractor
    {
        public const int MaxChars = 40000;

        public static bool TryExtract(string path, out string text, out bool truncated)
        {
            text = null;
            truncated = false;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return false;

            string ext = Path.GetExtension(path).ToLowerInvariant();
            string raw;
            switch (ext)
            {
                case ".txt":
                case ".csv":
                case ".tsv":
                case ".md":
                case ".json":
                case ".xml":
                case ".log":
                    raw = PlainTextReader.ReadText(path);
                    break;
                case ".html":
                case ".htm":
                    raw = PlainTextReader.ReadHtml(path);
                    break;
                case ".docx":
                    raw = OpenXmlReader.ReadWord(path);
                    break;
                case ".xlsx":
                    raw = OpenXmlReader.ReadExcel(path);
                    break;
                case ".pptx":
                    raw = OpenXmlReader.ReadPowerPoint(path);
                    break;
                default:
                    return false;
            }

            if (raw == null) return false;
            if (raw.Length > MaxChars)
            {
                raw = raw.Substring(0, MaxChars);
                truncated = true;
            }
            text = raw;
            return true;
        }
    }
}
