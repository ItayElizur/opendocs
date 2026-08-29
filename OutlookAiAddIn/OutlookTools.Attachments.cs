using System;
using System.IO;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using OfficeAi.Shared.AttachmentText;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Saves the attachment to disk and, for text-family and OpenXML types
        // only, returns extracted text. PDF / images / other binaries return
        // the path only - all parsing lives in OfficeAi.Shared.AttachmentText,
        // the swappable module (later: an HTTP call to the user's parser API).
        private static ToolResult GetAttachment(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            int index = Int(input, "attachment_index", -1);
            if (index < 1)
                return new ToolResult { Output = "attachment_index (1-based) is required. See get_email's attachments list.", IsError = true, Summary = "get_attachment" };

            Outlook.MailItem mail = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (mail == null)
                return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "get_attachment" };
            if (mail.Attachments == null || index > mail.Attachments.Count)
                return new ToolResult { Output = "No attachment at index " + index + " on that message.", IsError = true, Summary = "get_attachment" };

            Outlook.Attachment att = mail.Attachments[index];
            string typeName = AttachmentTypeName(att.Type);

            if (att.Type == Outlook.OlAttachmentType.olByReference)
                return new ToolResult { Output = "Attachment " + index + " is a link (olByReference) with no downloadable content.", IsError = true, Summary = "get_attachment" };
            if (att.Type == Outlook.OlAttachmentType.olOLE)
                return new ToolResult { Output = "Attachment " + index + " is an OLE object that cannot be saved to a file.", IsError = true, Summary = "get_attachment" };

            string saveDir = Str(input, "save_dir", null);
            if (string.IsNullOrEmpty(saveDir))
                saveDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAiAddIn", "Attachments");
            Directory.CreateDirectory(saveDir);

            string fileName = att.FileName;
            if (string.IsNullOrEmpty(fileName)) fileName = att.DisplayName;
            if (string.IsNullOrEmpty(fileName)) fileName = "attachment-" + index;
            if (att.Type == Outlook.OlAttachmentType.olEmbeddeditem && !fileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                fileName += ".msg";

            string path = Path.Combine(saveDir, SanitizeFileName(fileName));
            att.SaveAsFile(path);

            long size = 0;
            try { size = new FileInfo(path).Length; } catch { }

            var sb = new StringBuilder();
            sb.AppendLine("path: " + path);
            sb.AppendLine("file_name: " + Path.GetFileName(path));
            sb.AppendLine("type: " + typeName);
            sb.AppendLine("size: " + size);

            string text;
            bool truncated;
            if (AttachmentTextExtractor.TryExtract(path, out text, out truncated))
            {
                sb.AppendLine("truncated: " + truncated);
                sb.AppendLine();
                sb.AppendLine("extracted_text:");
                sb.Append(text);
            }
            else
            {
                sb.AppendLine("extracted_text: (not extracted - PDF, image, or other binary. The file is saved at the path above for an external parser.)");
            }

            return new ToolResult { Output = sb.ToString(), Summary = "get_attachment" };
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name;
        }
    }
}
