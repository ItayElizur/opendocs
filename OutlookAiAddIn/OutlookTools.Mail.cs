using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        private const string PrHasAttach = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";

        // Fast bulk read via Folder.GetTable - an in-memory rowset, no per-item
        // COM object. Named-property columns plus the has-attachments proptag.
        private static ToolResult ListEmails(JsonElement input)
        {
            string folderName = Str(input, "folder", "inbox");
            int limit = Math.Max(1, Int(input, "limit", 20));
            bool unreadOnly = Bool(input, "unread_only", false);

            Outlook.Folder folder = ResolveFolder(folderName);
            Outlook.Table table = unreadOnly
                ? folder.GetTable("[UnRead] = true", Outlook.OlTableContents.olUserItems)
                : folder.GetTable(Type.Missing, Outlook.OlTableContents.olUserItems);

            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");
            table.Columns.Add("Subject");
            table.Columns.Add("ReceivedTime");
            table.Columns.Add("SenderName");
            table.Columns.Add("UnRead");
            table.Columns.Add("MessageClass");
            table.Columns.Add(PrHasAttach);
            table.Sort("ReceivedTime", true);

            var sb = new StringBuilder();
            int n = 0;
            while (!table.EndOfTable && n < limit)
            {
                Outlook.Row row = table.GetNextRow();
                string cls = Convert.ToString(row["MessageClass"], CultureInfo.InvariantCulture) ?? "";
                if (!cls.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)) continue;

                n++;
                string received = "";
                try { received = Iso(Convert.ToDateTime(row["ReceivedTime"], CultureInfo.InvariantCulture)); } catch { }
                bool unread = false;
                try { unread = Convert.ToBoolean(row["UnRead"]); } catch { }
                bool hasAtt = false;
                try { hasAtt = Convert.ToBoolean(row[PrHasAttach]); } catch { }

                sb.AppendLine("- message_id: " + Convert.ToString(row["EntryID"], CultureInfo.InvariantCulture));
                sb.AppendLine("  subject: " + (Convert.ToString(row["Subject"], CultureInfo.InvariantCulture) ?? ""));
                sb.AppendLine("  from: " + (Convert.ToString(row["SenderName"], CultureInfo.InvariantCulture) ?? ""));
                sb.AppendLine("  received: " + received + "  unread: " + unread + "  has_attachments: " + hasAtt);
            }

            if (n == 0) return new ToolResult { Output = "No messages in " + folder.Name + (unreadOnly ? " (unread only)." : "."), Summary = "list_emails" };
            return new ToolResult { Output = "Folder: " + folder.Name + "\n" + sb, Summary = "list_emails" };
        }

        private static ToolResult GetEmail(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string folderName = Str(input, "folder", null);
            string storeId = folderName != null ? ResolveFolder(folderName).StoreID : null;

            Outlook.MailItem mail = ItemById(id, storeId) as Outlook.MailItem;
            if (mail == null) return new ToolResult { Output = "That message_id does not resolve to a mail item.", IsError = true, Summary = "get_email" };

            var to = new List<string>();
            var cc = new List<string>();
            foreach (Outlook.Recipient r in mail.Recipients)
            {
                string label = r.Name;
                string addr = SmtpOf(r.AddressEntry);
                string entry = string.IsNullOrEmpty(addr) ? label : label + " <" + addr + ">";
                if (r.Type == (int)Outlook.OlMailRecipientType.olCC) cc.Add(entry);
                else if (r.Type == (int)Outlook.OlMailRecipientType.olTo) to.Add(entry);
            }

            var sb = new StringBuilder();
            sb.AppendLine("subject: " + (mail.Subject ?? ""));
            string fromAddr = "";
            try { fromAddr = SmtpOf(mail.Sender); } catch { }
            if (string.IsNullOrEmpty(fromAddr)) fromAddr = mail.SenderEmailAddress ?? "";
            sb.AppendLine("from: " + (mail.SenderName ?? "") + " <" + fromAddr + ">");
            sb.AppendLine("to: " + string.Join("; ", to));
            if (cc.Count > 0) sb.AppendLine("cc: " + string.Join("; ", cc));
            try { sb.AppendLine("received: " + Iso(mail.ReceivedTime)); } catch { }
            sb.AppendLine("conversation_id: " + (mail.ConversationID ?? ""));
            sb.AppendLine("conversation_topic: " + (mail.ConversationTopic ?? ""));
            sb.AppendLine("importance: " + mail.Importance);
            sb.AppendLine("unread: " + mail.UnRead);

            int attCount = mail.Attachments != null ? mail.Attachments.Count : 0;
            if (attCount > 0)
            {
                sb.AppendLine("attachments:");
                for (int i = 1; i <= attCount; i++)
                {
                    Outlook.Attachment a = mail.Attachments[i];
                    sb.AppendLine("  - index: " + a.Index + "  name: " + (a.FileName ?? a.DisplayName ?? "") +
                                  "  type: " + AttachmentTypeName(a.Type) + "  size: " + a.Size);
                }
            }

            sb.AppendLine();
            sb.AppendLine("body:");
            sb.Append(Truncate(mail.Body ?? "", 40000));
            return new ToolResult { Output = sb.ToString(), Summary = "get_email" };
        }

        internal static string AttachmentTypeName(Outlook.OlAttachmentType t)
        {
            switch (t)
            {
                case Outlook.OlAttachmentType.olByValue: return "byValue";
                case Outlook.OlAttachmentType.olByReference: return "reference";
                case Outlook.OlAttachmentType.olEmbeddeditem: return "embeddedItem";
                case Outlook.OlAttachmentType.olOLE: return "ole";
                default: return t.ToString();
            }
        }

        private static ToolResult MarkEmail(JsonElement input, bool unread)
        {
            string id = ReqStr(input, "message_id");
            Outlook.MailItem mail = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (mail == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "mark_email" };
            mail.UnRead = unread;
            mail.Save();
            return new ToolResult { Output = (unread ? "Marked unread: " : "Marked read: ") + (mail.Subject ?? ""), Mutated = true, Summary = unread ? "mark_email_unread" : "mark_email_read" };
        }

        private static ToolResult FlagEmailImportant(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            bool important = Bool(input, "important", true);
            Outlook.MailItem mail = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (mail == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "flag_email_important" };
            mail.Importance = important ? Outlook.OlImportance.olImportanceHigh : Outlook.OlImportance.olImportanceNormal;
            mail.Save();
            return new ToolResult { Output = (important ? "High importance: " : "Normal importance: ") + (mail.Subject ?? ""), Mutated = true, Summary = "flag_email_important" };
        }

        private static ToolResult MoveEmail(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string dest = ReqStr(input, "destination");
            Outlook.Folder target = ResolveFolder(dest);

            object item = ItemById(id, StoreOf(input));
            dynamic d = item;
            string oldSubject = "";
            try { oldSubject = d.Subject; } catch { }
            dynamic moved = d.Move(target);
            string newId = "";
            try { newId = moved.EntryID; } catch { }

            return new ToolResult
            {
                Output = "Moved \"" + oldSubject + "\" to " + target.Name + ".\nmessage_id: " + newId + "\nold_message_id: " + id,
                Mutated = true,
                Summary = "move_email",
            };
        }

        private static ToolResult DeleteEmail(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            bool permanent = Bool(input, "permanent", false);

            object item = ItemById(id, StoreOf(input));
            dynamic d = item;
            string subject = "";
            try { subject = d.Subject; } catch { }

            Outlook.Folder deleted = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            dynamic moved = d.Move(deleted);
            if (permanent)
            {
                try { moved.Delete(); }
                catch (Exception ex) { DebugLog.WriteException("DeleteEmail permanent", ex); }
                return new ToolResult { Output = "Permanently deleted \"" + subject + "\" (removed from Deleted Items; may still be server-recoverable).", Mutated = true, Summary = "delete_email" };
            }
            string newId = "";
            try { newId = moved.EntryID; } catch { }
            return new ToolResult { Output = "Moved \"" + subject + "\" to Deleted Items.\nmessage_id: " + newId, Mutated = true, Summary = "delete_email" };
        }

        internal static string StoreOf(JsonElement input)
        {
            string folderName = Str(input, "folder", null);
            return folderName != null ? ResolveFolder(folderName).StoreID : null;
        }
    }
}
