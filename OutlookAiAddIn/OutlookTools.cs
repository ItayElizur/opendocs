using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // COM-backed tool surface mirroring C:\dev\mcp-outlook, Explorer-only.
    // Dispatcher + shared helpers live here; per-area handlers in the partials.
    public static partial class OutlookTools
    {
        private static readonly Dictionary<string, EditingMode> ModeByMailbox = new Dictionary<string, EditingMode>();

        public static void SetMode(string mbxKey, EditingMode mode)
        {
            ModeByMailbox[mbxKey] = mode;
        }

        private static EditingMode ModeFor(string mbxKey)
        {
            EditingMode m;
            return ModeByMailbox.TryGetValue(mbxKey, out m) ? m : EditingMode.FullAutonomy;
        }

        // Comment only / Track changes have no meaning for mail - the mode gate
        // treats them as read-only. Only Full autonomy permits mutation.
        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "list_emails", "search_emails", "get_email", "list_folders", "search_contacts",
            "list_events", "get_event", "list_tasks", "get_attachment", "find_meeting_slots",
        };

        public static ToolResult Execute(string mbxKey, string name, JsonElement input)
        {
            try
            {
                DebugLog.Write("OutlookTools.Execute " + name);
                EditingMode mode = ModeFor(mbxKey);
                if (!AlwaysAllowedTools.Contains(name) && mode != EditingMode.FullAutonomy)
                {
                    return new ToolResult
                    {
                        Output = "Blocked: the assistant is in a read-only mode. Switch to Full autonomy to send drafts, move, delete, flag, create tasks/reminders, or respond to invites.",
                        IsError = true,
                        Summary = name,
                    };
                }

                switch (name)
                {
                    case "list_emails": return ListEmails(input);
                    case "search_emails": return SearchEmails(input);
                    case "get_email": return GetEmail(input);
                    case "list_folders": return ListFolders(input);
                    case "search_contacts": return SearchContacts(input);
                    case "list_events": return ListEvents(input);
                    case "get_event": return GetEvent(input);
                    case "find_meeting_slots": return FindMeetingSlots(input);
                    case "list_tasks": return ListTasks(input);
                    case "get_attachment": return GetAttachment(input);

                    case "mark_email_read": return MarkEmail(input, false);
                    case "mark_email_unread": return MarkEmail(input, true);
                    case "flag_email_important": return FlagEmailImportant(input);
                    case "move_email": return MoveEmail(input);
                    case "delete_email": return DeleteEmail(input);
                    case "accept_meeting": return RespondMeeting(input, true);
                    case "decline_meeting": return RespondMeeting(input, false);
                    case "create_task": return CreateTask(input);
                    case "update_task": return UpdateTask(input);
                    case "set_reminder": return SetReminder(input);
                    case "set_email_reminder": return SetEmailReminder(input);

                    case "draft_email": return DraftEmail(input);
                    case "reply_email": return ReplyEmail(input, false);
                    case "reply_all_email": return ReplyEmail(input, true);
                    case "forward_email": return ForwardEmail(input);
                    case "draft_event": return DraftEvent(input);

                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("OutlookTools.Execute " + name, ex);
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        // ---- shared COM helpers ----

        internal static Outlook.Application App { get { return Globals.ThisAddIn.Application; } }

        internal static Outlook.NameSpace Ns { get { return Globals.ThisAddIn.Application.Session; } }

        internal static object ItemById(string entryId, string storeId)
        {
            if (string.IsNullOrEmpty(entryId)) throw new ArgumentException("message_id / event_id / task_id is required.");
            return string.IsNullOrEmpty(storeId) ? Ns.GetItemFromID(entryId) : Ns.GetItemFromID(entryId, storeId);
        }

        private static readonly Dictionary<string, Outlook.OlDefaultFolders> WellKnown =
            new Dictionary<string, Outlook.OlDefaultFolders>(StringComparer.OrdinalIgnoreCase)
            {
                { "inbox", Outlook.OlDefaultFolders.olFolderInbox },
                { "sent", Outlook.OlDefaultFolders.olFolderSentMail },
                { "sent items", Outlook.OlDefaultFolders.olFolderSentMail },
                { "sent mail", Outlook.OlDefaultFolders.olFolderSentMail },
                { "drafts", Outlook.OlDefaultFolders.olFolderDrafts },
                { "deleted", Outlook.OlDefaultFolders.olFolderDeletedItems },
                { "deleted items", Outlook.OlDefaultFolders.olFolderDeletedItems },
                { "trash", Outlook.OlDefaultFolders.olFolderDeletedItems },
                { "junk", Outlook.OlDefaultFolders.olFolderJunk },
                { "junk email", Outlook.OlDefaultFolders.olFolderJunk },
                { "outbox", Outlook.OlDefaultFolders.olFolderOutbox },
            };

        internal static Outlook.Folder ResolveFolder(string name)
        {
            if (string.IsNullOrEmpty(name))
                return (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.OlDefaultFolders def;
            if (WellKnown.TryGetValue(name.Trim(), out def))
                return (Outlook.Folder)Ns.GetDefaultFolder(def);

            Outlook.Folder found = FindFolderByName(Ns.Folders, name.Trim(), 0);
            if (found == null) throw new ArgumentException("Folder not found: " + name + ". Call list_folders to see available names.");
            return found;
        }

        private static Outlook.Folder FindFolderByName(Outlook.Folders folders, string name, int depth)
        {
            if (folders == null || depth > 8) return null;
            foreach (Outlook.Folder f in folders)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase)) return f;
                Outlook.Folder child = FindFolderByName(f.Folders, name, depth + 1);
                if (child != null) return child;
            }
            return null;
        }

        internal static string SmtpOf(Outlook.AddressEntry ae)
        {
            if (ae == null) return "";
            try
            {
                if (ae.Type == "EX")
                {
                    Outlook.ExchangeUser eu = ae.GetExchangeUser();
                    if (eu != null && !string.IsNullOrEmpty(eu.PrimarySmtpAddress)) return eu.PrimarySmtpAddress;
                }
                object v = ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
                string s = v as string;
                if (!string.IsNullOrEmpty(s)) return s;
            }
            catch { }
            try { return ae.Address ?? ""; } catch { return ""; }
        }

        // The anti-slow-scan seam: search_emails pushes this DASL into
        // Items.Restrict rather than iterating Items. The builder itself is pure
        // and unit-tested in OfficeAi.Shared (OutlookDasl.BuildSearchFilter).
        internal static string BuildSearchDasl(string query, DateTime? start, DateTime? end, string sender)
        {
            return OutlookDasl.BuildSearchFilter(query, start, end, sender);
        }

        // ---- JSON arg readers ----

        internal static string Str(JsonElement o, string name, string dflt)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind == JsonValueKind.String)
                return v.GetString();
            return dflt;
        }

        internal static string ReqStr(JsonElement o, string name)
        {
            string s = Str(o, name, null);
            if (string.IsNullOrEmpty(s)) throw new ArgumentException("Required field \"" + name + "\" is missing.");
            return s;
        }

        internal static int Int(JsonElement o, string name, int dflt)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind == JsonValueKind.Number)
            {
                int n;
                if (v.TryGetInt32(out n)) return n;
            }
            return dflt;
        }

        internal static bool Bool(JsonElement o, string name, bool dflt)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v))
            {
                if (v.ValueKind == JsonValueKind.True) return true;
                if (v.ValueKind == JsonValueKind.False) return false;
            }
            return dflt;
        }

        internal static DateTime? DateArg(JsonElement o, string name)
        {
            string s = Str(o, name, null);
            if (string.IsNullOrEmpty(s)) return null;
            DateTime d;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out d)) return d;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out d)) return d;
            throw new ArgumentException("Field \"" + name + "\" is not a valid date/time: " + s);
        }

        internal static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            return s.Length > max ? s.Substring(0, max) + "\n...[truncated]" : s;
        }

        internal static string Iso(DateTime d)
        {
            return d.ToString("o", CultureInfo.InvariantCulture);
        }
    }
}
