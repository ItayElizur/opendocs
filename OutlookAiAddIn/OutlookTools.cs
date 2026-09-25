using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Threading.Tasks;
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

        // Falls back to CommentOnly (Outlook's "Draft only" tier - see
        // AlwaysAllowedTools' comment below), matching the client's own new
        // default (chat-ui.ts's defaultModeFor()), for the brief window
        // before the client's first explicit "set-mode" bridge message
        // arrives (only sent on a user-initiated mode change, never on
        // mount). Previously fell back to FullAutonomy, which was safe only
        // because FullAutonomy was also the client's own default at the
        // time - no longer true.
        private static EditingMode ModeFor(string mbxKey)
        {
            EditingMode m;
            return ModeByMailbox.TryGetValue(mbxKey, out m) ? m : EditingMode.CommentOnly;
        }

        // Outlook repurposes two EditingMode slots Word/Excel/PowerPoint use
        // for real editing concepts that have no meaning for mail:
        // CommentOnly -> "Draft only" (draft/mutate freely, nothing sends),
        // TrackChanges -> "Automate approvals" (adds accept/decline, which
        // already auto-notify the organizer via resp.Send()). FullAutonomy
        // adds the tools that compose and send/create new content with no
        // review step. Ordinal check below relies on the enum's declared
        // order (ReadOnly < CommentOnly < TrackChanges < FullAutonomy)
        // matching this escalation exactly - see OfficeAi.Shared.EditingMode.
        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "list_emails", "search_emails", "get_email", "list_folders", "search_contacts",
            "list_events", "get_event", "list_tasks", "get_attachment", "find_meeting_slots",
            "list_color_categories",
        };

        // Tier 2 ("Draft only" / CommentOnly): mutates the mailbox or opens a
        // draft, but never leaves it unreviewed. set_event_categories/
        // set_category_color belong here, not in SendTierTools below - both
        // are purely local (appt.Categories/.Save(), cats.Add()/.Color),
        // never call .Send(), and carry the same risk profile as
        // move_email/flag_email_important right next to them. An earlier
        // version of this fix put them in SendTierTools to match their old
        // (pre-four-tier) Full-Autonomy-only gate, but that was restoring
        // the OLD binary model rather than applying this PR's own tiering
        // logic - every other local-only mutation here was deliberately
        // downgraded from Full-Autonomy-only, and these two were simply
        // missed, not deliberately kept stricter. Must stay in sync with
        // entry.ts's commentOnlyExtraTools.
        private static readonly HashSet<string> DraftTierTools = new HashSet<string>
        {
            "mark_email_read", "mark_email_unread", "flag_email_important", "move_email", "delete_email",
            "create_task", "update_task", "set_reminder", "set_email_reminder",
            "draft_email", "reply_email", "reply_all_email", "forward_email", "draft_event",
            "set_event_categories", "set_category_color",
        };

        // Tier 3 ("Automate approvals" / TrackChanges): already calls
        // resp.Send() to notify the organizer (RespondMeeting) - kept out of
        // the draft tier so "Draft only" honestly means nothing sends.
        private static readonly HashSet<string> ApprovalTierTools = new HashSet<string>
        {
            "accept_meeting", "decline_meeting",
        };

        // Tier 4 (Full autonomy only): composes and sends/creates brand-new
        // content with no review step at all.
        private static readonly HashSet<string> SendTierTools = new HashSet<string>
        {
            "send_email", "send_reply", "send_reply_all", "send_forward", "create_event",
        };

        private static string TierLabel(EditingMode mode)
        {
            switch (mode)
            {
                case EditingMode.FullAutonomy: return "Full autonomy";
                case EditingMode.TrackChanges: return "Automate approvals";
                case EditingMode.CommentOnly: return "Draft only";
                default: return "Read only";
            }
        }

        public static async Task<ToolResult> ExecuteAsync(string mbxKey, string name, JsonElement input)
        {
            try
            {
                DebugLog.Write("OutlookTools.Execute " + name);
                EditingMode mode = ModeFor(mbxKey);
                if (!AlwaysAllowedTools.Contains(name))
                {
                    EditingMode required =
                        SendTierTools.Contains(name) ? EditingMode.FullAutonomy :
                        ApprovalTierTools.Contains(name) ? EditingMode.TrackChanges :
                        EditingMode.CommentOnly; // DraftTierTools, and anything else not otherwise classified
                    if ((int)mode < (int)required)
                    {
                        return new ToolResult
                        {
                            Output = "Blocked: this action requires " + TierLabel(required) + " mode or higher (currently " + TierLabel(mode) + ").",
                            IsError = true,
                            Summary = name,
                        };
                    }
                }

                switch (name)
                {
                    case "list_emails": return ListEmails(input);
                    case "search_emails": return SearchEmails(input);
                    case "get_email": return GetEmail(input);
                    case "list_folders": return ListFolders(input);
                    case "search_contacts": return await SearchContactsAsync(input);
                    case "list_events": return ListEvents(input);
                    case "get_event": return GetEvent(input);
                    case "find_meeting_slots": return await FindMeetingSlotsAsync(input);
                    case "list_tasks": return ListTasks(input);
                    case "get_attachment": return GetAttachment(input);
                    case "list_color_categories": return ListColorCategories(input);

                    case "mark_email_read": return MarkEmail(input, false);
                    case "mark_email_unread": return MarkEmail(input, true);
                    case "flag_email_important": return FlagEmailImportant(input);
                    case "move_email": return MoveEmail(input);
                    case "delete_email": return DeleteEmail(input);
                    case "accept_meeting": return RespondMeeting(input, true);
                    case "decline_meeting": return RespondMeeting(input, false);
                    case "set_event_categories": return SetEventCategories(input);
                    case "set_category_color": return SetCategoryColor(input);
                    case "create_task": return CreateTask(input);
                    case "update_task": return UpdateTask(input);
                    case "set_reminder": return SetReminder(input);
                    case "set_email_reminder": return SetEmailReminder(input);

                    case "draft_email": return DraftEmail(input);
                    case "reply_email": return ReplyEmail(input, false);
                    case "reply_all_email": return ReplyEmail(input, true);
                    case "forward_email": return ForwardEmail(input);
                    case "draft_event": return DraftEvent(input);

                    case "send_email": return SendEmail(input);
                    case "send_reply": return SendReply(input, false);
                    case "send_reply_all": return SendReply(input, true);
                    case "send_forward": return SendForward(input);
                    case "create_event": return CreateEvent(input);

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

        // Like Int, but the default itself can be fractional - for
        // find_meeting_slots' start_hour/end_hour, whose default comes from a
        // mailbox's real (possibly non-hour-aligned, e.g. 08:30) EWS working
        // hours. The argument itself is still schema'd as a whole-hour
        // integer, so only the default needs the fractional path.
        internal static double Double(JsonElement o, string name, double dflt)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind == JsonValueKind.Number)
            {
                double n;
                if (v.TryGetDouble(out n)) return n;
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
