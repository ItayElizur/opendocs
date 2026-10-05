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

        // Falls back to CommentOnly ("Draft only"), matching the client's own
        // default, for the brief window before its first explicit "set-mode"
        // message arrives. See OutlookTools.cs.md for why the old
        // FullAutonomy fallback is no longer safe.
        private static EditingMode ModeFor(string mbxKey)
        {
            EditingMode m;
            return ModeByMailbox.TryGetValue(mbxKey, out m) ? m : EditingMode.CommentOnly;
        }

        // Outlook repurposes two EditingMode slots Word/Excel/PowerPoint use
        // for real editing concepts that have no meaning for mail:
        // CommentOnly -> "Draft only", TrackChanges -> "Automate approvals".
        // The ordinal check below relies on the enum's declared order
        // (ReadOnly < CommentOnly < TrackChanges < FullAutonomy) matching
        // this escalation exactly - see OfficeAi.Shared.EditingMode.
        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "list_emails", "search_emails", "get_email", "list_folders", "search_contacts",
            "list_events", "get_event", "list_tasks", "get_attachment", "find_meeting_slots", "open_email",
            "list_color_categories",
        };

        // Tier 2 ("Draft only" / CommentOnly): mutates the mailbox or opens a
        // draft, but never leaves it unreviewed. set_event_categories/
        // set_category_color/set_event_availability belong here rather than
        // SendTierTools - see OutlookTools.cs.md for why. apply_search is
        // here because it has a real, visible side effect (hijacks the
        // user's Explorer window) despite never mutating data - see
        // OutlookTools.cs.md. Must stay in sync with entry.ts's
        // commentOnlyExtraTools.
        private static readonly HashSet<string> DraftTierTools = new HashSet<string>
        {
            "mark_email_read", "mark_email_unread", "flag_email_important", "move_email", "delete_email",
            "create_task", "update_task", "set_reminder", "set_email_reminder",
            "draft_email", "draft_event",
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_cancel_event", "draft_edit_event",
            // undo/redo only replay the assistant's own recorded actions (see
            // OutlookTools.Undo.cs) and never send anything - sends and
            // meeting responses are barriers, not replayable entries - so
            // they sit at the Draft tier like the actions they mostly reverse.
            "undo_last_action", "redo_last_action",
        };

        // Tier 3 ("Automate approvals" / TrackChanges): already calls
        // resp.Send() to notify the organizer (RespondMeeting) - kept out of
        // the draft tier so "Draft only" honestly means nothing sends.
        private static readonly HashSet<string> ApprovalTierTools = new HashSet<string>
        {
            "respond_meeting",
        };

        // Tier 4 (Full autonomy only): composes and sends/creates brand-new
        // content with no review step at all.
        private static readonly HashSet<string> SendTierTools = new HashSet<string>
        {
            "send_email", "create_event", "cancel_event", "edit_event",
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

                    // delete_email spans two risk classes: permanent:false is
                    // fully reversible (Draft tier, like move_email);
                    // permanent:true is irreversible (needs the SendTierTools
                    // gate too, despite the tool NAME sitting in
                    // DraftTierTools). This is name-based gating's one
                    // input-aware exception - see OutlookTools.cs.md.
                    if (name == "delete_email" && Bool(input, "permanent", false) && (int)mode < (int)EditingMode.FullAutonomy)
                    {
                        return new ToolResult
                        {
                            Output = "Blocked: permanent delete requires " + TierLabel(EditingMode.FullAutonomy) + " mode or higher (currently " + TierLabel(mode) + "). Omit permanent, or set it to false, to move the message to Deleted Items instead.",
                            IsError = true,
                            Summary = name,
                        };
                    }
                }

                switch (name)
                {
                    case "list_emails": return ListEmails(input);
                    case "search_emails": return SearchEmails(input);
                    case "apply_search": return ApplySearch(input);
                    case "get_email": return GetEmail(input);
                    case "open_email": return OpenEmail(input);
                    case "list_folders": return ListFolders(input);
                    case "search_contacts": return await SearchContactsAsync(input);
                    case "list_events": return await ListEventsAsync(input);
                    case "get_event": return GetEvent(input);
                    case "find_meeting_slots": return await FindMeetingSlotsAsync(input);
                    case "list_tasks": return ListTasks(input);
                    case "get_attachment": return GetAttachment(input);
                    case "list_color_categories": return ListColorCategories(input);

                    case "mark_email_read": return MarkEmail(mbxKey, input, false);
                    case "mark_email_unread": return MarkEmail(mbxKey, input, true);
                    case "flag_email_important": return FlagEmailImportant(mbxKey, input);
                    case "move_email": return MoveEmail(mbxKey, input);
                    case "delete_email": return DeleteEmail(mbxKey, input);
                    case "undo_last_action": return UndoLastAction(mbxKey);
                    case "redo_last_action": return RedoLastAction(mbxKey);
                    case "respond_meeting": return RespondMeeting(mbxKey, input);
                    case "set_event_categories": return SetEventCategories(mbxKey, input);
                    case "set_category_color": return SetCategoryColor(mbxKey, input);
                    case "set_event_availability": return SetEventAvailability(mbxKey, input);
                    case "create_task": return CreateTask(mbxKey, input);
                    case "update_task": return UpdateTask(mbxKey, input);
                    case "set_reminder": return SetReminder(mbxKey, input);
                    case "set_email_reminder": return SetEmailReminder(mbxKey, input);

                    case "draft_email": return ComposeEmail(mbxKey, input, false);
                    case "draft_event": return DraftEvent(input);
                    case "draft_cancel_event": return DraftCancelEvent(input);
                    case "draft_edit_event": return DraftEditEvent(input);

                    case "send_email": return ComposeEmail(mbxKey, input, true);
                    case "create_event": return CreateEvent(mbxKey, input);
                    case "cancel_event": return CancelEvent(mbxKey, input);
                    case "edit_event": return EditEvent(mbxKey, input);

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

            // Default store only, same reasoning as list_folders: Ns.Folders
            // spans every store in the profile (shared mailboxes, Public
            // Folders, SharePoint lists), none of which is what a caller means
            // by a plain folder name.
            Outlook.Folder root = (Outlook.Folder)Ns.DefaultStore.GetRootFolder();
            Outlook.Folder found = FindFolderByName(root.Folders, name.Trim(), 0);
            if (found == null) throw new ArgumentException("Folder not found in your mailbox: " + name + ". Call list_folders to see available names.");
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

        internal static int? OptInt(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind == JsonValueKind.Number)
            {
                int n;
                if (v.TryGetInt32(out n)) return n;
            }
            return null;
        }

        internal static string[] StrArray(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind != JsonValueKind.Object || !o.TryGetProperty(name, out v) || v.ValueKind != JsonValueKind.Array)
                return null;
            var list = new List<string>();
            foreach (JsonElement item in v.EnumerateArray())
                if (item.ValueKind == JsonValueKind.String) list.Add(item.GetString());
            return list.ToArray();
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
