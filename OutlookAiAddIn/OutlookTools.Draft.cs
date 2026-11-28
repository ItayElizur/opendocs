using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAiAddIn
{
    // The item window's own tools, called by ItemPaneHost with the one item its pane holds; the model never passes an
    // item id. Nothing here saves or sends: every change stays in the open window until the user sends or saves it.
    // See OutlookTools.Draft.cs.md.
    public static partial class OutlookTools
    {
        private const int CurrentBodyMaxChars = 30000;

        // Indexed by Recipient.Type: OlMailRecipientType (1-3) and OlMeetingRecipientType (0-3).
        private static readonly string[] MailKinds = { null, "to", "cc", "bcc" };
        private static readonly string[] MeetingKinds = { "organizer", "required", "optional", "resource" };

        private static readonly string[] AppointmentOnlyFields = { "start", "end", "all_day", "location", "recurrence", "remove_recurrence", "reminder_minutes", "show_as", "categories" };

        private static ToolResult ItemErr(string tool, string message)
        {
            return new ToolResult { Output = message, IsError = true, Summary = tool };
        }

        // ---- get_current_item ----

        internal static ToolResult GetCurrentItem(object item)
        {
            const string tool = "get_current_item";
            var sb = new StringBuilder();
            if (item is Outlook.MailItem mail)
            {
                sb.AppendLine("An email being written in this window (not sent).");
                sb.AppendLine("subject: " + (mail.Subject ?? ""));
                sb.AppendLine("importance: " + ImportanceName(mail.Importance));
                AppendRecipients(sb, mail.Recipients, MailKinds, false);
            }
            else if (item is Outlook.AppointmentItem appt)
            {
                bool meeting = appt.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting;
                sb.AppendLine(IsReceivedMeeting(appt) ? "A meeting the user was invited to (organizer: " + (appt.Organizer ?? "") + "). Read-only: the user can only answer it, from the main Outlook window."
                    : IsCanceledMeeting(appt) ? "A canceled meeting."
                    : meeting ? "A meeting the user organizes."
                    : "An appointment (no attendees).");
                sb.AppendLine("subject: " + (appt.Subject ?? ""));
                sb.AppendLine("start: " + Iso(appt.Start) + "  end: " + Iso(appt.End) + "  all_day: " + appt.AllDayEvent);
                sb.AppendLine("location: " + (appt.Location ?? ""));
                sb.AppendLine("recurrence: " + DescribeRecurrence(appt));
                sb.AppendLine("reminder: " + (appt.ReminderSet ? appt.ReminderMinutesBeforeStart + " minutes before" : "none"));
                sb.AppendLine("show_as: " + BusyStatusName(appt.BusyStatus));
                sb.AppendLine("categories: " + (appt.Categories ?? ""));
                AppendRecipients(sb, appt.Recipients, MeetingKinds, true);
            }
            else
            {
                return ItemErr(tool, "This window holds neither an email nor an appointment.");
            }

            string body;
            try { body = TextAnchors.Clean(ReadBody(BodyDocument(item), out _), true).Text.Trim(); }
            catch (Exception ex)
            {
                DebugLog.WriteException(tool + " body", ex);
                body = ((dynamic)item).Body ?? "";
            }
            sb.AppendLine("body (" + body.Length + " characters):");
            sb.Append(body.Length > CurrentBodyMaxChars ? body.Substring(0, CurrentBodyMaxChars) + "\n...[the rest of the body is cut]" : body);
            return new ToolResult { Output = sb.ToString(), Summary = tool };
        }

        private static string ImportanceName(Outlook.OlImportance i)
        {
            return i == Outlook.OlImportance.olImportanceHigh ? "high" : i == Outlook.OlImportance.olImportanceLow ? "low" : "normal";
        }

        private static List<RecipientEntry> ReadRecipients(Outlook.Recipients rs, string[] kinds)
        {
            var list = new List<RecipientEntry>();
            for (int i = 1; i <= rs.Count; i++)
            {
                Outlook.Recipient r = rs[i];
                string address = SmtpOf(r.AddressEntry);
                list.Add(new RecipientEntry
                {
                    Name = r.Name ?? "",
                    Address = string.IsNullOrEmpty(address) ? r.Address ?? "" : address,
                    Kind = r.Type >= 0 && r.Type < kinds.Length ? kinds[r.Type] ?? "other" : "other",
                });
            }
            return list;
        }

        private static void AppendRecipients(StringBuilder sb, Outlook.Recipients rs, string[] kinds, bool withResponse)
        {
            List<RecipientEntry> list = ReadRecipients(rs, kinds);
            sb.AppendLine(withResponse ? "attendees:" : "recipients:");
            if (list.Count == 0) sb.AppendLine("  (none)");
            for (int i = 0; i < list.Count; i++)
            {
                string response = withResponse && list[i].Kind != "organizer"
                    ? "  response: " + rs[i + 1].MeetingResponseStatus.ToString().Replace("olResponse", "")
                    : "";
                sb.AppendLine("  " + list[i].Kind + ": " + list[i].Name + " <" + list[i].Address + ">" + response);
            }
        }

        // Only a series has a pattern to read: GetRecurrencePattern on a single appointment would make it one.
        private static string DescribeRecurrence(Outlook.AppointmentItem appt)
        {
            if (appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence || appt.RecurrenceState == Outlook.OlRecurrenceState.olApptException)
                return "this window holds one occurrence of a series";
            if (appt.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster) return "none";
            Outlook.RecurrencePattern p = appt.GetRecurrencePattern();
            string type = RecurrenceTypeMap.FirstOrDefault(kv => kv.Value == p.RecurrenceType).Key ?? p.RecurrenceType.ToString();
            string text = type + ", every " + p.Interval;
            if (p.RecurrenceType == Outlook.OlRecurrenceType.olRecursWeekly || p.RecurrenceType == Outlook.OlRecurrenceType.olRecursMonthNth || p.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearNth)
                text += ", on " + string.Join(" ", DayFlagMap.Where(kv => (p.DayOfWeekMask & kv.Value) != 0).Select(kv => kv.Key));
            return text + (p.NoEndDate ? ", no end date" : ", until " + p.PatternEndDate.ToString("yyyy-MM-dd") + " (" + p.Occurrences + " occurrences)");
        }

        // ---- the body, through the window's Word editor ----

        private static Word.Document BodyDocument(object item)
        {
            Outlook.Inspector inspector = item is Outlook.MailItem m ? m.GetInspector : ((Outlook.AppointmentItem)item).GetInspector;
            if (inspector.EditorType != Outlook.OlEditorType.olEditorWord)
                throw new InvalidOperationException("This window's body editor is not Word's (EditorType " + inspector.EditorType + "), so it cannot be edited in place.");
            Word.Document doc = inspector.WordEditor as Word.Document;
            if (doc == null) throw new InvalidOperationException("This window's body is not available yet.");
            return doc;
        }

        // The whole body as Range.Text, read with field codes and hidden text so every character has its document
        // position (TextAnchors.DocumentPositions); storyEnd is the body's end position, to check that the two agree.
        private static string ReadBody(Word.Document doc, out int storyEnd)
        {
            Word.Range story = doc.Content;
            story.TextRetrievalMode.IncludeFieldCodes = true;
            story.TextRetrievalMode.IncludeHiddenText = true;
            storyEnd = story.End;
            return story.Text ?? "";
        }

        private static string ReadRange(Word.Document doc, int start, int end)
        {
            Word.Range r = doc.Range(start, end);
            r.TextRetrievalMode.IncludeFieldCodes = true;
            r.TextRetrievalMode.IncludeHiddenText = true;
            return r.Text ?? "";
        }

        private static string ModelText(string s)
        {
            return (s ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Replace(' ', ' ');
        }

        // ---- edit_current_body ----

        internal static ToolResult EditCurrentBody(object item, JsonElement input)
        {
            const string tool = "edit_current_body";
            string find = Str(input, "find", null), replaceWith = Str(input, "replace_with", null);
            string insert = Str(input, "insert", null), at = Str(input, "at", null);
            if (find == "") find = null;
            if (insert == "") insert = null;
            if ((find == null) == (insert == null))
                return ItemErr(tool, "Pass either find with replace_with (to change a passage), or insert with at (to add text).");
            if (find != null && replaceWith == null)
                return ItemErr(tool, "replace_with is required with find. An empty string deletes the passage.");
            if (insert != null && at != "start" && at != "end" && at != "cursor")
                return ItemErr(tool, "at must be start, end or cursor.");

            var timer = Stopwatch.StartNew();
            Word.Document doc = BodyDocument(item);
            Word.Range target;
            string what;
            if (find != null)
            {
                int storyEnd;
                string raw = ReadBody(doc, out storyEnd);
                int[] pos = TextAnchors.DocumentPositions(raw);
                if (raw.Length > 0 && pos[raw.Length - 1] + 1 != storyEnd)
                    return ItemErr(tool, "The body's text cannot be mapped to positions reliably, so it was not changed. Use insert with at instead, or tell the user what to change.");
                CleanText clean = TextAnchors.Clean(raw, true);
                string needle = ModelText(find);
                List<int> starts = TextAnchors.FindAll(clean.Text, needle, 20);
                if (starts.Count == 0)
                    return ItemErr(tool, "The passage was not found in the body. The user may have changed it since you read it: call get_current_item and use the current text.");
                int? match = OptInt(input, "match");
                if (match.HasValue && (match.Value < 1 || match.Value > starts.Count))
                    return ItemErr(tool, "match " + match.Value + " does not exist: the passage appears " + starts.Count + " time(s).");
                if (starts.Count > 1 && !match.HasValue)
                {
                    var shown = new List<string>();
                    for (int i = 0; i < starts.Count && i < 5; i++)
                        shown.Add((i + 1) + ") \"" + TextAnchors.Snippet(clean.Text, starts[i], needle.Length, 30) + "\"");
                    return ItemErr(tool, "The passage appears " + starts.Count + " times (" + string.Join("; ", shown) + "). Each match is marked [[like this]]. Retry with a longer passage, or pass match: N.");
                }
                int s = starts[(match ?? 1) - 1];
                int startPos = pos[clean.Map[s]];
                int endPos = pos[clean.Map[s + needle.Length - 1]] + 1;
                // A last check that the positions hold the passage, so an edit never lands anywhere else.
                if (!string.Equals(TextAnchors.Clean(ReadRange(doc, startPos, endPos), true).Text, needle, StringComparison.OrdinalIgnoreCase))
                    return ItemErr(tool, "The passage's position in the body could not be confirmed, so it was not changed. Call get_current_item and try again.");
                target = doc.Range(startPos, endPos);
                what = replaceWith.Length == 0 ? "Deleted the passage" : "Replaced the passage";
            }
            else
            {
                int p = at == "start" ? 0 : at == "end" ? doc.Content.End - 1 : doc.Windows[1].Selection.Start;
                target = doc.Range(p, p);
                what = "Inserted the text at the " + (at == "cursor" ? "user's cursor" : at);
            }

            // One Range.Text write is one step on the window's own undo list, so the user's Ctrl+Z takes it back.
            target.Text = ModelText(replaceWith ?? insert).Replace('\n', '\r');
            DebugLog.Write(tool + " took " + timer.ElapsedMilliseconds + " ms");
            return new ToolResult { Output = what + ". The user can undo it with Ctrl+Z in the body.", Mutated = true, Summary = tool };
        }

        // ---- set_current_fields ----

        private static bool? OptBool(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind != JsonValueKind.Object || !o.TryGetProperty(name, out v)) return null;
            if (v.ValueKind == JsonValueKind.True) return true;
            if (v.ValueKind == JsonValueKind.False) return false;
            return null;
        }

        private static bool Has(JsonElement o, string name)
        {
            JsonElement v;
            return o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind != JsonValueKind.Null;
        }

        // add_recipients: { kind: [names or addresses] }. Null when valid, else the error text.
        private static string ReadRecipientAdds(JsonElement input, string[] kinds, List<KeyValuePair<string, string>> add)
        {
            JsonElement obj;
            if (input.ValueKind != JsonValueKind.Object || !input.TryGetProperty("add_recipients", out obj) || obj.ValueKind == JsonValueKind.Null) return null;
            string valid = kinds == MailKinds ? "to, cc, bcc (this is an email)" : "required, optional, resource (this is a meeting)";
            if (obj.ValueKind != JsonValueKind.Object) return "add_recipients must be an object such as {\"to\": [\"Dana\"]}; its keys are " + valid + ".";
            foreach (JsonProperty kind in obj.EnumerateObject())
            {
                if (kind.Name == "organizer" || Array.IndexOf(kinds, kind.Name) < 0) return "add_recipients takes " + valid + ", not \"" + kind.Name + "\".";
                if (kind.Value.ValueKind != JsonValueKind.Array) return "add_recipients." + kind.Name + " must be a list of names or addresses.";
                foreach (JsonElement who in kind.Value.EnumerateArray())
                    if (who.ValueKind == JsonValueKind.String) add.Add(new KeyValuePair<string, string>(who.GetString(), kind.Name));
            }
            return null;
        }

        // ChangeKind, then Remove (highest index first), then Add: the plan's indexes are into the list as it was read.
        // A name Outlook cannot resolve to exactly one address is taken off again and reported, never guessed.
        private static void ApplyRecipients(Outlook.Recipients rs, string[] kinds, List<KeyValuePair<string, string>> add, string[] remove, List<string> changes, List<string> problems)
        {
            if (add.Count == 0 && (remove == null || remove.Length == 0)) return;
            List<RecipientEntry> current = ReadRecipients(rs, kinds);
            RecipientEditPlan plan = RecipientEdits.Plan(current, add, remove);
            foreach (KeyValuePair<int, string> k in plan.ChangeKind)
            {
                rs[k.Key + 1].Type = Array.IndexOf(kinds, k.Value);
                changes.Add(current[k.Key].Name + " moved to " + k.Value);
            }
            foreach (int i in plan.Remove)
            {
                rs.Remove(i + 1);
                changes.Add("removed " + current[i].Name);
            }
            foreach (KeyValuePair<string, string> a in plan.Add)
            {
                Outlook.Recipient r = rs.Add(RecipientEdits.AddressText(a.Key));
                r.Type = Array.IndexOf(kinds, a.Value);
                try { r.Resolve(); } catch { }
                if (r.Resolved) { changes.Add("added " + r.Name + " (" + a.Value + ")"); continue; }
                r.Delete();
                problems.Add("\"" + a.Key + "\" did not resolve to exactly one address (unknown, or more than one match), so it was not added. Find the address with search_contacts, or ask the user.");
            }
            foreach (string q in plan.NotFound) problems.Add("\"" + q + "\" is not on this item, so nothing was removed for it.");
        }

        internal static ToolResult SetCurrentFields(object item, JsonElement input)
        {
            const string tool = "set_current_fields";
            var mail = item as Outlook.MailItem;
            var appt = item as Outlook.AppointmentItem;
            if (mail == null && appt == null) return ItemErr(tool, "This window holds neither an email nor an appointment.");

            string[] kinds = mail != null ? MailKinds : MeetingKinds;
            var add = new List<KeyValuePair<string, string>>();
            string addError = ReadRecipientAdds(input, kinds, add);
            if (addError != null) return ItemErr(tool, addError);
            string[] remove = StrArray(input, "remove_recipients");
            string subject = Str(input, "subject", null);
            var changes = new List<string>();
            var problems = new List<string>();

            if (mail != null)
            {
                string misplaced = AppointmentOnlyFields.FirstOrDefault(f => Has(input, f));
                if (misplaced != null) return ItemErr(tool, misplaced + " is for appointments and meetings; this window holds an email.");
                string importance = Str(input, "importance", null);
                if (importance != null && importance != "low" && importance != "normal" && importance != "high")
                    return ItemErr(tool, "importance must be low, normal or high.");

                if (subject != null) { mail.Subject = subject; changes.Add("subject"); }
                if (importance != null)
                {
                    mail.Importance = importance == "high" ? Outlook.OlImportance.olImportanceHigh : importance == "low" ? Outlook.OlImportance.olImportanceLow : Outlook.OlImportance.olImportanceNormal;
                    changes.Add("importance " + importance);
                }
                ApplyRecipients(mail.Recipients, kinds, add, remove, changes, problems);
            }
            else
            {
                if (IsCanceledMeeting(appt)) return ItemErr(tool, CanceledMeetingError(appt));
                if (IsReceivedMeeting(appt))
                    return ItemErr(tool, "\"" + (appt.Subject ?? "") + "\" is organized by " + (string.IsNullOrEmpty(appt.Organizer) ? "someone else" : appt.Organizer) +
                                         "; the user is only an attendee, so it cannot be changed here. They can answer it (respond_meeting) or propose a new time from the main Outlook window.");
                if (Has(input, "importance")) return ItemErr(tool, "importance is for emails; this window holds an appointment.");

                bool series = appt.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster;
                bool occurrence = appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence || appt.RecurrenceState == Outlook.OlRecurrenceState.olApptException;
                DateTime? start = DateArg(input, "start"), end = DateArg(input, "end");
                // Moving only the start keeps the duration.
                if (start.HasValue && !end.HasValue) end = start.Value + (appt.End - appt.Start);
                if (end.HasValue && !start.HasValue) start = appt.Start;
                if (start.HasValue && end.Value <= start.Value) return ItemErr(tool, "end must be after start.");
                bool? allDay = OptBool(input, "all_day");
                string location = Str(input, "location", null);
                bool removeRecurrence = Bool(input, "remove_recurrence", false);
                RecurrenceSpec recurrence = null;
                if (Has(input, "recurrence"))
                {
                    if (occurrence) return ItemErr(tool, "This window holds one occurrence of a series; its repeat pattern is changed from the series itself.");
                    if (removeRecurrence) return ItemErr(tool, "Pass recurrence or remove_recurrence, not both.");
                    string recurrenceError;
                    recurrence = ReadRecurrence(input, start ?? appt.Start, out recurrenceError);
                    if (recurrenceError != null) return ItemErr(tool, recurrenceError);
                }
                if (removeRecurrence && !series) return ItemErr(tool, "This is not a recurring series, so there is no repeat pattern to remove.");
                int? reminder = OptInt(input, "reminder_minutes");
                string showAs = Str(input, "show_as", null);
                Outlook.OlBusyStatus busy = Outlook.OlBusyStatus.olBusy;
                if (showAs != null && !TryParseBusyStatus(showAs, out busy))
                    return ItemErr(tool, "show_as must be free, tentative, busy, out_of_office or working_elsewhere.");
                string categories = Str(input, "categories", null);

                if (allDay.HasValue) { appt.AllDayEvent = allDay.Value; changes.Add(allDay.Value ? "all day" : "not all day"); }
                ApplyEventFields(new EventEdit { Master = appt, Appt = appt, Start = start, End = end, Subject = subject, Location = location }, series && start.HasValue);
                if (start.HasValue) changes.Add("time " + Iso(start.Value) + " - " + Iso(end.Value));
                if (subject != null) changes.Add("subject");
                if (location != null) changes.Add("location");
                if (removeRecurrence) { appt.ClearRecurrencePattern(); changes.Add("no longer repeats"); }
                if (recurrence != null) { ApplyRecurrence(appt, recurrence); changes.Add("repeats " + recurrence.Type); }
                if (reminder.HasValue)
                {
                    appt.ReminderSet = reminder.Value >= 0;
                    if (reminder.Value >= 0) appt.ReminderMinutesBeforeStart = reminder.Value;
                    changes.Add(reminder.Value >= 0 ? "reminder " + reminder.Value + " minutes before" : "no reminder");
                }
                if (showAs != null) { appt.BusyStatus = busy; changes.Add("show as " + BusyStatusName(busy)); }
                if (categories != null) { appt.Categories = categories; changes.Add("categories"); }
                int before = changes.Count;
                ApplyRecipients(appt.Recipients, kinds, add, remove, changes, problems);
                if (changes.Count > before) SyncMeetingStatus(appt);
            }

            if (changes.Count == 0 && problems.Count == 0) return ItemErr(tool, "Nothing to change: pass at least one field.");
            return new ToolResult
            {
                Output = (changes.Count > 0 ? "Changed in the open window (not saved or sent): " + string.Join("; ", changes) + "." : "Nothing changed.") +
                         (problems.Count > 0 ? " " + string.Join(" ", problems) : ""),
                Mutated = changes.Count > 0,
                IsError = changes.Count == 0,
                Summary = tool,
            };
        }

        // ---- read_main_chat and find_meeting_slots' default ----

        internal static ToolResult ReadMainChat()
        {
            return new ToolResult
            {
                Output = MainChatDigest.Format(ChatStore.LoadSinceLastDivider("OutlookAiAddIn", TaskPaneHost.MailboxChatId())),
                Summary = "read_main_chat",
            };
        }

        // The meeting's required and optional attendees, comma-separated, for find_meeting_slots called without any.
        internal static string CurrentAttendees(object item)
        {
            var appt = item as Outlook.AppointmentItem;
            if (appt == null) return "";
            return string.Join(", ", ReadRecipients(appt.Recipients, MeetingKinds)
                .Where(r => r.Kind == "required" || r.Kind == "optional")
                .Select(r => r.Address).Where(a => a.Length > 0));
        }
    }
}
