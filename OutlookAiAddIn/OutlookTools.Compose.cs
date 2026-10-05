using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Draft-and-display only. These NEVER call .Send() - a native Outlook
        // compose/appointment window is opened for the user to review and send.
        private const string Signature = "\n\n— Created with OpenDocs";

        private static string SeedSignature(string body)
        {
            if (string.IsNullOrEmpty(body)) return body ?? "";
            return body + Signature;
        }

        private static string PrependHtml(string body, string existingHtml)
        {
            if (string.IsNullOrEmpty(body)) return existingHtml;
            string html = WebUtility.HtmlEncode(body).Replace("\n", "<br>");
            return html + "<br><br>" + (existingHtml ?? "");
        }

        private static readonly Dictionary<string, Outlook.OlRecurrenceType> RecurrenceTypeMap = new Dictionary<string, Outlook.OlRecurrenceType>
        {
            { "daily", Outlook.OlRecurrenceType.olRecursDaily },
            { "weekly", Outlook.OlRecurrenceType.olRecursWeekly },
            { "monthly", Outlook.OlRecurrenceType.olRecursMonthly },
            { "monthlyNth", Outlook.OlRecurrenceType.olRecursMonthNth },
            { "yearly", Outlook.OlRecurrenceType.olRecursYearly },
            { "yearlyNth", Outlook.OlRecurrenceType.olRecursYearNth },
        };

        private static readonly Dictionary<string, Outlook.OlDaysOfWeek> DayFlagMap = new Dictionary<string, Outlook.OlDaysOfWeek>
        {
            { "sunday", Outlook.OlDaysOfWeek.olSunday },
            { "monday", Outlook.OlDaysOfWeek.olMonday },
            { "tuesday", Outlook.OlDaysOfWeek.olTuesday },
            { "wednesday", Outlook.OlDaysOfWeek.olWednesday },
            { "thursday", Outlook.OlDaysOfWeek.olThursday },
            { "friday", Outlook.OlDaysOfWeek.olFriday },
            { "saturday", Outlook.OlDaysOfWeek.olSaturday },
        };

        // Applies a validated RecurrenceSpec (OfficeAi.Shared) to an
        // AppointmentItem, converting to the Outlook enums/flags
        // OfficeAi.Shared can't reference directly. Called after attendees
        // are set but before .Save()/.Send(). See OutlookTools.Compose.cs.md
        // for the call-order caveat still worth verifying live.
        private static void ApplyRecurrence(Outlook.AppointmentItem a, RecurrenceSpec spec)
        {
            Outlook.RecurrencePattern pattern = a.GetRecurrencePattern();
            pattern.RecurrenceType = RecurrenceTypeMap[spec.Type];
            pattern.Interval = spec.Interval;

            // Only write the fields the chosen recurrence type actually uses -
            // Parse() doesn't reject irrelevant extras (e.g. day_of_month with
            // type "weekly"), so skip anything Outlook may not expect for the
            // active RecurrenceType.
            bool usesDaysOfWeek = spec.Type == "weekly" || spec.Type == "monthlyNth" || spec.Type == "yearlyNth";
            bool usesDayOfMonth = spec.Type == "monthly" || spec.Type == "yearly";
            bool usesInstance = spec.Type == "monthlyNth" || spec.Type == "yearlyNth";
            bool usesMonthOfYear = spec.Type == "yearly" || spec.Type == "yearlyNth";

            if (usesDaysOfWeek && spec.DaysOfWeek != null)
            {
                Outlook.OlDaysOfWeek mask = 0;
                foreach (string d in spec.DaysOfWeek) mask |= DayFlagMap[d];
                pattern.DayOfWeekMask = mask;
            }
            if (usesDayOfMonth && spec.DayOfMonth.HasValue) pattern.DayOfMonth = spec.DayOfMonth.Value;
            if (usesInstance && spec.Instance.HasValue) pattern.Instance = spec.Instance.Value;
            if (usesMonthOfYear && spec.MonthOfYear.HasValue) pattern.MonthOfYear = spec.MonthOfYear.Value;

            if (spec.Count.HasValue) pattern.Occurrences = spec.Count.Value;
            else if (spec.Until.HasValue) pattern.PatternEndDate = spec.Until.Value;
            else pattern.NoEndDate = true;
        }

        // Reads and validates the optional "recurrence" object from
        // create_event/draft_event's input. Returns null (with error left
        // null) when the field is omitted entirely; returns a populated
        // RecurrenceSpec on success; returns null with error set to a
        // model-facing IsError message on any validation failure.
        private static RecurrenceSpec ReadRecurrence(JsonElement input, DateTime start, out string error)
        {
            error = null;
            JsonElement rec;
            if (input.ValueKind != JsonValueKind.Object || !input.TryGetProperty("recurrence", out rec) || rec.ValueKind == JsonValueKind.Null)
                return null;
            if (rec.ValueKind != JsonValueKind.Object)
            {
                error = "recurrence must be an object, e.g. {\"type\": \"weekly\"}.";
                return null;
            }

            string type = Str(rec, "type", null);
            int interval = Int(rec, "interval", 1);
            string[] days = StrArray(rec, "days_of_week");
            if (days != null && days.Length == 0) days = null;
            int? dayOfMonth = OptInt(rec, "day_of_month");
            int? instance = OptInt(rec, "instance");
            int? monthOfYear = OptInt(rec, "month_of_year");
            int? count = OptInt(rec, "count");
            DateTime? until = DateArg(rec, "until");

            if (until.HasValue && until.Value.Date < start.Date)
            {
                error = "recurrence.until (" + Iso(until.Value) + ") must be on or after start (" + Iso(start) + ").";
                return null;
            }

            // Fill in defaults from the event's own start date for whatever
            // the chosen type needs but the caller omitted; never overrides a
            // value the caller did provide. See OutlookTools.Compose.cs.md
            // for why (including the empty-array handling above).
            string startDay = start.DayOfWeek.ToString().ToLowerInvariant();
            switch (type)
            {
                case "weekly":
                    if (days == null) days = new[] { startDay };
                    break;
                case "monthly":
                    if (!dayOfMonth.HasValue) dayOfMonth = start.Day;
                    break;
                case "monthlyNth":
                    if (days == null) days = new[] { startDay };
                    if (!instance.HasValue) instance = RecurrenceValidator.NthWeekdayOfMonth(start.Day, DateTime.DaysInMonth(start.Year, start.Month));
                    break;
                case "yearly":
                    if (!dayOfMonth.HasValue) dayOfMonth = start.Day;
                    if (!monthOfYear.HasValue) monthOfYear = start.Month;
                    break;
                case "yearlyNth":
                    if (days == null) days = new[] { startDay };
                    if (!instance.HasValue) instance = RecurrenceValidator.NthWeekdayOfMonth(start.Day, DateTime.DaysInMonth(start.Year, start.Month));
                    if (!monthOfYear.HasValue) monthOfYear = start.Month;
                    break;
                    // "daily" needs none of these; an unrecognized type falls
                    // through unchanged to RecurrenceValidator.Parse's own
                    // "not valid" error.
            }

            return RecurrenceValidator.Parse(type, interval, days, dayOfMonth, instance, monthOfYear, count, until, out error);
        }

        // One compose path for draft_email (opens a window, the user reviews and
        // sends) and send_email (sends immediately, Full autonomy only).
        // action picks a new email, reply, reply-all or forward. Reply and
        // forward target either an email (message_id) or a calendar meeting
        // (event_id), which goes through the meeting's invitation message -
        // see FindInviteMessage.
        private static ToolResult ComposeEmail(string mbxKey, JsonElement input, bool send)
        {
            string tool = send ? "send_email" : "draft_email";
            string action = Str(input, "action", "new");
            string to = Str(input, "to", "");
            string body = Str(input, "body", "");

            if (action == "new")
            {
                // to is required when sending: there's no compose window for a
                // human to fill it in. See OutlookTools.Compose.cs.md for what
                // an unset recipient risks.
                if (send) to = ReqStr(input, "to");
                string subject = Str(input, "subject", "");
                Outlook.MailItem m = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
                if (!string.IsNullOrEmpty(to)) m.To = to;
                m.Subject = subject;
                m.Body = SeedSignature(body);
                return FinishCompose(mbxKey, tool, send, "email", to, subject, () => m.Display(false), () => m.Send());
            }

            if (action != "reply" && action != "reply_all" && action != "forward")
                return new ToolResult { Output = "action must be one of: new, reply, reply_all, forward.", IsError = true, Summary = tool };

            string messageId = Str(input, "message_id", null);
            string eventId = Str(input, "event_id", null);
            if ((messageId == null) == (eventId == null))
                return new ToolResult { Output = "Pass exactly one of message_id (an email) or event_id (a calendar meeting).", IsError = true, Summary = tool };
            if (send && action == "forward") to = ReqStr(input, "to");

            Outlook.MailItem orig = null;
            Outlook.MeetingItem invite = null;
            string origSubject;
            if (messageId != null)
            {
                orig = ItemById(messageId, StoreOf(input)) as Outlook.MailItem;
                if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = tool };
                origSubject = orig.Subject ?? "";
            }
            else
            {
                Outlook.AppointmentItem appt = ResolveMeetingAppointment(ItemById(eventId, Str(input, "store_id", null)));
                if (appt == null) return new ToolResult { Output = "event_id does not resolve to a calendar event.", IsError = true, Summary = tool };
                invite = FindInviteMessage(appt);
                if (invite == null) return new ToolResult { Output = NoInviteError(appt), IsError = true, Summary = tool };
                origSubject = appt.Subject ?? "";
            }

            if (action == "forward")
            {
                if (orig != null)
                {
                    Outlook.MailItem fwd = orig.Forward();
                    if (!string.IsNullOrEmpty(to)) fwd.To = to;
                    if (!string.IsNullOrEmpty(body)) fwd.HTMLBody = PrependHtml(body, fwd.HTMLBody);
                    return FinishCompose(mbxKey, tool, send, "forward", to, origSubject, () => fwd.Display(false), () => fwd.Send());
                }

                // A meeting forward is a MeetingItem, not a MailItem: recipients
                // go through Recipients, and the cast avoids the Forward/Send/
                // Close method-vs-event name clashes on the MeetingItem type.
                Outlook._MeetingItem mfwd = (Outlook._MeetingItem)((Outlook._MeetingItem)invite).Forward();
                var unresolved = new List<string>();
                foreach (string part in to.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string addr = part.Trim();
                    int lt = addr.LastIndexOf('<');
                    if (lt >= 0 && addr.EndsWith(">")) addr = addr.Substring(lt + 1, addr.Length - lt - 2).Trim();
                    if (addr.Length == 0) continue;
                    Outlook.Recipient r = mfwd.Recipients.Add(addr);
                    try { r.Resolve(); } catch { }
                    if (!r.Resolved) unresolved.Add(addr);
                }
                if (send && unresolved.Count > 0)
                {
                    mfwd.Close(Outlook.OlInspectorClose.olDiscard);
                    return new ToolResult { Output = "Not sent - couldn't resolve: " + string.Join(", ", unresolved) + ".", IsError = true, Summary = tool };
                }
                if (!string.IsNullOrEmpty(body))
                {
                    try { mfwd.Body = SeedSignature(body) + "\n\n" + mfwd.Body; }
                    catch (Exception ex) { DebugLog.WriteException("ComposeEmail meeting forward body", ex); }
                }
                ToolResult fr = FinishCompose(mbxKey, tool, send, "forward", to, origSubject, () => mfwd.Display(false), () => mfwd.Send());
                if (!send) fr.Output += FormatUnresolvedAttendeesNote(unresolved);
                return fr;
            }

            bool all = action == "reply_all";
            Outlook.MailItem reply = orig != null
                ? (all ? orig.ReplyAll() : orig.Reply())
                : (all ? invite.ReplyAll() : invite.Reply());
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            return FinishCompose(mbxKey, tool, send, all ? "reply-all" : "reply", reply.To, origSubject, () => reply.Display(false), () => reply.Send());
        }

        private static ToolResult FinishCompose(string mbxKey, string tool, bool send, string kind, string recipients, string subject, Action display, Action sendIt)
        {
            if (!send)
            {
                display();
                return new ToolResult
                {
                    Output = "Opened " + (kind == "email" ? "a draft" : "a " + kind + " draft") + " in Outlook for the user to review and send.",
                    Summary = tool,
                };
            }
            sendIt();
            RecordIrreversible(mbxKey, tool + " (" + kind + ") to " + recipients);
            return new ToolResult { Output = "Sent " + kind + " to " + recipients + ": \"" + subject + "\".", Mutated = true, Summary = tool };
        }

        private static ToolResult DraftEvent(JsonElement input)
        {
            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");

            JsonElement recField;
            bool hasRecurrenceField = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("recurrence", out recField) && recField.ValueKind != JsonValueKind.Null;

            RecurrenceSpec recurrence = null;
            if (hasRecurrenceField)
            {
                if (!start.HasValue)
                    return new ToolResult { Output = "start is required when recurrence is specified, so defaults (day of week, day of month, etc.) can be derived from it.", IsError = true, Summary = "draft_event" };
                string recurrenceError;
                recurrence = ReadRecurrence(input, start.Value, out recurrenceError);
                if (recurrenceError != null) return new ToolResult { Output = recurrenceError, IsError = true, Summary = "draft_event" };
            }

            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = SeedSignature(Str(input, "body", ""));

            if (start.HasValue) a.Start = start.Value;
            if (end.HasValue) a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            var unresolvedAttendees = new List<string>();
            if (!string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt))
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired, unresolvedAttendees);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional, unresolvedAttendees);
            }
            if (recurrence != null) ApplyRecurrence(a, recurrence);
            a.Display(false);
            return new ToolResult
            {
                Output = "Opened an appointment draft in Outlook for the user to review and send." + (recurrence != null ? " Set to repeat " + recurrence.Type + ". Note: Outlook won't visually show the recurrence pattern in this review window until you save it once (a known Outlook limitation for brand-new unsaved items) - it applies correctly once saved or sent." : "") + FormatUnresolvedAttendeesNote(unresolvedAttendees),
                Summary = "draft_event",
            };
        }

        private static void AddAttendees(Outlook.AppointmentItem a, string csv, Outlook.OlMeetingRecipientType type, List<string> unresolved)
        {
            if (string.IsNullOrEmpty(csv)) return;
            foreach (string part in csv.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string addr = part.Trim();
                int lt = addr.LastIndexOf('<');
                if (lt >= 0 && addr.EndsWith(">")) addr = addr.Substring(lt + 1, addr.Length - lt - 2).Trim();
                if (addr.Length == 0) continue;
                Outlook.Recipient r = a.Recipients.Add(addr);
                r.Type = (int)type;
                // Resolve each Recipient individually right here, NOT via the
                // collection-level Recipients.ResolveAll() callers used to
                // rely on - that never reliably resolved anything (confirmed
                // live via COM). See OutlookTools.Compose.cs.md.
                try { r.Resolve(); } catch { }
                // A genuinely unresolvable address still fails silently
                // otherwise - report it back so the caller can flag it as a
                // possible typo (matches find_meeting_slots' own pattern).
                if (!r.Resolved) unresolved.Add(addr);
            }
        }

        // Shared by create_event/draft_event/edit_event/draft_edit_event's
        // result text - see AddAttendees above.
        private static string FormatUnresolvedAttendeesNote(List<string> unresolved)
        {
            return unresolved.Count > 0 ? " Could not resolve: " + string.Join(", ", unresolved) + " - check for typos." : "";
        }

        // Send-and-dispatch, Full-autonomy-only counterparts to the
        // draft/reply/forward/event tools above. Same bodies, but .Send()/
        // .Save() instead of .Display(false) - no native window, no review
        // step. Gated in OutlookTools.cs's ExecuteAsync (SendTierTools),
        // never reachable below Full Autonomy.
        // Non-meeting appointments only .Save() - Outlook has nobody to send
        // them to. A meeting (attendees present) must .Send() instead:
        // .Save() alone would leave it sitting unset on the organizer's
        // calendar with invitees never notified. Both members confirmed via
        // reflection against the referenced Outlook PIA (_AppointmentItem
        // exposes both Send() and Save()) before writing this branch.
        private static ToolResult CreateEvent(string mbxKey, JsonElement input)
        {
            // start/end are required here, unlike draft_event: there's no
            // review window to catch Outlook's own appointment defaults
            // before a real, immediately-live calendar entry is created. See
            // OutlookTools.Compose.cs.md.
            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "create_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "create_event" };

            string recurrenceError;
            RecurrenceSpec recurrence = ReadRecurrence(input, start.Value, out recurrenceError);
            if (recurrenceError != null) return new ToolResult { Output = recurrenceError, IsError = true, Summary = "create_event" };

            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = SeedSignature(Str(input, "body", ""));
            a.Start = start.Value;
            a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            bool isMeeting = !string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt);
            var unresolvedAttendees = new List<string>();
            if (isMeeting)
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired, unresolvedAttendees);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional, unresolvedAttendees);
            }
            if (recurrence != null) ApplyRecurrence(a, recurrence);

            if (isMeeting)
            {
                a.Send();
                RecordIrreversible(mbxKey, "create_event invite for \"" + (a.Subject ?? "") + "\"");
                // This confirmation line is the only place the user sees who
                // an irreversible, unreviewed invite went to - do not let an
                // empty req (optional_attendees-only) produce a malformed
                // "to ; alice@example.com." leading separator.
                string attendeeList = string.IsNullOrEmpty(req) ? opt : string.IsNullOrEmpty(opt) ? req : req + "; " + opt;
                return new ToolResult
                {
                    Output = "Created and sent invite: \"" + (a.Subject ?? "") + "\" to " + attendeeList + "." + (recurrence != null ? " Repeats " + recurrence.Type + "." : "") + FormatUnresolvedAttendeesNote(unresolvedAttendees),
                    Mutated = true,
                    Summary = "create_event",
                };
            }

            a.Save();
            RecordCreated(mbxKey, "create_event", "event_id", a, a.Subject ?? "");
            return new ToolResult
            {
                Output = "Created event: \"" + (a.Subject ?? "") + "\"." + (recurrence != null ? " Repeats " + recurrence.Type + "." : ""),
                Mutated = true,
                Summary = "create_event",
            };
        }
    }
}
