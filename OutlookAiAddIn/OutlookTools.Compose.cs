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

        // Applies a validated RecurrenceSpec (OfficeAi.Shared - pure, unit
        // tested in RecurrenceValidatorTests) to an AppointmentItem, converting
        // to the actual Outlook enums/flags OfficeAi.Shared can't reference
        // directly (it has no Outlook PIA reference, same split as
        // ColorUtil/BusyStatus in OutlookTools.Categories.cs). Called after
        // attendees are set but before .Save()/.Send() in both create_event
        // and draft_event - unverified whether GetRecurrencePattern() before
        // vs. after setting MeetingStatus/attendees matters; this ordering
        // (identity first, schedule second) matches natural Outlook UI flow
        // and is the one to verify live in Task 8.
        private static void ApplyRecurrence(Outlook.AppointmentItem a, RecurrenceSpec spec)
        {
            Outlook.RecurrencePattern pattern = a.GetRecurrencePattern();
            pattern.RecurrenceType = RecurrenceTypeMap[spec.Type];
            pattern.Interval = spec.Interval;

            if (spec.DaysOfWeek != null)
            {
                Outlook.OlDaysOfWeek mask = 0;
                foreach (string d in spec.DaysOfWeek) mask |= DayFlagMap[d];
                pattern.DayOfWeekMask = mask;
            }
            if (spec.DayOfMonth.HasValue) pattern.DayOfMonth = spec.DayOfMonth.Value;
            if (spec.Instance.HasValue) pattern.Instance = spec.Instance.Value;
            if (spec.MonthOfYear.HasValue) pattern.MonthOfYear = spec.MonthOfYear.Value;

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
            if (input.ValueKind != JsonValueKind.Object || !input.TryGetProperty("recurrence", out rec) || rec.ValueKind != JsonValueKind.Object)
                return null;

            string type = Str(rec, "type", null);
            int interval = Int(rec, "interval", 1);
            string[] days = StrArray(rec, "days_of_week");
            int? dayOfMonth = OptInt(rec, "day_of_month");
            int? instance = OptInt(rec, "instance");
            int? monthOfYear = OptInt(rec, "month_of_year");
            int? count = OptInt(rec, "count");
            DateTime? until = DateArg(rec, "until");

            // Fill in defaults from the event's own start date when the
            // caller omits a field the chosen type needs, rather than
            // forcing every call to spell out values already implied by
            // "recurring starting from <start>". Only fills what's
            // genuinely missing - an explicitly-provided value is never
            // overridden. RecurrenceValidator.Parse stays the strict, pure
            // validator; by the time it runs here the spec already looks
            // complete for whichever type was requested.
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

        private static ToolResult DraftEmail(JsonElement input)
        {
            string to = Str(input, "to", "");
            string subject = Str(input, "subject", "");
            string body = Str(input, "body", "");

            Outlook.MailItem m = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
            if (!string.IsNullOrEmpty(to)) m.To = to;
            m.Subject = subject;
            m.Body = SeedSignature(body);
            m.Display(false);
            return new ToolResult { Output = "Opened a draft in Outlook for the user to review and send.", Summary = "draft_email" };
        }

        private static ToolResult ReplyEmail(JsonElement input, bool all)
        {
            string id = ReqStr(input, "message_id");
            string body = Str(input, "body", "");
            string tool = all ? "reply_all_email" : "reply_email";

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = tool };

            Outlook.MailItem reply = all ? orig.ReplyAll() : orig.Reply();
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            reply.Display(false);
            return new ToolResult { Output = "Opened a " + (all ? "reply-all" : "reply") + " draft in Outlook for the user to review and send.", Summary = tool };
        }

        private static ToolResult ForwardEmail(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string to = Str(input, "to", "");
            string body = Str(input, "body", "");

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "forward_email" };

            Outlook.MailItem fwd = orig.Forward();
            if (!string.IsNullOrEmpty(to)) fwd.To = to;
            if (!string.IsNullOrEmpty(body)) fwd.HTMLBody = PrependHtml(body, fwd.HTMLBody);
            fwd.Display(false);
            return new ToolResult { Output = "Opened a forward draft in Outlook for the user to review and send.", Summary = "forward_email" };
        }

        private static ToolResult DraftEvent(JsonElement input)
        {
            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");

            JsonElement recField;
            bool hasRecurrenceField = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("recurrence", out recField) && recField.ValueKind == JsonValueKind.Object;

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
            if (!string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt))
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
            }
            if (recurrence != null) ApplyRecurrence(a, recurrence);
            a.Display(false);
            return new ToolResult
            {
                Output = "Opened an appointment draft in Outlook for the user to review and send." + (recurrence != null ? " Set to repeat " + recurrence.Type + "." : ""),
                Summary = "draft_event",
            };
        }

        private static void AddAttendees(Outlook.AppointmentItem a, string csv, Outlook.OlMeetingRecipientType type)
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
            }
        }

        // Send-and-dispatch, Full-autonomy-only counterparts to the
        // draft/reply/forward/event tools above. Same bodies, but .Send()/
        // .Save() instead of .Display(false) - no native window, no review
        // step. Gated in OutlookTools.cs's ExecuteAsync (SendTierTools),
        // never reachable below Full Autonomy.
        private static ToolResult SendEmail(string mbxKey, JsonElement input)
        {
            // to is required, unlike draft_email's - draft_email opens a
            // compose window where a human can add a missing recipient
            // before anything goes out; send_email has no such window, so
            // Send()-ing with no recipient set would either throw or, worse,
            // block on a native Outlook resolution prompt this add-in isn't
            // expecting a response to.
            string to = ReqStr(input, "to");
            string subject = Str(input, "subject", "");
            string body = Str(input, "body", "");

            Outlook.MailItem m = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
            m.To = to;
            m.Subject = subject;
            m.Body = SeedSignature(body);
            m.Send();
            RecordIrreversible(mbxKey, "send_email to " + to);
            return new ToolResult { Output = "Sent to " + to + ": \"" + subject + "\".", Mutated = true, Summary = "send_email" };
        }

        private static ToolResult SendReply(string mbxKey, JsonElement input, bool all)
        {
            string id = ReqStr(input, "message_id");
            string body = Str(input, "body", "");
            string tool = all ? "send_reply_all" : "send_reply";

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = tool };

            Outlook.MailItem reply = all ? orig.ReplyAll() : orig.Reply();
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            string to = reply.To;
            reply.Send();
            RecordIrreversible(mbxKey, tool + " to " + to);
            return new ToolResult { Output = "Sent " + (all ? "reply-all" : "reply") + " to " + to + ": \"" + (orig.Subject ?? "") + "\".", Mutated = true, Summary = tool };
        }

        private static ToolResult SendForward(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string to = ReqStr(input, "to");
            string body = Str(input, "body", "");

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "send_forward" };

            Outlook.MailItem fwd = orig.Forward();
            fwd.To = to;
            if (!string.IsNullOrEmpty(body)) fwd.HTMLBody = PrependHtml(body, fwd.HTMLBody);
            fwd.Send();
            RecordIrreversible(mbxKey, "send_forward to " + to);
            return new ToolResult { Output = "Forwarded to " + to + ": \"" + (orig.Subject ?? "") + "\".", Mutated = true, Summary = "send_forward" };
        }

        // Non-meeting appointments only .Save() - Outlook has nobody to send
        // them to. A meeting (attendees present) must .Send() instead:
        // .Save() alone would leave it sitting unset on the organizer's
        // calendar with invitees never notified. Both members confirmed via
        // reflection against the referenced Outlook PIA (_AppointmentItem
        // exposes both Send() and Save()) before writing this branch.
        private static ToolResult CreateEvent(string mbxKey, JsonElement input)
        {
            // Unlike draft_event (where an omitted start/end just leaves
            // Outlook's own new-appointment default - "now", 30 min - sitting
            // in a review window the user sees before sending/saving),
            // create_event has no review step at all: a caller that forgets
            // either would otherwise create a real, immediately-live calendar
            // entry at an unintended time with nothing surfacing that it was
            // defaulted. Required here specifically.
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
            if (isMeeting)
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
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
                    Output = "Created and sent invite: \"" + (a.Subject ?? "") + "\" to " + attendeeList + "." + (recurrence != null ? " Repeats " + recurrence.Type + "." : ""),
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
