using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Ordering is load-bearing: Sort("[Start]") -> IncludeRecurrences = true
        // -> Restrict. Any other order silently drops recurring instances, and
        // IncludeRecurrences rules out the faster GetTable path.
        private static ToolResult ListEvents(JsonElement input)
        {
            DateTime start = (DateArg(input, "start_date") ?? DateTime.Today).Date;
            DateTime end = (DateArg(input, "end_date") ?? DateTime.Today.AddDays(7)).Date.AddDays(1);
            int limit = Math.Max(1, Int(input, "limit", 50));

            Outlook.Folder cal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = cal.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            string filter = "[Start] <= '" + end.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [End] >= '" + start.ToString("g", CultureInfo.CurrentCulture) + "'";
            Outlook.Items restricted = items.Restrict(filter);

            var sb = new StringBuilder();
            int n = 0;
            foreach (object o in restricted)
            {
                if (n >= limit) break;
                Outlook.AppointmentItem appt = o as Outlook.AppointmentItem;
                if (appt == null) continue;
                n++;
                sb.AppendLine("- event_id: " + appt.EntryID);
                sb.AppendLine("  subject: " + (appt.Subject ?? ""));
                try { sb.AppendLine("  start: " + Iso(appt.Start) + "  end: " + Iso(appt.End)); } catch { }
                sb.AppendLine("  location: " + (appt.Location ?? ""));
                sb.AppendLine("  organizer: " + (appt.Organizer ?? "") + "  all_day: " + appt.AllDayEvent + "  recurring: " + appt.IsRecurring);
                sb.AppendLine("  response: " + appt.ResponseStatus + "  meeting_status: " + appt.MeetingStatus);
            }

            if (n == 0)
                return new ToolResult { Output = "No events between " + start.ToShortDateString() + " and " + end.AddDays(-1).ToShortDateString() + ".", Summary = "list_events" };
            return new ToolResult { Output = sb + "\n(Recurring instances share the master event_id.)", Summary = "list_events" };
        }

        private static ToolResult GetEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "get_event" };

            var sb = new StringBuilder();
            sb.AppendLine("subject: " + (appt.Subject ?? ""));
            try { sb.AppendLine("start: " + Iso(appt.Start) + "  end: " + Iso(appt.End)); } catch { }
            sb.AppendLine("location: " + (appt.Location ?? ""));
            sb.AppendLine("organizer: " + (appt.Organizer ?? ""));
            sb.AppendLine("required_attendees: " + (appt.RequiredAttendees ?? ""));
            sb.AppendLine("optional_attendees: " + (appt.OptionalAttendees ?? ""));
            sb.AppendLine("response: " + appt.ResponseStatus + "  recurring: " + appt.IsRecurring + "  meeting_status: " + appt.MeetingStatus);
            sb.AppendLine();
            sb.AppendLine("body:");
            sb.Append(Truncate(appt.Body ?? "", 40000));
            return new ToolResult { Output = sb.ToString(), Summary = "get_event" };
        }

        // Ranks candidate meeting times by attendee availability, using
        // Recipient.FreeBusy (a per-30-min status string). Pure ranking lives
        // in OfficeAi.Shared.MeetingSlots; this is the COM half.
        private static async Task<ToolResult> FindMeetingSlotsAsync(JsonElement input)
        {
            string attendeesRaw = ReqStr(input, "attendees");
            int duration = Int(input, "duration_minutes", 0);
            if (duration <= 0)
                return new ToolResult { Output = "duration_minutes is required and must be positive.", IsError = true, Summary = "find_meeting_slots" };

            OutlookEws.WorkWeekInfo? workWeek = await ResolveWorkWeekAsync();
            HashSet<DayOfWeek> workDays = workWeek.HasValue ? workWeek.Value.Days : FallbackWorkDays;
            string workDaysLabel = workWeek.HasValue ? string.Join(",", workDays) : "Sun-Thu (default - could not read the mailbox's actual work week)";

            double startHour = Double(input, "start_hour", workWeek.HasValue ? workWeek.Value.StartHour : 9);
            double endHour = Double(input, "end_hour", workWeek.HasValue ? workWeek.Value.EndHour : 18);
            if (endHour <= startHour)
                return new ToolResult { Output = "end_hour must be after start_hour.", IsError = true, Summary = "find_meeting_slots" };
            int limit = Math.Max(1, Int(input, "limit", 5));

            DateTime rangeStart, rangeEnd;
            DateTime? sd = DateArg(input, "start_date");
            DateTime? ed = DateArg(input, "end_date");
            if (sd.HasValue || ed.HasValue)
            {
                rangeStart = (sd ?? DateTime.Today).Date;
                rangeEnd = (ed ?? rangeStart.AddDays(4)).Date;
            }
            else
            {
                DefaultWorkRange(DateTime.Today, workDays, out rangeStart, out rangeEnd);
            }
            if (rangeEnd < rangeStart)
                return new ToolResult { Output = "end_date is before start_date.", IsError = true, Summary = "find_meeting_slots" };
            if ((rangeEnd - rangeStart).TotalDays > 28) rangeEnd = rangeStart.AddDays(28);

            List<DateTime> days = WorkDays(rangeStart, rangeEnd, workDays);
            if (days.Count == 0)
                return new ToolResult { Output = "No work days (" + workDaysLabel + ") between " + rangeStart.ToShortDateString() + " and " + rangeEnd.ToShortDateString() + ".", Summary = "find_meeting_slots" };

            DateTime anchor = days[0].Date;

            var freeBusy = new Dictionary<string, string>();
            var unresolved = new List<string>();

            try
            {
                Outlook.Recipient me = Ns.CurrentUser;
                string meLabel = string.IsNullOrEmpty(me.Name) ? "me" : me.Name;
                freeBusy[meLabel + " (organizer)"] = (string)me.FreeBusy(anchor, 30, true);
            }
            catch (Exception ex) { DebugLog.WriteException("FindMeetingSlots organizer FreeBusy", ex); }

            foreach (string part in attendeesRaw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string a = part.Trim();
                int lt = a.LastIndexOf('<');
                if (lt >= 0 && a.EndsWith(">")) a = a.Substring(lt + 1, a.Length - lt - 2).Trim();
                if (a.Length == 0) continue;
                try
                {
                    Outlook.Recipient r = Ns.CreateRecipient(a);
                    r.Resolve();
                    if (!r.Resolved) { unresolved.Add(a); continue; }
                    freeBusy[a] = (string)r.FreeBusy(anchor, 30, true);
                }
                catch (Exception ex)
                {
                    DebugLog.WriteException("FindMeetingSlots FreeBusy " + a, ex);
                    unresolved.Add(a);
                }
            }

            if (freeBusy.Count == 0)
                return new ToolResult { Output = "Could not get free/busy for anyone - attendees unresolved, or free/busy data is not available on this Exchange setup.", IsError = true, Summary = "find_meeting_slots" };

            List<FreeSlot> slots;
            try
            {
                slots = MeetingSlots.Rank(freeBusy, anchor, days, startHour, endHour, duration, 30, limit);
            }
            catch (ArgumentException ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = "find_meeting_slots" };
            }

            var sb = new StringBuilder();
            if (unresolved.Count > 0) sb.AppendLine("Could not resolve: " + string.Join(", ", unresolved));
            sb.AppendLine("Checked " + freeBusy.Count + " people, " + FormatHour(startHour) + "-" + FormatHour(endHour) + ", " + duration + " min slots:");
            foreach (FreeSlot s in slots)
            {
                sb.Append("- " + Iso(s.Start) + " to " + s.End.ToString("HH:mm", CultureInfo.InvariantCulture) +
                          "  (" + s.Available + "/" + s.Total + " free");
                if (s.Missing.Length > 0) sb.Append("; busy: " + string.Join(", ", s.Missing));
                sb.AppendLine(")");
            }
            sb.AppendLine("\nPass a slot's start/end to draft_event to send the invite.");
            return new ToolResult { Output = sb.ToString(), Summary = "find_meeting_slots" };
        }

        // Generalization of mcp-outlook's scheduling.default_range (originally
        // "today through Thursday of this work week, rolling to next Sun-Thu
        // if today is Fri/Sat") for an arbitrary work-days set: walk forward
        // from today to the next work day (today itself if it already is
        // one), then extend through the following contiguous run of work
        // days (capped at a week) - correct for any shape, including a
        // work week that wraps past a calendar-week boundary.
        private static void DefaultWorkRange(DateTime today, HashSet<DayOfWeek> workDays, out DateTime start, out DateTime end)
        {
            DateTime d = today.Date;
            for (int guard = 0; guard < 14 && !workDays.Contains(d.DayOfWeek); guard++) d = d.AddDays(1);
            start = d;

            DateTime e = start;
            for (int i = 0; i < 6; i++)
            {
                DateTime next = e.AddDays(1);
                if (!workDays.Contains(next.DayOfWeek)) break;
                e = next;
            }
            end = e;
        }

        private static List<DateTime> WorkDays(DateTime start, DateTime end, HashSet<DayOfWeek> workDays)
        {
            var days = new List<DateTime>();
            for (DateTime d = start.Date; d <= end.Date; d = d.AddDays(1))
            {
                if (workDays.Contains(d.DayOfWeek))
                    days.Add(d);
            }
            return days;
        }

        // start_hour/end_hour can be fractional (a mailbox's real EWS working
        // hours, e.g. 8.5 for 08:30) even though the tool's own arguments are
        // whole-hour integers - "00" formatting a fractional double would
        // silently round instead of showing the actual minutes.
        private static string FormatHour(double hour)
        {
            TimeSpan t = TimeSpan.FromHours(hour);
            return t.ToString(@"hh\:mm", CultureInfo.InvariantCulture);
        }

        private static ToolResult RespondMeeting(string mbxKey, JsonElement input, bool accept)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);

            Outlook.AppointmentItem appt = item as Outlook.AppointmentItem;
            if (appt == null)
            {
                Outlook.MeetingItem mi = item as Outlook.MeetingItem;
                if (mi != null) appt = mi.GetAssociatedAppointment(false);
            }
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = accept ? "accept_meeting" : "decline_meeting" };

            Outlook.OlMeetingResponse response = accept
                ? Outlook.OlMeetingResponse.olMeetingAccepted
                : Outlook.OlMeetingResponse.olMeetingDeclined;
            object respObj = appt.Respond(response, true, false);
            Outlook.MeetingItem resp = respObj as Outlook.MeetingItem;
            if (resp != null)
            {
                try { resp.Send(); } catch (Exception ex) { DebugLog.WriteException("RespondMeeting Send", ex); }
            }
            RecordIrreversible(mbxKey, (accept ? "accept_meeting" : "decline_meeting") + " for \"" + (appt.Subject ?? "") + "\"");
            return new ToolResult
            {
                Output = (accept ? "Accepted: " : "Declined: ") + (appt.Subject ?? ""),
                Mutated = true,
                Summary = accept ? "accept_meeting" : "decline_meeting",
            };
        }

        // Shared by draft_reschedule_event/reschedule_event: an olMeetingReceived
        // (or olMeetingReceivedAndCanceled) appointment is one the user only
        // attends, not organizes - Outlook gives attendees no authority to
        // unilaterally move someone else's meeting. Confirmed via .NET
        // reflection against the referenced Microsoft.Office.Interop.Outlook
        // 15.0.0.0 PIA that neither AppointmentItem/_AppointmentItem nor
        // MeetingItem/_MeetingItem expose any "Propose New Time" member -
        // _AppointmentItem.Respond only takes an OlMeetingResponse
        // (Accept/Decline/Tentative), no counter-proposal overload. Real
        // Outlook's "Propose New Time" is a ribbon/UI feature (MAPI
        // counter-proposal properties), not one the classic COM object model
        // exposes cleanly, so there is no authoritative or even semi-authoritative
        // way to honor a reschedule request on a received meeting here - both
        // tools refuse up front instead of silently attempting a Save()/Send()
        // Outlook wouldn't actually honor as a real reschedule.
        //
        // OlMeetingStatus has 5 values: olNonMeeting (0), olMeeting (1),
        // olMeetingReceived (3), olMeetingCanceled (5),
        // olMeetingReceivedAndCanceled (7) - checking only the exact
        // olMeetingReceived value here would let an attendee's copy of a
        // meeting the organizer has since canceled (olMeetingReceivedAndCanceled)
        // through silently, so both "received" statuses count.
        private static bool IsReceivedMeeting(Outlook.AppointmentItem appt)
        {
            return appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceived ||
                   appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;
        }

        // Separate from IsReceivedMeeting: covers the organizer's own
        // olMeetingCanceled copy too (not just the attendee-side
        // olMeetingReceivedAndCanceled) - rescheduling a canceled meeting is
        // nonsensical regardless of who canceled it or who's asking.
        private static bool IsCanceledMeeting(Outlook.AppointmentItem appt)
        {
            return appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled ||
                   appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;
        }

        private static string ReceivedMeetingError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" is a meeting organized by " +
                   (string.IsNullOrEmpty(appt.Organizer) ? "someone else" : appt.Organizer) +
                   " - you're only an attendee, not the organizer, so this tool has no authority to move it. " +
                   "Use Outlook's own \"Propose New Time\" option on the meeting instead.";
        }

        private static string CanceledMeetingError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" has been canceled, so there's nothing to reschedule.";
        }

        // Draft-tier: opens the appointment with the new Start/End already set but
        // NOT saved, exactly like draft_event - the user reviews the moved time in
        // the native window and decides whether to save it (and, if it's a
        // meeting, whether to send the update themselves). Never touches
        // attendees.
        private static ToolResult DraftRescheduleEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_reschedule_event" };

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "draft_reschedule_event" };
            if (IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "draft_reschedule_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "draft_reschedule_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "draft_reschedule_event" };

            bool isMeeting = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            appt.Start = start.Value;
            appt.End = end.Value;
            appt.Display(false);
            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\" with the new time (" + Iso(start.Value) + " to " + Iso(end.Value) +
                         ") in Outlook for the user to review and " + (isMeeting ? "save/send the update." : "save."),
                Summary = "draft_reschedule_event",
            };
        }

        // Full-autonomy-only counterpart to draft_reschedule_event above: sets the
        // new Start/End directly, then - like create_event's attendee-present
        // branch - .Send() if this is a meeting the user organizes (dispatching
        // the reschedule notice to attendees with no review step) or .Save() for
        // a plain appointment nobody needs to notify. Both members confirmed
        // present via reflection against the referenced PIA in CreateEvent above;
        // reused here rather than re-verified.
        private static readonly string[] RescheduleProps = { "Start", "End" };

        private static ToolResult RescheduleEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "reschedule_event" };

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "reschedule_event" };
            if (IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "reschedule_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "reschedule_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "reschedule_event" };

            string oldStart = Iso(appt.Start);
            string oldEnd = Iso(appt.End);
            object[] before = ReadProps(appt, RescheduleProps);
            appt.Start = start.Value;
            appt.End = end.Value;

            if (appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting)
            {
                appt.Send();
                // Barrier, not a snapshot: like create_event's invite branch, this
                // is an irreversible, unreviewed send - undo must stop here rather
                // than silently move the meeting back without telling attendees.
                RecordIrreversible(mbxKey, "reschedule_event invite update for \"" + (appt.Subject ?? "") + "\"");
                // This confirmation line is the only place the user sees that an
                // irreversible, unreviewed reschedule notice went out to attendees.
                return new ToolResult
                {
                    Output = "Rescheduled and sent update notice: \"" + (appt.Subject ?? "") + "\" from " + oldStart + " - " + oldEnd +
                             " to " + Iso(start.Value) + " - " + Iso(end.Value) + ".",
                    Mutated = true,
                    Summary = "reschedule_event",
                };
            }

            appt.Save();
            RecordSnapshot(mbxKey, "reschedule_event", appt, appt.Subject ?? "", RescheduleProps, before);
            return new ToolResult
            {
                Output = "Rescheduled: \"" + (appt.Subject ?? "") + "\" from " + oldStart + " - " + oldEnd +
                         " to " + Iso(start.Value) + " - " + Iso(end.Value) + ".",
                Mutated = true,
                Summary = "reschedule_event",
            };
        }
    }
}
