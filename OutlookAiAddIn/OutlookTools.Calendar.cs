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

        // Shared by RespondMeeting/DraftRespondMeeting: resolves event_id's
        // item to the underlying AppointmentItem, whether it's already one
        // or is a MeetingItem (a meeting request still sitting in the
        // Inbox) that needs GetAssociatedAppointment(false) first.
        private static Outlook.AppointmentItem ResolveMeetingAppointment(object item)
        {
            Outlook.AppointmentItem appt = item as Outlook.AppointmentItem;
            if (appt == null)
            {
                Outlook.MeetingItem mi = item as Outlook.MeetingItem;
                if (mi != null) appt = mi.GetAssociatedAppointment(false);
            }
            return appt;
        }

        // Shared by RespondMeeting/DraftRespondMeeting: Respond() only makes
        // sense on a meeting the user was actually invited to, not one they
        // organize themselves (olMeeting) or a plain appointment
        // (olNonMeeting) - "is not a meeting you were invited to" covers
        // both cases accurately, unlike a message that specifically says
        // "you organize this" (wrong for a plain appointment).
        private static string NotInvitedError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" is not a meeting you were invited to - there's nothing to respond to.";
        }

        // Shared by RespondMeeting/DraftRespondMeeting: distinct from
        // NotInvitedError - this is a meeting the user WAS invited to, but
        // the organizer has since canceled it (olMeetingCanceled or
        // olMeetingReceivedAndCanceled).
        private static string AlreadyCanceledRespondError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" has already been canceled - there's nothing to respond to.";
        }

        // response is the actual OlMeetingResponse to send (not just a bool)
        // so this one helper covers all three: accept_meeting, decline_meeting,
        // tentative_meeting. message is an optional comment attached to the
        // response before it's sent - UNVERIFIED live whether the organizer
        // actually sees this text on the delivered response (needs a real
        // received invite to test, not a self-organized item).
        private static ToolResult RespondMeeting(string mbxKey, JsonElement input, Outlook.OlMeetingResponse response, string toolName)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);
            Outlook.AppointmentItem appt = ResolveMeetingAppointment(item);
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = toolName };

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = AlreadyCanceledRespondError(appt), IsError = true, Summary = toolName };
            if (!IsReceivedMeeting(appt))
                return new ToolResult { Output = NotInvitedError(appt), IsError = true, Summary = toolName };

            string message = Str(input, "message", null);
            bool hasComment = !string.IsNullOrWhiteSpace(message);
            // Captured before Respond() - if Respond() replaces the item with
            // a new EntryID on accept/tentative (unverified, but plausible
            // per Respond()'s documented behavior), appt.Subject read after
            // that call could be a stale reference.
            string subject = appt.Subject ?? "";

            object respObj = appt.Respond(response, true, false);
            Outlook.MeetingItem resp = respObj as Outlook.MeetingItem;
            bool sendSucceeded = false;
            if (resp != null)
            {
                if (hasComment) resp.Body = message;
                try { resp.Send(); sendSucceeded = true; } catch (Exception ex) { DebugLog.WriteException(toolName + " Send", ex); }
            }
            bool commentSent = hasComment && sendSucceeded;
            RecordIrreversible(mbxKey, toolName + " for \"" + subject + "\"" + (commentSent ? " with a comment" : ""));
            string verb = response == Outlook.OlMeetingResponse.olMeetingAccepted ? "Accepted"
                        : response == Outlook.OlMeetingResponse.olMeetingTentative ? "Responded tentatively to"
                        : "Declined";
            // If Send() failed, the organizer was never notified (and any
            // comment was lost) even though the local Respond() already
            // went through - don't claim success without qualification.
            string sendNote = sendSucceeded ? (commentSent ? " (comment sent)" : "") : " - but the response could not be sent to the organizer.";
            return new ToolResult
            {
                Output = verb + ": " + subject + sendNote,
                Mutated = true,
                Summary = toolName,
            };
        }

        // Draft-tier counterpart to RespondMeeting - redesigned after a code
        // review confirmed (via Respond()'s documented behavior, and this
        // project's own prior history with draft_cancel_event hitting the
        // identical shape of bug) that calling appt.Respond() commits a real
        // calendar change at call time - a new EntryID on accept/tentative,
        // a move to Deleted Items on decline - independent of whether the
        // resulting response is ever sent or the window ever closed with an
        // action taken. That directly broke this codebase's "draft tools
        // persist nothing until the user acts" guarantee, the same way an
        // earlier version of draft_cancel_event did with an unsaved
        // MeetingStatus change. Fixed the same way that was: never mutate
        // the item at all here. Opens the original item completely
        // unchanged; the user picks Accept/Tentative/Decline themselves
        // from Outlook's own native ribbon buttons. Since this tool no
        // longer calls Respond() with a specific response type, one unified
        // tool replaces what used to be three separate ones
        // (draft_accept_meeting/draft_decline_meeting/draft_tentative_meeting).
        // message can no longer be pre-filled into a response body (that
        // would require calling Respond() to get the MeetingItem, the exact
        // call this redesign avoids) - it's returned in the output text
        // instead, for the user to paste in themselves if they use
        // Outlook's own "Edit response before sending" option.
        private static ToolResult DraftRespondMeeting(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);
            Outlook.AppointmentItem appt = ResolveMeetingAppointment(item);
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = "draft_respond_meeting" };

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = AlreadyCanceledRespondError(appt), IsError = true, Summary = "draft_respond_meeting" };
            if (!IsReceivedMeeting(appt))
                return new ToolResult { Output = NotInvitedError(appt), IsError = true, Summary = "draft_respond_meeting" };

            string message = Str(input, "message", null);
            bool hasComment = !string.IsNullOrWhiteSpace(message);

            appt.Display(false);

            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\" in Outlook - use the Accept/Tentative/Decline buttons there to respond." +
                         (hasComment ? " Suggested comment: \"" + message + "\" - paste it in if you use Outlook's \"Edit response before sending\" option." : ""),
                Summary = "draft_respond_meeting",
            };
        }

        // Shared by draft_edit_event/edit_event: an olMeetingReceived
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

        // Shared by all four occurrence-aware tools (reschedule/cancel, draft
        // and immediate). occurrence_date lets a caller target one instance
        // of a recurring series instead of the whole master.
        // RecurrencePattern.GetOccurrence(DateTime) confirmed present via
        // .NET reflection against the referenced PIA (not assumed) - this
        // had been undocumented capability until this addition, even though
        // every occurrence list_events returns shares the master's EntryID.
        private static ToolResult? ResolveOccurrenceTarget(Outlook.AppointmentItem master, string occurrenceDate, string toolName, out Outlook.AppointmentItem target)
        {
            target = master;
            if (occurrenceDate == null) return null;

            if (!master.IsRecurring)
                return new ToolResult { Output = "\"" + (master.Subject ?? "") + "\" is not a recurring event, so occurrence_date doesn't apply. Omit it to act on the event itself.", IsError = true, Summary = toolName };

            DateTime date;
            if (!DateTime.TryParse(occurrenceDate, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) &&
                !DateTime.TryParse(occurrenceDate, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out date))
                return new ToolResult { Output = "occurrence_date \"" + occurrenceDate + "\" is not a valid date/time.", IsError = true, Summary = toolName };

            Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
            try { target = pattern.GetOccurrence(date); }
            catch (Exception ex)
            {
                DebugLog.WriteException(toolName + " GetOccurrence", ex);
                return new ToolResult { Output = "No occurrence of \"" + (master.Subject ?? "") + "\" on " + occurrenceDate + ". Check list_events for this series' actual occurrence dates.", IsError = true, Summary = toolName };
            }
            return null;
        }

        // Checks whether moving an occurrence from `original` to `target` would
        // cross or land on the same day as another occurrence of the same series -
        // Outlook rejects both (confirmed live 2026-09-28 by reproducing the
        // specific error manually in Outlook's UI: "Cannot reschedule an occurrence
        // of the recurring appointment ... if it skips over a later occurrence of
        // the same appointment" - both cases surface only as a generic
        // COMException "Cannot save this item." from .Save(), with no way to
        // distinguish the cause from the exception itself). Checking this
        // ourselves first gives an exact, deterministic answer instead of relying
        // on Outlook's collapsed generic exception message.
        //
        // Queries the Calendar folder the same way ListEvents does (Sort("[Start]")
        // -> IncludeRecurrences = true -> Restrict, in that order - load-bearing,
        // see ListEvents' own comment) over the range between `original` and
        // `target` (inclusive of both endpoint days), then keeps only occurrences
        // of THIS series (matching master's EntryID) other than the one being
        // moved (excluded by day - it's still sitting at `original` since nothing
        // has been saved yet). If any remain, the closest one to `original` is the
        // binding obstruction; returns null if the move is clear.
        private static ToolResult? CheckOccurrenceReorderCollision(Outlook.AppointmentItem master, DateTime original, DateTime target, string subject, string toolName)
        {
            DateTime rangeStart = (original < target ? original : target).Date;
            DateTime rangeEnd = (original < target ? target : original).Date.AddDays(1);

            Outlook.Folder cal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = cal.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            string filter = "[Start] >= '" + rangeStart.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [Start] < '" + rangeEnd.ToString("g", CultureInfo.CurrentCulture) + "'";
            Outlook.Items restricted = items.Restrict(filter);

            Outlook.AppointmentItem nearest = null;
            foreach (object o in restricted)
            {
                Outlook.AppointmentItem candidate = o as Outlook.AppointmentItem;
                if (candidate == null) continue;
                if (!SameEntryId(candidate.EntryID, master.EntryID)) continue;
                if (candidate.Start.Date == original.Date) continue; // the occurrence being moved itself
                if (nearest == null || Math.Abs((candidate.Start - original).TotalMinutes) < Math.Abs((nearest.Start - original).TotalMinutes))
                    nearest = candidate;
            }

            if (nearest == null) return null;

            DateTime lowerBound = original < nearest.Start ? original : nearest.Start;
            DateTime upperBound = original < nearest.Start ? nearest.Start : original;
            return new ToolResult
            {
                Output = "Can't move \"" + subject + "\" to " + Iso(target) + " - it would cross or land on the same day as another occurrence of this series (on " +
                         nearest.Start.ToShortDateString() + "). Valid range: strictly between " + lowerBound.ToShortDateString() + " and " + upperBound.ToShortDateString() + ".",
                IsError = true,
                Summary = toolName,
            };
        }

        // Shared by edit_event/draft_edit_event: replaces the required and/or
        // optional attendee list wholesale (not a diff/merge - the caller
        // supplies the full new list each time, reading the current one
        // first via get_event if they need to preserve someone). The two
        // categories are independently optional: a null argument leaves that
        // attendee category completely untouched; a non-null argument
        // (including "") fully replaces it - clearing the existing entries
        // in that category and re-adding via the same AddAttendees helper
        // create_event/draft_event already use. The organizer recipient is
        // never touched. Recipients indices are 1-based (confirmed via .NET
        // reflection against the referenced PIA, matching every other
        // Outlook collection in this codebase); iterating downward from
        // Count avoids skipping an element after Remove shifts the rest down.
        private static void ReplaceAttendees(Outlook.AppointmentItem appt, string requiredCsv, string optionalCsv, List<string> unresolved)
        {
            for (int i = appt.Recipients.Count; i >= 1; i--)
            {
                int type = appt.Recipients[i].Type;
                if (requiredCsv != null && type == (int)Outlook.OlMeetingRecipientType.olRequired) { appt.Recipients.Remove(i); continue; }
                if (optionalCsv != null && type == (int)Outlook.OlMeetingRecipientType.olOptional) { appt.Recipients.Remove(i); continue; }
            }
            // AddAttendees resolves each recipient individually now - the
            // collection-level ResolveAll() that used to run here never
            // reliably resolved anything (confirmed live via COM).
            if (requiredCsv != null) AddAttendees(appt, requiredCsv, Outlook.OlMeetingRecipientType.olRequired, unresolved);
            if (optionalCsv != null) AddAttendees(appt, optionalCsv, Outlook.OlMeetingRecipientType.olOptional, unresolved);
        }

        // Counts real attendees only, excluding the organizer - used by
        // EditEvent/DraftEditEvent to decide whether clearing attendees
        // should revert MeetingStatus back to olNonMeeting. Confirmed live
        // via COM that the organizer does NOT appear in Recipients even
        // after a real .Send() (Recipients.Count matched exactly the number
        // of real invited attendees, organizer never counted) - this
        // exclusion is defensive insurance for any account/Exchange
        // configuration where that might differ, not a fix for an observed
        // bug.
        private static int CountAttendeeRecipients(Outlook.AppointmentItem appt)
        {
            int count = 0;
            for (int i = 1; i <= appt.Recipients.Count; i++)
                if (appt.Recipients[i].Type != (int)Outlook.OlMeetingRecipientType.olOrganizer) count++;
            return count;
        }

        // Shared by edit_event/draft_edit_event's result text: describes
        // which of the seven optional fields were actually touched, so the
        // caller sees a precise summary instead of a generic "updated".
        private static string DescribeEditEventChanges(DateTime? start, DateTime? end, string subject, string body, string location, bool attendeesChanged, string oldStart, string oldEnd)
        {
            var parts = new List<string>();
            if (start.HasValue) parts.Add("time changed from " + oldStart + " - " + oldEnd + " to " + Iso(start.Value) + " - " + Iso(end.Value));
            if (subject != null) parts.Add("subject changed");
            if (body != null) parts.Add("body changed");
            if (location != null) parts.Add("location changed");
            if (attendeesChanged) parts.Add("attendees updated");
            return parts.Count > 0 ? " (" + string.Join(", ", parts) + ")" : "";
        }

        // Full-autonomy-only, general edit: unlike the retired per-field
        // reschedule tool (Start/End only), this can touch start/end, subject, body, location, and
        // attendees in one call. One unified mutation flow decides .Send() vs
        // .Save() and whether the result is undo-able, rather than branching
        // per field - see the design's Section 2 for the reasoning.
        private static ToolResult EditEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "edit_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "edit_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            if (IsCanceledMeeting(master))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "edit_event" };
            if (IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "edit_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            string subject = Str(input, "subject", null);
            string body = Str(input, "body", null);
            string location = Str(input, "location", null);
            string requiredAttendees = Str(input, "required_attendees", null);
            string optionalAttendees = Str(input, "optional_attendees", null);

            if (!start.HasValue && !end.HasValue && subject == null && body == null && location == null &&
                requiredAttendees == null && optionalAttendees == null)
                return new ToolResult { Output = "At least one of start, end, subject, body, location, required_attendees, optional_attendees must be provided.", IsError = true, Summary = "edit_event" };

            if (start.HasValue != end.HasValue)
                return new ToolResult { Output = "start and end must be provided together.", IsError = true, Summary = "edit_event" };

            if (start.HasValue && end.Value <= start.Value)
                return new ToolResult { Output = "end must be after start.", IsError = true, Summary = "edit_event" };

            if ((requiredAttendees != null || optionalAttendees != null) && occDate != null)
                return new ToolResult { Output = "Attendee changes only apply to the whole series - omit occurrence_date.", IsError = true, Summary = "edit_event" };

            if (occDate != null && start.HasValue)
            {
                ToolResult? collision = CheckOccurrenceReorderCollision(master, appt.Start, start.Value, appt.Subject ?? "", "edit_event");
                if (collision != null) return collision.Value;
            }

            string scopeNote = occDate != null ? " (this occurrence only)" : "";
            bool wasMeetingBefore = master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            bool isRecurringWholeSeriesTimeChange = occDate == null && master.IsRecurring && start.HasValue;

            var props = new List<string>();
            if (start.HasValue && !isRecurringWholeSeriesTimeChange) { props.Add("Start"); props.Add("End"); }
            if (subject != null) props.Add("Subject");
            if (body != null) props.Add("Body");
            if (location != null) props.Add("Location");
            object[] before = ReadProps(appt, props.ToArray());

            string oldStart = Iso(appt.Start);
            string oldEnd = Iso(appt.End);

            // Outlook does not allow setting AppointmentItem.Start/.End directly
            // on a recurring master (confirmed live 2026-09-28, see the
            // retired reschedule tool's own history) - RecurrencePattern's
            // fields are the correct mechanism, same as create_event's
            // recurrence support and that retired tool's whole-series fix.
            if (isRecurringWholeSeriesTimeChange)
            {
                Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
                pattern.PatternStartDate = start.Value.Date;
                pattern.StartTime = start.Value;
                pattern.EndTime = end.Value;
            }
            else if (start.HasValue)
            {
                appt.Start = start.Value;
                appt.End = end.Value;
            }

            if (subject != null) appt.Subject = subject;
            if (body != null) appt.Body = body;
            if (location != null) appt.Location = location;

            bool attendeesChanged = false;
            bool revertToNonMeeting = false;
            var unresolvedAttendees = new List<string>();
            if (requiredAttendees != null || optionalAttendees != null)
            {
                int recipientsBefore = CountAttendeeRecipients(appt);
                ReplaceAttendees(appt, requiredAttendees, optionalAttendees, unresolvedAttendees);
                int recipientsAfter = CountAttendeeRecipients(appt);
                // If clearing brought the real attendee count to zero, the
                // event should revert to olNonMeeting - otherwise it stays
                // permanently "a meeting" (and thus a permanent undo barrier)
                // even after every attendee is removed. The actual flip is
                // deferred to just after .Send() below (not done here) - a
                // code review of this method noted that flipping MeetingStatus
                // to olNonMeeting BEFORE .Send() risks Outlook not treating
                // the send as a cancellation notice to the just-removed
                // attendees. Deferring is safe: reaching recipientsAfter == 0
                // from a meeting always means wasMeetingBefore was true, which
                // always forces isMeetingNow/mustBarrier true below, so the
                // .Send() branch always runs when this flag is set - there is
                // no code path where revertToNonMeeting is set but .Send() is
                // skipped.
                if (recipientsAfter == 0)
                {
                    if (appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting) revertToNonMeeting = true;
                }
                else if (appt.MeetingStatus != Outlook.OlMeetingStatus.olMeeting)
                {
                    appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                }
                // Confirmed default is already false; set explicitly so intent
                // doesn't depend on that default never changing. Outlook's own
                // mechanism for notifying only added/removed attendees rather
                // than everyone on the list.
                appt.ForceUpdateToAllAttendees = false;
                // A call supplying e.g. required_attendees="" on an event that
                // already has zero attendees is a true no-op - nothing was
                // added or removed, so it shouldn't count as an attendee
                // change (which would otherwise force isMeetingNow below and
                // burn a real .Send()/undo barrier for a call that changed
                // nothing).
                attendeesChanged = !(recipientsBefore == 0 && recipientsAfter == 0);
            }

            bool isMeetingNow = wasMeetingBefore || attendeesChanged;
            // Barrier whenever the call sends an invite OR touches whole-series
            // RecurrencePattern fields - the latter can't be snapshotted by
            // SnapshotEntry even with zero attendee involvement (see the
            // retired reschedule tool's own whole-series fix for the same rule).
            bool mustBarrier = isMeetingNow || isRecurringWholeSeriesTimeChange;
            string changeSummary = DescribeEditEventChanges(start, end, subject, body, location, attendeesChanged, oldStart, oldEnd);

            if (mustBarrier)
            {
                if (isMeetingNow)
                {
                    // Send while still flagged as a meeting so Outlook treats
                    // this as a real meeting update/cancellation to whoever
                    // was just removed, THEN apply the deferred revert to
                    // olNonMeeting (if attendees were cleared to zero), THEN
                    // Save. .Send() alone can leave the item's own Saved flag
                    // stuck False even after a successful send - confirmed
                    // live via COM (create a recurring meeting, clear its
                    // attendees, revert MeetingStatus, .Send(): Saved reads
                    // False even on a fresh re-fetch by EntryID, causing
                    // Outlook to prompt "save changes?" if the user later just
                    // opens and closes the item with nothing to change). An
                    // explicit .Save() right after .Send() clears it -
                    // confirmed the same sequence with the extra Save() reads
                    // Saved=True on a fresh re-fetch. The Save() is wrapped
                    // since it runs after the irreversible Send() has already
                    // succeeded - a failure here must not be reported as a
                    // failed edit_event call (the update already went out).
                    appt.Send();
                    if (revertToNonMeeting) appt.MeetingStatus = Outlook.OlMeetingStatus.olNonMeeting;
                    try { appt.Save(); }
                    catch (Exception ex) { DebugLog.WriteException("EditEvent post-send Save", ex); }
                }
                else
                {
                    appt.Save();
                }
                RecordIrreversible(mbxKey, "edit_event of \"" + (appt.Subject ?? "") + "\"" + scopeNote +
                                            (isMeetingNow ? " (update sent)" : "") +
                                            (isRecurringWholeSeriesTimeChange ? " (whole series time change)" : ""));
                return new ToolResult
                {
                    Output = "Updated" + (isMeetingNow ? " and sent update notice" : "") + scopeNote + ": \"" + (appt.Subject ?? "") + "\"." + changeSummary + FormatUnresolvedAttendeesNote(unresolvedAttendees),
                    Mutated = true,
                    Summary = "edit_event",
                };
            }

            try
            {
                appt.Save();
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("EditEvent Save", ex);
                if (occDate != null)
                {
                    // Same false-negative Save() risk confirmed for the
                    // retired reschedule tool's occurrence path: property setters
                    // write immediately via RPC, independent of Save()'s own
                    // finalize step, which can fail separately. Re-check by
                    // the NEW date if a time change was requested (the
                    // occurrence now sits there, not at occDate), else by the
                    // unchanged occDate.
                    Outlook.AppointmentItem recheck;
                    string checkDate = start.HasValue ? Iso(start.Value) : occDate;
                    ToolResult? stillMissing = ResolveOccurrenceTarget(master, checkDate, "edit_event", out recheck);
                    if (stillMissing == null)
                    {
                        appt = recheck;
                    }
                    else
                    {
                        return new ToolResult
                        {
                            Output = "Could not save changes to \"" + (appt.Subject ?? "") + "\": " + ex.Message +
                                     (start.HasValue ? " If this was a time change, it may have crossed another occurrence of the same series - check list_events for this series' other occurrence dates." : ""),
                            IsError = true,
                            Summary = "edit_event",
                        };
                    }
                }
                else
                {
                    return new ToolResult
                    {
                        Output = "Could not save changes to \"" + (appt.Subject ?? "") + "\": " + ex.Message,
                        IsError = true,
                        Summary = "edit_event",
                    };
                }
            }

            // RecordSnapshot reads appt.EntryID via ItemEntryIdOf(appt) - for an
            // occurrence, that's the real, resolvable EntryID GetOccurrence's
            // returned item gets once saved, so undo/redo works via the exact
            // same SnapshotEntry mechanism as the retired reschedule tool - no new
            // undo-entry type needed. props is always non-empty here: reaching
            // this branch requires mustBarrier == false, which means attendees
            // were never touched (that forces isMeetingNow, hence a barrier)
            // and, if a time change was requested, it's occurrence/non-recurring
            // (whole-series-recurring also forces a barrier) - so at least one
            // of Start/End/Subject/Body/Location is always in props.
            RecordSnapshot(mbxKey, "edit_event", appt, appt.Subject ?? "", props.ToArray(), before);
            return new ToolResult
            {
                Output = "Updated" + scopeNote + ": \"" + (appt.Subject ?? "") + "\"." + changeSummary + FormatUnresolvedAttendeesNote(unresolvedAttendees),
                Mutated = true,
                Summary = "edit_event",
            };
        }

        // Draft-tier counterpart to EditEvent: same validation and mutation
        // selection, but every change stays unsaved on the in-memory COM
        // object until appt.Display(false) opens it for the user to review
        // and save/send themselves. Never calls .Save()/.Send(), never
        // touches ForceUpdateToAllAttendees, never records undo/redo -
        // draft tools never persist anything, same contract as draft_event.
        private static ToolResult DraftEditEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_edit_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "draft_edit_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            if (IsCanceledMeeting(master))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "draft_edit_event" };
            if (IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "draft_edit_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            string subject = Str(input, "subject", null);
            string body = Str(input, "body", null);
            string location = Str(input, "location", null);
            string requiredAttendees = Str(input, "required_attendees", null);
            string optionalAttendees = Str(input, "optional_attendees", null);

            if (!start.HasValue && !end.HasValue && subject == null && body == null && location == null &&
                requiredAttendees == null && optionalAttendees == null)
                return new ToolResult { Output = "At least one of start, end, subject, body, location, required_attendees, optional_attendees must be provided.", IsError = true, Summary = "draft_edit_event" };

            if (start.HasValue != end.HasValue)
                return new ToolResult { Output = "start and end must be provided together.", IsError = true, Summary = "draft_edit_event" };

            if (start.HasValue && end.Value <= start.Value)
                return new ToolResult { Output = "end must be after start.", IsError = true, Summary = "draft_edit_event" };

            if ((requiredAttendees != null || optionalAttendees != null) && occDate != null)
                return new ToolResult { Output = "Attendee changes only apply to the whole series - omit occurrence_date.", IsError = true, Summary = "draft_edit_event" };

            if (occDate != null && start.HasValue)
            {
                ToolResult? collision = CheckOccurrenceReorderCollision(master, appt.Start, start.Value, appt.Subject ?? "", "draft_edit_event");
                if (collision != null) return collision.Value;
            }

            string scopeNote = occDate != null ? " (just this occurrence, not the whole series)" : "";
            bool isMeeting = master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;

            if (occDate == null && master.IsRecurring && start.HasValue)
            {
                Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
                pattern.PatternStartDate = start.Value.Date;
                pattern.StartTime = start.Value;
                pattern.EndTime = end.Value;
            }
            else if (start.HasValue)
            {
                appt.Start = start.Value;
                appt.End = end.Value;
            }

            if (subject != null) appt.Subject = subject;
            if (body != null) appt.Body = body;
            if (location != null) appt.Location = location;

            bool attendeesChanged = false;
            var unresolvedAttendees = new List<string>();
            if (requiredAttendees != null || optionalAttendees != null)
            {
                int recipientsBefore = CountAttendeeRecipients(appt);
                ReplaceAttendees(appt, requiredAttendees, optionalAttendees, unresolvedAttendees);
                int recipientsAfter = CountAttendeeRecipients(appt);
                // Same revert-to-non-meeting fix as EditEvent: if clearing
                // attendees brought the real attendee count to zero, don't
                // leave the item permanently marked as a meeting. No .Send()
                // ordering concern here - draft tools never send.
                if (recipientsAfter == 0)
                {
                    if (appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting) appt.MeetingStatus = Outlook.OlMeetingStatus.olNonMeeting;
                }
                else if (appt.MeetingStatus != Outlook.OlMeetingStatus.olMeeting)
                {
                    appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                }
                // Same no-op guard as EditEvent: an already-empty attendee
                // list touched with e.g. required_attendees="" changes nothing.
                attendeesChanged = !(recipientsBefore == 0 && recipientsAfter == 0);
                isMeeting = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            }

            appt.Display(false);
            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\"" + scopeNote + " with the requested changes in Outlook for the user to review and " +
                         (isMeeting ? "save/send." : "save.") + (attendeesChanged ? " Attendee list updated - review before sending." : "") + FormatUnresolvedAttendeesNote(unresolvedAttendees),
                Summary = "draft_edit_event",
            };
        }

        // Shared by draft_cancel_event/cancel_event - same organizer-authority
        // shape as edit_event above (IsCanceledMeeting/IsReceivedMeeting
        // reused, not duplicated), but an attendee's remedy for a meeting they
        // don't organize is decline_meeting, not "Propose New Time".
        private static string ReceivedMeetingCancelError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" is a meeting organized by " +
                   (string.IsNullOrEmpty(appt.Organizer) ? "someone else" : appt.Organizer) +
                   " - you're only an attendee, not the organizer, so this tool has no authority to cancel it. " +
                   "Use draft_respond_meeting (or decline_meeting, in Automate approvals or above) instead.";
        }

        // Only used by draft_cancel_event now - cancel_event treats an
        // already-canceled event as a cleanup (move to Deleted Items, no
        // resend) rather than refusing, so it never needs this message.
        private static string AlreadyCanceledError(Outlook.AppointmentItem appt)
        {
            return "\"" + (appt.Subject ?? "") + "\" has already been canceled - use cancel_event to remove it from your calendar.";
        }

        // Draft-tier: just opens the item unchanged, for both branches - the
        // user drives Outlook's own native "Cancel Meeting"/"Send
        // Cancellation" UI (or Delete, for a plain appointment) from there.
        //
        // Originally set MeetingStatus = olMeetingCanceled unsaved before
        // Display(false), mirroring the retired draft reschedule tool's unsaved
        // Start/End - removed 2026-09-28 after live testing showed Outlook
        // persists that change when the Inspector closes even without the
        // user clicking "Send Cancellation" (unlike Start/End, an unsaved
        // MeetingStatus change apparently isn't purely cosmetic here). That
        // silently canceled the meeting locally with attendees never
        // notified - worse than doing nothing, since it also means this tool
        // can no longer be un-done or reattempted (cancel_event/edit_event
        // both refuse on an already-canceled item). Never
        // mutates the item at all now.
        private static ToolResult DraftCancelEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_cancel_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "draft_cancel_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            if (IsCanceledMeeting(master))
                return new ToolResult { Output = AlreadyCanceledError(appt), IsError = true, Summary = "draft_cancel_event" };
            if (IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingCancelError(appt), IsError = true, Summary = "draft_cancel_event" };

            bool isMeeting = master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            string scopeNote = occDate != null ? " (just this occurrence)" : "";
            appt.Display(false);
            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\"" + scopeNote + " in Outlook for the user to review and " +
                         (isMeeting ? "cancel it themselves (Cancel Meeting, then Send Cancellation) if they want to proceed."
                                    : "delete from the calendar."),
                Summary = "draft_cancel_event",
            };
        }

        // Full-autonomy-only counterpart to draft_cancel_event above: for a
        // meeting the user organizes, sends the cancellation notice
        // immediately (no review step, mirroring edit_event's/
        // create_event's attendee-present branches) then removes it from the
        // user's own calendar; for a plain appointment, just removes it -
        // nobody to notify. Either way the item is moved to Deleted Items
        // (recoverable there), same as delete_email's non-permanent path, not
        // permanently deleted.
        //
        // An already-canceled event (IsCanceledMeeting - either the
        // organizer's own olMeetingCanceled copy, or an attendee's stale
        // olMeetingReceivedAndCanceled one) is treated as cleanup, not
        // refused: there's nothing new to notify anyone of, so it just moves
        // straight to Deleted Items like a plain appointment. This is the
        // only way to dismiss a canceled event at all - draft_cancel_event
        // still refuses on one (nothing to preview/review for a cleanup).
        private static ToolResult CancelEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "cancel_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "cancel_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            bool isOccurrence = occDate != null;
            // An occurrence has no "already canceled" state to clean up - a
            // deleted occurrence simply won't resolve via GetOccurrence()
            // again, surfacing as ResolveOccurrenceTarget's "no occurrence on
            // that date" error above instead of reaching this point.
            bool alreadyCanceled = !isOccurrence && IsCanceledMeeting(appt);
            if (!alreadyCanceled && IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingCancelError(appt), IsError = true, Summary = "cancel_event" };

            string subject = appt.Subject ?? "";
            bool isMeeting = !alreadyCanceled && master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            string scopeNote = isOccurrence ? " (this occurrence only)" : "";

            if (isMeeting)
            {
                appt.MeetingStatus = Outlook.OlMeetingStatus.olMeetingCanceled;
                appt.Send();
                // Barrier, not a move snapshot: the cancellation notice already
                // went out to attendees, so restoring this copy would only
                // half-undo the action and misleadingly imply it was fully
                // reversed.
                RecordIrreversible(mbxKey, "cancel_event invite cancellation for \"" + subject + "\"" + scopeNote);

                if (isOccurrence)
                {
                    // No COM API to un-delete a single occurrence
                    // (RecurrencePattern.Exceptions is entirely read-only -
                    // confirmed via reflection) - always a barrier above,
                    // regardless of plain-appointment vs. meeting, unlike
                    // whole-event cancellation below which stays undo-able
                    // for the plain-appointment case.
                    bool removedLocally = true;
                    try { appt.Delete(); }
                    catch (Exception ex)
                    {
                        DebugLog.WriteException("CancelEvent occurrence delete", ex);
                        // Same false-negative risk confirmed for the retired
                        // reschedule tool's occurrence Save() - re-check via
                        // ResolveOccurrenceTarget rather than assume the
                        // exception means the delete failed.
                        Outlook.AppointmentItem recheck;
                        ToolResult? stillGone = ResolveOccurrenceTarget(master, occDate, "cancel_event", out recheck);
                        removedLocally = stillGone != null; // non-null = "no occurrence found" = it's gone = delete succeeded despite the exception
                    }
                    return new ToolResult
                    {
                        Output = "Canceled and notified attendees" + scopeNote + ": \"" + subject + "\"." +
                                 (removedLocally ? "" : " Outlook would not remove this occurrence - delete it manually.") +
                                 " This occurrence cannot be undone.",
                        Mutated = true,
                        Summary = "cancel_event",
                    };
                }

                Outlook.Folder sourceFolder = (Outlook.Folder)appt.Parent;
                Outlook.Folder deletedFolder = (Outlook.Folder)sourceFolder.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                // Unverified whether Outlook still allows moving the item after
                // Send() on a now-canceled meeting - reflection confirms both
                // members exist, but chaining them hasn't been exercised
                // against a live client. Fails soft: the notice is the
                // irreversible part and already went out either way.
                bool removedFromCalendar = true;
                try { appt.Move(deletedFolder); }
                catch (Exception ex) { DebugLog.WriteException("CancelEvent post-send move", ex); removedFromCalendar = false; }
                return new ToolResult
                {
                    Output = "Canceled and notified attendees: \"" + subject + "\"" +
                             (removedFromCalendar
                                 ? " removed from your calendar."
                                 : " (Outlook would not remove it from your calendar after sending - delete it manually)."),
                    Mutated = true,
                    Summary = "cancel_event",
                };
            }

            if (isOccurrence)
            {
                // Plain-appointment occurrence: still a barrier, not a move -
                // deleting one occurrence has no folder-move equivalent to
                // record (it's dropped from the pattern's read-only Exceptions
                // list, not relocated to a recoverable folder).
                try { appt.Delete(); }
                catch (Exception ex)
                {
                    DebugLog.WriteException("CancelEvent occurrence delete", ex);
                    // Same false-negative risk confirmed for the retired
                    // reschedule tool's occurrence Save() - re-check before
                    // reporting failure.
                    Outlook.AppointmentItem recheck;
                    ToolResult? stillGone = ResolveOccurrenceTarget(master, occDate, "cancel_event", out recheck);
                    if (stillGone == null)
                        return new ToolResult { Output = "Could not cancel \"" + subject + "\": " + ex.Message, IsError = true, Summary = "cancel_event" };
                }
                RecordIrreversible(mbxKey, "cancel_event of \"" + subject + "\" occurrence");
                return new ToolResult { Output = "Canceled" + scopeNote + ": \"" + subject + "\". This occurrence cannot be undone.", Mutated = true, Summary = "cancel_event" };
            }

            {
                Outlook.Folder sourceFolder = (Outlook.Folder)appt.Parent;
                Outlook.Folder deletedFolder = (Outlook.Folder)sourceFolder.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                string oldId = appt.EntryID;
                dynamic moved = appt.Move(deletedFolder);
                string newId = moved.EntryID;
                RecordMove(mbxKey, "cancel_event", "event_id", subject, oldId, newId, sourceFolder, deletedFolder);
                return new ToolResult
                {
                    Output = (alreadyCanceled
                                 ? "Removed already-canceled event \"" + subject + "\""
                                 : "Canceled: \"" + subject + "\" moved") +
                             " to Deleted Items.\nevent_id: " + newId,
                    Mutated = true,
                    Summary = "cancel_event",
                };
            }
        }
    }
}
