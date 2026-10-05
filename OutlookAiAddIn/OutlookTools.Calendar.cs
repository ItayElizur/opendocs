using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        private static async Task<ToolResult> ListEventsAsync(JsonElement input)
        {
            DateTime start = (DateArg(input, "start_date") ?? DateTime.Today).Date;
            DateTime end = (DateArg(input, "end_date") ?? DateTime.Today.AddDays(7)).Date.AddDays(1);
            int limit = Math.Max(1, Int(input, "limit", 50));
            // An LLM tool caller omitting an optional string parameter often
            // sends "" rather than leaving it out entirely - treat that the
            // same as not having provided a mailbox at all, so it doesn't get
            // routed into the shared-calendar path (and ultimately into
            // Ns.CreateRecipient("") below) as if it were a real address.
            string mailbox = Str(input, "mailbox", null);
            if (string.IsNullOrWhiteSpace(mailbox)) mailbox = null;
            else mailbox = mailbox.Trim();

            if (mailbox == null)
            {
                Outlook.Folder ownCal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                StringBuilder ownSb;
                int ownN;
                QueryCalendarItems(ownCal, start, end, limit, null, null, out ownSb, out ownN);
                return BuildListEventsResult(ownSb, ownN, start, end, null);
            }

            Outlook.Recipient recipient = Ns.CreateRecipient(mailbox);
            bool resolved;
            try { resolved = recipient.Resolve(); } catch { resolved = false; }
            if (!resolved)
                return new ToolResult { Output = "Could not resolve \"" + mailbox + "\" - check the email address.", IsError = true, Summary = "list_events" };

            // Prefer the resolved recipient's own display name for output
            // text over the raw input, so e.g. mailbox: "dana" echoes back
            // who it actually resolved to rather than the ambiguous string
            // the caller typed.
            string displayName = !string.IsNullOrEmpty(recipient.Name) ? recipient.Name : mailbox;

            // This call is ONLY used to read .Store.StoreID now (get_event's
            // store_id parameter needs it) - the actual event data comes from
            // EWS below, off the UI thread. See OutlookTools.Calendar.cs.md
            // for why (a prior COM enumeration here froze Outlook).
            Outlook.Folder sharedCal;
            try
            {
                sharedCal = (Outlook.Folder)Ns.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderCalendar);
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("ListEvents GetSharedDefaultFolder", ex);
                return new ToolResult { Output = "Could not open " + mailbox + "'s calendar - you may not have been granted access to view it, or need to add it via Outlook's own \"Open Calendar\" first. (" + ex.Message + ")", IsError = true, Summary = "list_events" };
            }

            // Split from the GetSharedDefaultFolder try/catch above on purpose:
            // a StoreID read failure here means the folder itself opened fine,
            // so it should degrade to a missing store_id, not be misreported
            // as "you may not have been granted access". See
            // OutlookTools.Calendar.cs.md.
            string storeId;
            try { storeId = sharedCal.Store == null ? null : sharedCal.Store.StoreID; }
            catch (Exception ex) { DebugLog.WriteException("ListEvents shared calendar StoreID", ex); storeId = null; }

            string sharedSmtp = SmtpOf(recipient.AddressEntry);
            if (string.IsNullOrEmpty(sharedSmtp)) sharedSmtp = mailbox;

            Uri url;
            try
            {
                url = await ResolveEwsUrlAsync();
            }
            catch (InvalidOperationException ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = "list_events" };
            }

            IReadOnlyList<SharedCalendarEventRow> rows;
            try
            {
                rows = await OutlookEws.GetSharedCalendarEventsAsync(url, sharedSmtp, start, end, limit);
            }
            catch (Exception ex) when (OutlookEws.IsTimeout(ex))
            {
                DebugLog.WriteException("list_events shared calendar timeout", ex);
                return new ToolResult { Output = "The Exchange calendar lookup for " + mailbox + " timed out after 15s. Try again, or check your network / VPN connection.", IsError = true, Summary = "list_events" };
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException ex)
            {
                DebugLog.WriteException("list_events shared calendar ServiceRequestException", ex);
                return new ToolResult { Output = "Exchange rejected the calendar lookup for " + mailbox + ": " + ex.Message + " (Windows authentication to Exchange may have failed - are you on the domain network?)", IsError = true, Summary = "list_events" };
            }
            catch (WebException ex)
            {
                DebugLog.WriteException("list_events shared calendar WebException", ex);
                return new ToolResult
                {
                    Output = ex.Status == WebExceptionStatus.ProtocolError
                        ? "Windows authentication to Exchange failed (are you connected to the domain network / VPN?)."
                        : "Could not reach the Exchange server (" + ex.Status + ").",
                    IsError = true,
                    Summary = "list_events",
                };
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("list_events shared calendar (unexpected)", ex);
                return new ToolResult { Output = "Could not read " + mailbox + "'s calendar: " + ex.Message, IsError = true, Summary = "list_events" };
            }

            string text = SharedCalendarEventFormat.Format(rows, limit, displayName, storeId);
            int n = Math.Min(rows.Count, limit);
            return BuildListEventsResult(new StringBuilder(text), n, start, end, displayName);
        }

        // Ordering is load-bearing: Sort("[Start]") -> IncludeRecurrences = true
        // -> Restrict. Any other order silently drops recurring instances, and
        // IncludeRecurrences rules out the faster GetTable path.
        private static void QueryCalendarItems(Outlook.Folder cal, DateTime start, DateTime end, int limit, string mailbox, string storeId, out StringBuilder sb, out int n)
        {
            Outlook.Items items = cal.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            string filter = "[Start] <= '" + end.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [End] >= '" + start.ToString("g", CultureInfo.CurrentCulture) + "'";
            Outlook.Items restricted = items.Restrict(filter);

            sb = new StringBuilder();
            n = 0;
            foreach (object o in restricted)
            {
                if (n >= limit) break;
                Outlook.AppointmentItem appt = o as Outlook.AppointmentItem;
                if (appt == null) continue;
                n++;
                // Whatever Outlook itself resolves for Subject/Location is
                // passed through as-is, including for private items - this
                // add-in adds no visibility restriction of its own on top of
                // the caller's real Exchange permissions. See
                // OutlookTools.Calendar.cs.md.
                sb.AppendLine("- event_id: " + appt.EntryID);
                sb.AppendLine("  subject: " + (appt.Subject ?? ""));
                try { sb.AppendLine("  start: " + Iso(appt.Start) + "  end: " + Iso(appt.End)); } catch { }
                sb.AppendLine("  location: " + (appt.Location ?? ""));
                sb.AppendLine("  organizer: " + (appt.Organizer ?? "") + "  all_day: " + appt.AllDayEvent + "  recurring: " + appt.IsRecurring);
                sb.AppendLine("  response: " + appt.ResponseStatus + "  meeting_status: " + appt.MeetingStatus);
                if (mailbox != null)
                {
                    sb.AppendLine("  calendar_owner: " + mailbox);
                    // get_event needs this to resolve an event_id outside
                    // the caller's own default store - see GetEvent's
                    // store_id parameter.
                    sb.AppendLine("  store_id: " + storeId);
                }
            }
        }

        private static ToolResult BuildListEventsResult(StringBuilder sb, int n, DateTime start, DateTime end, string mailbox)
        {
            // mailbox == null must produce the exact pre-feature message
            // ("No events between X and Y.") with zero format change - see
            // the mailbox != null case for the only place "on <mailbox>'s
            // calendar" is introduced.
            string whoseCalendar = mailbox != null ? " on " + mailbox + "'s calendar" : "";
            if (n == 0)
                return new ToolResult { Output = "No events" + whoseCalendar + " between " + start.ToShortDateString() + " and " + end.AddDays(-1).ToShortDateString() + ".", Summary = "list_events" };
            return new ToolResult { Output = sb + "\n(Recurring instances share the master event_id.)", Summary = "list_events" };
        }

        private static ToolResult GetEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events.
            string storeId = Str(input, "store_id", null);
            Outlook.AppointmentItem appt = ItemById(id, storeId) as Outlook.AppointmentItem;
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

        // Default range for an arbitrary work-days set: walk forward from
        // today to the next work day, then extend through the following
        // contiguous run of work days (capped at a week). See
        // OutlookTools.Calendar.cs.md for the mcp-outlook lineage this
        // generalizes.
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

        // response is the actual OlMeetingResponse to send, so this one
        // helper covers all three: accept_meeting, decline_meeting,
        // tentative_meeting. message is an optional comment attached before
        // sending - UNVERIFIED live whether the organizer actually sees it
        // on the delivered response.
        private static ToolResult RespondMeeting(string mbxKey, JsonElement input, Outlook.OlMeetingResponse response, string toolName)
        {
            string id = ReqStr(input, "event_id");
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            object item = ItemById(id, storeId);
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

        // Draft-tier counterpart to RespondMeeting: never mutates the item.
        // Opens it unchanged; the user picks Accept/Tentative/Decline from
        // Outlook's own native ribbon buttons, and a suggested comment is
        // returned in the output text instead of pre-filled. See
        // OutlookTools.Calendar.cs.md for why appt.Respond() itself can't be
        // called here (it was tried and broke the draft-tools contract).
        private static ToolResult DraftRespondMeeting(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            object item = ItemById(id, storeId);
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

        // AppointmentItem has no Reply/Forward in Outlook's object model - only
        // MeetingItem, the invitation message in the Inbox, does. This finds
        // that message for a calendar event: first through the event's
        // conversation (verified live: the invite shows up there), then by
        // matching PidLidGlobalObjectId against meeting messages in the Inbox.
        // Null when the invite is gone (deleted/archived) or the user
        // organizes the event (an organizer has no invite message).
        private static Outlook.MeetingItem FindInviteMessage(Outlook.AppointmentItem appt)
        {
            try
            {
                Outlook.Conversation conv = appt.GetConversation();
                if (conv != null)
                {
                    foreach (object o in conv.GetRootItems())
                    {
                        Outlook.MeetingItem mi = o as Outlook.MeetingItem;
                        if (mi != null && (mi.MessageClass ?? "").StartsWith("IPM.Schedule.Meeting.Request", StringComparison.OrdinalIgnoreCase)) return mi;
                    }
                }
            }
            catch (Exception ex) { DebugLog.WriteException("FindInviteMessage conversation", ex); }

            try
            {
                const string gidTag = "http://schemas.microsoft.com/mapi/id/{6ED8DA90-450B-101B-98DA-00AA003F1305}/00030102";
                byte[] want = appt.PropertyAccessor.GetProperty(gidTag) as byte[];
                if (want == null) return null;
                Outlook.Folder inbox = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                int scanned = 0;
                foreach (object o in inbox.Items)
                {
                    if (++scanned > 3000) break;
                    Outlook.MeetingItem mi = o as Outlook.MeetingItem;
                    if (mi == null || !(mi.MessageClass ?? "").StartsWith("IPM.Schedule.Meeting.Request", StringComparison.OrdinalIgnoreCase)) continue;
                    try
                    {
                        byte[] got = mi.PropertyAccessor.GetProperty(gidTag) as byte[];
                        if (got != null && got.Length == want.Length && System.Linq.Enumerable.SequenceEqual(got, want)) return mi;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { DebugLog.WriteException("FindInviteMessage global id", ex); }
            return null;
        }

        private static string NoInviteError(Outlook.AppointmentItem appt)
        {
            return "Couldn't find the invitation email for \"" + (appt.Subject ?? "") + "\" - it may have been deleted or archived, or you organize this event (an organizer has no invitation to reply to or forward).";
        }

        // Opens Outlook's own reply (or reply-all) on the event's invitation
        // message, so it is threaded and quoted exactly like a reply made by
        // hand. Only opened for review, never sent. Errors when the invitation
        // message can't be found rather than building a lookalike.
        private static ToolResult DraftReplyEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string storeId = Str(input, "store_id", null);
            bool all = Bool(input, "reply_all", false);
            string body = Str(input, "body", "");
            const string tool = "draft_reply_event";

            Outlook.AppointmentItem appt = ResolveMeetingAppointment(ItemById(id, storeId));
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a calendar event.", IsError = true, Summary = tool };

            Outlook.MeetingItem invite = FindInviteMessage(appt);
            if (invite == null)
                return new ToolResult { Output = NoInviteError(appt), IsError = true, Summary = tool };

            Outlook.MailItem reply = all ? invite.ReplyAll() : invite.Reply();
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            reply.Display(false);
            return new ToolResult
            {
                Output = "Opened Outlook's own " + (all ? "reply-all" : "reply") + " to the invitation for \"" + (appt.Subject ?? "") + "\" for the user to review and send.",
                Summary = tool,
            };
        }

        // Opens Outlook's own meeting forward of the event's invitation
        // message (not an .ics attachment). Only opened for review, never sent.
        private static ToolResult DraftForwardEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string storeId = Str(input, "store_id", null);
            string to = Str(input, "to", "");
            string body = Str(input, "body", "");
            const string tool = "draft_forward_event";

            Outlook.AppointmentItem appt = ResolveMeetingAppointment(ItemById(id, storeId));
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a calendar event.", IsError = true, Summary = tool };

            Outlook.MeetingItem invite = FindInviteMessage(appt);
            if (invite == null)
                return new ToolResult { Output = NoInviteError(appt), IsError = true, Summary = tool };

            // _MeetingItem: the plain interface name avoids the Forward
            // method/event name clash on MeetingItem.
            Outlook.MeetingItem fwd = ((Outlook._MeetingItem)invite).Forward();
            var unresolved = new List<string>();
            foreach (string part in to.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string addr = part.Trim();
                int lt = addr.LastIndexOf('<');
                if (lt >= 0 && addr.EndsWith(">")) addr = addr.Substring(lt + 1, addr.Length - lt - 2).Trim();
                if (addr.Length == 0) continue;
                Outlook.Recipient r = fwd.Recipients.Add(addr);
                try { r.Resolve(); } catch { }
                if (!r.Resolved) unresolved.Add(addr);
            }
            if (!string.IsNullOrEmpty(body))
            {
                try { fwd.Body = SeedSignature(body) + "\n\n" + fwd.Body; }
                catch (Exception ex) { DebugLog.WriteException("DraftForwardEvent body", ex); }
            }
            fwd.Display(false);
            return new ToolResult
            {
                Output = "Opened Outlook's own forward of the invitation for \"" + (appt.Subject ?? "") + "\" for the user to review and send." + FormatUnresolvedAttendeesNote(unresolved),
                Summary = tool,
            };
        }

        // Shared by draft_edit_event/edit_event: an olMeetingReceived (or
        // olMeetingReceivedAndCanceled) appointment is one the user only
        // attends, not organizes - no COM API exists to honor a reschedule
        // request on it, so both tools refuse up front. Both "received"
        // statuses count, not just the exact one - see
        // OutlookTools.Calendar.cs.md for why (and for the reflection
        // evidence behind the "no COM API" claim).
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

        // Shared by all four occurrence-aware tools. occurrence_date lets a
        // caller target one instance of a recurring series instead of the
        // whole master, via RecurrencePattern.GetOccurrence(DateTime)
        // (confirmed present via reflection - see OutlookTools.Calendar.cs.md).
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
        // cross or land on the same day as another occurrence of the same
        // series - Outlook rejects both, but only via a generic COMException
        // with no way to distinguish the cause. Checking this ourselves first
        // gives an exact, deterministic answer instead. See
        // OutlookTools.Calendar.cs.md for the confirmed repro and the query
        // approach's details.
        private static ToolResult? CheckOccurrenceReorderCollision(Outlook.AppointmentItem master, DateTime original, DateTime target, string subject, string toolName)
        {
            DateTime rangeStart = (original < target ? original : target).Date;
            DateTime rangeEnd = (original < target ? target : original).Date.AddDays(1);

            // master's own Parent folder, not the caller's default calendar -
            // those differ for a shared-calendar event resolved via store_id.
            // See OutlookTools.Calendar.cs.md for the bug this fixed.
            Outlook.Folder cal = (Outlook.Folder)master.Parent;
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
        // optional attendee list wholesale, not a diff/merge. A null argument
        // leaves that category untouched; non-null (including "") fully
        // replaces it. The organizer is never touched. See
        // OutlookTools.Calendar.cs.md for the iteration-order note.
        private static void ReplaceAttendees(Outlook.AppointmentItem appt, string requiredCsv, string optionalCsv, List<string> unresolved)
        {
            for (int i = appt.Recipients.Count; i >= 1; i--)
            {
                int type = appt.Recipients[i].Type;
                if (requiredCsv != null && type == (int)Outlook.OlMeetingRecipientType.olRequired) { appt.Recipients.Remove(i); continue; }
                if (optionalCsv != null && type == (int)Outlook.OlMeetingRecipientType.olOptional) { appt.Recipients.Remove(i); continue; }
            }
            if (requiredCsv != null) AddAttendees(appt, requiredCsv, Outlook.OlMeetingRecipientType.olRequired, unresolved);
            if (optionalCsv != null) AddAttendees(appt, optionalCsv, Outlook.OlMeetingRecipientType.olOptional, unresolved);
        }

        // Counts real attendees only, excluding the organizer - used to
        // decide whether clearing attendees should revert MeetingStatus back
        // to olNonMeeting. See OutlookTools.Calendar.cs.md for why the
        // exclusion is defensive insurance rather than an observed-bug fix.
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
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, storeId) as Outlook.AppointmentItem;
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
                // If clearing brought attendee count to zero, revert to
                // olNonMeeting - but only AFTER .Send() below, not here (see
                // OutlookTools.Calendar.cs.md for why flipping it early risks
                // Outlook not treating the send as a cancellation notice, and
                // for why deferring is provably safe).
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
                // required_attendees="" on an event already at zero attendees
                // is a true no-op - don't let it force isMeetingNow below and
                // burn a real .Send()/undo barrier for nothing.
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
                    // Send while still flagged as a meeting, THEN apply the
                    // deferred revert to olNonMeeting, THEN Save - .Send()
                    // alone can leave Saved stuck False (confirmed live via
                    // COM); see OutlookTools.Calendar.cs.md. The Save() is
                    // wrapped since it runs after the irreversible Send() has
                    // already succeeded - a failure here isn't a failed call.
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
                    // Property setters write immediately via RPC, independent
                    // of Save()'s own finalize step, which can fail
                    // separately - re-check rather than trust the exception.
                    // Re-check by the NEW date if a time change was
                    // requested, else by the unchanged occDate.
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

            // props is always non-empty here: reaching this branch requires
            // mustBarrier == false, which rules out an untouched attendee
            // list and a whole-series recurring time change - see
            // OutlookTools.Calendar.cs.md for the full argument and for why
            // RecordSnapshot needs no new undo-entry type for an occurrence.
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
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, storeId) as Outlook.AppointmentItem;
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
        // Never mutates the item. See OutlookTools.Calendar.cs.md for the
        // live incident that ruled out pre-setting MeetingStatus unsaved.
        private static ToolResult DraftCancelEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, storeId) as Outlook.AppointmentItem;
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

        // Full-autonomy-only counterpart to draft_cancel_event: for an
        // organized meeting, sends the cancellation notice immediately then
        // moves it to Deleted Items (recoverable); for a plain appointment,
        // just moves it. An already-canceled event is treated as cleanup,
        // not refused - see OutlookTools.Calendar.cs.md for why that's the
        // only way to dismiss one at all.
        private static ToolResult CancelEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            // Needed only for an event_id from someone else's shared calendar
            // (list_events' mailbox parameter); omit for your own events. See
            // OutlookTools.Calendar.cs.md for the authorization-scope note.
            string storeId = Str(input, "store_id", null);
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, storeId) as Outlook.AppointmentItem;
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
                    // (RecurrencePattern.Exceptions is read-only - confirmed
                    // via reflection), so this is always a barrier, unlike
                    // whole-event cancellation below.
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
