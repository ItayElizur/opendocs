using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;
using OfficeAi.Shared;

namespace OutlookAiAddIn
{
    // The non-COM calls in the Outlook add-in - the raw EWS wire layer.
    // Tool-facing orchestration built on top of this (which account/endpoint
    // to use, turning a result into a ToolResult) lives in
    // OutlookTools.Ews.cs, not here; this file only wraps the EWS Managed API
    // itself. Two operations: ResolveNamesAsync (search_contacts - a
    // server-side Ambiguous Name Resolution over Contacts then the GAL,
    // exactly what the native Address Book dialog does and what the
    // mcp-outlook reference's account.protocol.resolve_names does) and
    // GetWorkingHoursAsync (find_meeting_slots' work-week default - see its
    // own comment). Both are deliberately NOT Microsoft.Office.Interop.Outlook:
    // the COM object model can do neither a multi-result directory search nor
    // expose a mailbox's configured work week, and EWS is plain HTTP with no
    // STA affinity so both run off the UI thread (Task.Run) and never freeze
    // Outlook.
    //
    // Auth is ExchangeService.UseDefaultCredentials (Windows Integrated Auth as
    // the signed-in user) - the .NET equivalent of mcp-outlook's auth_type=sspi.
    // No stored credentials, no impersonation. On-prem Exchange only.
    //
    // This file has no `using Microsoft.Office.Interop.Outlook`: the two
    // namespaces collide on Contact, EmailAddress, Folder, Task, Item, ...
    internal static class OutlookEws
    {
        private const int TimeoutMs = 15000;

        static OutlookEws()
        {
            // EWS over HTTPS to on-prem Exchange fails with "The underlying connection
            // was closed: An unexpected error occurred on a send" when the process's
            // default security protocol doesn't offer TLS 1.2 - the .NET Framework
            // default and the EWS Managed API 2.2 (2014) vintage can both leave it off.
            // OR it in without disturbing anything already enabled.
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
        }

        // Resolved once per process. Written on the UI thread after the first
        // successful discovery (the agent loop runs tool calls sequentially, so
        // there is never a concurrent writer).
        internal static Uri CachedUrl { get; set; }

        private static bool AllowHttpsRedirection(string redirectionUrl)
        {
            return redirectionUrl != null &&
                   redirectionUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }

        private static Ews.ExchangeService NewService(Uri url)
        {
            var svc = new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2010_SP2)
            {
                UseDefaultCredentials = true, // do NOT also set Credentials
                Timeout = TimeoutMs,
                Url = url,
            };
            return svc;
        }

        // Autodiscover fallback for when Account.AutoDiscoverXml was empty or
        // unparseable. Runs the AD SCP lookup + HTTP off the UI thread.
        public static Task<Uri> DiscoverUrlAsync(string smtp)
        {
            return Task.Run(() =>
            {
                var svc = new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2010_SP2)
                {
                    UseDefaultCredentials = true,
                    Timeout = TimeoutMs,
                };
                svc.AutodiscoverUrl(smtp, AllowHttpsRedirection);
                return svc.Url;
            });
        }

        public static Task<IReadOnlyList<ContactMatch>> ResolveNamesAsync(Uri url, string query)
        {
            return Task.Run(() => ResolveNames(url, query));
        }

        internal struct WorkWeekInfo
        {
            public HashSet<DayOfWeek> Days;
            public double StartHour;
            public double EndHour;
        }

        // GetUserAvailability's WorkingHours is EWS's documented, server-side
        // source for a mailbox's configured work days/hours - the same data
        // Outlook itself uses to shade "outside working hours" in the
        // scheduling assistant. Unlike Outlook's local Calendar Options
        // dialog, this has no COM equivalent at all: confirmed via .NET
        // reflection against the referenced Microsoft.Office.Interop.Outlook
        // PIA that no Application.CalendarOptions property (or any
        // WorkDay*/FirstDayOfWeek member) exists anywhere in that assembly,
        // so this EWS call is the only real, non-hardcoded source for it.
        public static Task<WorkWeekInfo?> GetWorkingHoursAsync(Uri url, string smtp)
        {
            return Task.Run(() => GetWorkingHours(url, smtp));
        }

        private static WorkWeekInfo? GetWorkingHours(Uri url, string smtp)
        {
            Ews.ExchangeService svc = NewService(url);
            var window = new Ews.TimeWindow(DateTime.Today, DateTime.Today.AddDays(7));
            Ews.GetUserAvailabilityResults results;
            try
            {
                results = svc.GetUserAvailability(
                    new[] { new Ews.AttendeeInfo(smtp) },
                    window,
                    Ews.AvailabilityData.FreeBusy);
            }
            catch (Ews.ServiceResponseException)
            {
                return null;
            }

            foreach (Ews.AttendeeAvailability a in results.AttendeesAvailability)
            {
                if (a.WorkingHours == null) continue;
                var days = new HashSet<DayOfWeek>();
                foreach (Ews.DayOfTheWeek d in a.WorkingHours.DaysOfTheWeek)
                {
                    switch (d)
                    {
                        case Ews.DayOfTheWeek.Sunday: days.Add(DayOfWeek.Sunday); break;
                        case Ews.DayOfTheWeek.Monday: days.Add(DayOfWeek.Monday); break;
                        case Ews.DayOfTheWeek.Tuesday: days.Add(DayOfWeek.Tuesday); break;
                        case Ews.DayOfTheWeek.Wednesday: days.Add(DayOfWeek.Wednesday); break;
                        case Ews.DayOfTheWeek.Thursday: days.Add(DayOfWeek.Thursday); break;
                        case Ews.DayOfTheWeek.Friday: days.Add(DayOfWeek.Friday); break;
                        case Ews.DayOfTheWeek.Saturday: days.Add(DayOfWeek.Saturday); break;
                        // Day/Weekday/WeekendDay are input-only aggregate
                        // values per the enum's own shape (confirmed via
                        // reflection) - never expected back from the server,
                        // so deliberately not expanded here.
                    }
                }
                if (days.Count == 0) continue;
                return new WorkWeekInfo
                {
                    Days = days,
                    // Kept as fractional hours, not truncated to an int: a
                    // mailbox configured for e.g. 08:30-17:30 would otherwise
                    // silently become 08:00-17:00 (offering a slot before the
                    // real start, and dropping the valid 17:00-17:30 slot).
                    StartHour = a.WorkingHours.StartTime.TotalHours,
                    EndHour = a.WorkingHours.EndTime.TotalHours,
                };
            }
            return null;
        }

        // list_events' shared-calendar path (mailbox parameter): EWS CalendarView
        // expands recurring appointments server-side within [start, end) - the EWS
        // equivalent of the COM path's IncludeRecurrences=true, without that path's
        // "expand everything, then filter" cost, and off the UI thread like every
        // other EWS call in this file. See docs/superpowers/specs/
        // 2026-09-30-outlook-shared-calendar-ews-design.md for why this exists (the
        // COM path froze Outlook, confirmed live even for a one-day range).
        public static Task<IReadOnlyList<SharedCalendarEventRow>> GetSharedCalendarEventsAsync(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit)
        {
            return Task.Run(() => GetSharedCalendarEvents(url, mailboxSmtp, start, end, limit));
        }

        private static IReadOnlyList<SharedCalendarEventRow> GetSharedCalendarEvents(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit)
        {
            Ews.ExchangeService svc = NewService(url);
            var folderId = new Ews.FolderId(Ews.WellKnownFolderName.Calendar, new Ews.Mailbox(mailboxSmtp));
            var view = new Ews.CalendarView(start, end, Math.Max(1, limit));
            view.PropertySet = new Ews.PropertySet(
                Ews.BasePropertySet.IdOnly,
                Ews.ItemSchema.Subject,
                Ews.AppointmentSchema.Start,
                Ews.AppointmentSchema.End,
                Ews.AppointmentSchema.Location,
                Ews.AppointmentSchema.Organizer,
                Ews.AppointmentSchema.IsAllDayEvent,
                Ews.AppointmentSchema.AppointmentType,
                Ews.AppointmentSchema.MyResponseType,
                Ews.AppointmentSchema.IsMeeting,
                Ews.AppointmentSchema.IsCancelled);

            Ews.FindItemsResults<Ews.Appointment> found = svc.FindAppointments(folderId, view);

            var results = new List<SharedCalendarEventRow>();
            // A ConvertId timeout on one item means the connection is broken for
            // the rest of this call too - once that happens, stop calling ConvertId
            // entirely rather than letting every remaining item pay its own full
            // timeout (up to ~12 minutes for 50 items on a mid-loop connection
            // drop, with list_events just hanging the whole time). A non-timeout
            // failure on a single item is left alone (still tried for subsequent
            // items) since that's plausibly just a one-off glitch for that item.
            bool convertIdBroken = false;
            foreach (Ews.Appointment appt in found.Items)
            {
                string entryId = null;
                if (!convertIdBroken)
                {
                    try
                    {
                        entryId = ConvertToEntryId(svc, appt, mailboxSmtp);
                    }
                    catch (Exception ex) when (IsTimeout(ex))
                    {
                        DebugLog.WriteException("GetSharedCalendarEvents ConvertId timeout - aborting further ConvertId calls", ex);
                        convertIdBroken = true;
                        entryId = null;
                    }
                }
                Ews.AppointmentType apptType = TryGetAppointmentType(appt);
                results.Add(new SharedCalendarEventRow
                {
                    EntryId = entryId,
                    Subject = appt.Subject,
                    // Normalized to Unspecified (from EWS's Local) so the shape
                    // printed by "o" formatting (no timezone offset) matches the
                    // COM own-calendar path's DateTime.Kind exactly.
                    Start = DateTime.SpecifyKind(appt.Start, DateTimeKind.Unspecified),
                    End = DateTime.SpecifyKind(appt.End, DateTimeKind.Unspecified),
                    Location = appt.Location,
                    Organizer = TryGetOrganizerLabel(appt),
                    AllDay = TryGetBool(appt, Ews.AppointmentSchema.IsAllDayEvent, false),
                    Recurring = apptType == Ews.AppointmentType.RecurringMaster ||
                                apptType == Ews.AppointmentType.Occurrence ||
                                apptType == Ews.AppointmentType.Exception,
                    ResponseStatus = TryGetResponseType(appt).ToString(),
                    MeetingStatus = TryGetBool(appt, Ews.AppointmentSchema.IsCancelled, false) ? "Cancelled"
                                    : (TryGetBool(appt, Ews.AppointmentSchema.IsMeeting, false) ? "Meeting" : "NonMeeting"),
                });
            }
            return results;
        }

        // Some PropertySet-requested properties aren't reliably populated by every
        // Exchange server for every item type returned from CalendarView - confirmed
        // live 2026-10-01: Appointment.IsCancelled threw ServiceObjectPropertyException
        // ("This property was requested, but it wasn't returned by the server") for a
        // plain, non-meeting appointment, despite IsCancelled being in the PropertySet
        // above. TryGetProperty is EWS's documented safe-read for exactly this case -
        // these helpers wrap it so a missing property degrades to a sensible default
        // instead of crashing the whole list_events call. Applied to every
        // meeting-specific property read here (IsAllDayEvent/AppointmentType/
        // MyResponseType/IsMeeting/IsCancelled/Organizer) since they're all plausibly
        // absent for some item type, unlike Subject/Start/End/Location which are core
        // enough to every calendar item that direct access is kept as-is.
        private static bool TryGetBool(Ews.Appointment appt, Ews.PropertyDefinition prop, bool fallback)
        {
            object value;
            if (appt.TryGetProperty(prop, out value) && value is bool) return (bool)value;
            return fallback;
        }

        private static Ews.MeetingResponseType TryGetResponseType(Ews.Appointment appt)
        {
            object value;
            if (appt.TryGetProperty(Ews.AppointmentSchema.MyResponseType, out value) && value is Ews.MeetingResponseType)
                return (Ews.MeetingResponseType)value;
            return Ews.MeetingResponseType.Unknown;
        }

        private static Ews.AppointmentType TryGetAppointmentType(Ews.Appointment appt)
        {
            object value;
            if (appt.TryGetProperty(Ews.AppointmentSchema.AppointmentType, out value) && value is Ews.AppointmentType)
                return (Ews.AppointmentType)value;
            return Ews.AppointmentType.Single;
        }

        private static string TryGetOrganizerLabel(Ews.Appointment appt)
        {
            object value;
            if (appt.TryGetProperty(Ews.AppointmentSchema.Organizer, out value) && value is Ews.EmailAddress)
            {
                Ews.EmailAddress organizer = (Ews.EmailAddress)value;
                return organizer.Name ?? organizer.Address ?? "";
            }
            return "";
        }

        // Converts this item's EWS id to the classic Outlook/MAPI EntryID format so
        // it stays resolvable via the existing Ns.GetItemFromID(entryId, storeId)
        // every read/write tool already uses (get_event's own store_id parameter) -
        // list_events' shared-calendar output must not change shape just because
        // this path now fetches data via EWS instead of COM. Failure here (should
        // be rare - reasoned from the EWS Managed API surface, not yet verified
        // live) degrades to an empty event_id rather than dropping the whole event:
        // the caller still sees the event exists, just can't act on it via
        // get_event/edit_event until this is investigated.
        //
        // A timeout is deliberately NOT caught here - it propagates to
        // GetSharedCalendarEvents' loop, which needs to distinguish "this was a
        // timeout" (stop calling ConvertId for the rest of the batch) from any
        // other failure (fine to keep trying subsequent items).
        private static string ConvertToEntryId(Ews.ExchangeService svc, Ews.Appointment appt, string mailboxSmtp)
        {
            try
            {
                // HexEntryId, not EntryId: EntryId returns a base64-encoded
                // PR_ENTRYID, but Outlook.AppointmentItem.EntryID / Namespace.
                // GetItemFromID (which every ItemById/get_event/edit_event call in
                // this add-in goes through) expect the hex-encoded form - EWS's
                // own docs call HexEntryId "the format used by Microsoft Outlook".
                var converted = svc.ConvertId(new Ews.AlternateId(Ews.IdFormat.EwsId, appt.Id.UniqueId, mailboxSmtp), Ews.IdFormat.HexEntryId);
                return ((Ews.AlternateId)converted).UniqueId;
            }
            catch (Exception ex) when (!IsTimeout(ex))
            {
                DebugLog.WriteException("GetSharedCalendarEvents ConvertId", ex);
                return null;
            }
        }

        // Queries Contacts and the Directory (GAL) separately and merges both,
        // rather than the single-call ResolveNameSearchLocation.ContactsThenDirectory
        // this used to use. ContactsThenDirectory short-circuits: Exchange's ANR
        // match against a personal Contacts folder only checks DisplayName (not
        // GivenName/Surname/company/etc.), and if that phase returns anything at
        // all, the fuller GAL ANR (which does cover first/last name) never runs -
        // confirmed as the cause of search_contacts matching display name only.
        // ContactSearchFormat.Format (the caller) already dedupes by email/name
        // and applies the limit, so merging both lists here is safe.
        private static IReadOnlyList<ContactMatch> ResolveNames(Uri url, string query)
        {
            var results = new List<ContactMatch>();
            Ews.ExchangeService svc = NewService(url);
            AppendResolutions(svc, query, Ews.ResolveNameSearchLocation.ContactsOnly, results);
            AppendResolutions(svc, query, Ews.ResolveNameSearchLocation.DirectoryOnly, results);
            return results;
        }

        private static void AppendResolutions(Ews.ExchangeService svc, string query, Ews.ResolveNameSearchLocation location, List<ContactMatch> results)
        {
            Ews.NameResolutionCollection col;
            try
            {
                col = svc.ResolveName(query, location, true);
            }
            catch (Ews.ServiceResponseException)
            {
                // ErrorNameResolutionNoResults / NoMailbox surface here on some
                // servers rather than as an empty collection - treat as "no matches"
                // at this search location (the other location may still match).
                return;
            }

            foreach (Ews.NameResolution nr in col)
            {
                string email = SmtpFrom(nr);
                if (string.IsNullOrEmpty(email)) continue; // mirror mcp-outlook: drop entries with no address

                results.Add(new ContactMatch
                {
                    FullName = FullNameFrom(nr),
                    DisplayName = DisplayNameFrom(nr, email),
                    Email = email,
                });
            }
        }

        private static string SmtpFrom(Ews.NameResolution nr)
        {
            Ews.EmailAddress mb = nr.Mailbox;
            if (mb != null && !string.IsNullOrEmpty(mb.Address) && mb.Address.Contains("@") &&
                (string.IsNullOrEmpty(mb.RoutingType) || mb.RoutingType.Equals("SMTP", StringComparison.OrdinalIgnoreCase)))
            {
                return mb.Address;
            }

            // GAL hits often carry an X500/legacyExchangeDN Address (RoutingType
            // "EX", no "@"); fall back to the resolved contact's own addresses.
            Ews.Contact contact = nr.Contact;
            if (contact != null && contact.EmailAddresses != null)
            {
                foreach (Ews.EmailAddressKey key in new[]
                {
                    Ews.EmailAddressKey.EmailAddress1,
                    Ews.EmailAddressKey.EmailAddress2,
                    Ews.EmailAddressKey.EmailAddress3,
                })
                {
                    Ews.EmailAddress ea;
                    if (contact.EmailAddresses.TryGetValue(key, out ea) &&
                        ea != null && !string.IsNullOrEmpty(ea.Address) && ea.Address.Contains("@"))
                    {
                        return ea.Address;
                    }
                }
            }

            return null;
        }

        // The directory's own display name - may be org-formatted (e.g. a GAL
        // entry whose AD displayName attribute reads "Dept/Unit/Title") and look
        // nothing like the person's actual name. Kept separate from
        // FullNameFrom (below) specifically so a caller can show both when they
        // disagree - confirmed live 2026-10-01: an agent saw only this field,
        // it didn't resemble the searched name, and it wrongly concluded the
        // search had failed even though the match was correct.
        private static string DisplayNameFrom(Ews.NameResolution nr, string emailFallback)
        {
            Ews.Contact contact = nr.Contact;
            if (contact != null && !string.IsNullOrEmpty(contact.DisplayName)) return contact.DisplayName;
            if (nr.Mailbox != null && !string.IsNullOrEmpty(nr.Mailbox.Name)) return nr.Mailbox.Name;
            return emailFallback;
        }

        // The directory's GivenName/Surname (AD's givenName/sn attributes) - the
        // actual person's name, independent of however DisplayName happens to be
        // formatted. Empty for a shared/role mailbox with no such fields (e.g.
        // "IT Helpdesk"), in which case ContactSearchFormat falls back to
        // DisplayName alone.
        private static string FullNameFrom(Ews.NameResolution nr)
        {
            Ews.Contact contact = nr.Contact;
            if (contact == null) return "";
            string given = contact.GivenName ?? "";
            string surname = contact.Surname ?? "";
            return (given + " " + surname).Trim();
        }

        // Timeouts surface inconsistently across EWS Managed API paths - as a
        // bare TimeoutException, or a WebException(Timeout), or either wrapped in
        // a ServiceRequestException. Walk the whole inner chain.
        public static bool IsTimeout(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                if (e is TimeoutException) return true;
                var web = e as WebException;
                if (web != null && web.Status == WebExceptionStatus.Timeout) return true;
            }
            return false;
        }
    }
}
