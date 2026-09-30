using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeAi.Shared
{
    // One row of list_events' output, backend-agnostic - populated from
    // Outlook COM properties for the caller's own calendar, or from EWS
    // Appointment properties for a shared calendar (see OutlookEws.cs's
    // GetSharedCalendarEventsAsync). Keeping this a plain struct (no Outlook
    // COM or EWS type in its shape) is what makes Format below unit-testable
    // without a live Outlook/Exchange session - same role ContactSearchFormat
    // plays for search_contacts.
    public struct SharedCalendarEventRow
    {
        public string EntryId;
        public string Subject;
        public DateTime Start;
        public DateTime End;
        public string Location;
        public string Organizer;
        public bool AllDay;
        public bool Recurring;
        public string ResponseStatus;
        public string MeetingStatus;
    }

    // Pure formatting for list_events' shared-calendar (EWS) path only -
    // reproduces the exact per-event text shape OutlookTools.Calendar.cs's
    // QueryCalendarItems already writes inline for the own-calendar (COM)
    // path, so the tool's output doesn't reveal which backend answered it.
    // QueryCalendarItems itself is deliberately NOT refactored to call this -
    // see the design doc's Section 1 for why (its try/catch around
    // appt.Start/appt.End is load-bearing and has no equivalent here).
    public static class SharedCalendarEventFormat
    {
        public static string Format(IReadOnlyList<SharedCalendarEventRow> rows, int limit, string calendarOwner, string storeId)
        {
            var sb = new StringBuilder();
            int n = 0;
            foreach (SharedCalendarEventRow row in rows)
            {
                if (n >= limit) break;
                n++;
                sb.AppendLine("- event_id: " + (row.EntryId ?? ""));
                sb.AppendLine("  subject: " + (row.Subject ?? ""));
                sb.AppendLine("  start: " + Iso(row.Start) + "  end: " + Iso(row.End));
                sb.AppendLine("  location: " + (row.Location ?? ""));
                sb.AppendLine("  organizer: " + (row.Organizer ?? "") + "  all_day: " + row.AllDay + "  recurring: " + row.Recurring);
                sb.AppendLine("  response: " + (row.ResponseStatus ?? "") + "  meeting_status: " + (row.MeetingStatus ?? ""));
                if (calendarOwner != null)
                {
                    sb.AppendLine("  calendar_owner: " + calendarOwner);
                    sb.AppendLine("  store_id: " + storeId);
                }
            }
            return sb.ToString();
        }

        private static string Iso(DateTime d)
        {
            return d.ToString("o", CultureInfo.InvariantCulture);
        }
    }
}
