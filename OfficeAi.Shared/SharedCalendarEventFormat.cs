using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeAi.Shared
{
    // One row of list_events' output, backend-agnostic - populated from
    // either Outlook COM properties (own calendar) or EWS Appointment
    // properties (shared calendar). Plain struct (no COM/EWS type) so Format
    // below is unit-testable without a live session. See
    // SharedCalendarEventFormat.cs.md.
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

    // Pure formatting for list_events' shared-calendar (EWS) path - reproduces
    // the same per-event text shape QueryCalendarItems writes inline for the
    // own-calendar (COM) path, so the tool's output doesn't reveal which
    // backend answered it. Deliberately not unified by refactoring
    // QueryCalendarItems to call this - see SharedCalendarEventFormat.cs.md.
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
