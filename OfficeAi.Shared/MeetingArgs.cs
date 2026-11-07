using System;

namespace OfficeAi.Shared
{
    // Small decisions of the Outlook meeting tools, kept free of COM so they can be unit tested.
    public static class MeetingArgs
    {
        // RecurrencePattern.GetOccurrence needs the occurrence's start date AND time; a date alone (midnight) never
        // matches, so it is combined with the series' time of day.
        public static DateTime OccurrenceStart(DateTime date, TimeSpan seriesTimeOfDay)
        {
            return date.TimeOfDay == TimeSpan.Zero ? date.Date + seriesTimeOfDay : date;
        }

        // A message_id / event_id with no action (or action "new") would silently create a new email or event and
        // ignore the target. An empty string counts as absent.
        public static bool TargetNeedsAction(string action, params string[] ids)
        {
            if (!string.IsNullOrEmpty(action) && action != "new") return false;
            foreach (string id in ids)
                if (!string.IsNullOrEmpty(id)) return true;
            return false;
        }
    }
}
