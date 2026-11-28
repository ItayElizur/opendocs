using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeAi.Shared
{
    // The line Outlook adds to the model's prompt about the user's work week. Mailboxes usually have the right work
    // days but Outlook's untouched default hours, so the days are stated as fact and the hours as something to confirm.
    public static class WorkWeekPrompt
    {
        public static string Describe(IEnumerable<DayOfWeek> days, double startHour, double endHour)
        {
            string dayList = string.Join(", ", days.Distinct().OrderBy(d => d).Select(d => CultureInfo.InvariantCulture.DateTimeFormat.GetDayName(d)));
            return "Work days from the user's mailbox: " + dayList + ". Use them for \"this week\", \"the next work day\" " +
                   "and \"the weekend\". The mailbox's work hours are " + Hour(startHour) + "-" + Hour(endHour) + ", but " +
                   "work hours are often left at Outlook's default rather than set by the user: when hours matter (a " +
                   "meeting time, \"during work hours\", find_meeting_slots), ask the user to confirm their hours first " +
                   "unless they already gave them.";
        }

        public static string Unknown()
        {
            return "The user's work week could not be read from their mailbox. When work days or hours matter (\"this " +
                   "week\", a meeting time, find_meeting_slots), ask the user which days and hours they work instead of " +
                   "assuming Monday to Friday.";
        }

        // Outlook's WorkDay registry value (File > Options > Calendar > Work time): one bit per day, Sunday = 128 down to
        // Saturday = 2, so Monday to Friday is 124 and Sunday to Thursday is 248.
        public static List<DayOfWeek> DaysFromOutlookMask(int mask)
        {
            return Enum.GetValues(typeof(DayOfWeek)).Cast<DayOfWeek>().Where(d => (mask & (128 >> (int)d)) != 0).ToList();
        }

        // 8.5 -> "08:30"; a fraction that is not a whole minute is rounded to one.
        public static string Hour(double hour)
        {
            int minutes = (int)Math.Round(hour * 60);
            return (minutes / 60).ToString("00", CultureInfo.InvariantCulture) + ":" + (minutes % 60).ToString("00", CultureInfo.InvariantCulture);
        }
    }
}
