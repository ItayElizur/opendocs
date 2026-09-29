using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// A validated recurrence request, independent of the Outlook COM enums
    /// (OlRecurrenceType/OlDaysOfWeek) - this project doesn't reference the
    /// Outlook PIA, same split as OutlookDasl/MeetingSlots. OutlookAiAddIn
    /// converts this into RecurrencePattern fields.
    /// </summary>
    public sealed class RecurrenceSpec
    {
        public string Type;
        public int Interval;
        public string[] DaysOfWeek;
        public int? DayOfMonth;
        public int? Instance;
        public int? MonthOfYear;
        public int? Count;
        public DateTime? Until;
    }
}
