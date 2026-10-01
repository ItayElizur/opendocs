using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure validation for create_event/draft_event's optional "recurrence"
    /// object - every failure names the specific field and the fix, since
    /// this is the only feedback the calling model gets to correct its next
    /// call. Mirrors set_event_availability's TryParseBusyStatus pattern
    /// (OutlookAiAddIn/OutlookTools.Categories.cs).
    /// </summary>
    public static class RecurrenceValidator
    {
        public static readonly string[] ValidTypes =
        {
            "daily", "weekly", "monthly", "monthlyNth", "yearly", "yearlyNth",
        };

        public static readonly string[] ValidDays =
        {
            "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday",
        };

        // Which occurrence of its weekday `dayOfMonth` represents within a
        // month of `daysInMonth` days - 1-4 for 1st-4th, 5 if it's the last
        // such weekday in the month (matches RecurrencePattern.Instance's
        // own 1-4/5-for-"last" convention). Pure integer math so callers
        // (OutlookAiAddIn) can pass DateTime.Day/DateTime.DaysInMonth
        // without this project needing a DateTime dependency beyond what
        // it already has.
        public static int NthWeekdayOfMonth(int dayOfMonth, int daysInMonth)
        {
            int n = (dayOfMonth - 1) / 7 + 1;
            return dayOfMonth + 7 > daysInMonth ? 5 : n;
        }

        public static RecurrenceSpec Parse(string type, int interval, string[] daysOfWeek, int? dayOfMonth,
            int? instance, int? monthOfYear, int? count, DateTime? until, out string error)
        {
            error = null;

            if (string.IsNullOrEmpty(type) || Array.IndexOf(ValidTypes, type) < 0)
            {
                error = "recurrence.type \"" + type + "\" is not valid. Valid values: " + string.Join(", ", ValidTypes) + ".";
                return null;
            }
            if (interval < 1)
            {
                error = "recurrence.interval must be 1 or greater (got " + interval + ").";
                return null;
            }
            if (count.HasValue && until.HasValue)
            {
                error = "recurrence.count and recurrence.until are mutually exclusive - specify at most one (omit both for no end date).";
                return null;
            }
            if (count.HasValue && count.Value < 1)
            {
                error = "recurrence.count must be 1 or greater (got " + count.Value + ").";
                return null;
            }

            string[] normalizedDays = null;
            if (daysOfWeek != null && daysOfWeek.Length > 0)
            {
                normalizedDays = new string[daysOfWeek.Length];
                for (int i = 0; i < daysOfWeek.Length; i++)
                {
                    string d = (daysOfWeek[i] ?? "").Trim().ToLowerInvariant();
                    if (Array.IndexOf(ValidDays, d) < 0)
                    {
                        error = "recurrence.days_of_week has an unrecognized day \"" + daysOfWeek[i] + "\". Valid values: " + string.Join(", ", ValidDays) + ".";
                        return null;
                    }
                    normalizedDays[i] = d;
                }
            }

            switch (type)
            {
                case "weekly":
                    if (normalizedDays == null)
                    {
                        error = "recurrence.days_of_week is required for type \"weekly\" - e.g. [\"monday\", \"wednesday\"].";
                        return null;
                    }
                    break;

                case "monthly":
                    if (!dayOfMonth.HasValue || dayOfMonth.Value < 1 || dayOfMonth.Value > 31)
                    {
                        error = "recurrence.day_of_month is required for type \"monthly\" and must be 1-31 (got " +
                                (dayOfMonth.HasValue ? dayOfMonth.Value.ToString() : "none") + ").";
                        return null;
                    }
                    break;

                case "monthlyNth":
                case "yearlyNth":
                    if (!instance.HasValue || instance.Value < 1 || instance.Value > 5)
                    {
                        error = "recurrence.instance is required for type \"" + type + "\" and must be 1-4 (1st-4th) or 5 (last) (got " +
                                (instance.HasValue ? instance.Value.ToString() : "none") + ").";
                        return null;
                    }
                    if (normalizedDays == null || normalizedDays.Length != 1)
                    {
                        error = "recurrence.days_of_week must have exactly one day for type \"" + type + "\" - e.g. [\"tuesday\"] for \"2nd Tuesday\".";
                        return null;
                    }
                    if (type == "yearlyNth" && (!monthOfYear.HasValue || monthOfYear.Value < 1 || monthOfYear.Value > 12))
                    {
                        error = "recurrence.month_of_year is required for type \"yearlyNth\" and must be 1-12 (got " +
                                (monthOfYear.HasValue ? monthOfYear.Value.ToString() : "none") + ").";
                        return null;
                    }
                    break;

                case "yearly":
                    if (!monthOfYear.HasValue || monthOfYear.Value < 1 || monthOfYear.Value > 12)
                    {
                        error = "recurrence.month_of_year is required for type \"yearly\" and must be 1-12 (got " +
                                (monthOfYear.HasValue ? monthOfYear.Value.ToString() : "none") + ").";
                        return null;
                    }
                    if (!dayOfMonth.HasValue || dayOfMonth.Value < 1 || dayOfMonth.Value > 31)
                    {
                        error = "recurrence.day_of_month is required for type \"yearly\" and must be 1-31 (got " +
                                (dayOfMonth.HasValue ? dayOfMonth.Value.ToString() : "none") + ").";
                        return null;
                    }
                    break;

                    // "daily" needs none of the above.
            }

            return new RecurrenceSpec
            {
                Type = type,
                Interval = interval,
                DaysOfWeek = normalizedDays,
                DayOfMonth = dayOfMonth,
                Instance = instance,
                MonthOfYear = monthOfYear,
                Count = count,
                Until = until,
            };
        }
    }
}
