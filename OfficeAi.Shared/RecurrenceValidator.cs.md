## RecurrenceValidator.NthWeekdayOfMonth

```
// Which occurrence of its weekday `dayOfMonth` represents within a
// month of `daysInMonth` days - 1-4 for 1st-4th, 5 if it's the last
// such weekday in the month (matches RecurrencePattern.Instance's
// own 1-4/5-for-"last" convention). Pure integer math so callers
// (OutlookAiAddIn) can pass DateTime.Day/DateTime.DaysInMonth
// without this project needing a DateTime dependency beyond what
// it already has.
```
