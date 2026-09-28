# Outlook Recurring-Series Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add recurring-series support to Outlook's calendar tools — create a series (all 6 `OlRecurrenceType` patterns), edit/cancel a single occurrence vs. the whole series, including recurring meetings with attendees.

**Architecture:** Extend the four existing calendar tools (`create_event`/`draft_event` gain an optional `recurrence` object; `reschedule_event`/`draft_reschedule_event`/`cancel_event`/`draft_cancel_event` gain an optional `occurrence_date` string) rather than adding new tools. A pure, unit-tested validator (`OfficeAi.Shared.RecurrenceValidator`) parses and validates the `recurrence` object before any COM call; a new `ResolveOccurrenceTarget` helper in `OutlookTools.Calendar.cs` resolves `occurrence_date` via `RecurrencePattern.GetOccurrence()` into the item every existing check/mutation then operates on unchanged.

**Tech Stack:** C# / .NET Framework 4.8, `Microsoft.Office.Interop.Outlook` 15.0.0.0 (COM interop), xUnit (`OfficeAi.Shared.Tests`), TypeScript (`entry.ts` tool schemas, no build tooling beyond esbuild).

**Spec:** `docs/superpowers/specs/2026-09-28-outlook-recurring-series-design.md`

## Global Constraints

- All 6 `OlRecurrenceType` values must be supported for creation: `daily`, `weekly`, `monthly`, `monthlyNth`, `yearly`, `yearlyNth` (spec Scope decision 2).
- Recurring meetings (with attendees) are in scope, not just plain appointments (spec Scope decision 1).
- Extend the 4 existing tools; do not add new tool names (spec Scope decision 3).
- Every recurrence validation failure must name the exact field and the fix — no generic "invalid recurrence" messages (spec Scope decision 6, confirmed with user).
- Occurrence-level cancellation is **always** a barrier in the undo stack (plain appointment or meeting, no exception) — there is no COM API to reverse a deleted occurrence (spec Ground truth: `Exception`/`Exceptions` are read-only). Occurrence-level reschedule stays undo-able.
- Tier placement is unchanged: `create_event`/`reschedule_event`/`cancel_event` stay Full-autonomy-only, `draft_*` stay Draft-tier, regardless of `recurrence`/`occurrence_date` being present.
- Human-in-the-loop chat confirmation is explicitly out of scope for this plan (separate future spec).
- Implementation order follows the spec's risk ranking (Section 6): occurrence-level operations built and manually verified before series creation.

---

## Task 1: `RecurrenceSpec` + `RecurrenceValidator` (pure, TDD)

**Files:**
- Create: `OfficeAi.Shared/RecurrenceSpec.cs`
- Create: `OfficeAi.Shared/RecurrenceValidator.cs`
- Test: `OfficeAi.Shared.Tests/RecurrenceValidatorTests.cs`

**Interfaces:**
- Produces: `RecurrenceSpec` (public fields: `Type` (string), `Interval` (int), `DaysOfWeek` (string[], lowercase day names or null), `DayOfMonth`/`Instance`/`MonthOfYear`/`Count` (int?), `Until` (DateTime?)). `RecurrenceValidator.Parse(string type, int interval, string[] daysOfWeek, int? dayOfMonth, int? instance, int? monthOfYear, int? count, DateTime? until, out string error) → RecurrenceSpec` (returns null and sets `error` on any validation failure; returns a populated spec and `error = null` on success). `RecurrenceValidator.ValidTypes` / `RecurrenceValidator.ValidDays` (public `string[]`, reused by Task 7's error messages and Task 1's own tests).

- [ ] **Step 1: Write the failing tests**

Create `OfficeAi.Shared.Tests/RecurrenceValidatorTests.cs`:

```csharp
using System;
using Xunit;
using OfficeAi.Shared;

public class RecurrenceValidatorTests
{
    [Fact]
    public void Daily_Valid_NoOtherFieldsRequired()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.NotNull(spec);
        Assert.Equal("daily", spec.Type);
        Assert.Equal(1, spec.Interval);
    }

    [Fact]
    public void InvalidType_ListsValidValues()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("hourly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("hourly", error);
        Assert.Contains("daily", error);
        Assert.Contains("weekly", error);
        Assert.Contains("monthlyNth", error);
        Assert.Contains("yearlyNth", error);
    }

    [Fact]
    public void IntervalLessThanOne_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 0, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("interval", error);
    }

    [Fact]
    public void CountAndUntilBothSet_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 5, DateTime.Today, out error);
        Assert.Null(spec);
        Assert.Contains("mutually exclusive", error);
    }

    [Fact]
    public void CountLessThanOne_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 0, null, out error);
        Assert.Null(spec);
        Assert.Contains("count", error);
    }

    [Fact]
    public void Weekly_MissingDaysOfWeek_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("days_of_week", error);
        Assert.Contains("weekly", error);
    }

    [Fact]
    public void Weekly_Valid_NormalizesDayCase()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 2, new[] { "Monday", "WEDNESDAY" }, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(new[] { "monday", "wednesday" }, spec.DaysOfWeek);
        Assert.Equal(2, spec.Interval);
    }

    [Fact]
    public void UnknownDayName_ListsValidDays()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 1, new[] { "Funday" }, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("Funday", error);
        Assert.Contains("sunday", error);
        Assert.Contains("saturday", error);
    }

    [Fact]
    public void Monthly_MissingDayOfMonth_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("day_of_month", error);
    }

    [Fact]
    public void Monthly_DayOfMonthOutOfRange_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, 32, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("1-31", error);
    }

    [Fact]
    public void Monthly_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, 15, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(15, spec.DayOfMonth);
    }

    [Fact]
    public void MonthlyNth_MissingInstanceAndDay_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("instance", error);
    }

    [Fact]
    public void MonthlyNth_InstanceOutOfRange_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "tuesday" }, null, 6, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("1-4", error);
    }

    [Fact]
    public void MonthlyNth_MultipleDays_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "monday", "tuesday" }, null, 2, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("exactly one day", error);
    }

    [Fact]
    public void MonthlyNth_Valid_SecondTuesday()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "tuesday" }, null, 2, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(2, spec.Instance);
        Assert.Equal(new[] { "tuesday" }, spec.DaysOfWeek);
    }

    [Fact]
    public void MonthlyNth_LastInstance_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "friday" }, null, 5, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(5, spec.Instance);
    }

    [Fact]
    public void Yearly_MissingMonthAndDay_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("month_of_year", error);
    }

    [Fact]
    public void Yearly_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearly", 1, null, 25, null, 12, null, null, out error);
        Assert.Null(error);
        Assert.Equal(25, spec.DayOfMonth);
        Assert.Equal(12, spec.MonthOfYear);
    }

    [Fact]
    public void YearlyNth_MissingMonthOfYear_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearlyNth", 1, new[] { "monday" }, null, 1, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("month_of_year", error);
    }

    [Fact]
    public void YearlyNth_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearlyNth", 1, new[] { "monday" }, null, 1, 9, null, null, out error);
        Assert.Null(error);
        Assert.Equal(1, spec.Instance);
        Assert.Equal(9, spec.MonthOfYear);
    }

    [Fact]
    public void Count_SetsEndCondition()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 10, null, out error);
        Assert.Null(error);
        Assert.Equal(10, spec.Count);
        Assert.Null(spec.Until);
    }

    [Fact]
    public void Until_SetsEndCondition()
    {
        DateTime until = new DateTime(2027, 1, 1);
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, until, out error);
        Assert.Null(error);
        Assert.Equal(until, spec.Until);
        Assert.Null(spec.Count);
    }

    [Fact]
    public void NeitherCountNorUntil_NoEndDate()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Null(spec.Count);
        Assert.Null(spec.Until);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --filter RecurrenceValidatorTests`
Expected: compile error (`RecurrenceSpec`/`RecurrenceValidator` don't exist yet).

- [ ] **Step 3: Implement `RecurrenceSpec`**

Create `OfficeAi.Shared/RecurrenceSpec.cs`:

```csharp
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
```

- [ ] **Step 4: Implement `RecurrenceValidator`**

Create `OfficeAi.Shared/RecurrenceValidator.cs`:

```csharp
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
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --filter RecurrenceValidatorTests`
Expected: all 21 tests PASS.

- [ ] **Step 6: Run the full shared test suite to confirm no regressions**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo`
Expected: 208/208 passing (187 existing + 21 new).

- [ ] **Step 7: Commit**

```bash
git add OfficeAi.Shared/RecurrenceSpec.cs OfficeAi.Shared/RecurrenceValidator.cs OfficeAi.Shared.Tests/RecurrenceValidatorTests.cs
git commit -m "feat(outlook): add RecurrenceValidator for recurring-series creation"
```

---

## Task 2: JSON argument helpers (`StrArray`, `OptInt`)

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.cs` (add helpers near the existing `Str`/`Int`/`Bool`/`DateArg` block, ~line 340)

**Interfaces:**
- Consumes: nothing new.
- Produces: `internal static string[] StrArray(JsonElement o, string name)` — returns `null` if the field is absent or not a JSON array; `internal static int? OptInt(JsonElement o, string name)` — returns `null` if absent or not a number. Both consumed by Task 7's `ReadRecurrence`.

- [ ] **Step 1: Add the helpers**

In `OutlookAiAddIn/OutlookTools.cs`, immediately after the existing `Int` method (around line 324, before `Double`):

```csharp
        internal static int? OptInt(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out v) && v.ValueKind == JsonValueKind.Number)
            {
                int n;
                if (v.TryGetInt32(out n)) return n;
            }
            return null;
        }

        internal static string[] StrArray(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind != JsonValueKind.Object || !o.TryGetProperty(name, out v) || v.ValueKind != JsonValueKind.Array)
                return null;
            var list = new List<string>();
            foreach (JsonElement item in v.EnumerateArray())
                if (item.ValueKind == JsonValueKind.String) list.Add(item.GetString());
            return list.ToArray();
        }
```

- [ ] **Step 2: Build to confirm it compiles**

Run: `"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Debug -nologo -v:minimal`
Expected: clean build, 0 errors/warnings. (`System.Collections.Generic` is already imported at the top of `OutlookTools.cs`.)

- [ ] **Step 3: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.cs
git commit -m "feat(outlook): add StrArray/OptInt JSON argument helpers"
```

---

## Task 3: Occurrence targeting + wire into `reschedule_event`/`draft_reschedule_event`

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (add `ResolveOccurrenceTarget` after `CanceledMeetingError`, ~line 298; rewrite `DraftRescheduleEvent`/`RescheduleEvent`, ~lines 305-393)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (add `occurrence_date` to both tools' schemas, ~lines 385-397 and 446-458; update both tools' descriptions)

**Interfaces:**
- Consumes: `RecurrenceSpec`/`RecurrenceValidator` not needed here (this task is occurrence *targeting*, not creation). Uses existing `ItemById`, `Str`, `IsCanceledMeeting`, `IsReceivedMeeting`, `RecordSnapshot`, `RecordIrreversible`, `ReadProps` unchanged.
- Produces: `private static ToolResult ResolveOccurrenceTarget(Outlook.AppointmentItem master, string occurrenceDate, string toolName, out Outlook.AppointmentItem target)` — returns `null` and sets `target` on success (to `master` if `occurrenceDate` is null, to the resolved occurrence otherwise); returns a populated `ToolResult` (caller must `return` it immediately) on any failure. Consumed by Task 4.

- [ ] **Step 1: Add `ResolveOccurrenceTarget`**

In `OutlookAiAddIn/OutlookTools.Calendar.cs`, immediately after `CanceledMeetingError` (after line 298, before the `DraftRescheduleEvent` comment):

```csharp
        // Shared by all four occurrence-aware tools (reschedule/cancel, draft
        // and immediate). occurrence_date lets a caller target one instance
        // of a recurring series instead of the whole master.
        // RecurrencePattern.GetOccurrence(DateTime) confirmed present via
        // .NET reflection against the referenced PIA (not assumed) - this
        // had been undocumented capability until this addition, even though
        // every occurrence list_events returns shares the master's EntryID.
        private static ToolResult ResolveOccurrenceTarget(Outlook.AppointmentItem master, string occurrenceDate, string toolName, out Outlook.AppointmentItem target)
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
```

- [ ] **Step 2: Rewrite `DraftRescheduleEvent`**

Replace the existing `DraftRescheduleEvent` method (lines 305-332) with:

```csharp
        private static ToolResult DraftRescheduleEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_reschedule_event" };

            Outlook.AppointmentItem appt;
            ToolResult occurrenceError = ResolveOccurrenceTarget(master, occDate, "draft_reschedule_event", out appt);
            if (occurrenceError != null) return occurrenceError;

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "draft_reschedule_event" };
            if (IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "draft_reschedule_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "draft_reschedule_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "draft_reschedule_event" };

            bool isMeeting = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            string scopeNote = occDate != null ? " (just this occurrence, not the whole series)" : "";
            appt.Start = start.Value;
            appt.End = end.Value;
            appt.Display(false);
            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\"" + scopeNote + " with the new time (" + Iso(start.Value) + " to " + Iso(end.Value) +
                         ") in Outlook for the user to review and " + (isMeeting ? "save/send the update." : "save."),
                Summary = "draft_reschedule_event",
            };
        }
```

- [ ] **Step 3: Rewrite `RescheduleEvent`**

Replace the existing `RescheduleEvent` method (lines 343-393) with:

```csharp
        private static ToolResult RescheduleEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "reschedule_event" };

            Outlook.AppointmentItem appt;
            ToolResult occurrenceError = ResolveOccurrenceTarget(master, occDate, "reschedule_event", out appt);
            if (occurrenceError != null) return occurrenceError;

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "reschedule_event" };
            if (IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "reschedule_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "reschedule_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "reschedule_event" };

            string scopeNote = occDate != null ? " (this occurrence only)" : "";
            string oldStart = Iso(appt.Start);
            string oldEnd = Iso(appt.End);
            object[] before = ReadProps(appt, RescheduleProps);
            appt.Start = start.Value;
            appt.End = end.Value;

            if (appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting)
            {
                appt.Send();
                // Barrier, not a snapshot: like create_event's invite branch, this
                // is an irreversible, unreviewed send - undo must stop here rather
                // than silently move the meeting back without telling attendees.
                RecordIrreversible(mbxKey, "reschedule_event invite update for \"" + (appt.Subject ?? "") + "\"" + scopeNote);
                return new ToolResult
                {
                    Output = "Rescheduled and sent update notice" + scopeNote + ": \"" + (appt.Subject ?? "") + "\" from " + oldStart + " - " + oldEnd +
                             " to " + Iso(start.Value) + " - " + Iso(end.Value) + ".",
                    Mutated = true,
                    Summary = "reschedule_event",
                };
            }

            appt.Save();
            // RecordSnapshot reads appt.EntryID via ItemEntryIdOf(appt) - for an
            // occurrence, that's the real, resolvable EntryID GetOccurrence's
            // returned item gets once saved (an occurrence becomes a distinct,
            // independently-addressable "exception" item, not a phantom only
            // reachable through the pattern), so undo/redo works via the exact
            // same SnapshotEntry mechanism as every other reschedule - no new
            // undo-entry type needed. Verify this holds live in Task 6.
            RecordSnapshot(mbxKey, "reschedule_event", appt, appt.Subject ?? "", RescheduleProps, before);
            return new ToolResult
            {
                Output = "Rescheduled" + scopeNote + ": \"" + (appt.Subject ?? "") + "\" from " + oldStart + " - " + oldEnd +
                         " to " + Iso(start.Value) + " - " + Iso(end.Value) + ".",
                Mutated = true,
                Summary = "reschedule_event",
            };
        }
```

- [ ] **Step 4: Update `entry.ts` schemas**

In `OutlookAiAddIn/web-src/entry.ts`, find the `draft_reschedule_event` tool definition and add `occurrence_date` to its `properties`, and update its description:

```typescript
  {
    name: 'draft_reschedule_event',
    description:
      'Opens an existing calendar event with a new start/end already filled in, unsaved, for the user to review and save/send. Never touches attendees. Omit occurrence_date to act on the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value for this series) to move just that one occurrence instead. Recurring events share one EntryID for the whole series - occurrence_date is the only way to target a single instance.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        start: { type: 'string', description: 'New start date-time, e.g. "2026-09-01T14:00".' },
        end: { type: 'string', description: 'New end date-time.' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to move (from list_events\' start value). Omit to move the whole series.' },
      },
      required: ['event_id', 'start', 'end'],
    },
  },
```

Find the `reschedule_event` tool definition and make the equivalent change:

```typescript
  {
    name: 'reschedule_event',
    description:
      'Moves an existing calendar event to a new start/end immediately - NO review window. If the user organizes it (has attendees), sends the reschedule notice to them right away. Only available in Full autonomy. Prefer draft_reschedule_event unless the user clearly wants this moved right now, with no chance to review it first. Only works on events the user organizes or a plain appointment - if it\'s a meeting the user only attends (not the organizer), this returns an error instead of attempting an unauthoritative change; use Outlook\'s own "Propose New Time" for those. Omit occurrence_date to act on the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value for this series) to move just that one occurrence instead.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        start: { type: 'string', description: 'New start date-time, e.g. "2026-09-01T14:00".' },
        end: { type: 'string', description: 'New end date-time.' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to move (from list_events\' start value). Omit to move the whole series.' },
      },
      required: ['event_id', 'start', 'end'],
    },
  },
```

- [ ] **Step 5: Build to confirm it compiles**

Run:
```bash
cd OutlookAiAddIn
/c/dev/opendocs/OutlookAiAddIn/node_modules/.bin/esbuild.cmd web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap
cd ..
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Release -nologo -v:minimal
```
Expected: both clean, 0 errors/warnings.

- [ ] **Step 6: Run the shared test suite (regression check — this task has no new pure logic)**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo`
Expected: 208/208 passing, unchanged from Task 1.

- [ ] **Step 7: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add occurrence_date targeting to reschedule_event/draft_reschedule_event"
```

---

## Task 4: Wire `occurrence_date` into `cancel_event`/`draft_cancel_event`

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (rewrite `DraftCancelEvent`/`CancelEvent`, current lines ~430-528)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (add `occurrence_date` to both tools' schemas; update `undo_last_action`'s description)

**Interfaces:**
- Consumes: `ResolveOccurrenceTarget` from Task 3.
- Produces: nothing new consumed by later tasks.

- [ ] **Step 1: Rewrite `DraftCancelEvent`**

Replace the existing `DraftCancelEvent` method with:

```csharp
        private static ToolResult DraftCancelEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_cancel_event" };

            Outlook.AppointmentItem appt;
            ToolResult occurrenceError = ResolveOccurrenceTarget(master, occDate, "draft_cancel_event", out appt);
            if (occurrenceError != null) return occurrenceError;

            if (IsCanceledMeeting(appt))
                return new ToolResult { Output = AlreadyCanceledError(appt), IsError = true, Summary = "draft_cancel_event" };
            if (IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingCancelError(appt), IsError = true, Summary = "draft_cancel_event" };

            bool isMeeting = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
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
```

- [ ] **Step 2: Rewrite `CancelEvent`**

Replace the existing `CancelEvent` method with:

```csharp
        private static ToolResult CancelEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "cancel_event" };

            Outlook.AppointmentItem appt;
            ToolResult occurrenceError = ResolveOccurrenceTarget(master, occDate, "cancel_event", out appt);
            if (occurrenceError != null) return occurrenceError;

            bool isOccurrence = occDate != null;
            // An occurrence has no "already canceled" state to clean up - a
            // deleted occurrence simply won't resolve via GetOccurrence()
            // again, surfacing as ResolveOccurrenceTarget's "no occurrence on
            // that date" error above instead of reaching this point.
            bool alreadyCanceled = !isOccurrence && IsCanceledMeeting(appt);
            if (!alreadyCanceled && IsReceivedMeeting(appt))
                return new ToolResult { Output = ReceivedMeetingCancelError(appt), IsError = true, Summary = "cancel_event" };

            string subject = appt.Subject ?? "";
            bool isMeeting = !alreadyCanceled && appt.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
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
                    catch (Exception ex) { DebugLog.WriteException("CancelEvent occurrence delete", ex); removedLocally = false; }
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
                appt.Delete();
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
```

- [ ] **Step 3: Update `entry.ts` schemas**

Find `draft_cancel_event`'s tool definition and update:

```typescript
  {
    name: 'draft_cancel_event',
    description:
      'Opens an existing calendar event for the user to review before canceling it themselves - for a meeting the user organizes, via Outlook\'s own Cancel Meeting/Send Cancellation buttons; for a plain appointment, via Delete. Never sends, deletes, or changes anything itself. Only works on events the user organizes or a plain appointment - on a meeting the user only attends, use decline_meeting instead. Omit occurrence_date to act on the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value) to preview canceling just that one occurrence instead.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to preview canceling (from list_events\' start value). Omit to act on the whole series.' },
      },
      required: ['event_id'],
    },
  },
```

Find `cancel_event`'s tool definition and update:

```typescript
  {
    name: 'cancel_event',
    description:
      'Cancels an existing calendar event immediately - NO review window. If the user organizes it (has attendees), sends the cancellation notice to them right away, then removes it from the calendar; a plain appointment is just removed. An event that\'s already canceled is also just removed - nothing new to notify, this is the only way to dismiss one. Only available in Full autonomy. Prefer draft_cancel_event unless the user clearly wants this canceled right now, with no chance to review it first. Only refuses on a still-active meeting the user only attends (not the organizer) - use decline_meeting for those. Omit occurrence_date to cancel the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value) to cancel just that one occurrence instead - occurrence cancellation can NEVER be undone (unlike whole-series cancellation of a plain appointment), since Outlook has no API to restore a deleted occurrence.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to cancel (from list_events\' start value). Omit to cancel the whole series. Cannot be undone.' },
      },
      required: ['event_id'],
    },
  },
```

Find `undo_last_action`'s tool description (the one listing covered tools). It currently has a line reading (approximately):

```typescript
      'set_category_color, set_event_availability, reschedule_event without attendees, cancel_event on a plain appointment (moved back out of Deleted Items), and create_event without attendees (moved to Deleted Items). Sends, meeting invites, accept/decline_meeting, cancel_event on an organized meeting, and permanent deletes ' +
```

Replace that entire line (this is the whole-line replacement, not an insertion) with:

```typescript
      'set_category_color, set_event_availability, reschedule_event without attendees, cancel_event on a plain or already-canceled whole event (moved back out of Deleted Items), and create_event without attendees (moved to Deleted Items). Sends, meeting invites, accept/decline_meeting, cancel_event on a still-active organized meeting, cancel_event on any single occurrence (always, whether plain or meeting - Outlook has no API to restore a deleted occurrence), and permanent deletes ' +
```

- [ ] **Step 4: Build to confirm it compiles**

Run the same esbuild + MSBuild commands as Task 3, Step 5.
Expected: both clean.

- [ ] **Step 5: Run the shared test suite (regression check)**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo`
Expected: 208/208 passing.

- [ ] **Step 6: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add occurrence_date targeting to cancel_event/draft_cancel_event"
```

---

## Task 5: Package, install, and manually verify occurrence-level operations (highest risk — spec Section 6 item 1)

This task has no automatable steps — it produces a build for the user to test live, per this session's established pattern (the agent has no live Outlook access). Do not proceed to Task 6 until this task's manual verification passes, per the spec's required risk order.

**Files:** none (build/package only).

- [ ] **Step 1: Package and install**

```bash
cd deploy
powershell -ExecutionPolicy Bypass -File ./package.ps1 -App Outlook
cd dist
powershell -ExecutionPolicy Bypass -File ./install.ps1 -App Outlook
```
Restart Outlook afterward.

- [ ] **Step 2: Hand the user this manual test checklist**

Create a recurring meeting first (a plain `create_event` call still works — recurrence creation isn't wired until Task 7/8, so use Outlook's own UI to make a recurring meeting with test-only attendees for this task):

1. In Outlook's UI, create a recurring meeting (e.g., weekly, 3 occurrences) with only your own test accounts as attendees.
2. `FORCE_TOOL:list_events` — copy the `event_id` and note two of the occurrence `start` dates.
3. **Occurrence reschedule (highest risk):**
   ```
   FORCE_TOOL:reschedule_event{"event_id": "<id>", "occurrence_date": "<first occurrence date>", "start": "<new date>T10:00", "end": "<new date>T10:30"}
   ```
   Expect: `"Rescheduled and sent update notice (this occurrence only): ..."`. **Critically check**: did attendees' calendars show only that one occurrence moved, or did the whole series shift? Did the other occurrence(s) stay untouched?
4. **Occurrence reschedule undo:**
   ```
   FORCE_TOOL:undo_last_action
   ```
   If the occurrence was a plain (no-attendee) reschedule instead, confirm undo succeeds and re-resolves the occurrence correctly (see Task 3 Step 3's code comment on this). If it was the meeting-with-attendees case above, confirm undo refuses (barrier).
5. **Occurrence cancellation (highest risk):**
   ```
   FORCE_TOOL:cancel_event{"event_id": "<id>", "occurrence_date": "<second occurrence date>"}
   ```
   Expect: `"Canceled and notified attendees (this occurrence only): ..."`. **Critically check**: did attendees get a cancellation for just that date, and does the series still show its other occurrence(s) as active? Confirm `undo_last_action` refuses (this must be a barrier — expected, not a bug).
6. **Non-recurring regression check:** repeat `reschedule_event`/`cancel_event` on a **non-recurring** event with no `occurrence_date` — confirm identical behavior to before this change (whole-event, no `scopeNote` text).

- [ ] **Step 3: Record findings**

If either step 3 or step 5 shows a whole-series side effect instead of a single-occurrence one, **stop and fix `ResolveOccurrenceTarget`/the calling code before proceeding to Task 6** — this is the scenario the spec's risk ranking exists to catch before it reaches series creation. If both pass, update `docs/superpowers/specs/2026-09-28-outlook-recurring-series-design.md`'s Section 6 risk list to mark item 1 as verified, with the date and what was confirmed.

---

## Task 6: `ApplyRecurrence` + wire `recurrence` into `create_event`/`draft_event`

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Compose.cs` (add `RecurrenceTypeMap`, `DayFlagMap`, `ApplyRecurrence`, `ReadRecurrence` near the top of the file after `PrependHtml`, ~line 26; rewrite `DraftEvent`/`CreateEvent`)

**Interfaces:**
- Consumes: `RecurrenceSpec`/`RecurrenceValidator` (Task 1), `StrArray`/`OptInt` (Task 2).
- Produces: `private static void ApplyRecurrence(Outlook.AppointmentItem a, RecurrenceSpec spec)`, `private static RecurrenceSpec ReadRecurrence(JsonElement input, out string error)` — not consumed by other tasks, but keep the names for Task 8's doc references.

- [ ] **Step 1: Add the recurrence-application helpers**

In `OutlookAiAddIn/OutlookTools.Compose.cs`, immediately after `PrependHtml` (after line 26, before `DraftEmail`):

```csharp
        private static readonly Dictionary<string, Outlook.OlRecurrenceType> RecurrenceTypeMap = new Dictionary<string, Outlook.OlRecurrenceType>
        {
            { "daily", Outlook.OlRecurrenceType.olRecursDaily },
            { "weekly", Outlook.OlRecurrenceType.olRecursWeekly },
            { "monthly", Outlook.OlRecurrenceType.olRecursMonthly },
            { "monthlyNth", Outlook.OlRecurrenceType.olRecursMonthNth },
            { "yearly", Outlook.OlRecurrenceType.olRecursYearly },
            { "yearlyNth", Outlook.OlRecurrenceType.olRecursYearNth },
        };

        private static readonly Dictionary<string, Outlook.OlDaysOfWeek> DayFlagMap = new Dictionary<string, Outlook.OlDaysOfWeek>
        {
            { "sunday", Outlook.OlDaysOfWeek.olSunday },
            { "monday", Outlook.OlDaysOfWeek.olMonday },
            { "tuesday", Outlook.OlDaysOfWeek.olTuesday },
            { "wednesday", Outlook.OlDaysOfWeek.olWednesday },
            { "thursday", Outlook.OlDaysOfWeek.olThursday },
            { "friday", Outlook.OlDaysOfWeek.olFriday },
            { "saturday", Outlook.OlDaysOfWeek.olSaturday },
        };

        // Applies a validated RecurrenceSpec (OfficeAi.Shared - pure, unit
        // tested in RecurrenceValidatorTests) to an AppointmentItem, converting
        // to the actual Outlook enums/flags OfficeAi.Shared can't reference
        // directly (it has no Outlook PIA reference, same split as
        // ColorUtil/BusyStatus in OutlookTools.Categories.cs). Called after
        // attendees are set but before .Save()/.Send() in both create_event
        // and draft_event - unverified whether GetRecurrencePattern() before
        // vs. after setting MeetingStatus/attendees matters; this ordering
        // (identity first, schedule second) matches natural Outlook UI flow
        // and is the one to verify live in Task 8.
        private static void ApplyRecurrence(Outlook.AppointmentItem a, RecurrenceSpec spec)
        {
            Outlook.RecurrencePattern pattern = a.GetRecurrencePattern();
            pattern.RecurrenceType = RecurrenceTypeMap[spec.Type];
            pattern.Interval = spec.Interval;

            if (spec.DaysOfWeek != null)
            {
                Outlook.OlDaysOfWeek mask = 0;
                foreach (string d in spec.DaysOfWeek) mask |= DayFlagMap[d];
                pattern.DayOfWeekMask = mask;
            }
            if (spec.DayOfMonth.HasValue) pattern.DayOfMonth = spec.DayOfMonth.Value;
            if (spec.Instance.HasValue) pattern.Instance = spec.Instance.Value;
            if (spec.MonthOfYear.HasValue) pattern.MonthOfYear = spec.MonthOfYear.Value;

            if (spec.Count.HasValue) pattern.Occurrences = spec.Count.Value;
            else if (spec.Until.HasValue) pattern.PatternEndDate = spec.Until.Value;
            else pattern.NoEndDate = true;
        }

        // Reads and validates the optional "recurrence" object from
        // create_event/draft_event's input. Returns null (with error left
        // null) when the field is omitted entirely; returns a populated
        // RecurrenceSpec on success; returns null with error set to a
        // model-facing IsError message on any validation failure.
        private static RecurrenceSpec ReadRecurrence(JsonElement input, out string error)
        {
            error = null;
            JsonElement rec;
            if (input.ValueKind != JsonValueKind.Object || !input.TryGetProperty("recurrence", out rec) || rec.ValueKind != JsonValueKind.Object)
                return null;

            string type = Str(rec, "type", null);
            int interval = Int(rec, "interval", 1);
            string[] days = StrArray(rec, "days_of_week");
            int? dayOfMonth = OptInt(rec, "day_of_month");
            int? instance = OptInt(rec, "instance");
            int? monthOfYear = OptInt(rec, "month_of_year");
            int? count = OptInt(rec, "count");
            DateTime? until = DateArg(rec, "until");

            return RecurrenceValidator.Parse(type, interval, days, dayOfMonth, instance, monthOfYear, count, until, out error);
        }
```

Add `using OfficeAi.Shared;` if not already present at the top of the file — it already is (line 4).

- [ ] **Step 2: Rewrite `DraftEvent`**

Replace the existing `DraftEvent` method with:

```csharp
        private static ToolResult DraftEvent(JsonElement input)
        {
            string recurrenceError;
            RecurrenceSpec recurrence = ReadRecurrence(input, out recurrenceError);
            if (recurrenceError != null) return new ToolResult { Output = recurrenceError, IsError = true, Summary = "draft_event" };

            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = SeedSignature(Str(input, "body", ""));

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (start.HasValue) a.Start = start.Value;
            if (end.HasValue) a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            if (!string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt))
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
            }
            if (recurrence != null) ApplyRecurrence(a, recurrence);
            a.Display(false);
            return new ToolResult
            {
                Output = "Opened an appointment draft in Outlook for the user to review and send." + (recurrence != null ? " Set to repeat " + recurrence.Type + "." : ""),
                Summary = "draft_event",
            };
        }
```

- [ ] **Step 3: Rewrite `CreateEvent`**

Replace the existing `CreateEvent` method with:

```csharp
        private static ToolResult CreateEvent(string mbxKey, JsonElement input)
        {
            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (!start.HasValue) return new ToolResult { Output = "start is required.", IsError = true, Summary = "create_event" };
            if (!end.HasValue) return new ToolResult { Output = "end is required.", IsError = true, Summary = "create_event" };

            string recurrenceError;
            RecurrenceSpec recurrence = ReadRecurrence(input, out recurrenceError);
            if (recurrenceError != null) return new ToolResult { Output = recurrenceError, IsError = true, Summary = "create_event" };

            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = SeedSignature(Str(input, "body", ""));
            a.Start = start.Value;
            a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            bool isMeeting = !string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt);
            if (isMeeting)
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
            }
            if (recurrence != null) ApplyRecurrence(a, recurrence);

            if (isMeeting)
            {
                a.Send();
                RecordIrreversible(mbxKey, "create_event invite for \"" + (a.Subject ?? "") + "\"");
                string attendeeList = string.IsNullOrEmpty(req) ? opt : string.IsNullOrEmpty(opt) ? req : req + "; " + opt;
                return new ToolResult
                {
                    Output = "Created and sent invite: \"" + (a.Subject ?? "") + "\" to " + attendeeList + "." + (recurrence != null ? " Repeats " + recurrence.Type + "." : ""),
                    Mutated = true,
                    Summary = "create_event",
                };
            }

            a.Save();
            RecordCreated(mbxKey, "create_event", "event_id", a, a.Subject ?? "");
            return new ToolResult
            {
                Output = "Created event: \"" + (a.Subject ?? "") + "\"." + (recurrence != null ? " Repeats " + recurrence.Type + "." : ""),
                Mutated = true,
                Summary = "create_event",
            };
        }
```

- [ ] **Step 4: Build to confirm it compiles**

Run: `"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Release -nologo -v:minimal`
Expected: clean, 0 errors/warnings.

- [ ] **Step 5: Run the shared test suite (regression check)**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo`
Expected: 208/208 passing, unchanged.

- [ ] **Step 6: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Compose.cs
git commit -m "feat(outlook): wire recurrence creation into create_event/draft_event"
```

---

## Task 7: `entry.ts` schema for `recurrence` on `create_event`/`draft_event`

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts` (`draft_event` schema ~lines 340-360, `create_event` schema ~lines 428-444 — exact line numbers shifted by Task 3/4's edits, locate by tool `name`)

- [ ] **Step 1: Add the `recurrence` property to both schemas**

Both `draft_event` and `create_event` get the same new `recurrence` property added to their `inputSchema.properties`:

```typescript
        recurrence: {
          type: 'object',
          description:
            'Makes this a recurring series. type is required: "daily", "weekly", "monthly" (needs day_of_month), "monthlyNth"/"yearlyNth" (needs instance 1-4 or 5 for "last", plus a single days_of_week entry; yearlyNth also needs month_of_year), or "yearly" (needs month_of_year and day_of_month). interval defaults to 1 (every N days/weeks/months/years). At most one of count (end after N occurrences) or until (end by date) - omit both for no end date.',
          properties: {
            type: { type: 'string', enum: ['daily', 'weekly', 'monthly', 'monthlyNth', 'yearly', 'yearlyNth'] },
            interval: { type: 'number', description: 'Every N days/weeks/months/years. Default 1.' },
            days_of_week: { type: 'array', items: { type: 'string' }, description: 'e.g. ["monday","wednesday"] for weekly; exactly one day for monthlyNth/yearlyNth.' },
            day_of_month: { type: 'number', description: '1-31. Required for monthly and yearly.' },
            instance: { type: 'number', description: '1-4 for 1st-4th, 5 for "last". Required for monthlyNth/yearlyNth.' },
            month_of_year: { type: 'number', description: '1-12. Required for yearly and yearlyNth.' },
            count: { type: 'number', description: 'End after N occurrences. Mutually exclusive with until.' },
            until: { type: 'string', description: 'End by this date. Mutually exclusive with count.' },
          },
          required: ['type'],
        },
```

Add this property to `draft_event`'s `inputSchema.properties` (alongside `subject`/`location`/`body`/`start`/`end`/`required_attendees`/`optional_attendees`) and update its description to mention `recurrence`; do the same for `create_event`.

- [ ] **Step 2: Update the two tools' top-level descriptions**

`draft_event`'s description gains: `' Pass recurrence to make it a repeating series.'` appended. `create_event`'s description gains the same, plus a note that recurrence creation has the same no-review-window caveat as everything else in that tool.

- [ ] **Step 3: Build to confirm it compiles**

Run the esbuild command from Task 3 Step 5.
Expected: clean.

- [ ] **Step 4: Commit**

```bash
git add OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add recurrence schema to create_event/draft_event"
```

---

## Task 8: `docs/ai-tool-surface.md` updates

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1: Add a dated update block**

Add a new `> **Update <today's date> (Outlook gains recurring-series support):**` block (matching the style of every prior dated block this session added — see the `2026-09-28 (Outlook gains cancel tools)` block for the template) summarizing: `RecurrenceValidator` (pure, unit-tested, `OfficeAi.Shared`), `ResolveOccurrenceTarget` (occurrence targeting via `RecurrencePattern.GetOccurrence`, confirmed via reflection), which 4 tools gained `occurrence_date`/`recurrence` params, and the undo/redo asymmetry (occurrence cancellation is always a barrier; occurrence reschedule reuses the existing `SnapshotEntry` mechanism, confirmed live in Task 5).

- [ ] **Step 2: Correct the "Structural fragility" section**

Find the sentence claiming `get_event`/`accept_meeting`/`decline_meeting`/`reschedule_event`/`draft_reschedule_event`/`cancel_event`/`draft_cancel_event` "cannot target a single occurrence unambiguously." Correct it: those tools that don't take `occurrence_date` still can't (unchanged), but `reschedule_event`/`draft_reschedule_event`/`cancel_event`/`draft_cancel_event` now can, via `occurrence_date` — note this was previously documented as a fundamental COM limitation, which reflection against the referenced PIA proved wrong (`RecurrencePattern.GetOccurrence` exists and works).

- [ ] **Step 3: Update the Mutating/Draft-and-display/Auto-send tool tables**

Add a note to each of the 4 extended tools' existing table rows (`reschedule_event`, `draft_reschedule_event`, `cancel_event`, `draft_cancel_event`, `create_event`, `draft_event`) describing the new optional param, matching the terse style already used for other params in that table.

- [ ] **Step 4: Update "Unproven at runtime"**

Add an entry for the two ordering/behavior risks flagged in Task 6's code comments (does `GetRecurrencePattern()` before/after `MeetingStatus` matter; does the whole-series `Start`/`End` reassignment on a recurring master behave as assumed — this is PR #21's still-open Finding #2, now load-bearing for a real feature) and note whether Task 5's manual verification resolved the occurrence-level risks.

- [ ] **Step 5: Commit**

```bash
git add docs/ai-tool-surface.md
git commit -m "docs(outlook): document recurring-series support"
```

---

## Task 9: Package, install, and manually verify series creation + whole-series operations (spec Section 6 items 2-4)

**Files:** none (build/package only).

- [ ] **Step 1: Package and install**

Same commands as Task 5, Step 1.

- [ ] **Step 2: Hand the user this manual test checklist**

1. **Series creation, each of the 6 types** — for each, call `create_event` with a distinct `recurrence` object (one daily, one weekly with 2 days, one monthly by day-of-month, one monthlyNth "2nd Tuesday", one yearly, one yearlyNth) and confirm via `list_events` (or Outlook's own UI) that the resulting series matches what was requested — right days, right interval, right end condition (`count` vs. `until` vs. no end date).
2. **Series creation validation errors** — try a few invalid `recurrence` objects (missing `days_of_week` for `weekly`, both `count` and `until` set, unrecognized `type`) and confirm each returns the specific, actionable `IsError` message from `RecurrenceValidator`, not a generic failure or raw COM exception.
3. **Recurring meeting creation** — `create_event` with `recurrence` and `required_attendees` together; confirm the invite sends correctly and shows as a recurring series to attendees.
4. **Whole-series reschedule** (PR #21's Finding #2, now load-bearing): `reschedule_event` with no `occurrence_date` on a recurring event — confirm it shifts every occurrence's time consistently, doesn't corrupt the pattern, and (for a meeting) sends one series-level update notice, not one per occurrence.
5. **Whole-series cancellation**: `cancel_event` with no `occurrence_date` on a recurring meeting — confirm attendees get a single "series canceled" notice, not one per occurrence, and the whole series disappears from the calendar.
6. **Series creation undo/redo**: `create_event` with `recurrence` (no attendees), then `undo_last_action` — confirm the whole series moves to Deleted Items, and `redo_last_action` brings it all back.

- [ ] **Step 3: Record findings**

Update `docs/superpowers/specs/2026-09-28-outlook-recurring-series-design.md`'s Section 6 risk list to mark items 2-4 verified (or not), with what was confirmed and any follow-up fixes needed. If any check fails, fix the underlying code and repeat Tasks 6-9 as needed before considering the feature complete.

---

## Self-Review Notes

- **Spec coverage:** Section 1 (recurrence schema) → Tasks 1, 6, 7. Section 2 (occurrence targeting) → Task 3. Section 3 (whole-series, no new code) → verified in Task 9. Section 4 (occurrence edit/cancel) → Tasks 3-4. Section 5 (undo/redo) → Tasks 3-4 (reuses existing `SnapshotEntry`/`RecordIrreversible`, simpler than the spec's original "new snapshot-entry variant" speculation — see Task 3 Step 3's code comment). Section 6 (risk/verification) → Tasks 5, 9, in the required order. Section 7 (docs/tier wiring) → Tasks 3, 4, 7, 8; tier placement confirmed unchanged (no `DraftTierTools`/`SendTierTools` edits anywhere in this plan).
- **Deviation from spec worth flagging to the user:** Section 5 originally proposed a new undo-entry type keyed by `(master EntryID, StoreID, original occurrence date)` for occurrence reschedule. Reading `OutlookTools.Undo.cs`'s actual `RecordSnapshot`/`SnapshotEntry` implementation during planning showed this is unnecessary — an occurrence becomes a real, independently-addressable item with its own `EntryID` once saved, so the existing generic mechanism just works. Behavior matches the spec's intent (occurrence reschedule stays undo-able); only the mechanism is simpler. Flagged for live confirmation in Task 5.
- **Placeholder scan:** no TBD/TODO; every step has concrete code or an exact command.
- **Type consistency:** `RecurrenceSpec` fields/`RecurrenceValidator.Parse` signature used identically across Tasks 1, 6. `ResolveOccurrenceTarget`'s signature used identically across Tasks 3, 4. `ToolResult`/`Mutated`/`IsError`/`Summary` match the existing type throughout (`OfficeAi.Shared.ToolResult`, unchanged).
