# Design: Outlook recurring-series support (create, edit, cancel — series and single occurrence)

**Status:** approved design, not yet implemented. Branch: none yet — to be created for implementation.

## Problem

Outlook's calendar tools (`create_event`/`draft_event`, `reschedule_event`/`draft_reschedule_event`, `cancel_event`/`draft_cancel_event`) have no concept of recurrence at all:

- `create_event`/`draft_event` never touch `RecurrencePattern` — every created event is a one-off.
- Every mutating calendar tool resolves `event_id` via `ItemById`, which for a recurring series always returns the **master** item. `docs/ai-tool-surface.md` currently documents this as a fundamental limitation ("all occurrences share one EntryID, so tools act on the whole series, not a single occurrence") — **this claim is wrong**, corrected by this design (see Ground truth below).

This design covers: creating a recurring series (with real recurrence parameters, not just a flag), editing a single occurrence vs. the whole series, and canceling a single occurrence vs. the whole series — including recurring **meetings** with attendees, not just plain appointments.

## Ground truth (verified via .NET reflection against the referenced PIA, not assumed)

Checked against `Microsoft.Office.Interop.Outlook, Version=15.0.0.0` (the exact version this project references), loaded from the GAC (`C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\`), matching this repo's established verification discipline (e.g. `reschedule_event`'s "no Propose New Time member" check).

- **`RecurrencePattern` properties** (settable): `RecurrenceType` (`OlRecurrenceType`), `Interval` (int), `DayOfWeekMask` (`OlDaysOfWeek`, flags), `DayOfMonth` (int), `Instance` (int), `MonthOfYear` (int), `Occurrences` (int), `PatternStartDate`/`PatternEndDate` (DateTime), `NoEndDate` (bool), `StartTime`/`EndTime` (DateTime — time-of-day component), plus read-only `Exceptions`.
- **`OlRecurrenceType`** has exactly 6 values: `olRecursDaily`, `olRecursWeekly`, `olRecursMonthly`, `olRecursMonthNth`, `olRecursYearly`, `olRecursYearNth`.
- **`AppointmentItem`** exposes `IsRecurring` (bool, read-only), `RecurrenceState` (`OlRecurrenceState`: `olApptNotRecurring`/`olApptMaster`/`olApptOccurrence`/`olApptException`), `GetRecurrencePattern()` (returns/creates a `RecurrencePattern`, converting a plain appointment into a recurring master), `ClearRecurrencePattern()`.
- **`RecurrencePattern.GetOccurrence(DateTime startDate) → AppointmentItem`** — this is the critical finding. **Single-occurrence targeting is possible.** The current "shared EntryID" limitation documented in this repo is not a COM API limitation — it's simply that no existing tool calls this method.
- **`Exception` (singular, one entry in `RecurrencePattern.Exceptions`) is entirely read-only**: `AppointmentItem`, `Deleted` (bool), `OriginalDate` (DateTime), `ItemProperties` — getters only, no setters, no `Delete`/`Remove`/`Add` method anywhere on `Exception` or the `Exceptions` collection. **There is no COM API to undo a deleted occurrence.** This is load-bearing for the undo/redo design below.

## Scope decisions (confirmed with the user during brainstorming)

1. **Recurring meetings (with attendees) are in scope**, not just plain appointments — full complexity accepted up front rather than deferred.
2. **All 6 `OlRecurrenceType` values** are supported for creation (Daily, Weekly, Monthly, Monthly-Nth, Yearly, Yearly-Nth) — no YAGNI cut here, despite Monthly-Nth/Yearly-Nth being rarer and harder to specify via chat.
3. **Extend existing tools**, don't add new ones: `create_event`/`draft_event` gain an optional `recurrence` param; `reschedule_event`/`draft_reschedule_event`/`cancel_event`/`draft_cancel_event` gain an optional `occurrence_date` param. Matches this repo's existing pattern of one tool spanning risk/scope classes via an input-aware branch (e.g. `delete_email`'s `permanent` flag) rather than growing the tool count.
4. **Undo/redo is in scope for this round**, with one deliberate exception forced by the `Exception` read-only finding above: occurrence-level cancellation is a **barrier**, not undo-able, full stop — there is no COM path to reverse it, regardless of plain-appointment vs. meeting.
5. **Human-in-the-loop chat confirmation** (the user's idea, floated late in brainstorming, for requiring an explicit approval click even at Full Autonomy for whole-series cancel/edit) is **explicitly out of scope for this design** — confirmed via investigation that no such mechanism exists anywhere in the codebase, and building it would touch shared `agent-core`/`chat-ui` infrastructure used by all four apps (Word/Excel/PowerPoint/Outlook), not just Outlook. The chokepoint, for whenever that gets its own design: `AgentLoop.finishTurn()` in `shared/web-src/agent-core/loop.ts:540-593`, between the `onToolStart` event (loop.ts:567) and the `skill.executeTool()` call (loop.ts:571) — nothing today awaits anything UI-controllable there. Tracked as a future, separate spec.
6. **Verbose, actionable validation errors are a hard requirement**, not a nice-to-have — every recurrence-field validation failure must name the exact problem and the fix, matching `set_event_availability`'s `TryParseBusyStatus` pattern ("Unknown availability \"X\". Valid values: ...").

## Design

### 1. Recurrence schema (creation)

`create_event`/`draft_event` gain an optional `recurrence` object:

```
recurrence: {
  type: 'daily' | 'weekly' | 'monthly' | 'monthlyNth' | 'yearly' | 'yearlyNth',
  interval?: number,          // every N days/weeks/months/years, default 1
  days_of_week?: string[],    // weekly: e.g. ['monday','wednesday']; monthlyNth/yearlyNth: single day
  day_of_month?: number,      // monthly: 1-31
  instance?: number,          // monthlyNth/yearlyNth: 1-4 for 1st/2nd/3rd/4th, 5 for "last" (matches RecurrencePattern.Instance's own convention directly, verify via reflection during implementation - not yet checked)
  month_of_year?: number,     // yearly/yearlyNth: 1-12
  count?: number,             // end after N occurrences
  until?: string,             // end by date (mutually exclusive with count)
}
```

Implementation: build the `AppointmentItem` exactly as `create_event`/`draft_event` do today (subject, location, body, attendees, `Start`/`End`), then — if `recurrence` is present — call `.GetRecurrencePattern()` and set the relevant fields *before* `.Save()`/`.Send()`. `Start`/`End` on the appointment itself become the first occurrence's time and the pattern's `StartTime`/`EndTime` (time-of-day); `PatternStartDate` derives from the same `start` value.

**Validation (all required — see Scope decision 6):**
- `type` must be one of the 6 values; unrecognized value → `IsError` listing all 6.
- `type: 'weekly'` requires non-empty `days_of_week`; missing → `IsError` naming the field and giving an example.
- `type: 'monthly'` requires `day_of_month` (1-31); out of range or missing → `IsError` with the valid range.
- `type: 'monthlyNth'` / `'yearlyNth'` require `instance` and exactly one `days_of_week` entry; `type: 'yearly'` / `'yearlyNth'` additionally require `month_of_year` (1-12).
- `interval < 1` → `IsError`.
- Both `count` and `until` present → `IsError` explaining they're mutually exclusive.
- Unknown day names in `days_of_week` → `IsError` listing valid day names (reuse the pattern from `TryParseBusyStatus`-style case-insensitive matching).
- Every validation error names the specific field and what would fix it — never a generic "invalid recurrence" message.

### 2. Occurrence targeting mechanism

`reschedule_event`/`draft_reschedule_event`/`cancel_event`/`draft_cancel_event` gain an optional `occurrence_date` param (a date string, matching the `start` value `list_events` already shows per-occurrence — that becomes the documented way to obtain a valid value). Shared resolution helper (new, in `OutlookTools.Calendar.cs`):

```csharp
private static ToolResult ResolveOccurrenceTarget(Outlook.AppointmentItem master, string occurrenceDate, string toolName, out Outlook.AppointmentItem target)
{
    target = master;
    if (occurrenceDate == null) return null; // whole-series path, unchanged
    if (!master.IsRecurring)
        return new ToolResult { Output = "\"" + (master.Subject ?? "") + "\" is not a recurring event, so occurrence_date doesn't apply. Omit it to act on the event itself.", IsError = true, Summary = toolName };
    Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
    try { target = pattern.GetOccurrence(ParsedDate(occurrenceDate)); }
    catch (Exception ex)
    {
        DebugLog.WriteException(toolName + " GetOccurrence", ex);
        return new ToolResult { Output = "No occurrence of \"" + (master.Subject ?? "") + "\" on " + occurrenceDate + ". Check list_events for this series' actual occurrence dates.", IsError = true, Summary = toolName };
    }
    return null; // success, target is set
}
```

Every existing check (organizer-authority, `.Save()`/`.Send()` branch selection) then runs against `target` instead of unconditionally `appt` — same code, resolved differently. `list_events`'s tool description gains a note that its `start` value is what to pass back as `occurrence_date`.

### 3. Whole-series edit/cancel

Omitting `occurrence_date` on a recurring event resolves `target = master` — exactly today's existing `reschedule_event`/`cancel_event` behavior, unchanged. No new implementation needed here; this is purely a verification item (see Risk section).

### 4. Occurrence-level edit/cancel

**Occurrence reschedule:** resolve `occurrence` via Section 2, set `occurrence.Start`/`occurrence.End`, then `.Save()` (plain) or `.Send()` (meeting) — structurally identical to `reschedule_event`'s existing branches, operating on `occurrence` instead of `master`.

**Occurrence cancellation:** resolve `occurrence`, then `.Delete()` directly (plain) or `MeetingStatus = olMeetingCanceled` + `.Send()` + `.Delete()` (meeting) — same shape as `cancel_event`, operating on `occurrence`. No "already-canceled occurrence" cleanup case exists (unlike whole-event `cancel_event`'s extension) — a deleted occurrence simply won't resolve via `GetOccurrence()` again, so a repeat call surfaces as the "no occurrence on that date" error from Section 2, not a distinct already-canceled state.

### 5. Undo/redo

| Action | Mechanism |
|---|---|
| Series creation (`create_event`, no attendees, with `recurrence`) | `RecordCreated`-style entry on the master — undo moves the whole series to Deleted Items, redo moves it back. Unchanged from today's non-recurring `create_event`. |
| Whole-series reschedule | Unchanged — `Start`/`End` snapshot on the master. |
| Whole-series cancellation (organized meeting) | Barrier, as today. |
| Occurrence reschedule | **New** snapshot-entry variant keyed by `(master EntryID, master StoreID, original occurrence date)`, not the occurrence's own EntryID — re-resolves via `master.GetRecurrencePattern().GetOccurrence(originalDate)` on undo/redo. Props: `Start`/`End`, same as `reschedule_event`'s existing snapshot. |
| Occurrence cancellation | **Barrier, always** — plain appointment or meeting, no exception. Forced by the read-only `Exception` API (Ground truth above): there is no way to reverse a deleted occurrence. Must be called out explicitly in the tool's model-facing description so the model doesn't assume occurrence cancellation is as safe as whole-event cancellation is at Draft tier. |

### 6. Risk areas & verification plan (ranked by consequence of being wrong)

1. **Occurrence-level `.Send()` semantics** (reschedule or cancel a single occurrence of a meeting) — highest risk. A wrong assumption sends a real, wrong notice to real attendees, with no undo for the cancel case. **Must be verified against a live mailbox with test-only attendees, on throwaway recurring test events, before merge.**
2. **Whole-series `.Send()` on the master** (reschedule/cancel a recurring meeting) — existing code, never tested against a *recurring* item specifically. Medium risk.
3. **Whole-series `Start`/`End` reassignment on the master** (PR #21's pre-existing, still-open Finding #2) — does it shift the whole pattern's time-of-day correctly, or corrupt/no-op the pattern? Medium risk; failure mode is "wrong behavior," not "wrong notice sent."
4. **`RecurrencePattern` field-setting for creation** — lower risk (a malformed series is wrong but not sent anywhere), needs live confirmation that all 6 pattern types produce the recurrence Outlook's own UI would show for the same input.

**Implementation and testing should proceed in this order** — occurrence-level meeting operations first (highest risk, get it right or don't ship it), series creation last (safest to get wrong, easiest to fix after the fact).

### 7. Docs / entry.ts / tier wiring

- `entry.ts` schema descriptions bake in the validation rules up front (Section 1), not just on rejection.
- `docs/ai-tool-surface.md`: dated update block; corrections to "Mutating tools"/"Auto-send tools"/"Draft-and-display tools" table rows for the four extended tools; **correction** to the "Structural fragility" section's now-inaccurate claim that recurring occurrences can't be targeted individually; updates to "Unproven at runtime" for the new risk areas (Section 6).
- Tier placement is **unchanged**: `create_event`/`reschedule_event`/`cancel_event` stay Full-autonomy-only, `draft_*` stay Draft-tier, regardless of `recurrence`/`occurrence_date` being present. Recurrence is additional *scope* for existing tools, not a new risk category requiring new tiers.

## Out of scope for this design

- Human-in-the-loop chat confirmation (Scope decision 5) — separate future spec, cross-app shared infrastructure.
- Any UI/chat-ui changes — this is entirely a C#/`entry.ts` schema change, same shape as every prior Outlook tool addition this session.
- Modifying an occurrence's attendee list independently of the series (adding/removing attendees for just one occurrence) — not raised during brainstorming, treated as not-yet-requested rather than deliberately excluded; flag if it turns out to matter during implementation.
