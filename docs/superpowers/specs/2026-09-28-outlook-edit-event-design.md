# Design: `edit_event`/`draft_edit_event` — replacing `reschedule_event`/`draft_reschedule_event`

**Status:** approved design, not yet implemented. Branch: `feature/outlook-recurring-series` (continues directly on top of the recurring-series work already merged there).

## Problem

`reschedule_event`/`draft_reschedule_event` can only change an event's `start`/`end`. There is no way to edit a calendar event's subject, body, location, or attendee list once it exists — the only other mutating calendar tools are narrow single-purpose ones (`set_event_categories`, `set_category_color`, `set_event_availability`, `set_reminder`). This design generalizes reschedule into a full edit capability, reusing everything the recurring-series work already built (occurrence targeting, the whole-series `RecurrencePattern` fix, the collision check, the false-negative Save() retry) rather than duplicating it.

## Ground truth (verified this session, not assumed)

- **`AppointmentItem.RequiredAttendees`/`.OptionalAttendees` are directly settable string properties** (confirmed via .NET reflection on `_AppointmentItem`: `CanWrite = True` for both) — but `create_event`/`draft_event`'s existing `AddAttendees` helper doesn't use this path; it manipulates the `Recipients` collection directly (`Recipients.Add()` + `.Type` + `.ResolveAll()`). Which mechanism correctly *edits* an existing item's attendees (vs. just cosmetically changing a display string) is unconfirmed — this design uses the `Recipients`-collection approach, matching the already-proven creation path, rather than the unverified string-set path.
- **`AppointmentItem.ForceUpdateToAllAttendees`** (boolean) exists (confirmed via reflection) and is Outlook's own mechanism for exactly the "notify only added/removed attendees, not everyone" behavior visible in Outlook's UI. **Confirmed by the user: its default is `false`** (the "delta" behavior — newly added attendees get a fresh invite, removed attendees get a cancellation, unchanged attendees get nothing). This design sets it explicitly to `false` before `.Send()` whenever attendees change, matching the confirmed default rather than relying on it implicitly.
- All ground truth from the recurring-series design (`docs/superpowers/specs/2026-09-28-outlook-recurring-series-design.md`) carries forward unchanged: `RecurrencePattern.GetOccurrence`, the read-only `Exception`/`Exceptions` API (no undo for occurrence deletion), and — critically — that `AppointmentItem.Start`/`.End` cannot be set directly on a recurring master (must go through `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` instead, confirmed live and already fixed on this branch).

## Scope decisions (confirmed with the user during brainstorming)

1. **Full replacement**, not an alias: `reschedule_event`/`draft_reschedule_event` are removed entirely from `entry.ts` and the C# switch; `edit_event`/`draft_edit_event` are the only way to change an event's time (or anything else) going forward.
2. **Editable fields:** `start`, `end` (must be given together if either is — no partial time edits), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` (all independently optional; at least one required overall).
3. **Attendee editing is whole-series/non-recurring only** — `occurrence_date` + attendee fields together is an error. Subject/body/location/start/end remain editable at the occurrence level, same as today's reschedule.
4. **Send semantics:** any `edit_event` call that results in the event being (or becoming) a meeting sends an update via `.Send()`; a plain event/occurrence with no attendees before or after just `.Save()`s. `ForceUpdateToAllAttendees = false` is set explicitly whenever attendees are touched.
5. **No-op validation:** at least one of the seven optional fields must be provided, else a clean `IsError` naming all seven.
6. **Undo/redo:** reuses the existing generic `SnapshotEntry` mechanism for `subject`/`body`/`location` (and `start`/`end` where that already works today), with a **dynamically-built props list per call** instead of the fixed `RescheduleProps` constant — only the fields actually touched are snapshotted. Whole-series-recurring time changes remain a barrier (existing, approved design); if other fields are *also* changed in that same call, the whole call stays a barrier (no partial-undo of a mixed pattern-field + item-property change).
7. **Converting a plain event into a meeting is supported**: if `required_attendees`/`optional_attendees` are provided and the event has no attendees yet, `MeetingStatus` flips to `olMeeting` and the resulting call sends an invite — this is a natural consequence of "attendee editing," not a separate feature.

## Design

### 1. Schema

```
edit_event / draft_edit_event:
  event_id (required)
  occurrence_date (optional — targets a single occurrence, same semantics as today)
  start, end (optional, but required together if either is given)
  subject, body, location (optional, independent)
  required_attendees, optional_attendees (optional, independent — rejected together with occurrence_date)
```

Validation, in order:
1. `event_id` resolves to an appointment → existing check, unchanged.
2. Resolve `occurrence_date` via `ResolveOccurrenceTarget` → existing, unchanged.
3. Existing `IsCanceledMeeting(master)`/`IsReceivedMeeting(master)` organizer-authority checks → unchanged, still run against `master`.
4. At least one of the seven optional fields provided, else `IsError`: `"At least one of start, end, subject, body, location, required_attendees, optional_attendees must be provided."`
5. `start`/`end` given together or neither, else `IsError`: `"start and end must be provided together."`
6. If (`required_attendees` or `optional_attendees`) is given and `occurrence_date` is also given → `IsError`: `"Attendee changes only apply to the whole series - omit occurrence_date."`
7. If `occurrence_date` given and `start` given → existing `CheckOccurrenceReorderCollision`, unchanged.

### 2. Mutation and send/save decision

One unified flow, not per-field branches:

```
wasMeetingBefore = master.MeetingStatus == olMeeting   // reliable, pre-mutation (per the earlier master-vs-appt fix)
isRecurringWholeSeriesTimeChange = occDate == null && master.IsRecurring && start.HasValue

props = new List<string>()
if (start.HasValue && !isRecurringWholeSeriesTimeChange) props.AddRange("Start", "End")
if (subject != null) props.Add("Subject")
if (body != null) props.Add("Body")
if (location != null) props.Add("Location")
before = ReadProps(appt, props.ToArray())   // captured before any mutation; empty array is fine if only attendees/pattern change

// Apply mutations
if (isRecurringWholeSeriesTimeChange):
    pattern = master.GetRecurrencePattern()
    pattern.PatternStartDate = start.Value.Date; pattern.StartTime = start.Value; pattern.EndTime = end.Value
elif (start.HasValue):
    appt.Start = start.Value; appt.End = end.Value   // existing occurrence/non-recurring path, unchanged

if (subject != null) appt.Subject = subject
if (body != null) appt.Body = body
if (location != null) appt.Location = location

attendeesChanged = false
if (requiredAttendees != null || optionalAttendees != null):
    ReplaceAttendees(appt, requiredAttendees ?? "", optionalAttendees ?? "")   // new helper, see below
    if (appt.MeetingStatus != olMeeting) appt.MeetingStatus = olMeeting
    appt.ForceUpdateToAllAttendees = false
    attendeesChanged = true

isMeetingNow = wasMeetingBefore || attendeesChanged
mustBarrier = isMeetingNow || isRecurringWholeSeriesTimeChange   // whole-series RecurrencePattern fields are never undo-able, even with no attendees

if (mustBarrier):
    if (isMeetingNow) appt.Send() else appt.Save()
    RecordIrreversible(mbxKey, "edit_event of \"" + subject + "\"" + scopeNote)   // barrier
else:
    try { appt.Save() } catch { /* existing false-negative re-verify via ResolveOccurrenceTarget, unchanged */ }
    RecordSnapshot(mbxKey, "edit_event", appt, appt.Subject ?? "", props.ToArray(), before)
```

`mustBarrier` is true whenever the call sends an invite (`isMeetingNow`) OR touches whole-series `RecurrencePattern` fields (`isRecurringWholeSeriesTimeChange`) — the latter matches the existing recurring-series design exactly, just generalized to "any edit_event call that includes a whole-series-recurring time change," even one with no attendees and no other meeting involvement (e.g. a plain recurring series being time-shifted, which still must `.Save()` to persist the pattern change, but as an irreversible barrier rather than a snapshot, since `RecurrencePattern` fields aren't reachable by `SnapshotEntry`).

### 3. `ReplaceAttendees` helper (new)

```
ReplaceAttendees(AppointmentItem appt, string requiredCsv, string optionalCsv):
    // Remove every non-organizer recipient (iterate backwards since Remove shifts indices)
    for i = appt.Recipients.Count downto 1:
        if appt.Recipients[i].Type != olOrganizer:
            appt.Recipients.Remove(i)
    AddAttendees(appt, requiredCsv, olRequired)   // reuse existing helper
    AddAttendees(appt, optionalCsv, olOptional)
    try { appt.Recipients.ResolveAll() } catch { }
```

This is new, unverified COM interaction — flagged in Risk Areas below.

### 4. Draft variant (`draft_edit_event`)

Same validation and mutation-selection logic as `edit_event`, except:
- Time changes: `appt.Start`/`.End` set unsaved (occurrence/non-recurring), or `RecurrencePattern` fields set unsaved (whole-series-recurring) — mirrors the existing `draft_reschedule_event` pattern exactly (already proven safe for `Start`/`End`; the `RecurrencePattern` fields are the already-fixed `draft_reschedule_event` whole-series case, just generalized).
- `subject`/`body`/`location`/attendee changes applied the same way as the immediate version, but the call ends in `appt.Display(false)` instead of `.Save()`/`.Send()` — **never touches `ForceUpdateToAllAttendees` or calls `.Send()`**, since nothing is sent from a draft tool. The user reviews everything (including the pending attendee list) in the opened Outlook window and decides whether to send/save it themselves.
- No `RecordSnapshot`/`RecordIrreversible` calls anywhere (draft tools never persist, matching `draft_reschedule_event`'s existing contract).

### 5. Risk areas & verification order (ranked by consequence of being wrong)

1. **`ReplaceAttendees`'s recipient-clearing** — highest risk. Getting the organizer-exclusion wrong could remove the user's own organizer entry or fail to actually clear old attendees, and this feeds directly into a `.Send()` call with real people. Test on throwaway test-account attendees first.
2. **Converting a plain event into a meeting via `edit_event`** — untested combination (attendees added to an *existing*, previously-saved item, not a brand-new one like `create_event`).
3. **`subject`/`body`/`location` on a recurring master** — likely fine (not part of `RecurrencePattern`), but unverified.
4. **`ForceUpdateToAllAttendees`'s exact delta behavior** when combined with a simultaneous time/subject change (does it still correctly suppress notifying unchanged attendees, or does *any* other field change force a full notify regardless of this flag?).

Implementation and testing should proceed in this order — attendee mechanics first (highest risk, get it right or don't ship it), whole-series-recurring content edits last (safest to get wrong).

### 6. Full replacement mechanics

- Remove `reschedule_event`/`draft_reschedule_event` tool definitions from `entry.ts` entirely; add `edit_event`/`draft_edit_event` in their place.
- Rename `RescheduleEvent`/`DraftRescheduleEvent` methods to `EditEvent`/`DraftEditEvent` in `OutlookAiAddIn/OutlookTools.Calendar.cs` (or keep as new methods and delete the old ones — implementer's call, whichever keeps the diff more readable), updating the `OutlookTools.cs` switch statement (`"reschedule_event"`/`"draft_reschedule_event"` cases → `"edit_event"`/`"draft_edit_event"`), tier sets (`DraftTierTools`/`SendTierTools` — same tier placement as today, just the tool name changes), and every doc reference (`docs/ai-tool-surface.md` has many — the dated update blocks, table rows, the "Structural fragility" section, "Unproven at runtime", `undo_last_action`'s covered-tools list, the system prompt in `entry.ts`).
- `RescheduleProps` (the fixed `{"Start","End"}` constant) is removed — replaced by the dynamically-built props list described in Section 2.

## Out of scope for this design

- Occurrence-level attendee editing (Scope decision 3) — explicitly deferred.
- Adding/removing individual attendees without replacing the whole list (diffing) — the model achieves this today via `get_event` (read current attendees) + `edit_event` (write the full new list), per the user's own reasoning during brainstorming.
- Any UI/chat-ui changes — this is entirely a C#/`entry.ts` schema change, same shape as every prior addition this session.
- The human-in-the-loop chat confirmation idea (raised once already this session, already deferred to its own future spec) — not revisited here.
