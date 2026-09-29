# Outlook edit_event Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `reschedule_event`/`draft_reschedule_event` with `edit_event`/`draft_edit_event`, a general pair that can edit an existing calendar event's `start`/`end`, `subject`, `body`, `location`, and attendees (not just time), reusing the recurring-series occurrence-targeting, collision-check, and undo/redo machinery already on this branch.

**Architecture:** Two new methods in `OutlookAiAddIn/OutlookTools.Calendar.cs` (`EditEvent`, `DraftEditEvent`) plus a new `ReplaceAttendees` helper, replacing `RescheduleEvent`/`DraftRescheduleEvent`/`RescheduleProps` entirely. Both reuse the existing `ResolveOccurrenceTarget`/`IsCanceledMeeting`/`IsReceivedMeeting`/`CheckOccurrenceReorderCollision` helpers unchanged. A single unified mutation flow (not per-field branches) decides `.Send()` vs `.Save()` vs draft `.Display(false)`, and whether the result is a `RecordSnapshot` (undo-able) or `RecordIrreversible` (barrier).

**Tech Stack:** C# / .NET Framework 4.8, `Microsoft.Office.Interop.Outlook` COM interop, TypeScript (`entry.ts` tool schemas).

**Spec:** `docs/superpowers/specs/2026-09-28-outlook-edit-event-design.md`

## Global Constraints

- **Full replacement**: `reschedule_event`/`draft_reschedule_event` are removed entirely (tool definitions, switch cases, tier-set entries, doc references) — `edit_event`/`draft_edit_event` are the only way to change an event's time or anything else after Task 3.
- **Attendee editing is whole-series/non-recurring only**: `occurrence_date` given together with `required_attendees`/`optional_attendees` is an `IsError`, both on `edit_event` and `draft_edit_event`.
- **No-op validation**: at least one of `start`, `end`, `subject`, `body`, `location`, `required_attendees`, `optional_attendees` must be provided, else a clean `IsError` naming all seven.
- **`start`/`end` pairing**: must be provided together or neither — a partial time edit is an `IsError`.
- **Barrier rule**: an `edit_event` call is a `RecordIrreversible` barrier (not `RecordSnapshot`) whenever it results in a `.Send()` (the event is or becomes a meeting) **or** it changes a whole recurring series' time (`RecurrencePattern` fields aren't reachable by `SnapshotEntry`) — this matches even when neither condition alone would require it (e.g. a whole-series time-only change on a plain event with no attendees is still a barrier).
- **`ForceUpdateToAllAttendees = false`**: set explicitly on `appt` immediately before any `.Send()` that follows an attendee change in `EditEvent` — confirmed default, set explicitly rather than relied on implicitly.
- **`draft_edit_event` never sends or persists**: no `.Save()`, `.Send()`, `ForceUpdateToAllAttendees`, `RecordSnapshot`, or `RecordIrreversible` call anywhere in `DraftEditEvent` — it only mutates the in-memory COM object and ends in `appt.Display(false)`, exactly like `draft_event`/the old `draft_reschedule_event`.
- **Tier placement unchanged**: `edit_event` occupies exactly the tier `reschedule_event` occupied (`SendTierTools`, Full autonomy only); `draft_edit_event` occupies exactly the tier `draft_reschedule_event` occupied (`DraftTierTools`, Draft only or higher).

---

### Task 1: `ReplaceAttendees` helper + `EditEvent` (send-tier)

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (insert new code; `RescheduleEvent`/`DraftRescheduleEvent`/`RescheduleProps` stay untouched for now — removed in Task 3)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (add switch case + `SendTierTools` entry)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (add `edit_event` schema, additive — `reschedule_event`'s schema stays for now)

**Interfaces:**
- Consumes (all pre-existing, unchanged): `ResolveOccurrenceTarget(Outlook.AppointmentItem master, string occurrenceDate, string toolName, out Outlook.AppointmentItem target) -> ToolResult?`; `IsCanceledMeeting(Outlook.AppointmentItem)`/`IsReceivedMeeting(Outlook.AppointmentItem) -> bool`; `CanceledMeetingError`/`ReceivedMeetingError(Outlook.AppointmentItem) -> string`; `CheckOccurrenceReorderCollision(Outlook.AppointmentItem master, DateTime original, DateTime target, string subject, string toolName) -> ToolResult?`; `AddAttendees(Outlook.AppointmentItem a, string csv, Outlook.OlMeetingRecipientType type)` (in `OutlookTools.Compose.cs`); `ReadProps(object item, string[] props) -> object[]`, `RecordSnapshot(string mbxKey, string toolName, object item, string subject, string[] props, object[] before)`, `RecordIrreversible(string mbxKey, string description)` (in `OutlookTools.Undo.cs`); `ReqStr`/`Str`/`DateArg`/`Iso` (in `OutlookTools.cs`).
- Produces: `ReplaceAttendees(Outlook.AppointmentItem appt, string requiredCsv, string optionalCsv)` (reused by Task 2's `DraftEditEvent`); `EditEvent(string mbxKey, JsonElement input) -> ToolResult`; `DescribeEditEventChanges(DateTime? start, DateTime? end, string subject, string body, string location, bool attendeesChanged, string oldStart, string oldEnd) -> string` (reused by Task 2).

- [ ] **Step 1: Insert `ReplaceAttendees` and `DescribeEditEventChanges` helpers into `OutlookAiAddIn/OutlookTools.Calendar.cs`**

Find this exact comment block (the one immediately before `DraftRescheduleEvent`):

```csharp
        // Draft-tier: opens the appointment with the new Start/End already set but
        // NOT saved, exactly like draft_event - the user reviews the moved time in
        // the native window and decides whether to save it (and, if it's a
        // meeting, whether to send the update themselves). Never touches
        // attendees.
        private static ToolResult DraftRescheduleEvent(JsonElement input)
```

Insert immediately before it (keeping `DraftRescheduleEvent` itself untouched):

```csharp
        // Shared by edit_event/draft_edit_event: replaces the attendee list
        // wholesale (not a diff/merge - the caller supplies the full new
        // list each time, reading the current one first via get_event if
        // they need to preserve someone). Removes every recipient except
        // the organizer, then re-adds via the same AddAttendees helper
        // create_event/draft_event already use. Recipients indices are
        // 1-based (confirmed via .NET reflection against the referenced
        // PIA, matching every other Outlook collection in this codebase);
        // iterating downward from Count avoids skipping an element after
        // Remove shifts the rest down. OlMeetingRecipientType.olOrganizer
        // == 0 (confirmed via reflection) is the only type excluded.
        private static void ReplaceAttendees(Outlook.AppointmentItem appt, string requiredCsv, string optionalCsv)
        {
            for (int i = appt.Recipients.Count; i >= 1; i--)
            {
                if (appt.Recipients[i].Type != (int)Outlook.OlMeetingRecipientType.olOrganizer)
                    appt.Recipients.Remove(i);
            }
            AddAttendees(appt, requiredCsv, Outlook.OlMeetingRecipientType.olRequired);
            AddAttendees(appt, optionalCsv, Outlook.OlMeetingRecipientType.olOptional);
            try { appt.Recipients.ResolveAll(); } catch { }
        }

        // Shared by edit_event/draft_edit_event's result text: describes
        // which of the seven optional fields were actually touched, so the
        // caller sees a precise summary instead of a generic "updated".
        private static string DescribeEditEventChanges(DateTime? start, DateTime? end, string subject, string body, string location, bool attendeesChanged, string oldStart, string oldEnd)
        {
            var parts = new List<string>();
            if (start.HasValue) parts.Add("time changed from " + oldStart + " - " + oldEnd + " to " + Iso(start.Value) + " - " + Iso(end.Value));
            if (subject != null) parts.Add("subject changed");
            if (body != null) parts.Add("body changed");
            if (location != null) parts.Add("location changed");
            if (attendeesChanged) parts.Add("attendees updated");
            return parts.Count > 0 ? " (" + string.Join(", ", parts) + ")" : "";
        }

        // Draft-tier: opens the appointment with the new Start/End already set but
        // NOT saved, exactly like draft_event - the user reviews the moved time in
        // the native window and decides whether to save it (and, if it's a
        // meeting, whether to send the update themselves). Never touches
        // attendees.
        private static ToolResult DraftRescheduleEvent(JsonElement input)
```

- [ ] **Step 2: Insert `EditEvent` into `OutlookAiAddIn/OutlookTools.Calendar.cs`**

Find this exact comment block (the one immediately before `RescheduleEvent`, right after the `RescheduleProps` field):

```csharp
        private static readonly string[] RescheduleProps = { "Start", "End" };

        private static ToolResult RescheduleEvent(string mbxKey, JsonElement input)
```

Insert immediately before it (keeping `RescheduleProps`/`RescheduleEvent` untouched):

```csharp
        // Full-autonomy-only, general edit: unlike reschedule_event (Start/End
        // only), this can touch start/end, subject, body, location, and
        // attendees in one call. One unified mutation flow decides .Send() vs
        // .Save() and whether the result is undo-able, rather than branching
        // per field - see the design's Section 2 for the reasoning.
        private static ToolResult EditEvent(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "edit_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "edit_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            if (IsCanceledMeeting(master))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "edit_event" };
            if (IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "edit_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            string subject = Str(input, "subject", null);
            string body = Str(input, "body", null);
            string location = Str(input, "location", null);
            string requiredAttendees = Str(input, "required_attendees", null);
            string optionalAttendees = Str(input, "optional_attendees", null);

            if (!start.HasValue && !end.HasValue && subject == null && body == null && location == null &&
                requiredAttendees == null && optionalAttendees == null)
                return new ToolResult { Output = "At least one of start, end, subject, body, location, required_attendees, optional_attendees must be provided.", IsError = true, Summary = "edit_event" };

            if (start.HasValue != end.HasValue)
                return new ToolResult { Output = "start and end must be provided together.", IsError = true, Summary = "edit_event" };

            if ((requiredAttendees != null || optionalAttendees != null) && occDate != null)
                return new ToolResult { Output = "Attendee changes only apply to the whole series - omit occurrence_date.", IsError = true, Summary = "edit_event" };

            if (occDate != null && start.HasValue)
            {
                ToolResult? collision = CheckOccurrenceReorderCollision(master, appt.Start, start.Value, appt.Subject ?? "", "edit_event");
                if (collision != null) return collision.Value;
            }

            string scopeNote = occDate != null ? " (this occurrence only)" : "";
            bool wasMeetingBefore = master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;
            bool isRecurringWholeSeriesTimeChange = occDate == null && master.IsRecurring && start.HasValue;

            var props = new List<string>();
            if (start.HasValue && !isRecurringWholeSeriesTimeChange) { props.Add("Start"); props.Add("End"); }
            if (subject != null) props.Add("Subject");
            if (body != null) props.Add("Body");
            if (location != null) props.Add("Location");
            object[] before = ReadProps(appt, props.ToArray());

            string oldStart = Iso(appt.Start);
            string oldEnd = Iso(appt.End);

            // Outlook does not allow setting AppointmentItem.Start/.End directly
            // on a recurring master (confirmed live 2026-09-28, see
            // reschedule_event's own history) - RecurrencePattern's fields are
            // the correct mechanism, same as create_event's recurrence support
            // and reschedule_event's whole-series fix.
            if (isRecurringWholeSeriesTimeChange)
            {
                Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
                pattern.PatternStartDate = start.Value.Date;
                pattern.StartTime = start.Value;
                pattern.EndTime = end.Value;
            }
            else if (start.HasValue)
            {
                appt.Start = start.Value;
                appt.End = end.Value;
            }

            if (subject != null) appt.Subject = subject;
            if (body != null) appt.Body = body;
            if (location != null) appt.Location = location;

            bool attendeesChanged = false;
            if (requiredAttendees != null || optionalAttendees != null)
            {
                ReplaceAttendees(appt, requiredAttendees ?? "", optionalAttendees ?? "");
                if (appt.MeetingStatus != Outlook.OlMeetingStatus.olMeeting) appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                // Confirmed default is already false; set explicitly so intent
                // doesn't depend on that default never changing. Outlook's own
                // mechanism for notifying only added/removed attendees rather
                // than everyone on the list.
                appt.ForceUpdateToAllAttendees = false;
                attendeesChanged = true;
            }

            bool isMeetingNow = wasMeetingBefore || attendeesChanged;
            // Barrier whenever the call sends an invite OR touches whole-series
            // RecurrencePattern fields - the latter can't be snapshotted by
            // SnapshotEntry even with zero attendee involvement (see
            // reschedule_event's own whole-series fix for the same rule).
            bool mustBarrier = isMeetingNow || isRecurringWholeSeriesTimeChange;
            string changeSummary = DescribeEditEventChanges(start, end, subject, body, location, attendeesChanged, oldStart, oldEnd);

            if (mustBarrier)
            {
                if (isMeetingNow) appt.Send(); else appt.Save();
                RecordIrreversible(mbxKey, "edit_event of \"" + (appt.Subject ?? "") + "\"" + scopeNote +
                                            (isMeetingNow ? " (update sent)" : "") +
                                            (isRecurringWholeSeriesTimeChange ? " (whole series time change)" : ""));
                return new ToolResult
                {
                    Output = "Updated" + (isMeetingNow ? " and sent update notice" : "") + scopeNote + ": \"" + (appt.Subject ?? "") + "\"." + changeSummary,
                    Mutated = true,
                    Summary = "edit_event",
                };
            }

            try
            {
                appt.Save();
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("EditEvent Save", ex);
                if (occDate != null)
                {
                    // Same false-negative Save() risk confirmed for
                    // reschedule_event's occurrence path: property setters
                    // write immediately via RPC, independent of Save()'s own
                    // finalize step, which can fail separately. Re-check by
                    // the NEW date if a time change was requested (the
                    // occurrence now sits there, not at occDate), else by the
                    // unchanged occDate.
                    Outlook.AppointmentItem recheck;
                    string checkDate = start.HasValue ? Iso(start.Value) : occDate;
                    ToolResult? stillMissing = ResolveOccurrenceTarget(master, checkDate, "edit_event", out recheck);
                    if (stillMissing == null)
                    {
                        appt = recheck;
                    }
                    else
                    {
                        return new ToolResult
                        {
                            Output = "Could not save changes to \"" + (appt.Subject ?? "") + "\": " + ex.Message +
                                     (start.HasValue ? " If this was a time change, it may have crossed another occurrence of the same series - check list_events for this series' other occurrence dates." : ""),
                            IsError = true,
                            Summary = "edit_event",
                        };
                    }
                }
                else
                {
                    return new ToolResult
                    {
                        Output = "Could not save changes to \"" + (appt.Subject ?? "") + "\": " + ex.Message,
                        IsError = true,
                        Summary = "edit_event",
                    };
                }
            }

            // RecordSnapshot reads appt.EntryID via ItemEntryIdOf(appt) - for an
            // occurrence, that's the real, resolvable EntryID GetOccurrence's
            // returned item gets once saved, so undo/redo works via the exact
            // same SnapshotEntry mechanism as reschedule_event - no new
            // undo-entry type needed. props is always non-empty here: reaching
            // this branch requires mustBarrier == false, which means attendees
            // were never touched (that forces isMeetingNow, hence a barrier)
            // and, if a time change was requested, it's occurrence/non-recurring
            // (whole-series-recurring also forces a barrier) - so at least one
            // of Start/End/Subject/Body/Location is always in props.
            RecordSnapshot(mbxKey, "edit_event", appt, appt.Subject ?? "", props.ToArray(), before);
            return new ToolResult
            {
                Output = "Updated" + scopeNote + ": \"" + (appt.Subject ?? "") + "\"." + changeSummary,
                Mutated = true,
                Summary = "edit_event",
            };
        }

        private static readonly string[] RescheduleProps = { "Start", "End" };

        private static ToolResult RescheduleEvent(string mbxKey, JsonElement input)
```

- [ ] **Step 3: Wire `edit_event` into `OutlookAiAddIn/OutlookTools.cs`**

In the `SendTierTools` set, change:

```csharp
        private static readonly HashSet<string> SendTierTools = new HashSet<string>
        {
            "send_email", "send_reply", "send_reply_all", "send_forward", "create_event", "reschedule_event", "cancel_event",
        };
```

to:

```csharp
        private static readonly HashSet<string> SendTierTools = new HashSet<string>
        {
            "send_email", "send_reply", "send_reply_all", "send_forward", "create_event", "reschedule_event", "cancel_event", "edit_event",
        };
```

In the `switch (name)` block, change:

```csharp
                    case "create_event": return CreateEvent(mbxKey, input);
                    case "reschedule_event": return RescheduleEvent(mbxKey, input);
                    case "cancel_event": return CancelEvent(mbxKey, input);
```

to:

```csharp
                    case "create_event": return CreateEvent(mbxKey, input);
                    case "reschedule_event": return RescheduleEvent(mbxKey, input);
                    case "cancel_event": return CancelEvent(mbxKey, input);
                    case "edit_event": return EditEvent(mbxKey, input);
```

- [ ] **Step 4: Add the `edit_event` schema to `OutlookAiAddIn/web-src/entry.ts`**

Find this exact block:

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

Insert immediately after it (leaving the `reschedule_event` block itself untouched for now — removed in Task 3):

```typescript
  {
    name: 'edit_event',
    description:
      'Edits an existing calendar event immediately - NO review window. Any combination of start+end (together), subject, body, location, required_attendees, optional_attendees - at least one must be given. If the result is (or becomes) a meeting, sends the update notice right away. Only available in Full autonomy. Prefer draft_edit_event unless the user clearly wants this applied right now, with no chance to review it first. Only works on events the user organizes or a plain appointment - if it\'s a meeting the user only attends (not the organizer), this returns an error instead of attempting an unauthoritative change; use Outlook\'s own "Propose New Time" for those. Omit occurrence_date to act on the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value) to target one occurrence instead - occurrence-level edits can only change start/end/subject/body/location, not attendees (attendee changes only apply to the whole series). To add/remove specific attendees while keeping others, read the current list with get_event first and pass the full new list here.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to edit (from list_events\' start value). Omit to act on the whole series. Cannot be combined with required_attendees/optional_attendees.' },
        start: { type: 'string', description: 'New start date-time, e.g. "2026-09-01T14:00". Must be given together with end.' },
        end: { type: 'string', description: 'New end date-time. Must be given together with start.' },
        subject: { type: 'string' },
        body: { type: 'string' },
        location: { type: 'string' },
        required_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>". Replaces the whole required-attendee list. Whole-series/non-recurring only.' },
        optional_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>". Replaces the whole optional-attendee list. Whole-series/non-recurring only.' },
      },
      required: ['event_id'],
    },
  },
```

- [ ] **Step 5: Build and verify compilation**

Run: `"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Release -nologo -v:minimal` (from `C:\dev\opendocs`). If it fails with a NuGet `RuntimeIdentifier` error, run the same command with `-t:Restore` first, then retry `-t:Build`.

Expected: Build succeeds with 0 errors.

Run: `/c/dev/opendocs/OutlookAiAddIn/node_modules/.bin/esbuild.cmd web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap` (from `OutlookAiAddIn/`).

Expected: esbuild reports success with no errors.

- [ ] **Step 6: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add edit_event (attendees, subject, body, location, start/end)"
```

---

### Task 2: `DraftEditEvent` (draft-tier)

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (insert new code)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (add switch case + `DraftTierTools` entry)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (add `draft_edit_event` schema, additive)

**Interfaces:**
- Consumes: everything Task 1 consumes, plus `ReplaceAttendees` and `DescribeEditEventChanges` (both from Task 1, same file).
- Produces: `DraftEditEvent(JsonElement input) -> ToolResult`.

- [ ] **Step 1: Insert `DraftEditEvent` into `OutlookAiAddIn/OutlookTools.Calendar.cs`**

Find this exact block (immediately before `RescheduleProps`/`EditEvent`, inserted in Task 1):

```csharp
        private static readonly string[] RescheduleProps = { "Start", "End" };
```

Insert immediately before it:

```csharp
        // Draft-tier counterpart to EditEvent: same validation and mutation
        // selection, but every change stays unsaved on the in-memory COM
        // object until appt.Display(false) opens it for the user to review
        // and save/send themselves. Never calls .Save()/.Send(), never
        // touches ForceUpdateToAllAttendees, never records undo/redo -
        // draft tools never persist anything, same contract as draft_event.
        private static ToolResult DraftEditEvent(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string occDate = Str(input, "occurrence_date", null);
            Outlook.AppointmentItem master = ItemById(id, null) as Outlook.AppointmentItem;
            if (master == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "draft_edit_event" };

            Outlook.AppointmentItem appt;
            ToolResult? occurrenceError = ResolveOccurrenceTarget(master, occDate, "draft_edit_event", out appt);
            if (occurrenceError != null) return occurrenceError.Value;

            if (IsCanceledMeeting(master))
                return new ToolResult { Output = CanceledMeetingError(appt), IsError = true, Summary = "draft_edit_event" };
            if (IsReceivedMeeting(master))
                return new ToolResult { Output = ReceivedMeetingError(appt), IsError = true, Summary = "draft_edit_event" };

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            string subject = Str(input, "subject", null);
            string body = Str(input, "body", null);
            string location = Str(input, "location", null);
            string requiredAttendees = Str(input, "required_attendees", null);
            string optionalAttendees = Str(input, "optional_attendees", null);

            if (!start.HasValue && !end.HasValue && subject == null && body == null && location == null &&
                requiredAttendees == null && optionalAttendees == null)
                return new ToolResult { Output = "At least one of start, end, subject, body, location, required_attendees, optional_attendees must be provided.", IsError = true, Summary = "draft_edit_event" };

            if (start.HasValue != end.HasValue)
                return new ToolResult { Output = "start and end must be provided together.", IsError = true, Summary = "draft_edit_event" };

            if ((requiredAttendees != null || optionalAttendees != null) && occDate != null)
                return new ToolResult { Output = "Attendee changes only apply to the whole series - omit occurrence_date.", IsError = true, Summary = "draft_edit_event" };

            if (occDate != null && start.HasValue)
            {
                ToolResult? collision = CheckOccurrenceReorderCollision(master, appt.Start, start.Value, appt.Subject ?? "", "draft_edit_event");
                if (collision != null) return collision.Value;
            }

            string scopeNote = occDate != null ? " (just this occurrence, not the whole series)" : "";
            bool isMeeting = master.MeetingStatus == Outlook.OlMeetingStatus.olMeeting;

            if (occDate == null && master.IsRecurring && start.HasValue)
            {
                Outlook.RecurrencePattern pattern = master.GetRecurrencePattern();
                pattern.PatternStartDate = start.Value.Date;
                pattern.StartTime = start.Value;
                pattern.EndTime = end.Value;
            }
            else if (start.HasValue)
            {
                appt.Start = start.Value;
                appt.End = end.Value;
            }

            if (subject != null) appt.Subject = subject;
            if (body != null) appt.Body = body;
            if (location != null) appt.Location = location;

            bool attendeesChanged = false;
            if (requiredAttendees != null || optionalAttendees != null)
            {
                ReplaceAttendees(appt, requiredAttendees ?? "", optionalAttendees ?? "");
                if (appt.MeetingStatus != Outlook.OlMeetingStatus.olMeeting) appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                attendeesChanged = true;
                isMeeting = true;
            }

            appt.Display(false);
            return new ToolResult
            {
                Output = "Opened \"" + (appt.Subject ?? "") + "\"" + scopeNote + " with the requested changes in Outlook for the user to review and " +
                         (isMeeting ? "save/send." : "save.") + (attendeesChanged ? " Attendee list updated - review before sending." : ""),
                Summary = "draft_edit_event",
            };
        }

        private static readonly string[] RescheduleProps = { "Start", "End" };
```

- [ ] **Step 2: Wire `draft_edit_event` into `OutlookAiAddIn/OutlookTools.cs`**

In `DraftTierTools`, change:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_reschedule_event", "draft_cancel_event",
```

to:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_reschedule_event", "draft_cancel_event", "draft_edit_event",
```

In the `switch (name)` block, change:

```csharp
                    case "draft_event": return DraftEvent(input);
                    case "draft_reschedule_event": return DraftRescheduleEvent(input);
                    case "draft_cancel_event": return DraftCancelEvent(input);
```

to:

```csharp
                    case "draft_event": return DraftEvent(input);
                    case "draft_reschedule_event": return DraftRescheduleEvent(input);
                    case "draft_cancel_event": return DraftCancelEvent(input);
                    case "draft_edit_event": return DraftEditEvent(input);
```

- [ ] **Step 3: Add the `draft_edit_event` schema to `OutlookAiAddIn/web-src/entry.ts`**

Find this exact block:

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

Insert immediately after it (leaving the `draft_reschedule_event` block itself untouched for now — removed in Task 3):

```typescript
  {
    name: 'draft_edit_event',
    description:
      'Opens an existing calendar event with requested changes already applied, unsaved, for the user to review and save/send. Any combination of start+end (together), subject, body, location, required_attendees, optional_attendees - at least one must be given. Omit occurrence_date to act on the whole series (or a non-recurring event); pass occurrence_date (a date from list_events\' start value) to target one occurrence instead - occurrence-level edits can only change start/end/subject/body/location, not attendees. To add/remove specific attendees while keeping others, read the current list with get_event first and pass the full new list here.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        occurrence_date: { type: 'string', description: 'For a recurring event: the date of the single occurrence to edit (from list_events\' start value). Omit to act on the whole series. Cannot be combined with required_attendees/optional_attendees.' },
        start: { type: 'string', description: 'New start date-time, e.g. "2026-09-01T14:00". Must be given together with end.' },
        end: { type: 'string', description: 'New end date-time. Must be given together with start.' },
        subject: { type: 'string' },
        body: { type: 'string' },
        location: { type: 'string' },
        required_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>". Replaces the whole required-attendee list. Whole-series/non-recurring only.' },
        optional_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>". Replaces the whole optional-attendee list. Whole-series/non-recurring only.' },
      },
      required: ['event_id'],
    },
  },
```

- [ ] **Step 4: Build and verify compilation**

Same two build commands as Task 1 Step 5.

Expected: both succeed with 0 errors.

- [ ] **Step 5: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add draft_edit_event"
```

---

### Task 3: Full replacement — remove reschedule_event/draft_reschedule_event

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (delete `RescheduleProps`, `RescheduleEvent`, `DraftRescheduleEvent`)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (remove switch cases + tier-set entries)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (remove schemas, update display entries/system prompt/tier arrays)

**Interfaces:**
- Consumes: nothing new — this task only removes code Tasks 1-2 made obsolete.
- Produces: nothing new — `edit_event`/`draft_edit_event` (Tasks 1-2) become the sole calendar-edit tools.

This task is mechanical: `edit_event`/`draft_edit_event` are already fully functional (Tasks 1-2), so this is deletion plus reference updates, not new logic.

- [ ] **Step 1: Delete `RescheduleProps`, `RescheduleEvent`, and `DraftRescheduleEvent` from `OutlookAiAddIn/OutlookTools.Calendar.cs`**

Delete the entire `DraftRescheduleEvent` method (from its leading comment block `// Draft-tier: opens the appointment with the new Start/End already set but` through its closing `}`), and the entire `RescheduleProps` field plus `RescheduleEvent` method (from the comment block starting `// Full-autonomy-only counterpart to draft_reschedule_event above` through `RescheduleEvent`'s closing `}`). Leave every helper they used (`ResolveOccurrenceTarget`, `IsCanceledMeeting`, `IsReceivedMeeting`, `CanceledMeetingError`, `ReceivedMeetingError`, `CheckOccurrenceReorderCollision`, `ReplaceAttendees`, `DescribeEditEventChanges`, `EditEvent`, `DraftEditEvent`) untouched — they're still used by `EditEvent`/`DraftEditEvent` and by `cancel_event`/`draft_cancel_event` below them in the same file.

Verify nothing else in the file references `RescheduleProps`/`RescheduleEvent`/`DraftRescheduleEvent` after deletion:

Run: `grep -n "RescheduleProps\|RescheduleEvent\|DraftRescheduleEvent" OutlookAiAddIn/OutlookTools.Calendar.cs`

Expected: no output (empty).

- [ ] **Step 2: Remove `reschedule_event`/`draft_reschedule_event` from `OutlookAiAddIn/OutlookTools.cs`**

In `DraftTierTools`, change:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_reschedule_event", "draft_cancel_event", "draft_edit_event",
```

to:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_cancel_event", "draft_edit_event",
```

In `SendTierTools`, change:

```csharp
            "send_email", "send_reply", "send_reply_all", "send_forward", "create_event", "reschedule_event", "cancel_event", "edit_event",
```

to:

```csharp
            "send_email", "send_reply", "send_reply_all", "send_forward", "create_event", "cancel_event", "edit_event",
```

In the `switch (name)` block, change:

```csharp
                    case "draft_event": return DraftEvent(input);
                    case "draft_reschedule_event": return DraftRescheduleEvent(input);
                    case "draft_cancel_event": return DraftCancelEvent(input);
                    case "draft_edit_event": return DraftEditEvent(input);
```

to:

```csharp
                    case "draft_event": return DraftEvent(input);
                    case "draft_cancel_event": return DraftCancelEvent(input);
                    case "draft_edit_event": return DraftEditEvent(input);
```

and change:

```csharp
                    case "create_event": return CreateEvent(mbxKey, input);
                    case "reschedule_event": return RescheduleEvent(mbxKey, input);
                    case "cancel_event": return CancelEvent(mbxKey, input);
                    case "edit_event": return EditEvent(mbxKey, input);
```

to:

```csharp
                    case "create_event": return CreateEvent(mbxKey, input);
                    case "cancel_event": return CancelEvent(mbxKey, input);
                    case "edit_event": return EditEvent(mbxKey, input);
```

- [ ] **Step 3: Remove `reschedule_event`/`draft_reschedule_event` schemas from `OutlookAiAddIn/web-src/entry.ts`**

Delete the `draft_reschedule_event` tool-definition block entirely (the one Task 2 Step 3 left untouched, immediately before the `draft_edit_event` block it inserted):

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

Delete the `reschedule_event` tool-definition block entirely (the one Task 1 Step 4 left untouched, immediately before the `edit_event` block it inserted):

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

- [ ] **Step 4: Update `OUTLOOK_TOOL_DISPLAY` in `entry.ts`**

Change:

```typescript
  draft_reschedule_event: d('Draft new time', 'טיוטת שינוי מועד', 'Opens an event with a new time to review and save/send.', 'פותח אירוע עם מועד חדש לבדיקה ולשמירה/שליחה.'),
  draft_cancel_event: d('Draft cancellation', 'טיוטת ביטול', 'Opens an event to review before canceling it.', 'פותח אירוע לבדיקה לפני ביטולו.'),
  send_email: d('Send email (auto-send)', 'שליחת הודעה (שליחה אוטומטית)', 'Sends an email immediately, no review window.', 'שולח הודעה מיידית, ללא חלון בדיקה.'),
  send_reply: d('Send reply (auto-send)', 'שליחת תשובה (שליחה אוטומטית)', 'Sends a reply immediately, no review window.', 'שולח תשובה מיידית, ללא חלון בדיקה.'),
  send_reply_all: d('Send reply all (auto-send)', 'שליחת תשובה לכולם (שליחה אוטומטית)', 'Sends a reply-to-all immediately, no review window.', 'שולח תשובה-לכולם מיידית, ללא חלון בדיקה.'),
  send_forward: d('Send forward (auto-send)', 'שליחת העברה (שליחה אוטומטית)', 'Forwards a message immediately, no review window.', 'מעביר הודעה מיידית, ללא חלון בדיקה.'),
  create_event: d('Create event (auto-send)', 'יצירת אירוע (שליחה אוטומטית)', 'Creates/sends a calendar event immediately, no review window.', 'יוצר/שולח אירוע יומן מיידית, ללא חלון בדיקה.'),
  reschedule_event: d('Reschedule event (auto-send)', 'שינוי מועד אירוע (שליחה אוטומטית)', 'Moves an event and sends the update immediately, no review window.', 'מזיז אירוע ושולח עדכון מיידית, ללא חלון בדיקה.'),
  cancel_event: d('Cancel event (auto-send)', 'ביטול אירוע (שליחה אוטומטית)', 'Cancels an event and notifies attendees immediately, no review window.', 'מבטל אירוע ומודיע למוזמנים מיידית, ללא חלון בדיקה.'),
}
```

to:

```typescript
  draft_cancel_event: d('Draft cancellation', 'טיוטת ביטול', 'Opens an event to review before canceling it.', 'פותח אירוע לבדיקה לפני ביטולו.'),
  draft_edit_event: d('Draft event edit', 'טיוטת עריכת אירוע', 'Opens an event with the requested changes to review and save/send.', 'פותח אירוע עם השינויים המבוקשים לבדיקה ולשמירה/שליחה.'),
  send_email: d('Send email (auto-send)', 'שליחת הודעה (שליחה אוטומטית)', 'Sends an email immediately, no review window.', 'שולח הודעה מיידית, ללא חלון בדיקה.'),
  send_reply: d('Send reply (auto-send)', 'שליחת תשובה (שליחה אוטומטית)', 'Sends a reply immediately, no review window.', 'שולח תשובה מיידית, ללא חלון בדיקה.'),
  send_reply_all: d('Send reply all (auto-send)', 'שליחת תשובה לכולם (שליחה אוטומטית)', 'Sends a reply-to-all immediately, no review window.', 'שולח תשובה-לכולם מיידית, ללא חלון בדיקה.'),
  send_forward: d('Send forward (auto-send)', 'שליחת העברה (שליחה אוטומטית)', 'Forwards a message immediately, no review window.', 'מעביר הודעה מיידית, ללא חלון בדיקה.'),
  create_event: d('Create event (auto-send)', 'יצירת אירוע (שליחה אוטומטית)', 'Creates/sends a calendar event immediately, no review window.', 'יוצר/שולח אירוע יומן מיידית, ללא חלון בדיקה.'),
  cancel_event: d('Cancel event (auto-send)', 'ביטול אירוע (שליחה אוטומטית)', 'Cancels an event and notifies attendees immediately, no review window.', 'מבטל אירוע ומודיע למוזמנים מיידית, ללא חלון בדיקה.'),
  edit_event: d('Edit event (auto-send)', 'עריכת אירוע (שליחה אוטומטית)', 'Edits an event and sends the update immediately, no review window.', 'עורך אירוע ושולח עדכון מיידית, ללא חלון בדיקה.'),
}
```

- [ ] **Step 5: Update `commentOnlyExtraTools` in `entry.ts`**

Change:

```typescript
    'draft_event',
    'draft_reschedule_event',
    'draft_cancel_event',
```

to:

```typescript
    'draft_event',
    'draft_cancel_event',
    'draft_edit_event',
```

- [ ] **Step 6: Update `autoSendTools` in `entry.ts`**

Change:

```typescript
  autoSendTools: ['send_email', 'send_reply', 'send_reply_all', 'send_forward', 'create_event', 'reschedule_event', 'cancel_event'],
```

to:

```typescript
  autoSendTools: ['send_email', 'send_reply', 'send_reply_all', 'send_forward', 'create_event', 'cancel_event', 'edit_event'],
```

- [ ] **Step 7: Update `undo_last_action`'s description in `entry.ts`**

Change:

```typescript
      'set_category_color, set_event_availability, reschedule_event without attendees, cancel_event on a plain or already-canceled whole event (moved back out of Deleted Items), and create_event without attendees (moved to Deleted Items). Sends, meeting invites, accept/decline_meeting, cancel_event on a still-active organized meeting, cancel_event on any single occurrence (always, whether plain or meeting - Outlook has no API to restore a deleted occurrence), and permanent deletes ' +
```

to:

```typescript
      'set_category_color, set_event_availability, edit_event on a non-recurring event or a single occurrence (start/end/subject/body/location only) without attendee changes, cancel_event on a plain or already-canceled whole event (moved back out of Deleted Items), and create_event without attendees (moved to Deleted Items). Sends, meeting invites, accept/decline_meeting, cancel_event on a still-active organized meeting, cancel_event on any single occurrence (always, whether plain or meeting - Outlook has no API to restore a deleted occurrence), edit_event whenever it sends an update or changes a whole recurring series\' time, and permanent deletes ' +
```

- [ ] **Step 8: Update the system prompt in `entry.ts`**

Change:

```typescript
    'You can read and search mail, open a specific message in its own Outlook window, read attachments, triage messages (mark read/unread, flag importance, move, delete), manage the calendar (list/read events, accept/decline invitations, reschedule or cancel events, color events with tags via list_color_categories/set_event_categories/set_category_color, set an event\'s Free/Busy/Tentative/Out of Office/Working Elsewhere status via set_event_availability), ' +
    'manage tasks and reminders, and draft replies/forwards/new mail and calendar events. ' +
    'Drafting tools (draft_email, reply_email, reply_all_email, forward_email, draft_event, draft_reschedule_event, draft_cancel_event) open a normal Outlook compose or appointment window pre-filled - they never send or create directly; the user reviews and sends. ' +
    'send_email/send_reply/send_reply_all/send_forward/create_event/reschedule_event/cancel_event are different: they send or create IMMEDIATELY, with no review window at all - only available in Full autonomy, and only worth using when the user has clearly asked for something to go out right now with no chance to check it first. Default to the drafting tools otherwise. ' +
    'reschedule_event/draft_reschedule_event only work on events the user organizes (or a plain appointment with no attendees) - on a meeting the user only attends, they return an error instead of an unauthoritative change; point the user at Outlook\'s own "Propose New Time" for those. Recurring events share one event_id for the whole series - omit occurrence_date to act on the whole series, or pass one (a date from list_events\' start value) to target a single occurrence instead, for reschedule_event/draft_reschedule_event/cancel_event/draft_cancel_event. ' +
    'cancel_event/draft_cancel_event have the same organizer-only restriction on a still-active meeting the user only attends - use decline_meeting instead. Canceling a meeting the user organizes sends a cancellation notice to attendees (Full autonomy for cancel_event, or reviewed first via draft_cancel_event); canceling a plain appointment, or any already-canceled event, just removes it from the calendar (moved to Deleted Items, recoverable), nobody to notify - cancel_event is the only way to dismiss an already-canceled event. ' +
    'message_id / event_id / task_id values are Outlook EntryIDs. When the user has one or more messages selected, that selection (with its message_id) is in your context - prefer it over searching. ' +
    'Prefer list_emails / search_emails / list_tasks (fast, server-side) over reading items one by one. ' +
    "Once you've found the relevant messages, apply_search can show the same results in the user's own Outlook window instead of only listing them in chat. " +
    "undo_last_action/redo_last_action step back and forward through your own actions in this chat (not the user's manual Outlook actions). Anything that sent something (emails, invites, meeting responses, cancellation notices) or a permanent delete can't be undone and blocks undo past it - say so rather than claim it was reversed. " +
    "Your available tools depend on the user's editing mode, from least to most permissive: Read only (read/search only) -> Draft only (also triage, tasks, reminders, and drafting replies/forwards/new mail/events/reschedules/cancellations) -> Automate approvals (also auto-accept/decline meeting invitations, which notifies the organizer) -> Full autonomy (also send_email/send_reply/send_reply_all/send_forward/create_event/reschedule_event/cancel_event, which send/create immediately).",
```

to:

```typescript
    'You can read and search mail, open a specific message in its own Outlook window, read attachments, triage messages (mark read/unread, flag importance, move, delete), manage the calendar (list/read events, accept/decline invitations, edit or cancel events, color events with tags via list_color_categories/set_event_categories/set_category_color, set an event\'s Free/Busy/Tentative/Out of Office/Working Elsewhere status via set_event_availability), ' +
    'manage tasks and reminders, and draft replies/forwards/new mail and calendar events. ' +
    'Drafting tools (draft_email, reply_email, reply_all_email, forward_email, draft_event, draft_edit_event, draft_cancel_event) open a normal Outlook compose or appointment window pre-filled - they never send or create directly; the user reviews and sends. ' +
    'send_email/send_reply/send_reply_all/send_forward/create_event/edit_event/cancel_event are different: they send or create IMMEDIATELY, with no review window at all - only available in Full autonomy, and only worth using when the user has clearly asked for something to go out right now with no chance to check it first. Default to the drafting tools otherwise. ' +
    'edit_event/draft_edit_event can change start/end, subject, body, location, and/or attendees in one call (at least one field required) - only work on events the user organizes (or a plain appointment with no attendees); on a meeting the user only attends, they return an error instead of an unauthoritative change, point the user at Outlook\'s own "Propose New Time" for those. Recurring events share one event_id for the whole series - omit occurrence_date to act on the whole series, or pass one (a date from list_events\' start value) to target a single occurrence instead, for edit_event/draft_edit_event/cancel_event/draft_cancel_event; attendee changes only apply to the whole series, never a single occurrence. To add/remove specific attendees while keeping the rest, read the current list with get_event first and pass the full new list to edit_event/draft_edit_event. ' +
    'cancel_event/draft_cancel_event have the same organizer-only restriction on a still-active meeting the user only attends - use decline_meeting instead. Canceling a meeting the user organizes sends a cancellation notice to attendees (Full autonomy for cancel_event, or reviewed first via draft_cancel_event); canceling a plain appointment, or any already-canceled event, just removes it from the calendar (moved to Deleted Items, recoverable), nobody to notify - cancel_event is the only way to dismiss an already-canceled event. ' +
    'message_id / event_id / task_id values are Outlook EntryIDs. When the user has one or more messages selected, that selection (with its message_id) is in your context - prefer it over searching. ' +
    'Prefer list_emails / search_emails / list_tasks (fast, server-side) over reading items one by one. ' +
    "Once you've found the relevant messages, apply_search can show the same results in the user's own Outlook window instead of only listing them in chat. " +
    "undo_last_action/redo_last_action step back and forward through your own actions in this chat (not the user's manual Outlook actions). Anything that sent something (emails, invites, meeting responses, cancellation notices) or a permanent delete can't be undone and blocks undo past it - say so rather than claim it was reversed. " +
    "Your available tools depend on the user's editing mode, from least to most permissive: Read only (read/search only) -> Draft only (also triage, tasks, reminders, and drafting replies/forwards/new mail/events/edits/cancellations) -> Automate approvals (also auto-accept/decline meeting invitations, which notifies the organizer) -> Full autonomy (also send_email/send_reply/send_reply_all/send_forward/create_event/edit_event/cancel_event, which send/create immediately).",
```

- [ ] **Step 9: Build, verify, and confirm no dangling references**

Run: `grep -rn "reschedule_event\|draft_reschedule_event\|RescheduleEvent\|DraftRescheduleEvent\|RescheduleProps" OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/web-src/entry.ts`

Expected: no output (empty) — every reference removed.

Run both build commands from Task 1 Step 5.

Expected: both succeed with 0 errors.

- [ ] **Step 10: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "refactor(outlook): remove reschedule_event/draft_reschedule_event, replaced by edit_event"
```

---

### Task 4: Update docs/ai-tool-surface.md

**Files:**
- Modify: `docs/ai-tool-surface.md`

**Interfaces:**
- Consumes: nothing — documentation only, describing Tasks 1-3's already-committed code.
- Produces: nothing — no code changes.

- [ ] **Step 1: Read the current file and update every `reschedule_event`/`draft_reschedule_event` reference**

Run: `grep -n "reschedule_event" docs/ai-tool-surface.md` to get the current line numbers (they will have shifted since this plan was written), then read each surrounding section and apply these rules:

1. **Tool reference table rows** (`draft_reschedule_event (added 2026-09-27)` and `reschedule_event (added 2026-09-27)`, currently two separate table rows): replace both with a single new row `edit_event (added 2026-09-28, replaces reschedule_event)` and `draft_edit_event (added 2026-09-28, replaces draft_reschedule_event)`, describing: resolves `event_id`/`occurrence_date` via the existing `ResolveOccurrenceTarget`; validates at least one of the seven optional fields is present and `start`/`end` are given together; applies `start`/`end` via `RecurrencePattern` fields (whole-series-recurring) or `appt.Start`/`.End` directly (occurrence/non-recurring), `subject`/`body`/`location` directly, and attendees via the new `ReplaceAttendees` helper (whole-series/non-recurring only, rejects `occurrence_date` + attendee fields together); `.Send()`s if the result is/becomes a meeting, else `.Save()`s; barrier (`RecordIrreversible`) whenever it sends or changes a whole recurring series' time, else `RecordSnapshot` with a dynamically-built props list (only the fields actually touched). Note `ForceUpdateToAllAttendees = false` is set explicitly before any attendee-driven `.Send()`. `draft_edit_event` never sends/saves/records — ends in `Display(false)`.
2. **`undo_last_action` table row**: update to describe `edit_event` in place of `reschedule_event` — undo-able cases (non-recurring or occurrence-level, no attendee change) vs. barrier cases (sends, or whole-series-recurring time change) — mirroring the existing wording style for `cancel_event`'s occurrence-level barrier note right next to it.
3. **Every prose reference** to `reschedule_event`/`draft_reschedule_event` (organizer-authority checks, occurrence targeting, the whole-series `RecurrencePattern` fix, the false-negative `Save()` retry, the collision check, tier placement, "Structural fragility"/"Unproven at runtime" sections): rename to `edit_event`/`draft_edit_event`, keeping the described mechanism identical (all of it is reused unchanged) and noting where it now also applies to `subject`/`body`/`location`/attendees, e.g. "same false-negative `Save()` risk `reschedule_event` had, now in `edit_event`'s occurrence path" rather than restating the whole mechanism from scratch each time.
4. **New note**: add one short dated update block (`**Update 2026-09-28:**`, matching this file's existing convention for dated update blocks near the top) stating that `reschedule_event`/`draft_reschedule_event` were fully replaced by `edit_event`/`draft_edit_event`, which additionally supports `subject`/`body`/`location`/attendee edits; the three new risk areas from the design spec (`ReplaceAttendees`'s recipient-clearing, converting a plain event into a meeting via `edit_event`, `subject`/`body`/`location` edits on a recurring master) are unverified at runtime pending manual testing, same caveat style already used for the recurring-series feature's own "Verification status: unproven at runtime" notes.

- [ ] **Step 2: Verify no dangling references**

Run: `grep -n "reschedule_event" docs/ai-tool-surface.md`

Expected: no output, OR only appears inside historical/dated note text that explicitly describes a past state (e.g. inside an already-existing `**Corrected 2026-09-28:**` block describing what used to be true) — never as a description of current tool behavior.

- [ ] **Step 3: Commit**

```bash
git add docs/ai-tool-surface.md
git commit -m "docs(outlook): document edit_event/draft_edit_event, remove reschedule_event references"
```

---

### Task 5: Build, test, package, install for manual testing

**Files:** none (build/test/deploy only)

**Interfaces:** none — verification task.

- [ ] **Step 1: Run the full C# test suite**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo` (from `C:\dev\opendocs`).

Expected: 216/216 passing (unaffected — this feature touches no code the test suite covers; `RecurrenceValidator`/`RecurrenceValidatorTests` are untouched by this plan).

- [ ] **Step 2: Rebuild the add-in**

Run both commands from Task 1 Step 5 (MSBuild + esbuild) again from a clean state to confirm the final combined diff builds.

Expected: both succeed with 0 errors.

- [ ] **Step 3: Check Outlook is closed, then package and install**

Run: `tasklist | grep -i outlook`

Expected: no Outlook process running. If one is running, ask the user to close Outlook before continuing.

Run:
```bash
cd deploy && powershell.exe -ExecutionPolicy Bypass -File ./package.ps1 -App Outlook && cd dist && powershell.exe -ExecutionPolicy Bypass -File ./install.ps1 -App Outlook
```

Expected: both scripts complete without error.

- [ ] **Step 4: Hand off for manual testing**

Tell the user the build is installed and ask them to restart Outlook, then provide a list of `FORCE_TOOL` commands covering (in the spec's risk order): `edit_event`/`draft_edit_event` attendee changes on a whole series and a non-recurring meeting (add/remove/replace attendees, converting a plain event into a meeting), then `subject`/`body`/`location` edits on an occurrence and on a whole recurring series, then combined `start`/`end` + other-field edits, then the existing occurrence-collision and organizer-authority checks (reused unchanged, but worth one confirming pass through the new tool name).

No commit for this task — it's verification only.
