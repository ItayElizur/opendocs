# Inject Today's Date/Time Into the System Prompt — Implementation Plan

> **Status (2026-09-19): Task 1 and Task 2 both implemented and verified.** Task 2's first attempt (`Application.CalendarOptions`) was correctly abandoned after verification showed the API doesn't exist — but a second, better source was found afterward: EWS's `GetUserAvailability` → `WorkingHours` (the same documented mechanism Outlook's own scheduling assistant uses), reusing the EWS plumbing `search_contacts` already built. See Task 2 below for the corrected, shipped version.

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-reported — asking the model (via the Outlook add-in) to schedule something "on Tuesday" consistently produced a Wednesday event.

**Goal:** Give the model actual ground truth for "today" (date, weekday, local time, timezone) once per conversation, in every app (Word/Excel/PowerPoint/Outlook), so relative-date tools (`draft_event`'s `start`/`end`, `find_meeting_slots`' `start_date`/`end_date`, any "next Tuesday" / "tomorrow" phrasing) are resolved against a known value instead of the model's own guess. Outlook additionally gets its mailbox's *actual* configured work week/hours (Task 2), replacing the previously-hardcoded Sun–Thu 9–18 default.

**Root cause (confirmed, not assumed):** traced the full request path — `shared/web-src/agent-core/loop.ts:468` builds `system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? '')`; `systemSuffix` is defined once, shared by all four apps, in `shared/web-src/app-shell/bootstrap.ts:654-662`, and only ever carries the user's own "Document guidelines" text (Task 8). Every app's own `systemPrompt` (each `entry.ts`) is a static string with no date in it either. `shared/web-src/ai-provider/stream.ts` sends this system string straight to whichever provider (Anthropic/OpenAI/Gemini/custom) with nothing else injected. **The model is never told what today's date is, anywhere, in any app.** `OutlookTools.Calendar.cs`'s `DefaultWorkRange`/`WorkDays` (Sun-Thu workweek) only affect `find_meeting_slots`' *default* date range when no dates are given — `draft_event`'s `start`/`end` are plain ISO strings the model must produce unaided, with zero day-of-week logic anywhere in this repo. So the Tuesday→Wednesday failure is not a workweek-config bug and not specific to Outlook — it's a missing-ground-truth bug that affects every date-sensitive tool in every app equally.

**Tech Stack:** TypeScript (`shared/web-src/app-shell/bootstrap.ts`), shared across all four add-ins.

## Answers to the two open questions this plan started with

**Q: can the agent get the actual work week (Mon–Fri vs. Sun–Thu) instead of a hardcoded guess?**
**Yes — via EWS, not COM.** The first attempt (`Application.CalendarOptions`) was wrong: that property does not exist in this repo's referenced Outlook PIA (confirmed via .NET reflection). The Windows registry does have the setting, but in an undocumented bitmask this session couldn't safely decode. The real answer: `ExchangeService.GetUserAvailability(...)`'s response includes `AttendeeAvailability.WorkingHours` (`DaysOfTheWeek`, `StartTime`, `EndTime`) — the same documented, server-side data Outlook's own scheduling assistant uses to shade "outside working hours." Confirmed via .NET reflection against the already-referenced `Microsoft.Exchange.WebServices` 2.2.0 assembly (the same one `search_contacts` uses) before writing any code. Implemented in Task 2 below, reusing `search_contacts`'s existing EWS connection/autodiscover plumbing (`OutlookEws.cs`, `FindExchangeAccountInfo`) rather than duplicating it.

**Q: is recomputing this every single turn too aggressive?**
Not in terms of cost — the entire system prompt and full conversation history is already resent on every turn regardless (the chat completion APIs this app talks to are stateless; see `stream.ts`), so one more ~20-token sentence riding along is negligible next to that. But the *design* was needlessly repetitive for no benefit: nothing in a normal chat session needs the date re-checked turn-by-turn. **Revised to compute once per conversation**, using the exact mechanism this codebase already has for exactly this shape of thing — `activeDocMessage`, frozen by `beginConversation()` at conversation-start boundaries only (`bootstrap.ts:395-406`, "initial load, New chat... never read live per-turn"). The only cost is a conversation that happens to span midnight keeps yesterday's date for its remainder — an accepted, explicit tradeoff, not an oversight.

## Global Constraints

- Use **local** date/time (the user's machine, via `Date`'s local getters / `Intl.DateTimeFormat().resolvedOptions().timeZone`), not UTC — `Date.prototype.toISOString()` is UTC and would silently shift the date near midnight in most timezones (including Israel, UTC+2/+3). Do not use it for the date/weekday fields.
- Keep the string short and unambiguous: explicit weekday name (the actual bug symptom) + ISO date + local time + IANA timezone name, e.g. `Today is Thursday, 2026-09-18, 14:32 (Asia/Jerusalem).` — do not rely on the model inferring the weekday from the date itself, since that's the exact computation that was already failing silently.
- Wrap in try/catch: if `Intl` throws for any reason in a given WebView2/locale configuration, fall back to a plain date string rather than breaking the whole system prompt.
- Task 2 does not touch `find_meeting_slots`' *ranking* logic (`OfficeAi.Shared/MeetingSlots.cs`) — only the hardcoded Sun–Thu day-set it's fed, replacing it with the real one.

---

### Task 1: Add the date-context line, computed once per conversation

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`

- [ ] **Step 1: Add a small formatting helper**, near the top of `bootstrap.ts` (module scope):

```ts
// Fix for: relative-date tool args (draft_event, find_meeting_slots, etc.)
// were resolved by the model with no ground truth for "today" anywhere in
// the system prompt - it had to guess both the date and the weekday from
// training data, which is exactly how "next Tuesday" turned into
// Wednesday. Computed once per conversation (see beginConversation()) -
// not worth recomputing every turn for a value that only changes at
// midnight; a conversation spanning midnight keeps its start-of-chat date.
function todayContextLine(): string {
  try {
    const now = new Date()
    const weekday = now.toLocaleDateString('en-US', { weekday: 'long' })
    const y = now.getFullYear()
    const m = String(now.getMonth() + 1).padStart(2, '0')
    const d = String(now.getDate()).padStart(2, '0')
    const hh = String(now.getHours()).padStart(2, '0')
    const mm = String(now.getMinutes()).padStart(2, '0')
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone
    return `Today is ${weekday}, ${y}-${m}-${d}, ${hh}:${mm} (${tz}).`
  } catch {
    return `Today is ${new Date().toDateString()}.`
  }
}
```

- [ ] **Step 2: Freeze it alongside `activeDocMessage`**, in the exact same spot and the exact same way (`bootstrap.ts:395-406`):

```ts
let savedDocMessage = ''
let activeDocMessage = ''
let activeDateContext = ''
function beginConversation(): void {
  activeDocMessage = savedDocMessage
  activeDateContext = todayContextLine()
}
```

- [ ] **Step 3: Use the frozen value in `systemSuffix`** (`bootstrap.ts:662`), not a live call:

```ts
systemSuffix: () => '\n\n' + activeDateContext + (activeDocMessage ? '\n\nDocument guidelines from the user:\n' + activeDocMessage : ''),
```

Update the comment above it (currently Task-8-specific) to note it now carries two independent pieces, both frozen at conversation start: the date line and the conditional document-guidelines text.

**Verification:**
- [ ] `tsc --noEmit` clean in all four apps (they all import `bootstrap.ts` via `@officeai/app-shell`).
- [ ] All four bundles rebuild (esbuild command in `docs/superpowers/plans/STATUS.md`).
- [ ] Manual, in whichever app is reachable: ask the model "what's today's date?" at the start of a fresh chat — it answers with the actual local date/weekday, not a guess.
- [ ] Manual, Outlook specifically (the reported repro): ask it to draft an event "next Tuesday at 3pm" on a day where that previously produced Wednesday — confirm `draft_event`'s `start` now lands on the correct date. Note the exact local date/weekday this was tested on, since the bug is date-dependent and a fixed regression test can't reproduce the same day every time.

---

### Task 2 (Outlook only) — implemented via EWS `GetUserAvailability`, not COM

**First attempt (`Application.CalendarOptions`) — genuinely doesn't exist, verified before writing code:** no `CalendarOptions` property on `Application`/`_Application`, no type named `*CalendarOptions*` anywhere in the referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA, no `WorkDay*`/`FirstDayOfWeek`/`WorkWeek` member anywhere in it. The registry does have the setting (`HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Calendar\WorkDay`, a DWORD bitmask — confirmed present, `248` on this machine), but its bit layout doesn't match Outlook's real `OlDaysOfWeek` flags (`olSunday=1...olSaturday=64`, max sum 127 — `248` exceeds that), and decoding an undocumented format would be an unverified guess with a silent-wrong-behavior failure mode. Correctly abandoned at the time.

**Second attempt — found by asking "is there really no non-hardcoded way?" instead of stopping at the first dead end.** EWS's Availability service has exactly this, documented and already reachable: `ExchangeService.GetUserAvailability(attendees, timeWindow, requestedData)` → `GetUserAvailabilityResults.AttendeesAvailability[i].WorkingHours`, a `WorkingHours` object with `DaysOfTheWeek: Collection<DayOfTheWeek>` (a plain Sunday..Saturday enum, no bitmask ambiguity), `StartTime`/`EndTime: TimeSpan`. This is the same data Outlook's own scheduling assistant uses to shade "outside working hours" — not a guess, confirmed via .NET reflection against the already-referenced `Microsoft.Exchange.WebServices` 2.2.0 assembly (`GetUserAvailability`'s two overloads, `GetUserAvailabilityResults`, `AttendeeAvailability`, `WorkingHours`, `AvailabilityData`, `TimeWindow`, `AttendeeInfo` constructors — every member this task's code calls) before writing any of it.

**Files:**
- Modified: `OutlookAiAddIn/OutlookEws.cs` — new `GetWorkingHoursAsync(Uri url, string smtp)` / `WorkWeekInfo` (Days/StartHour/EndHour), following `ResolveNamesAsync`'s exact existing shape (`NewService(url)`, `Task.Run`, same `ExchangeService` construction).
- Modified: `OutlookAiAddIn/OutlookTools.Calendar.cs`:
  - New `ResolveWorkWeekAsync()`: resolves the EWS endpoint the same way `search_contacts` does (reuses `FindExchangeAccountInfo()`/`OutlookEws.CachedUrl` directly — same partial class, no duplication), calls `GetWorkingHoursAsync`, caches the result for the process (`_cachedWorkWeek`/`_workWeekResolved` — a failed lookup is also cached as "resolved: unavailable" so it isn't retried every call). Any failure (no Exchange account, EWS unreachable, etc.) returns `null` rather than throwing — the caller falls back to the old hardcoded Sun–Thu/9–18 default, same graceful-degradation posture `search_contacts` already has for on-prem-only EWS.
  - `FindMeetingSlots` → `FindMeetingSlotsAsync`: awaits `ResolveWorkWeekAsync()` first; `start_hour`/`end_hour` default to the resolved hours (still overridable by explicit args); the work-days `HashSet<DayOfWeek>` feeds `WorkDays`/`DefaultWorkRange`, both generalized from their old Sun-Thu-literal shape to accept an arbitrary set (a simple forward-walk that works for any contiguous or non-contiguous work-days shape, not just a calendar-week-aligned one).
- Modified: `OutlookAiAddIn/OutlookTools.cs` — dispatcher case now `return await FindMeetingSlotsAsync(input);` (second Outlook tool to go async, after `search_contacts`).
- Modified: `OutlookAiAddIn/web-src/entry.ts` — `find_meeting_slots`' description no longer states a fixed Sun-Thu default.
- Modified: `docs/ai-tool-surface.md` — `find_meeting_slots` row and the async-tool-execution note updated.

**Verification:** `dotnet build`/MSBuild clean, `tsc --noEmit` clean, bundle rebuilds. **Not yet exercised against a live Exchange mailbox** (none reachable in this environment) — first live use should confirm `WorkingHours` actually populates for a real on-prem account and that a non-default work week (e.g. Mon–Fri) changes `find_meeting_slots`' default range/hours accordingly; on failure, confirm it falls back to Sun–Thu/9–18 with no error surfaced to the model.
