# Inject Today's Date/Time Into the System Prompt — Implementation Plan

> **Status (2026-09-19): Task 1 implemented and verified. Task 2 (Outlook work week) investigated and abandoned — see its section below; the API it depended on doesn't exist in this repo's referenced PIA.**

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-reported — asking the model (via the Outlook add-in) to schedule something "on Tuesday" consistently produced a Wednesday event.

**Goal:** Give the model actual ground truth for "today" (date, weekday, local time, timezone) once per conversation, in every app (Word/Excel/PowerPoint/Outlook), so relative-date tools (`draft_event`'s `start`/`end`, `find_meeting_slots`' `start_date`/`end_date`, any "next Tuesday" / "tomorrow" phrasing) are resolved against a known value instead of the model's own guess. (A second goal — also giving Outlook the user's actual configured work week — was investigated as Task 2 and abandoned; see that section.)

**Root cause (confirmed, not assumed):** traced the full request path — `shared/web-src/agent-core/loop.ts:468` builds `system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? '')`; `systemSuffix` is defined once, shared by all four apps, in `shared/web-src/app-shell/bootstrap.ts:654-662`, and only ever carries the user's own "Document guidelines" text (Task 8). Every app's own `systemPrompt` (each `entry.ts`) is a static string with no date in it either. `shared/web-src/ai-provider/stream.ts` sends this system string straight to whichever provider (Anthropic/OpenAI/Gemini/custom) with nothing else injected. **The model is never told what today's date is, anywhere, in any app.** `OutlookTools.Calendar.cs`'s `DefaultWorkRange`/`WorkDays` (Sun-Thu workweek) only affect `find_meeting_slots`' *default* date range when no dates are given — `draft_event`'s `start`/`end` are plain ISO strings the model must produce unaided, with zero day-of-week logic anywhere in this repo. So the Tuesday→Wednesday failure is not a workweek-config bug and not specific to Outlook — it's a missing-ground-truth bug that affects every date-sensitive tool in every app equally.

**Tech Stack:** TypeScript (`shared/web-src/app-shell/bootstrap.ts`), shared across all four add-ins.

## Answers to the two open questions this plan started with

**Q: can the agent get the actual work week (Mon–Fri vs. Sun–Thu) instead of a hardcoded guess?**
**Corrected after verification, see Task 2 below — the original "yes" answer given here was wrong.** `Application.CalendarOptions` does not exist in this repo's referenced Outlook PIA (confirmed via .NET reflection, not just re-checked against docs). The setting does exist, but only in the Windows registry, in an undocumented bitmask this session could not safely decode. Not implemented — `find_meeting_slots` keeps its existing hardcoded Sun–Thu default.

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

### Task 2 (Outlook only) — ABANDONED after verification (2026-09-19), not implemented

**Original plan:** read the user's real work week from `Application.CalendarOptions` instead of the hardcoded Sun–Thu in `WorkDays()`/`DefaultWorkRange()` (`OutlookTools.Calendar.cs:169-193`).

**Step 0's own verification requirement caught a wrong assumption before any code was written** — exactly what it was there to do:

- **`Application.CalendarOptions` does not exist.** Checked via .NET reflection against the actual referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA (the same assembly `apply_search`'s plan verified `Explorer.Search` against): no `CalendarOptions` property on `Application`/`_Application`, no type named `*CalendarOptions*` anywhere in the assembly, no `WorkDay*`/`FirstDayOfWeek`/`WorkWeek` member anywhere in it either. The original answer to "can the agent get the workweek days" (given before this was checked) was **wrong** — corrected here.
- **The setting does exist, but only in the registry**, not via COM: `HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Calendar` has `WorkDay` (a DWORD bitmask), `CalDefStart`/`CalDefEnd` (start/end time in minutes). Confirmed present on this machine (`WorkDay = 248`).
- **The bitmask's day-to-bit mapping could not be confirmed.** Outlook's real, documented COM flags enum for this exact purpose, `OlDaysOfWeek` (confirmed via reflection: `olSunday=1, olMonday=2, olTuesday=4, olWednesday=8, olThursday=16, olFriday=32, olSaturday=64`, max sum 127), does **not** match — `248` exceeds 127, so the registry value uses some other, undocumented bit layout. Decoding it without an authoritative source would be exactly the kind of unverified guess this plan's own Task 0/Step 0 discipline exists to prevent, with a worse failure mode than a wrong COM member name: a bad decode produces silently wrong work-days, not a build error.

**Decision: do not implement.** `WorkDays()`/`DefaultWorkRange()` keep their existing hardcoded Sun–Thu behavior. If this is worth fixing later, the honest options are (a) a real, documented API if one turns up (not found in this session), (b) reverse-engineering the registry bit layout against a Microsoft source that actually states it (not attempted here), or (c) the much simpler route of exposing the work week as an explicit setting in the add-in's own UI instead of trying to auto-detect it from Outlook. None of these are in scope for this plan.

**No code changes in this task.** Task 1 above is unaffected and stands alone.
