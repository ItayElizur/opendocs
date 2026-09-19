# Inject Today's Date/Time (and Outlook's Work Week) Into the System Prompt — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-reported — asking the model (via the Outlook add-in) to schedule something "on Tuesday" consistently produced a Wednesday event.

**Goal:** Give the model actual ground truth for "today" (date, weekday, local time, timezone) once per conversation, in every app (Word/Excel/PowerPoint/Outlook), so relative-date tools (`draft_event`'s `start`/`end`, `find_meeting_slots`' `start_date`/`end_date`, any "next Tuesday" / "tomorrow" phrasing) are resolved against a known value instead of the model's own guess. Outlook additionally gets the user's *actual* configured work week (Task 2), replacing today's hardcoded Sun–Thu assumption.

**Root cause (confirmed, not assumed):** traced the full request path — `shared/web-src/agent-core/loop.ts:468` builds `system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? '')`; `systemSuffix` is defined once, shared by all four apps, in `shared/web-src/app-shell/bootstrap.ts:654-662`, and only ever carries the user's own "Document guidelines" text (Task 8). Every app's own `systemPrompt` (each `entry.ts`) is a static string with no date in it either. `shared/web-src/ai-provider/stream.ts` sends this system string straight to whichever provider (Anthropic/OpenAI/Gemini/custom) with nothing else injected. **The model is never told what today's date is, anywhere, in any app.** `OutlookTools.Calendar.cs`'s `DefaultWorkRange`/`WorkDays` (Sun-Thu workweek) only affect `find_meeting_slots`' *default* date range when no dates are given — `draft_event`'s `start`/`end` are plain ISO strings the model must produce unaided, with zero day-of-week logic anywhere in this repo. So the Tuesday→Wednesday failure is not a workweek-config bug and not specific to Outlook — it's a missing-ground-truth bug that affects every date-sensitive tool in every app equally.

**Tech Stack:** TypeScript (`shared/web-src/app-shell/bootstrap.ts`) for Task 1, shared across all four add-ins. C# (`OutlookAiAddIn/OutlookTools.Calendar.cs`, `OfficeAi.Shared/PaneHostBase.cs` or equivalent bridge) + TypeScript for Task 2, Outlook-only.

## Answers to the two open questions this plan started with

**Q: can the agent get the actual work week (Mon–Fri vs. Sun–Thu) instead of a hardcoded guess?**
Yes — Outlook exposes this directly: `Application.CalendarOptions` (a `Microsoft.Office.Interop.Outlook.CalendarOptions` object) has `WorkDayMonday`/`WorkDayTuesday`/.../`WorkDaySunday` booleans and `FirstDayOfWeek`, reflecting whatever the user actually configured in Outlook's own Calendar Options — not an assumption baked into this codebase. Today's `WorkDays()`/`DefaultWorkRange()` (`OutlookTools.Calendar.cs:169-193`) hardcode Sun–Thu unconditionally, which is wrong for any user (or shared mailbox) configured differently. **This API's exact member names need the same live-verification discipline as everything else uncertain in this codebase** (confirm via reflection against the referenced PIA before coding, same posture as the `apply_search` plan's Task 0 for `Explorer.Search`) — treated as fact here based on long-standing Outlook object model documentation, but not yet checked against *this* repo's referenced PIA version. See Task 2.

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

### Task 2 (Outlook only): Read the user's real work week instead of hardcoding Sun–Thu

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (`WorkDays`/`DefaultWorkRange`, `OutlookTools.cs:169-193`)
- Modify: wherever a per-conversation value can reach the Outlook system prompt (see Step 2 — this needs a small new bridge path; there is currently no mechanism carrying a COM-derived value into `systemSuffix`/`buildContext` for Outlook specifically)

- [ ] **Step 0 (do first): verify the API.** Confirm, via .NET reflection against the actually-referenced `Microsoft.Office.Interop.Outlook` PIA (not assumed from general documentation), that `Application.CalendarOptions` exposes `WorkDayMonday`...`WorkDaySunday` (booleans) and `FirstDayOfWeek`. Record the exact confirmed member names as a comment above Step 1's code.

- [ ] **Step 1: Replace the hardcoded work-day set.** `WorkDays()` (`OutlookTools.Calendar.cs:184-193`) currently does:
  ```csharp
  if (d.DayOfWeek != DayOfWeek.Friday && d.DayOfWeek != DayOfWeek.Saturday)
  ```
  Replace with a lookup against `Application.CalendarOptions`'s actual booleans (cache the `HashSet<DayOfWeek>` once per call, not per-day). `DefaultWorkRange` (`:169-182`)'s "roll to next work week" logic should walk forward using the same real set instead of the literal `idx <= 4` / Sun-Thu assumption baked into its day-index arithmetic.

- [ ] **Step 2: Surface it to the model, not just to `find_meeting_slots`' default.** Two options, pick one based on how much this matters beyond `find_meeting_slots` (which already gets it for free via Step 1):
  - **Minimal (recommended for v1):** do nothing further — `find_meeting_slots`' description already tells the model its default range logic; once Step 1 makes that logic honest, no prompt text needs to change, since the tool's *behavior* is now correct even if the model never explicitly reasons about the work week itself.
  - **Fuller:** add one clause to Outlook's own system prompt addition, e.g. `Your work week is {days} (first day: {day}).` — this needs a real value threaded from C# into the WebView2 layer once per conversation, which has no existing path today (unlike Word/Excel/PowerPoint's `buildContext`, which is wired for selection state, not arbitrary COM-derived facts). Building that path is real, non-trivial scope (a new bridge message + a per-app `buildContext`-like hook) — do not build it speculatively; only take this branch if Step 1 alone proves insufficient in practice (e.g. the model still gets "next work day" wrong for a non-default work week even though `find_meeting_slots` itself now defaults correctly).

**Verification:**
- [ ] `dotnet build`/MSBuild on `OutlookAiAddIn` clean.
- [ ] Manual, on a mailbox with a non-default work week configured in Outlook's own Calendar Options (e.g. Mon–Fri): call `find_meeting_slots` with no explicit date range and confirm the default range now matches that configuration, not Sun–Thu.
- [ ] Manual, on a mailbox left at Sun–Thu (or whatever this environment's actual default is): confirm no behavior change from today.
