# Inject Today's Date/Time Into Every Skill's System Prompt — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-reported — asking the model (via the Outlook add-in) to schedule something "on Tuesday" consistently produced a Wednesday event.

**Goal:** Give the model actual ground truth for "today" (date, weekday, local time, timezone) on every turn, in every app (Word/Excel/PowerPoint/Outlook), so relative-date tools (`draft_event`'s `start`/`end`, `find_meeting_slots`' `start_date`/`end_date`, any "next Tuesday" / "tomorrow" phrasing) are resolved against a known value instead of the model's own guess.

**Root cause (confirmed, not assumed):** traced the full request path — `shared/web-src/agent-core/loop.ts:468` builds `system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? '')`; `systemSuffix` is defined once, shared by all four apps, in `shared/web-src/app-shell/bootstrap.ts:654-662`, and only ever carries the user's own "Document guidelines" text (Task 8). Every app's own `systemPrompt` (each `entry.ts`) is a static string with no date in it either. `shared/web-src/ai-provider/stream.ts` sends this system string straight to whichever provider (Anthropic/OpenAI/Gemini/custom) with nothing else injected. **The model is never told what today's date is, anywhere, in any app.** `OutlookTools.Calendar.cs`'s `DefaultWorkRange`/`WorkDays` (Sun-Thu workweek) only affect `find_meeting_slots`' *default* date range when no dates are given — `draft_event`'s `start`/`end` are plain ISO strings the model must produce unaided, with zero day-of-week logic anywhera in this repo. So the Tuesday→Wednesday failure is not a workweek-config bug and not specific to Outlook — it's a missing-ground-truth bug that affects every date-sensitive tool in every app equally.

**Architecture:** Fix it once, centrally, since `systemSuffix` is already the one shared per-turn injection point every app already goes through. Add a small pure function that formats "today" from the browser's own `Date`/`Intl` APIs (no new dependency, no COM/native call needed — the WebView2 runtime has standard `Date`/`Intl` support), and prepend its output to the existing `systemSuffix` return value in `bootstrap.ts`.

**Tech Stack:** TypeScript (`shared/web-src/app-shell/bootstrap.ts`), shared across Word/Excel/PowerPoint/Outlook add-ins.

## Global Constraints

- Recompute every turn, not once per conversation — `systemSuffix` is already called every turn (per its own doc comment at `agent-core/types.ts:68`), and a long-running conversation can cross midnight; do not cache the value.
- Use **local** date/time (the user's machine, via `Date`'s local getters / `Intl.DateTimeFormat().resolvedOptions().timeZone`), not UTC — `Date.prototype.toISOString()` is UTC and would silently shift the date near midnight in most timezones (including Israel, UTC+2/+3). Do not use it for the date/weekday fields.
- Keep the string short and unambiguous: explicit weekday name (the actual bug symptom) + ISO date + local time + IANA timezone name, e.g. `Today is Thursday, 2026-09-18, 14:32 (Asia/Jerusalem).` — do not rely on the model inferring the weekday from the date itself, since that's the exact computation that was already failing silently.
- Wrap in try/catch: if `Intl` throws for any reason in a given WebView2/locale configuration, fall back to a plain date string rather than breaking the whole system prompt.
- Does not touch `find_meeting_slots`' Sun-Thu default-range logic (`OutlookTools.Calendar.cs`) — that is a separate, already-working feature (a *default* range, not date resolution), out of scope here.

---

### Task 1: Add the date-context helper and wire it into `systemSuffix`

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`

- [ ] **Step 1: Add a small formatting helper**, near the top of `bootstrap.ts` (module scope, no dependencies on any of the function-local state around it):

```ts
// Fix for: relative-date tool args (draft_event, find_meeting_slots, etc.)
// were resolved by the model with no ground truth for "today" anywhere in
// the system prompt - it had to guess both the date and the weekday from
// training data, which is exactly how "next Tuesday" turned into
// Wednesday. Recomputed every turn (see call site) since a conversation
// can span midnight.
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

- [ ] **Step 2: Prepend it to the existing `systemSuffix`** at `bootstrap.ts:662`:

```ts
systemSuffix: () => '\n\n' + todayContextLine() + (activeDocMessage ? '\n\nDocument guidelines from the user:\n' + activeDocMessage : ''),
```

Update the comment above it (currently Task-8-specific) to note this now carries two independent pieces: the always-present date line, and the conditional document-guidelines text.

**Verification:**
- [ ] `tsc --noEmit` clean in all four apps (they all import `bootstrap.ts` via `@officeai/app-shell`).
- [ ] All four bundles rebuild (esbuild command in `docs/superpowers/plans/STATUS.md`).
- [ ] Manual, in whichever app is reachable: ask the model "what's today's date?" — it answers with the actual local date/weekday, not a guess.
- [ ] Manual, Outlook specifically (the reported repro): ask it to draft an event "next Tuesday at 3pm" on a day where that previously produced Wednesday — confirm `draft_event`'s `start` now lands on the correct date. Note the exact local date/weekday this was tested on in the verification writeup, since the bug is date-dependent and a fixed regression test can't reproduce the same day every time.
