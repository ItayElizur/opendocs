# Outlook Auto-Send Mode (send email / create event without a review step) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-requested — "a way to create a complete event / send an email (maybe simply on full automation settings)."

**Goal:** Let the model actually send an email or create/send a calendar event end-to-end, with no native Outlook compose/appointment window for the user to review first.

**This is a genuine reversal of a standing safety invariant, not an additive feature.** Every draft/compose tool in this codebase carries the same comment: *"Draft-and-display only. These NEVER call `.Send()` - a native Outlook compose/appointment window is opened for the user to review and send"* (`OutlookAiAddIn/OutlookTools.Compose.cs:11-12`), and `docs/ai-tool-surface.md:518-519` states it as a documented guarantee. This plan narrows that guarantee for Full Autonomy mode specifically, rather than removing it.

## Revision note (2026-09-19, user feedback on the first draft)

The first draft of this plan proposed a **second**, separately-toggled gate (an `autoSendEnabled` setting) on top of Full Autonomy. User feedback: **drop the second gate — Full Autonomy alone is enough.** Rationale confirmed against the actual codebase: every existing mutating Outlook tool (`mark_email_read`, `move_email`, `delete_email`, `accept_meeting`, `create_task`, `draft_email`, `draft_event`, etc.) already requires Full Autonomy and nothing more (`OutlookTools.cs:28-50`'s `AlwaysAllowedTools`/mode-gate is the *only* tier below it — there is no existing "more permissive than Full Autonomy" concept in this codebase to draw a line under). Adding the five new send tools to that exact same, already-existing tier is consistent with how every other mutating tool already works, not a new pattern. This **removes the old Task 1 (settings plumbing) and Task 2 (bridge propagation) entirely** — see the new Task 1/2 below, which do different things (default mode + mode description copy, not a second permission gate).

**Two things the user asked for instead, to offset the removed second gate:**
1. Change the **default** editing mode (across all four apps, today hardcoded to Full Autonomy on every fresh session) to something safer, so reaching Full Autonomy — and therefore auto-send — is always a deliberate switch, never the silent starting state.
2. Make Full Autonomy's own description, in Outlook specifically, say plainly that it can send emails and calendar invites on the user's behalf.

## Task 0: Decisions (trimmed after the revision above)

| # | Question | Recommendation | Why |
|---|---|---|---|
| 1 | New distinct tools (`send_email`, `send_reply`, `send_reply_all`, `send_forward`, `create_event`) vs. a hidden branch inside the existing `draft_*` tools? | **New distinct tools.** | Keeps every existing tool's description permanently truthful ("draft_email never sends" stays true in every mode, forever); the model chooses explicitly between "draft" and "send" rather than one tool silently changing behavior based on invisible state. |
| 2 | Visible confirmation after an auto-send fires? | **Yes, mandatory** — the chat UI must show a distinct, unmissable line ("Sent to X about Y" / "Created and sent invite to Z") styled differently from a normal "opened a draft" tool result. | Removing the native review window removes the *only* place the user would otherwise see this happened before it's irreversible. Not optional polish. **User agreed with this in the first round.** |
| 3 | Scope: Outlook only, or a pattern other apps could reuse later? | **Outlook only**, matching the literal request. | Word/Excel/PowerPoint have no equivalent "send" concept. **User agreed with this in the first round.** |

Rows removed after the revision (no longer applicable — there is no second gate to decide the shape of): the old #2 (global toggle vs. allow-list) and #6 (does the toggle persist).

---

**Architecture (single-gate design):** The five new send tools are gated **exactly** the way `draft_email`/`move_email`/every other mutating tool already is — by simply *not* being in `AlwaysAllowedTools` (`OutlookTools.cs:30-34`), so `ExecuteAsync`'s existing check (`:41-50`) blocks them outside Full Autonomy with no new code. Client-side, they're added to `ALL_OUTLOOK_TOOLS` the normal way and **left out of `readOnlyTools`** (`entry.ts:337-348`) — `availableForMode()` (`bootstrap.ts:356-360`) already hides everything not in that list unless the mode is Full Autonomy, exactly like `draft_event` today. No new setting, no new bridge message, no new per-mailbox dictionary.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools*.cs`), TypeScript (`shared/chat-ui/chat-ui.ts`, `shared/web-src/app-shell/bootstrap.ts`, `OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- Every new send-tool's description must say, in plain language, that it sends/creates immediately with **no review step** — so the model doesn't reach for it casually; still prefer `draft_*` unless the user's own instruction this turn clearly wants immediate sending.
- Reuse every existing helper the `draft_*` tools already use (`AddAttendees`, `PrependHtml`, recipient resolution) — only the final `.Send()`/`.Save()` step and the tool name are new.
- The signature line ("— Created with OpenDocs") still applies — auto-sent mail/events get the same signature `draft_email`/`draft_event` do.
- No new UI surface beyond the mandatory confirmation styling (Task 0 #2) and the default-mode/description changes below — no "automation dashboard," no history log; the existing chat transcript is the record.

---

### Task 1: Change the default editing mode (all four apps)

**Problem:** `chat-ui.ts:423`'s `const defaultMode: EditingMode = menuModes.indexOf('fullAutonomy') !== -1 ? 'fullAutonomy' : menuModes[menuModes.length - 1]` and `bootstrap.ts:351`'s `let editingMode: EditingMode = 'fullAutonomy'` both currently start **every fresh session, in every app**, already in Full Autonomy — before the user has touched anything. These two defaults are independent today (nothing keeps them in sync beyond both happening to be hardcoded to the same value) and **both** need to change together, or the UI and the actual tool-filtering logic disagree about what's selected.

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts`

- [ ] **Step 1: One shared helper, so the two call sites can never drift.** In `chat-ui.ts`, next to `resolveModes` (`:239`), add and export:

```ts
const DEFAULT_MODE_PRIORITY: EditingMode[] = ['trackChanges', 'readOnly', 'commentOnly', 'fullAutonomy']

export function defaultModeFor(modes: EditingMode[]): EditingMode {
  return DEFAULT_MODE_PRIORITY.find((m) => modes.includes(m)) ?? modes[0]
}
```

This picks Track Changes when the app offers it (Word/Excel/PowerPoint — none of them pass an explicit `availableModes`, so they get all four); Outlook's `availableModes: ['readOnly', 'fullAutonomy']` (`entry.ts:351`) doesn't include Track Changes at all (deliberately — see its existing comment, "Comment only / Track changes have no meaning for mail"), so it falls through to **Read-only**, which is already the mode Outlook's own server-side gate treats identically to Track Changes anyway. No change to Outlook's `availableModes` is needed or proposed — Track Changes stays off its menu; only which of its *existing* two options is selected by default changes.

- [ ] **Step 2:** In `chat-ui.ts:423`, replace the inline ternary with `defaultModeFor(menuModes)`.
- [ ] **Step 3:** In `bootstrap.ts:351`, replace `let editingMode: EditingMode = 'fullAutonomy'` with `let editingMode: EditingMode = defaultModeFor(resolveModes(config.availableModes))` (import both helpers from `chat-ui.ts`; `resolveModes` already exists there per the comment at `chat-ui.ts:236-238` and is what `mountChatUI` itself calls at `:422`).

**Verification:**
- [ ] `tsc --noEmit` clean; all four bundles rebuild.
- [ ] `npx vitest run` in `shared/chat-ui` — add a test asserting `defaultModeFor(['readOnly','commentOnly','trackChanges','fullAutonomy'])` returns `'trackChanges'` and `defaultModeFor(['readOnly','fullAutonomy'])` returns `'readOnly'`.
- [ ] Manual, each app: a fresh pane opens with Track Changes selected (Word/Excel/PowerPoint) or Read-only selected (Outlook) — not Full Autonomy — and the model's first-turn tool list matches that mode (e.g. in Outlook, `draft_email` is not offered until the user switches modes).

---

### Task 2: Full Autonomy's description, Outlook-specific

**Problem:** `modeFullAutonomyDesc: { en: 'Edits applied directly', he: 'עריכות מוחלות ישירות' }` (`chat-ui.ts:33`) is a single string shared by all four apps' mode menu. Changing it in place to mention sending email would be wrong for Word/Excel/PowerPoint, where Full Autonomy never sends anything to anyone.

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts` (accept an optional per-mode description override)
- Modify: `shared/web-src/app-shell/bootstrap.ts` (pass it through from config)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (supply Outlook's own copy)

- [ ] **Step 1:** Add an optional `modeDescriptionOverrides?: Partial<Record<EditingMode, { en: string; he: string }>>` to `ChatUIOptions` (near `modes?: EditingMode[]`, `chat-ui.ts:174`); wherever the menu/label renders a mode's description string, check the override map first, falling back to the shared `STRINGS` entry.
- [ ] **Step 2:** Thread it through `bootstrap.ts`'s `startAddIn` config (`AppConfig`, near `availableModes?: EditingMode[]`, `:335`) into the `mountChatUI` call.
- [ ] **Step 3:** In `OutlookAiAddIn/web-src/entry.ts`'s `startAddIn(...)` call, add:
  ```ts
  modeDescriptionOverrides: {
    fullAutonomy: {
      en: 'Edits applied directly - including sending emails and calendar invites in your name',
      he: 'עריכות מוחלות ישירות - כולל שליחת הודעות והזמנות יומן בשמך',
    },
  },
  ```

**Verification:** `tsc --noEmit` clean; bundle rebuilds. Manual: Outlook's mode menu shows the extended Full Autonomy description in both languages; Word/Excel/PowerPoint are unchanged.

---

### Task 3: C# — the send tools themselves

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Compose.cs` (new handlers)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (dispatcher cases only — **no new gate, no new dictionary**; the existing mode check at `:41-50` already covers these since they're simply absent from `AlwaysAllowedTools`)

- [ ] **Step 1: `send_email`** — same body as `DraftEmail` (`OutlookTools.Compose.cs:28-40`) but `m.Send()` instead of `m.Display(false)`; result text states who it was sent to and the subject (this text is exactly what Task 5's distinct styling wraps).
- [ ] **Step 2: `send_reply` / `send_reply_all`** — same as `ReplyEmail` (`:42-55`) but `.Send()` on the reply item instead of `.Display(false)`.
- [ ] **Step 3: `send_forward`** — same as `ForwardEmail` (`:57-71`) but `.Send()`.
- [ ] **Step 4: `create_event`** — same as `DraftEvent` (`:73-96`) but calls `a.Save()` for a plain appointment, or (when attendees were added, `MeetingStatus = olMeeting`) `a.Send()` to actually dispatch the meeting invite rather than just opening it — **confirm via reflection against the referenced Outlook PIA** which call is correct for a meeting item before assuming `.Send()` behaves like `.Save()`'s non-meeting case; do not guess (same discipline as the `apply_search` plan's Task 0 for `Explorer.Search`, and this plan's own note on `Application.CalendarOptions` in the date-context plan).
- [ ] **Step 5:** wire all five into `ExecuteAsync`'s switch (`:52-84`), in the mutating block. Each handler's `ToolResult.Mutated = true` (unlike the `draft_*` tools, which never set it — these genuinely change external state).

**Verification:** `dotnet build`/MSBuild clean. Manual (real mailbox, low-stakes test recipient): in Full Autonomy, ask for an email to be sent — confirm it actually lands in the recipient's inbox with no compose window ever appearing, and that Read-only mode doesn't even offer the tool (per Task 4). Repeat for `create_event` with and without attendees.

---

### Task 4: TypeScript — tool schemas and display strings

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] Add `send_email`/`send_reply`/`send_reply_all`/`send_forward`/`create_event` to `ALL_OUTLOOK_TOOLS`, same input shapes as their `draft_*`/`reply_email`/etc. counterparts, with descriptions that state plainly they send/create **immediately, with no review step**.
- [ ] Add each to `OUTLOOK_TOOL_DISPLAY`, with label text that visually signals "sends immediately" (e.g. an explicit "(auto-send)" suffix in both languages) — this is the tool-step chip the user sees in the transcript in real time.
- [ ] **Do not** add them to `readOnlyTools` — that omission alone is what keeps them invisible outside Full Autonomy, per the Architecture section above. No other config field needed.
- [ ] Extend the system prompt (`entry.ts:324-331`) with one clear sentence: these five tools send/create immediately with no review step, so use them only when the user's request this turn clearly wants that — default to `draft_*` otherwise.

**Verification:** `tsc --noEmit` clean; bundle rebuilds; the five tools appear (with display strings) only when Full Autonomy is selected.

---

### Task 5: Distinct "sent automatically" UI treatment (Task 0 #2)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts` (tool-step rendering)
- Modify: `shared/chat-ui/chat-ui.test.ts` (new test: a step for one of the five auto-send tool names renders with the distinct marker/class; a step for `draft_email` does not)

- [ ] Give the five auto-send tool names a distinct visual marker in the completed-step rendering (e.g. a filled/warning-colored icon instead of the plain checkmark `draft_*` steps get) — driven by tool name, not a new field threaded through the whole result pipeline, keeping this a presentation-only change.

**Verification:** `npx vitest run` passes including the new test; manual check that a live auto-send call visibly stands out in the transcript from a `draft_*` call.

---

### Task 6: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] New subsection for the five auto-send tools, explicitly contrasted with the existing "Draft-and-display tools" section (`:508-519`) — state that they require Full Autonomy (nothing more), and correct the current blanket claim ("These never call `.Send()`... nothing in this repo sends mail or creates calendar events on its own" and the "Excluded / deferred" section that currently lists `send_email`/`create_event` as deliberately unsupported) now that it's no longer true unconditionally.
- [ ] Note the changed default mode (Task 1) and the Outlook-specific Full Autonomy description (Task 2) somewhere discoverable — e.g. near the mode-gate documentation this file already has.

**Verification:** doc accurately reflects the shipped behavior once Tasks 1-5 land.
