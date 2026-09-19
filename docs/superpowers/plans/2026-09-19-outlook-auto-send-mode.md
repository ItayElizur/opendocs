# Outlook Auto-Send Mode (send email / create event without a review step) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-requested — "a way to create a complete event / send an email (maybe simply on full automation settings)."

**Goal:** Let the model actually send an email or create/send a calendar event end-to-end, with no native Outlook compose/appointment window for the user to review first.

**This is a genuine reversal of a standing safety invariant, not an additive feature.** Every draft/compose tool in this codebase carries the same comment: *"Draft-and-display only. These NEVER call `.Send()` - a native Outlook compose/appointment window is opened for the user to review and send"* (`OutlookAiAddIn/OutlookTools.Compose.cs:11-12`), and `docs/ai-tool-surface.md:518-519` states it as a documented guarantee. This plan narrows that guarantee for Full Autonomy mode specifically, rather than removing it.

## Revision history

**Round 1 → Round 2:** dropped a proposed second, separately-toggled gate (`autoSendEnabled`) — Full Autonomy alone was made the single gate for the five new send tools, matching how every other mutating Outlook tool already works.

**Round 2 → Round 3 (this revision):** user proposal — instead of Outlook jumping straight from Read-only to "everything including auto-send" in one step, give it a real **third, middle tier**: draft/mutate freely, but never actually send anything. This reuses the `TrackChanges` `EditingMode` slot that already exists in the shared enum (`OfficeAi.Shared/PaneHostBase.cs:13`) but that Outlook has never used (its own comment says Track Changes "has no meaning for mail" — true for *editing*, but a perfectly good name-slot to repurpose as "drafts, never sends"). This is a better fit than Round 2's fallback (Outlook defaulting to Read-only because it had no third option) — Outlook now gets the same three-tier shape Word/Excel/PowerPoint already have, just with different meaning per tier.

**A finding that came out of designing this tier, not part of the original ask:** `accept_meeting`/`decline_meeting` (`OutlookTools.Calendar.cs:195-224`) already call `resp.Send()` to notify the organizer — they're an existing, easy-to-miss instance of automatic external communication, currently gated only behind Full Autonomy same as everything else. If the new middle tier is going to honestly mean "nothing leaves this mailbox without you reviewing it," these two belong in the **new** Full-Autonomy-only bucket alongside the five new send tools, not in the middle tier. Flagging this explicitly since it reclassifies two *existing* tools, not just adds new ones — see Task 0 below.

## Task 0: Decisions

| # | Question | Recommendation | Why |
|---|---|---|---|
| 1 | New distinct tools (`send_email`, `send_reply`, `send_reply_all`, `send_forward`, `create_event`) vs. a hidden branch inside the existing `draft_*` tools? | **New distinct tools.** | Keeps every existing tool's description permanently truthful ("draft_email never sends" stays true in every mode, forever). |
| 2 | Visible confirmation after an auto-send fires? | **Yes, mandatory.** | Removing the native review window removes the only place the user would otherwise see this happened before it's irreversible. **Agreed in round 1.** |
| 3 | Scope: Outlook only? | **Yes.** | Word/Excel/PowerPoint have no "send" concept. **Agreed in round 1.** |
| 4 | Should `accept_meeting`/`decline_meeting` move into the new Full-Autonomy-only tier (alongside the 5 send tools), since they already auto-notify the organizer via `resp.Send()`? | **Yes, recommended** — otherwise "the middle tier never sends anything" would be false the first time a user accepts a meeting in it. | Discovered while designing the tier, not requested — **needs your confirmation**, since it's a behavior change for two tools that exist today and currently work under Full Autonomy only (no regression either way — they'd simply stay Full-Autonomy-gated, just now for an explicit stated reason instead of incidentally). |

**Do not proceed past Task 0 #4 without confirming it** — everything below assumes yes.

---

**Architecture (three-tier design):**

| Tier | `EditingMode` | Outlook tools available |
|---|---|---|
| Read-only | `readOnly` | `AlwaysAllowedTools` only — `list_emails`, `search_emails`, `get_email`, `list_folders`, `search_contacts`, `list_events`, `get_event`, `list_tasks`, `get_attachment`, `find_meeting_slots` (`OutlookTools.cs:30-34`). Unchanged. |
| **Draft only** (repurposed `trackChanges`) | `trackChanges` | Everything Read-only has, **plus** every existing mutating/drafting tool that never leaves the mailbox without a review step: `mark_email_read`/`unread`, `flag_email_important`, `move_email`, `delete_email`, `create_task`, `update_task`, `set_reminder`, `set_email_reminder`, `draft_email`, `reply_email`, `reply_all_email`, `forward_email`, `draft_event`. **New tier — none of this changes what these tools do, only which mode unlocks them.** |
| Full autonomy | `fullAutonomy` | Everything Draft-only has, **plus** the tools that actually send something to someone outside the review loop: `accept_meeting`, `decline_meeting` (existing, reclassified per Task 0 #4), and the five new `send_email`/`send_reply`/`send_reply_all`/`send_forward`/`create_event`. |

Both layers enforce this, same defense-in-depth posture the existing mode gate already has:
1. **Client-side tool-list gating** — a new `config.fullAutonomyOnlyTools: string[]` (Outlook only) lists the 7 tools in the bottom row; `availableForMode()` (`bootstrap.ts:356-360`) excludes them whenever mode is `trackChanges`, on top of its existing behavior. **Word/Excel/PowerPoint don't set this field, so it's an empty set for them — their `trackChanges` behavior (show every tool, since real editing tools are legitimately usable under Word's own track-changes recording) is completely unchanged.**
2. **Server-side re-check in `OutlookTools.ExecuteAsync`** (`:36-50`) — the existing binary check (`mode != EditingMode.FullAutonomy`) becomes a three-way check: always-allowed tools pass regardless; the 7 full-autonomy-only tools require `mode == FullAutonomy`; everything else requires `mode == FullAutonomy || mode == TrackChanges`. **This is the actual guarantee**, same as before — client-side filtering is the UX nicety.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools*.cs`), TypeScript (`shared/chat-ui/chat-ui.ts`, `shared/web-src/app-shell/bootstrap.ts`, `OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- Every new send-tool's description must say, in plain language, that it sends/creates immediately with **no review step** — so the model doesn't reach for it casually; prefer `draft_*` unless the user's own instruction this turn clearly wants immediate sending.
- Reuse every existing helper the `draft_*` tools already use (`AddAttendees`, `PrependHtml`, recipient resolution) — only the final `.Send()`/`.Save()` step and the tool name are new.
- The signature line ("— Created with OpenDocs") still applies — auto-sent mail/events get the same signature `draft_email`/`draft_event` do.
- No new UI surface beyond the mandatory confirmation styling (Task 0 #2) and the mode default/label/description changes below — no "automation dashboard," no history log; the existing chat transcript is the record.

---

### Task 1: Give Outlook the Draft-only tier and make it the default (all four apps)

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts`

- [ ] **Step 1: Add `trackChanges` to Outlook's `availableModes`.** `entry.ts:351` currently reads `availableModes: ['readOnly', 'fullAutonomy']`; change to `['readOnly', 'trackChanges', 'fullAutonomy']`.

- [ ] **Step 2: Relabel it for Outlook specifically.** `modeTrackChanges`/`modeTrackChangesDesc` (`chat-ui.ts:30-31`, "Track changes" / "Edits as reviewable revisions") is shared text and fits Word/Excel/PowerPoint, not mail. Extend `ChatUIOptions` with an optional per-mode override (this generalizes what Round 2 already needed for Full Autonomy's description — do both in one mechanism):
  ```ts
  modeOverrides?: Partial<Record<EditingMode, { label: { en: string; he: string }; description: { en: string; he: string } }>>
  ```
  Wherever the mode menu renders a mode's label/description, check this map first, falling back to the shared `STRINGS` entries. Thread it through `bootstrap.ts`'s `AppConfig` (near `availableModes?`, `:335`) into `mountChatUI`.

- [ ] **Step 3: Outlook's overrides**, in `entry.ts`'s `startAddIn(...)` call:
  ```ts
  modeOverrides: {
    trackChanges: {
      label: { en: 'Draft only', he: 'טיוטות בלבד' },
      description: { en: 'Drafts emails and events for you to review and send yourself - nothing sends automatically.', he: 'מכין טיוטות הודעות ואירועים לבדיקה ושליחה על ידך - שום דבר לא נשלח אוטומטית.' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Edits applied directly - including sending emails and calendar invites in your name.', he: 'עריכות מוחלות ישירות - כולל שליחת הודעות והזמנות יומן בשמך.' },
    },
  },
  ```

- [ ] **Step 4: Make Draft-only the default, uniformly across all four apps.** `chat-ui.ts:423`'s `const defaultMode: EditingMode = menuModes.indexOf('fullAutonomy') !== -1 ? 'fullAutonomy' : menuModes[menuModes.length - 1]` and `bootstrap.ts:351`'s `let editingMode: EditingMode = 'fullAutonomy'` both currently start every fresh session already in Full Autonomy. Add one shared helper (`chat-ui.ts`, exported, next to `resolveModes`):
  ```ts
  export function defaultModeFor(modes: EditingMode[]): EditingMode {
    return modes.includes('trackChanges') ? 'trackChanges' : modes[0]
  }
  ```
  Now that Outlook offers `trackChanges` too (Step 1), this resolves the same way for all four apps — no per-app special-casing needed. Use it at both call sites: `chat-ui.ts:423` (`defaultModeFor(menuModes)`) and `bootstrap.ts:351` (`defaultModeFor(resolveModes(config.availableModes))`, importing both helpers from `chat-ui.ts` — `resolveModes` already exists there, `chat-ui.ts:236-238`). **Both sites must change together** — they're independent state today and must agree on what's selected before the user touches anything.

**Verification:**
- [ ] `tsc --noEmit` clean; all four bundles rebuild.
- [ ] `npx vitest run` in `shared/chat-ui` — add a test asserting `defaultModeFor(['readOnly','commentOnly','trackChanges','fullAutonomy'])` and `defaultModeFor(['readOnly','trackChanges','fullAutonomy'])` both return `'trackChanges'`.
- [ ] Manual, each app: a fresh pane opens with Draft only/Track changes selected, not Full Autonomy. Outlook's mode menu shows "Draft only" and "Full autonomy" with the extended copy, in both languages; Word/Excel/PowerPoint are unchanged.

---

### Task 2: Client-side tool gating for the new tier

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Add `fullAutonomyOnlyTools?: string[]` to `AppConfig` (near `readOnlyTools`).
- [ ] **Step 2:** In `bootstrap.ts`'s `availableForMode()` (`:356-360`):
  ```ts
  const fullAutonomyOnlySet = new Set(config.fullAutonomyOnlyTools ?? [])

  function availableForMode(): string[] {
    if (editingMode === 'readOnly') return config.tools.filter((t) => readOnlySet.has(t.name)).map((t) => t.name)
    if (editingMode === 'commentOnly') return config.tools.filter((t) => commentOnlySet.has(t.name)).map((t) => t.name)
    if (editingMode === 'trackChanges') return config.tools.filter((t) => !fullAutonomyOnlySet.has(t.name)).map((t) => t.name)
    return config.tools.map((t) => t.name)
  }
  ```
  An empty `fullAutonomyOnlySet` (every app but Outlook) makes the `trackChanges` branch identical to today's fallthrough — verified no behavior change for Word/Excel/PowerPoint.
- [ ] **Step 3:** In Outlook's `entry.ts`, set:
  ```ts
  fullAutonomyOnlyTools: ['accept_meeting', 'decline_meeting', 'send_email', 'send_reply', 'send_reply_all', 'send_forward', 'create_event'],
  ```
  (contingent on Task 0 #4's confirmation — if declined, drop `accept_meeting`/`decline_meeting` from this list and leave them where they are today).

**Verification:** `tsc --noEmit` clean; bundle rebuilds. Manual, Outlook: in Draft only mode, `draft_email`/`move_email`/etc. are offered but `send_email`/`accept_meeting` are not; switching to Full Autonomy reveals them.

---

### Task 3: Server-side three-tier gate

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.cs`

- [ ] **Step 1:** Add the mirrored server-side set:
  ```csharp
  private static readonly HashSet<string> FullAutonomyOnlyTools = new HashSet<string>
  {
      "accept_meeting", "decline_meeting", "send_email", "send_reply", "send_reply_all", "send_forward", "create_event",
  };
  ```
  (drop `accept_meeting`/`decline_meeting` if Task 0 #4 is declined — keep this list byte-for-byte in sync with `entry.ts`'s `fullAutonomyOnlyTools`, and say so in a comment on both sides.)
- [ ] **Step 2:** Replace `ExecuteAsync`'s mode check (`:41-50`):
  ```csharp
  if (!AlwaysAllowedTools.Contains(name))
  {
      // FullAutonomyOnlyTools always needs FullAutonomy; everything else needs at least TrackChanges.
      bool allowed = FullAutonomyOnlyTools.Contains(name)
          ? mode == EditingMode.FullAutonomy
          : (mode == EditingMode.FullAutonomy || mode == EditingMode.TrackChanges);
      if (!allowed)
      {
          return new ToolResult
          {
              Output = FullAutonomyOnlyTools.Contains(name)
                  ? "Blocked: this action sends something immediately and requires Full autonomy mode."
                  : "Blocked: the assistant is in a read-only mode. Switch to Draft only or Full autonomy to send drafts, move, delete, flag, create tasks/reminders, or respond to invites.",
              IsError = true,
              Summary = name,
          };
      }
  }
  ```

**Verification:** `dotnet build`/MSBuild clean. Manual: in Draft only mode, `draft_email` succeeds and `send_email` is blocked with the new message; in Full Autonomy, both succeed.

---

### Task 4: C# — the send tools themselves

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Compose.cs`
- Modify: `OutlookAiAddIn/OutlookTools.cs` (dispatcher cases only)

- [ ] **Step 1: `send_email`** — same body as `DraftEmail` (`OutlookTools.Compose.cs:28-40`) but `m.Send()` instead of `m.Display(false)`; result text states who it was sent to and the subject (this text is exactly what Task 6's distinct styling wraps).
- [ ] **Step 2: `send_reply` / `send_reply_all`** — same as `ReplyEmail` (`:42-55`) but `.Send()` instead of `.Display(false)`.
- [ ] **Step 3: `send_forward`** — same as `ForwardEmail` (`:57-71`) but `.Send()`.
- [ ] **Step 4: `create_event`** — same as `DraftEvent` (`:73-96`) but calls `a.Save()` for a plain appointment, or (attendees present, `MeetingStatus = olMeeting`) `a.Send()` to dispatch the invite — **confirm via reflection against the referenced Outlook PIA** which call is correct for a meeting item before assuming `.Send()` mirrors `.Save()`'s non-meeting case; do not guess.
- [ ] **Step 5:** wire all five into `ExecuteAsync`'s switch (`:52-84`). Each handler's `ToolResult.Mutated = true`.

**Verification:** `dotnet build`/MSBuild clean. Manual (real mailbox, low-stakes test recipient): in Full Autonomy, an email actually sends with no compose window; in Draft only, the same request opens a draft instead (model should reach for `draft_email`, since `send_email` isn't even offered there).

---

### Task 5: TypeScript — tool schemas and display strings

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] Add `send_email`/`send_reply`/`send_reply_all`/`send_forward`/`create_event` to `ALL_OUTLOOK_TOOLS`, same input shapes as their `draft_*`/`reply_email`/etc. counterparts, with descriptions stating plainly they send/create **immediately, with no review step**.
- [ ] Add each to `OUTLOOK_TOOL_DISPLAY`, with label text signaling "sends immediately" (e.g. an "(auto-send)" suffix in both languages).
- [ ] **Do not** add them to `readOnlyTools`. They're gated by Task 2/3's `fullAutonomyOnlyTools` mechanism instead.
- [ ] Extend the system prompt (`entry.ts:324-331`) to describe all three tiers now available (Read-only / Draft only / Full autonomy) and that the five send tools plus accept/decline exist only in Full Autonomy — prefer `draft_*` otherwise.

**Verification:** `tsc --noEmit` clean; bundle rebuilds; the five tools appear only in Full Autonomy.

---

### Task 6: Distinct "sent automatically" UI treatment (Task 0 #2)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/chat-ui/chat-ui.test.ts`

- [ ] Give the tools in `fullAutonomyOnlyTools` a distinct visual marker in the completed-step rendering (e.g. a filled/warning-colored icon instead of the plain checkmark other steps get) — driven by tool name (the app passes its own list down, same shape as `readOnlyTools`), not a new field threaded through the result pipeline.

**Verification:** `npx vitest run` passes including a new test (a full-autonomy-only tool step renders with the marker; a `draft_*` step does not); manual check in a live session.

---

### Task 7: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] Document the three-tier model for Outlook (Read-only / Draft only / Full autonomy) replacing the current binary description, including the `accept_meeting`/`decline_meeting` reclassification and why (they already `resp.Send()`).
- [ ] New subsection for the five auto-send tools, contrasted with "Draft-and-display tools" (`:508-519`) — correct the blanket "never call `.Send()`" claim and the "Excluded / deferred" section's `send_email`/`create_event` listing.

**Verification:** doc accurately reflects the shipped behavior once Tasks 1-6 land.
