# Outlook Auto-Send Mode (send email / create event without a review step) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-requested — "a way to create a complete event / send an email (maybe simply on full automation settings)."

**Goal:** Let the model actually send an email or create/send a calendar event end-to-end, with no native Outlook compose/appointment window for the user to review first.

**This is a genuine reversal of a standing safety invariant, not an additive feature.** Every draft/compose tool in this codebase carries the same comment: *"Draft-and-display only. These NEVER call `.Send()` - a native Outlook compose/appointment window is opened for the user to review and send"* (`OutlookAiAddIn/OutlookTools.Compose.cs:11-12`), and `docs/ai-tool-surface.md:518-519` states it as a documented guarantee. This plan narrows that guarantee for the most-permissive Outlook tier specifically, rather than removing it.

## Revision history

**Round 1 → 2:** dropped a proposed second, separately-toggled `autoSendEnabled` gate — Full Autonomy alone was made the single gate for the five new send tools.

**Round 2 → 3:** repurposed the unused `TrackChanges` mode as an Outlook-specific "Draft only" tier, sitting between Read-only and Full Autonomy — draft/mutate freely, never send. Surfaced a finding: `accept_meeting`/`decline_meeting` (`OutlookTools.Calendar.cs:195-224`) already call `resp.Send()` to notify the organizer, so they're an existing case of auto-sending; round 3 asked whether they should move into the same Full-Autonomy-only bucket as the five new send tools.

**Round 3 → 4 (this revision):** user proposal — don't lump `accept_meeting`/`decline_meeting` in with the five send tools at all; give them their **own**, third tier: "Automate approvals," between Draft only and Full autonomy. This **resolves round 3's open question by construction** rather than by picking yes/no — approvals get their own dedicated permission level instead of being folded into either neighbor.

**The clean part: this needs zero changes to the shared `EditingMode` enum.** It already has exactly four values (`OfficeAi.Shared/PaneHostBase.cs:13`: `ReadOnly, CommentOnly, TrackChanges, FullAutonomy`), declared in that escalating order. Outlook was only using two of them (`ReadOnly`, `FullAutonomy`) — round 3 put `TrackChanges` to work as "Draft only"; this round puts the **fourth, still-unused slot, `CommentOnly`**, to work too. Word/Excel/PowerPoint's own use of `CommentOnly` (their real "add comments, no edits" reviewer role) and `TrackChanges` (real edit-with-tracking) is completely untouched — this is all Outlook-side relabeling and Outlook-side tool-list config, reusing infrastructure that already exists for a different purpose in the other three apps.

## The four tiers

| # | `EditingMode` | Label (Outlook) | Outlook tools available |
|---|---|---|---|
| 1 | `readOnly` | Read only | `AlwaysAllowedTools` — `list_emails`, `search_emails`, `get_email`, `list_folders`, `search_contacts`, `list_events`, `get_event`, `list_tasks`, `get_attachment`, `find_meeting_slots` (`OutlookTools.cs:30-34`). Unchanged. |
| 2 | `commentOnly` | **Draft only** | Tier 1, plus every tool that mutates the mailbox or opens a draft but never leaves it unreviewed: `mark_email_read`/`unread`, `flag_email_important`, `move_email`, `delete_email`, `create_task`, `update_task`, `set_reminder`, `set_email_reminder`, `draft_email`, `reply_email`, `reply_all_email`, `forward_email`, `draft_event`. |
| 3 | `trackChanges` | **Automate approvals** | Tier 2, plus `accept_meeting`/`decline_meeting` — these already auto-notify the organizer via `resp.Send()`, so they get their own explicit opt-in instead of hiding in either neighbor. |
| 4 | `fullAutonomy` | Full autonomy | Tier 3, plus the five new tools that compose and send/create new content with no review: `send_email`, `send_reply`, `send_reply_all`, `send_forward`, `create_event`. |

Each tier is a strict superset of the one before it — this is a ladder, not four independent allowlists.

---

**Architecture:** Reuses the exact mechanism Word/Excel/PowerPoint's `commentOnly` already has (`bootstrap.ts:354`: `commentOnlySet = new Set([...readOnlyTools, ...(commentOnlyExtraTools ?? [])])`), and extends it one rung further for the new `trackChanges`-as-tier-3 case:

```ts
const commentOnlySet = new Set([...config.readOnlyTools, ...(config.commentOnlyExtraTools ?? [])])
const trackChangesSet = config.trackChangesExtraTools
  ? new Set([...commentOnlySet, ...config.trackChangesExtraTools])
  : null   // null = "no opinion", falls through to today's "everything" behavior

function availableForMode(): string[] {
  if (editingMode === 'readOnly') return config.tools.filter((t) => readOnlySet.has(t.name)).map((t) => t.name)
  if (editingMode === 'commentOnly') return config.tools.filter((t) => commentOnlySet.has(t.name)).map((t) => t.name)
  if (editingMode === 'trackChanges' && trackChangesSet) return config.tools.filter((t) => trackChangesSet.has(t.name)).map((t) => t.name)
  return config.tools.map((t) => t.name)  // fullAutonomy always; trackChanges too when trackChangesExtraTools is unset (Word/Excel/PowerPoint - unchanged)
}
```

`commentOnlyExtraTools`/`trackChangesExtraTools` are both **optional**, per-app fields that already exist in spirit (`commentOnlyExtraTools` literally already exists — `trackChangesExtraTools` is the one new field, this plan's only client-side schema change). Neither is set by Word/Excel/PowerPoint's `entry.ts` today, so their behavior is provably unchanged: `commentOnlySet` for them still means just `readOnlyTools` (their actual comment-only tools live wherever they already configure `commentOnlyExtraTools` today — unaffected either way, this plan doesn't touch their files), and `trackChangesSet` stays `null`, falling through to "everything," exactly as now.

**Server-side, `OutlookTools.ExecuteAsync`** (`:36-50`) mirrors this as an ordinal check — convenient because the enum's declared order already **is** the escalation order (`ReadOnly < CommentOnly < TrackChanges < FullAutonomy`, i.e. `(int)mode` increases with permission):

```csharp
private static readonly HashSet<string> DraftTierTools = new HashSet<string> { /* the 13 tools in row 2 above */ };
private static readonly HashSet<string> ApprovalTierTools = new HashSet<string> { "accept_meeting", "decline_meeting" };
private static readonly HashSet<string> SendTierTools = new HashSet<string> { "send_email", "send_reply", "send_reply_all", "send_forward", "create_event" };

// inside ExecuteAsync, replacing the current binary check:
if (!AlwaysAllowedTools.Contains(name))
{
    EditingMode required =
        SendTierTools.Contains(name) ? EditingMode.FullAutonomy :
        ApprovalTierTools.Contains(name) ? EditingMode.TrackChanges :
        EditingMode.CommentOnly; // DraftTierTools, and anything else not otherwise classified
    if ((int)mode < (int)required)
    {
        return new ToolResult { Output = "Blocked: ...", IsError = true, Summary = name };
    }
}
```

**Default mode:** stays the least-permissive *non-read-only* tier, per app — Word/Excel/PowerPoint keep `trackChanges` as their default (unchanged from round 3; their own real tier 2/3 split is a separate concern, out of scope here). Outlook's own default becomes `commentOnly` ("Draft only") — its least-permissive non-read-only tier under this new 4-tier scheme. Generalize the round-3 `defaultModeFor` helper to take an explicit preference instead of hardcoding `trackChanges`:

```ts
export function defaultModeFor(modes: EditingMode[], preferred: EditingMode = 'trackChanges'): EditingMode {
  return modes.includes(preferred) ? preferred : modes[0]
}
```

Word/Excel/PowerPoint call it with no second argument (unchanged). Outlook's `entry.ts` passes its own preferred default (`'commentOnly'`) through a new optional `AppConfig.defaultMode?: EditingMode` field, threaded into the `bootstrap.ts:351` call site alongside the existing `chat-ui.ts:423` one — both must still change together, per round 3.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools*.cs`), TypeScript (`shared/chat-ui/chat-ui.ts`, `shared/web-src/app-shell/bootstrap.ts`, `OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- Every new send-tool's description must say, in plain language, that it sends/creates immediately with **no review step** — prefer `draft_*` unless the user's own instruction this turn clearly wants immediate sending.
- Reuse every existing helper the `draft_*` tools already use (`AddAttendees`, `PrependHtml`, recipient resolution) — only the final `.Send()`/`.Save()` step and the tool name are new.
- The signature line ("— Created with OpenDocs") still applies — auto-sent mail/events get the same signature `draft_email`/`draft_event` do.
- No new UI surface beyond the mandatory confirmation styling (Task 0 below) and the mode label/default changes — no "automation dashboard," no history log; the existing chat transcript is the record.

## Task 0: Decisions (round 3's #4 is now moot — see above; these carry over unresolved)

| # | Question | Recommendation | Why |
|---|---|---|---|
| 1 | New distinct tools vs. a hidden branch inside `draft_*`? | **New distinct tools.** | Every existing tool's description stays permanently truthful. **Agreed round 1.** |
| 2 | Visible confirmation after an auto-send fires? | **Yes, mandatory.** | The native review window was the only place the user would see this before it's irreversible. **Agreed round 1.** |
| 3 | Scope: Outlook only? | **Yes.** | Word/Excel/PowerPoint have no "send" concept. **Agreed round 1.** |

---

### Task 1: Wire up the four tiers in Outlook's config

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] Step 1: `availableModes: ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']` (was `['readOnly', 'fullAutonomy']`).
- [ ] Step 2: `commentOnlyExtraTools: [ 'mark_email_read', 'mark_email_unread', 'flag_email_important', 'move_email', 'delete_email', 'create_task', 'update_task', 'set_reminder', 'set_email_reminder', 'draft_email', 'reply_email', 'reply_all_email', 'forward_email', 'draft_event' ]` — the full "Draft only" tool set (tier 2).
- [ ] Step 3: `trackChangesExtraTools: ['accept_meeting', 'decline_meeting']` — tier 3's addition on top of tier 2.
- [ ] Step 4: `defaultMode: 'commentOnly'`.
- [ ] Step 5: `modeOverrides` for all three non-`readOnly` labels (extends the mechanism `chat-ui.ts`/`bootstrap.ts` need to gain — see Task 2):
  ```ts
  modeOverrides: {
    commentOnly: {
      label: { en: 'Draft only', he: 'טיוטות בלבד' },
      description: { en: 'Drafts, flags, moves, and manages mail/tasks for your review - nothing sends.', he: 'מכין טיוטות, מסמן, מעביר ומנהל דואר/משימות לבדיקתך - שום דבר לא נשלח.' },
    },
    trackChanges: {
      label: { en: 'Automate approvals', he: 'אוטומציית אישורים' },
      description: { en: 'Everything in Draft only, plus auto-accepting/declining meeting invites (notifies the organizer).', he: 'כל מה שיש בטיוטות בלבד, בתוספת אישור/דחייה אוטומטיים של הזמנות לפגישה (מודיע למארגן).' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Everything above, plus sending emails and creating/sending calendar invites in your name.', he: 'כל מה שלמעלה, בתוספת שליחת הודעות ויצירה/שליחה של הזמנות יומן בשמך.' },
    },
  },
  ```

**Verification:** manual, once Task 2/3 land — Outlook's mode menu shows all four tiers with this copy, in order, defaulting to Draft only.

---

### Task 2: Client-side — the new `trackChangesExtraTools` rung, `defaultMode`, and `modeOverrides`

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts`

- [ ] Step 1: Add `trackChangesExtraTools?: string[]` to `AppConfig`, next to the existing `commentOnlyExtraTools?: string[]`.
- [ ] Step 2: In `bootstrap.ts`, add the `trackChangesSet` computation and updated `availableForMode()` exactly as shown in the Architecture section above (`:353-359`).
- [ ] Step 3: Add `defaultMode?: EditingMode` to `AppConfig`; generalize `defaultModeFor` as shown above (exported from `chat-ui.ts`, next to `resolveModes`); use it at both `chat-ui.ts:423` and `bootstrap.ts:351` (both sites change together, per round 3's own note on this).
- [ ] Step 4: Add `modeOverrides?: Partial<Record<EditingMode, { label: {en,he}; description: {en,he} }>>` to `ChatUIOptions`/`AppConfig`; wherever the mode menu renders a label/description, check this map before the shared `STRINGS` entries (`modeReadOnly`, `modeCommentOnly`, `modeTrackChanges`, `modeFullAutonomy` and their `*Desc` counterparts, `chat-ui.ts:26-33`).

**Verification:**
- [ ] `tsc --noEmit` clean; all four bundles rebuild.
- [ ] `npx vitest run` in `shared/chat-ui` — new tests: `defaultModeFor(['readOnly','trackChanges','fullAutonomy'])` still returns `'trackChanges'` (Word/Excel/PowerPoint unaffected); `defaultModeFor(['readOnly','commentOnly','trackChanges','fullAutonomy'], 'commentOnly')` returns `'commentOnly'`.
- [ ] Manual, Word: mode menu and default unchanged from before this plan.
- [ ] Manual, Outlook: four tiers, correct labels, correct tool availability at each (Draft only shows `draft_email` but not `accept_meeting` or `send_email`; Automate approvals adds `accept_meeting`/`decline_meeting` but still not `send_email`; Full autonomy shows everything).

---

### Task 3: Server-side four-tier gate

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.cs`

- [ ] Step 1: Add `DraftTierTools`, `ApprovalTierTools`, `SendTierTools` (the three `HashSet<string>`s from the Architecture section — keep `DraftTierTools`/`ApprovalTierTools`/`SendTierTools` byte-for-byte in sync with `entry.ts`'s `commentOnlyExtraTools`/`trackChangesExtraTools`/the five new tools respectively, and say so in a comment on both sides).
- [ ] Step 2: Replace `ExecuteAsync`'s binary mode check (`:41-50`) with the ordinal check shown in the Architecture section, with a per-tier blocked message (name the tier the user needs to switch to).

**Verification:** `dotnet build`/MSBuild clean. Manual, one call per tier boundary: `draft_email` blocked in Read-only, succeeds in Draft only; `accept_meeting` blocked in Draft only, succeeds in Automate approvals; `send_email` blocked in Automate approvals, succeeds in Full autonomy.

---

### Task 4: C# — the send tools themselves

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Compose.cs`
- Modify: `OutlookAiAddIn/OutlookTools.cs` (dispatcher cases only)

- [ ] Step 1: `send_email` — same body as `DraftEmail` (`OutlookTools.Compose.cs:28-40`) but `m.Send()` instead of `m.Display(false)`; result text states who it was sent to and the subject (feeds Task 6's distinct styling).
- [ ] Step 2: `send_reply` / `send_reply_all` — same as `ReplyEmail` (`:42-55`) but `.Send()`.
- [ ] Step 3: `send_forward` — same as `ForwardEmail` (`:57-71`) but `.Send()`.
- [ ] Step 4: `create_event` — same as `DraftEvent` (`:73-96`) but calls `a.Save()` for a plain appointment, or (attendees present, `MeetingStatus = olMeeting`) `a.Send()` to dispatch the invite — **confirm via reflection against the referenced Outlook PIA** which call is correct for a meeting item; do not guess.
- [ ] Step 5: wire all five into `ExecuteAsync`'s switch (`:52-84`). Each handler's `ToolResult.Mutated = true`.

**Verification:** `dotnet build`/MSBuild clean. Manual (real mailbox, low-stakes test recipient): Full Autonomy sends with no compose window; Automate approvals opens a draft instead (model reaches for `draft_email`, since `send_email` isn't offered).

---

### Task 5: TypeScript — tool schemas and display strings

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] Add `send_email`/`send_reply`/`send_reply_all`/`send_forward`/`create_event` to `ALL_OUTLOOK_TOOLS`, descriptions stating plainly they send/create **immediately, with no review step**.
- [ ] Add each to `OUTLOOK_TOOL_DISPLAY`, label text signaling "sends immediately" (e.g. an "(auto-send)" suffix, both languages).
- [ ] **Do not** add them to `commentOnlyExtraTools` or `trackChangesExtraTools` — that omission is what confines them to Full Autonomy.
- [ ] Extend the system prompt (`entry.ts:324-331`) to describe all four tiers and that the five send tools exist only in Full Autonomy.

**Verification:** `tsc --noEmit` clean; bundle rebuilds; the five tools appear only in Full Autonomy.

---

### Task 6: Distinct "sent automatically" UI treatment (Task 0 #2)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/chat-ui/chat-ui.test.ts`

- [ ] Give the five send-tier tool names a distinct visual marker in the completed-step rendering (e.g. a filled/warning-colored icon instead of the plain checkmark other steps get) — driven by tool name, not a new field threaded through the result pipeline. (`accept_meeting`/`decline_meeting` also auto-send, but that's pre-existing, unchanged behavior — out of scope to restyle here unless you want it included.)

**Verification:** `npx vitest run` passes including a new test; manual check in a live session.

---

### Task 7: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] Document the four-tier model for Outlook, replacing the current binary description.
- [ ] New subsection for the five auto-send tools, contrasted with "Draft-and-display tools" (`:508-519`) — correct the blanket "never call `.Send()`" claim and the "Excluded / deferred" section's `send_email`/`create_event` listing.

**Verification:** doc accurately reflects the shipped behavior once Tasks 1-6 land.
