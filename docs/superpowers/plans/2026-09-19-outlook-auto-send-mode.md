# Outlook Auto-Send Mode (send email / create event without a review step) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** user-requested — "a way to create a complete event / send an email (maybe simply on full automation settings)."

**Goal:** Let the model actually send an email or create/send a calendar event end-to-end, with no native Outlook compose/appointment window for the user to review first — as an explicit, separately-gated opt-in, not an automatic consequence of switching to Full Autonomy.

**This is a genuine reversal of a standing safety invariant, not an additive feature.** Every draft/compose tool in this codebase carries the same comment: *"Draft-and-display only. These NEVER call `.Send()` - a native Outlook compose/appointment window is opened for the user to review and send"* (`OutlookAiAddIn/OutlookTools.Compose.cs:11-12`), and `docs/ai-tool-surface.md:518-519` states it as a documented guarantee: *"These never call `.Send()` (mail) or save a calendar event. The user sends from the opened Outlook window."* This plan proposes narrowing, not removing, that guarantee — see the Task 0 decision gate below before writing any code.

## Task 0: Decisions needed before implementation (recommended defaults below; needs explicit user sign-off, not just an implementer's best guess — same posture `docs/superpowers/plans/STATUS.md`'s own "Decisions" table takes for its two product-owner gates)

| # | Question | Recommended default | Why |
|---|---|---|---|
| 1 | New distinct tools (`send_email`, `send_reply`, `send_reply_all`, `send_forward`, `create_event`) vs. a hidden branch inside the existing `draft_*` tools? | **New distinct tools.** | Keeps every existing tool's description permanently truthful ("draft_email never sends" stays true in every mode, forever) and makes gating a tool-list-membership question (reusing the exact mechanism `readOnlyTools`/`commentOnlyExtraTools` already use), not a runtime behavior switch buried inside a tool the model was told is safe. |
| 2 | Global on/off toggle, or a recipient allow-list (e.g. only auto-send to known contacts, never a new/external address)? | **Global toggle for v1.** Note the allow-list as a real, documented future enhancement — do not build it now (unscoped speculative complexity for a feature with zero live usage yet). | Matches "don't design for hypothetical requirements" — ship the simple version, learn from real use, narrow later if it proves necessary. |
| 3 | Does enabling auto-send require Full Autonomy mode to already be selected, or is it independent? | **Requires Full Autonomy**, checked at the moment a send-tool runs (not just at toggle time) — same posture as every other mutating tool. | Auto-send is strictly more permissive than Full Autonomy's existing mutations (irreversible, externally visible), so it must never be reachable from a less-trusted mode. |
| 4 | Visible confirmation after an auto-send fires? | **Yes, mandatory** — the chat UI must show a distinct, unmissable line ("Sent to X about Y" / "Created and sent invite to Z") styled differently from a normal "opened a draft" tool result. | Removing the native review window removes the *only* place the user would otherwise see this happened before it's irreversible. This is not optional polish. |
| 5 | Scope: Outlook only, or a pattern other apps could reuse later? | **Outlook only**, matching the literal request. | Word/Excel/PowerPoint have no equivalent "send" concept; forcing a shared abstraction now would be speculative. |
| 6 | Does turning off Full Autonomy (or closing/reopening the pane) revoke the auto-send toggle, or does it persist? | **Persists** (same storage tier as `skipTlsVerify`), but is **inert** whenever mode isn't Full Autonomy — Decision 3 already makes it unreachable outside Full Autonomy regardless of whether the stored flag is true. | Avoids forcing re-opt-in every session while keeping the actual safety boundary (Decision 3) as the only thing that matters. |

**Do not proceed past Task 0 without the user explicitly confirming or overriding each row** — these are product/safety calls, not implementation details.

---

**Architecture (assuming Task 0's defaults):** Two new, independent gates must both be true for a send-tool call to succeed, mirroring the existing `EditingMode` gate's own defense-in-depth posture (`bootstrap.ts`'s comment at `:370-373`: filtering happens both client-side in the active tool list *and* is re-checked structurally):

1. **Client-side tool-list gating** (which tools the model can even see/call this turn) — extends `bootstrap.ts`'s existing `availableForMode()` (`:356-360`), which already computes the tool set from `editingMode`; add a second filter for a new `config.autoSendTools: string[]` list, included only when a new persisted `autoSendEnabled` setting is true **and** `editingMode === 'fullAutonomy'`.
2. **Server-side re-check in `OutlookTools.ExecuteAsync`** (`OutlookTools.cs:36-50`) — the existing mode gate already blocks non-`AlwaysAllowedTools` calls outside Full Autonomy; add a second `AutoSendTools` `HashSet` and a second check requiring a new per-mailbox `AutoSendByMailbox` flag (mirroring `ModeByMailbox`, `OutlookTools.cs:15-26`) to be true, returning the same shape of "Blocked" `ToolResult` the mode gate already uses (`:44-50`) when it isn't. **This is the actual guarantee** — the client-side filter is a UX nicety, exactly as Task 4 of `2026-08-23-pp05-gateway-tool-schemas.md` already established for schema validation ("the schema is advisory... the actual guarantee lives" server-side).

The toggle itself flows the same path `editingMode` already does: `PanelSettings` (client, `shared/web-src/app-shell/settings.ts`) → a new WebView2 bridge message (mirroring `postMode`/`"set-mode"`, `bridge.ts:194-196` → `PaneHostBase.cs`'s `OnOtherMessage` switch, `:160-169`) → a new abstract `SetAutoSend(bool)` alongside the existing `SetEditingMode` (`PaneHostBase.cs:138`) → `OutlookTools.SetAutoSend(mbxKey, bool)` alongside the existing `SetMode` (`OutlookTools.cs:17-20`).

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools*.cs`, `OfficeAi.Shared/PaneHostBase.cs`), TypeScript (`shared/web-src/app-shell/{bootstrap,bridge,settings}.ts`, `shared/chat-ui/chat-ui.ts`, `OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- Every new send-tool's description must say, in plain language, that it sends/creates immediately with **no review step** — the model needs to know this is categorically different from `draft_email` so it doesn't reach for it casually (e.g. still prefer `draft_event` unless the user's own instruction clearly asked for full automation this turn, per the system prompt addition in Task 5).
- Reuse every existing helper the `draft_*` tools already use (`AddAttendees`, `PrependHtml`, recipient resolution) — only the final `.Send()`/`.Save()` step and the tool name/gating are new.
- The signature line (`— Created with OpenDocs`, see the branding rebrand work) still applies — auto-sent mail/events get the same signature `draft_email`/`draft_event` do.
- No new UI surface beyond one settings checkbox + the mandatory confirmation styling (Task 0 decisions 1-4) — do not build a separate "automation dashboard" or history log; the existing chat transcript already is the record (each send shows as a distinctly-styled tool result in it).

---

### Task 1: Settings plumbing — `autoSendEnabled`

**Files:**
- Modify: `shared/web-src/app-shell/settings.ts` (add `autoSendEnabled: boolean` to `PanelSettings`, default `false`, same load/persist shape as `skipTlsVerify` at `:19,61,68,72`)
- Modify: `shared/chat-ui/chat-ui.ts` (new checkbox in the settings panel, same DOM pattern as `[data-field="skipTlsVerify"]`; label text must state the risk plainly, e.g. "Auto-send emails and events (skips your review — sends/creates immediately). Requires Full Autonomy.")
- Modify: `shared/chat-ui/chat-ui.test.ts` (mirror the two existing `skipTlsVerify` tests: default unchecked, reports state on Save)

- [ ] **Step 1-3:** implement per the `skipTlsVerify` template cited above.
- [ ] **Step 4:** the checkbox is visually de-emphasized/disabled (not hidden) when the current editing mode isn't Full Autonomy, with a short inline note why — consistent with Decision 3, and stops a confused "why doesn't this do anything" report.

**Verification:** `npx vitest run` in `shared/chat-ui` passes including the two new tests.

---

### Task 2: Bridge — propagate the toggle to .NET

**Files:**
- Modify: `shared/web-src/app-shell/bridge.ts` (new `postAutoSend(enabled: boolean)`, mirroring `postMode` at `:194-196`)
- Modify: `shared/web-src/app-shell/bootstrap.ts` (call it from the settings-save handler alongside where `postMode`/theme are applied, and wire the new `config.autoSendTools` filter into `availableForMode()`, `:356-360`)
- Modify: `OfficeAi.Shared/PaneHostBase.cs` (new `case "set-auto-send":` in `OnOtherMessage`, `:160` onward, reading a boolean payload; new abstract `protected abstract void SetAutoSend(bool enabled);` alongside `SetEditingMode`, `:138`)
- Modify: `OutlookAiAddIn/TaskPaneHost.cs` (or wherever `SetEditingMode` is currently implemented for Outlook — implement the new abstract method, routing to `OutlookTools.SetAutoSend(mbxKey, enabled)`)

- [ ] Implement per the `set-mode`/`SetEditingMode` pattern cited above, end to end.

**Verification:** `tsc --noEmit` clean; `dotnet build`/MSBuild on `OutlookAiAddIn` and `OfficeAi.Shared` clean.

---

### Task 3: C# — the send tools themselves

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Compose.cs` (new handlers)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (new `AutoSendByMailbox` dict + `SetAutoSend`/`AutoSendFor`, mirroring `ModeByMailbox`/`SetMode`/`ModeFor` at `:15-26`; new `AutoSendTools` `HashSet`; dispatcher cases; the second gate check in `ExecuteAsync`)

- [ ] **Step 1: the gate**, added right after the existing mode-gate block in `ExecuteAsync` (`:41-50`):

```csharp
private static readonly HashSet<string> AutoSendTools = new HashSet<string>
{
    "send_email", "send_reply", "send_reply_all", "send_forward", "create_event",
};

// ... inside ExecuteAsync, after the existing mode check:
if (AutoSendTools.Contains(name) && !AutoSendFor(mbxKey))
{
    return new ToolResult
    {
        Output = "Blocked: auto-send is not enabled. Use draft_email/draft_event instead, or ask the user to enable auto-send in Settings.",
        IsError = true,
        Summary = name,
    };
}
```

- [ ] **Step 2: `send_email`** — same body as `DraftEmail` (`OutlookTools.Compose.cs:28-40`) but `m.Send()` instead of `m.Display(false)`; result text states who it was sent to and the subject (Task 0 Decision 4's confirmation content starts here — the *tool result itself* is what the chat UI renders distinctly, per Task 5).
- [ ] **Step 3: `send_reply` / `send_reply_all`** — same as `ReplyEmail` (`:42-55`) but `.Send()` on the reply item instead of `.Display(false)`.
- [ ] **Step 4: `send_forward`** — same as `ForwardEmail` (`:57-71`) but `.Send()`.
- [ ] **Step 5: `create_event`** — same as `DraftEvent` (`:73-96`) but calls `a.Save()` for a plain appointment, or (when attendees were added, `MeetingStatus = olMeeting`) `a.Send()` to actually dispatch the meeting invite rather than just opening it — confirm via reflection against the referenced Outlook PIA which call is correct for a meeting item before assuming `.Send()` works identically to `.Save()`'s non-meeting case (do not guess this one; it's exactly the kind of signature assumption this project's own history warns against, see `apply_search`'s Task 0 for the same discipline applied to a different uncertain API).
- [ ] **Step 6:** each handler's `ToolResult.Mutated = true` (unlike the `draft_*` tools, which never set it — these genuinely change external state).

**Verification:** `dotnet build`/MSBuild clean. Manual (real mailbox, low-stakes test recipient): with auto-send enabled and Full Autonomy active, ask for an email to be sent — confirm it actually lands in the recipient's inbox with no compose window ever appearing, and that the exact same request with auto-send OFF still only opens a draft. Repeat for `create_event` with and without attendees.

---

### Task 4: TypeScript — tool schemas, display strings, gating list

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] Add `send_email`/`send_reply`/`send_reply_all`/`send_forward`/`create_event` to `ALL_OUTLOOK_TOOLS`, same input shapes as their `draft_*`/`reply_email`/etc. counterparts, with descriptions that state plainly they send/create **immediately, with no review step**.
- [ ] Add each to `OUTLOOK_TOOL_DISPLAY`, with label text that visually signals "sends immediately" (e.g. an explicit "(auto-send)" suffix in both languages) — this is the tool-step chip the user sees in the transcript in real time, not just the final result text.
- [ ] Add a new `autoSendTools: string[]` array to the `startAddIn` config (new config field from Task 2) listing exactly these five tool names.
- [ ] Extend the system prompt (`entry.ts:324-331`) with one clear sentence: these tools exist only when the user has explicitly enabled auto-send, and even then the model should prefer them only when the user's request in this turn clearly wants immediate sending — default to `draft_*` otherwise.

**Verification:** `tsc --noEmit` clean; bundle rebuilds.

---

### Task 5: Distinct "sent automatically" UI treatment (Task 0 Decision 4)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts` (tool-step rendering — likely near wherever `.ai-applied-tag`/`.ai-step-icon` are set, per the existing tool-group rendering `chat-ui.test.ts` already exercises)
- Modify: `shared/chat-ui/chat-ui.test.ts` (new test: a step for one of the five auto-send tool names renders with the distinct marker/class; a step for `draft_email` does not)

- [ ] Give the five auto-send tool names a distinct visual marker in the completed-step rendering (e.g. a filled/warning-colored icon instead of the plain checkmark `draft_*` steps get) — driven by tool name, not a new field threaded through the whole result pipeline, keeping this a presentation-only change.

**Verification:** `npx vitest run` passes including the new test; manual check that a live auto-send call visibly stands out in the transcript from a `draft_*` call.

---

### Task 6: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] New subsection for the five auto-send tools, explicitly contrasted with the existing "Draft-and-display tools" section (`:508-519`) — state the dual gate (Full Autonomy + auto-send setting), and correct the current blanket claim ("These never call `.Send()`... nothing in this repo sends mail or creates calendar events on its own" — find and update the exact wording near `:518-523` and the "Excluded / deferred" section that currently lists `send_email`/`create_event` as deliberately unsupported) now that it's no longer true unconditionally.

**Verification:** doc accurately reflects the shipped gating, including the two-gate model, once Tasks 1-5 land.
