# Outlook `open_email` Tool — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give the model a way to open a specific email in Outlook's own reading UI — the native message window, not a draft — so it can surface "here's the email" after a `list_emails`/`search_emails`/`get_email` call, without the user having to go find it themselves. Read-only: it never mutates the item.

**Architecture:** Outlook already has the exact mechanism this needs, one layer down. `OutlookTools.Compose.cs:34-38`'s `DraftEmail` calls `m.Display(false)` on a **new** `MailItem` it just created. `open_email` is the same `.Display(false)` call on an **existing** item resolved the same way every other by-id tool resolves one: `ItemById(entryId, storeId)` (`OutlookTools.cs:99-103`), exactly as `GetEmail` does at the top of `OutlookTools.Mail.cs`. No new COM surface, no new pattern — just wiring an existing capability (`Display`) onto an existing resolution path (`ItemById`) for a new verb (open, not read/draft).

Because it only calls `.Display()` and never touches a property, it belongs in `AlwaysAllowedTools` (`OutlookTools.cs:30-34`) and `readOnlyTools` (`OutlookAiAddIn/web-src/entry.ts`'s `startAddIn` config) alongside `get_email` — available in Read-only mode, not gated behind Full Autonomy.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools.Mail.cs`, `OutlookTools.cs`), TypeScript (`OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- Read-only: `Mutated` stays unset/false on the `ToolResult`, matching `GetEmail`'s pattern, not `DraftEmail`'s.
- Must work for `folder` values the same way `get_email`/`reply_email` do — `StoreOf(input)` for the store id, per the existing `ItemById(id, StoreOf(input))` call shape used in `OutlookTools.Compose.cs:48`.
- No signature, no body edit, no window is pre-filled with anything — this opens the message exactly as Outlook would if the user double-clicked it.
- Build/bundle command: see `docs/superpowers/plans/STATUS.md`'s "Build commands" section (esbuild for `entry.ts`, MSBuild for the C# project — MSBuild requires the full Visual Studio toolchain, not available in every environment; `dotnet build`/`dotnet test` only covers `OfficeAi.Shared`).

---

### Task 1: C# — `open_email` handler

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Mail.cs` (add the handler near `GetEmail`)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (dispatcher + `AlwaysAllowedTools`)

- [ ] **Step 1: Add the handler**

```csharp
private static ToolResult OpenEmail(JsonElement input)
{
    string id = ReqStr(input, "message_id");
    Outlook.MailItem m = ItemById(id, StoreOf(input)) as Outlook.MailItem;
    if (m == null)
        return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "open_email" };

    m.Display(false);
    return new ToolResult { Output = "Opened \"" + (m.Subject ?? "") + "\" in Outlook.", Summary = "open_email" };
}
```

Place it directly below `GetEmail` in `OutlookTools.Mail.cs` — same file as the tool it most resembles.

- [ ] **Step 2: Wire the dispatcher**

In `OutlookTools.cs`'s `ExecuteAsync` switch (`:54-63`, the read-only block), add:
```csharp
case "open_email": return OpenEmail(input);
```

- [ ] **Step 3: Add to `AlwaysAllowedTools`** (`OutlookTools.cs:30-34`) — it's read-only, so it must work in Read-only mode too, not just Full Autonomy.

**Verification:** `dotnet build` (or full MSBuild) on `OutlookAiAddIn` succeeds. Manual: call `open_email` with a real `message_id` from a prior `list_emails` result — Outlook opens that message's own reading window, unmodified.

---

### Task 2: TypeScript — tool schema, display strings, read-only registration

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Add the schema entry to `ALL_OUTLOOK_TOOLS`**, directly after `get_email` (`entry.ts:48-53`):

```ts
{
  name: 'open_email',
  description: 'Opens a message in its own Outlook reading window (like double-clicking it) so the user can see it directly. Read-only.',
  inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, folder: FOLDER }, required: ['message_id'] },
},
```

- [ ] **Step 2: Add to `OUTLOOK_TOOL_DISPLAY`** (`entry.ts:291` onward), after `get_email`'s entry:

```ts
open_email: d('Open email', 'פתיחת הודעה', 'Opens a message in Outlook.', 'פותח הודעה ב-Outlook.'),
```

- [ ] **Step 3: Add `'open_email'` to the `readOnlyTools` array** (`entry.ts:337-348`), after `'get_email'`.

- [ ] **Step 4: Mention it in the system prompt** — extend the existing "read and search mail" sentence (`entry.ts:325-331`) to note the model can open a message directly in Outlook, not just read its contents inline.

**Verification:** `tsc --noEmit` clean; bundle rebuilds (esbuild command in STATUS.md); the tool shows up with its display strings in both languages in the running add-in's tool list/UI.

---

### Task 3: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] Add `open_email` to the Outlook read-only tools table (next to `get_email`), noting it calls `.Display(false)` on an existing item and is `Mutated: false`.

**Verification:** doc reads correctly against the implemented behavior (self-check, no build step).
