# FT-1: Full Settings Screen — Tool Registry and Document Guidelines

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source:** feature request, 2026-08-23. Not a PP item — this is new capability, not a defect fix.

**Goal:** A "More settings" button in the settings dropdown opens a full-panel settings screen *inside the Airchat Office frame* (replacing the chat view, not a separate OS window). It carries the connection and language fields, an edit-scope control, a list of every tool with a localized name and description and a per-tool registration toggle, and a free-text document system message that is injected into every new conversation for that document.

## Behavior contract (as specified, do not re-derive)

**Screen.** Same pane, different view. `.ai-chat` + `.ai-composer` are replaced by the settings view inside the existing `.ai-panel`; the header stays. While in the settings view, the header's gear button becomes a "back to current conversation" button.

**Tool registration.**
- Changing the edit scope — from the chat composer's mode menu *or* from the scope control on this screen — resets the registered set to that scope's default.
- After that, each toggle the user makes is retained until the next scope change or until the document is next opened.
- Registration is **unregister-within-scope**: the scope defines what is registerable; the user can switch tools off and back on within that set. A tool outside the current scope renders greyed and non-interactive, with a hover hint saying to change the scope — which is why the scope control lives on this screen.
- Not persisted between documents.

**Document system message.** Per document, persisted on disk beside the chat history so it survives closing and reopening the file. Injected into every **new** conversation; an in-flight conversation is not retroactively changed.

**Three different lifetimes — keep them straight, this is the main source of bugs in this feature:**

| Data | Lifetime | Storage |
|---|---|---|
| Connection settings, language | Global per app | `localStorage` (existing `airchat-settings`) |
| Document system message | Per document, survives reopen | **New** `DocSettingsStore`, keyed by the same id `ChatStore` uses |
| Tool registration overrides | Until scope change or document close | In-memory only, never written |

**Architecture:** Almost all of this lands in `shared/chat-ui/chat-ui.ts` + `.css`, which is already shared by all three apps — so it is built once. The tool list needs data the chat UI does not have (which tools exist, their localized labels, which are in scope), so it arrives through the mount options. The document system message needs a round trip to C#, which is two new bridge message kinds and one new store class.

One design point worth stating up front: **tool `description` fields in the schemas are the model's contract and stay English.** They are tuned for the model and translating them would change its behavior. The screen therefore renders a *separate*, UI-only localized label per tool. Supplying those labels for all ~40 tools across the three apps is the bulk of this feature's content work (Task 5).

**Tech Stack:** TypeScript + CSS in `shared/chat-ui/`; C# 7.3 / .NET Framework 4.8 for the per-document store; vitest/jsdom for tests.

## Dependencies

- **PP-0 (`2026-08-23-pp00-shared-app-shell.md`) first.** This feature adds per-app configuration (the tool display map) and new bridge messages; without the shell that is three copies of every wiring change. Task 5's map slots straight into `AddInConfig`.
- **PP-6 (`2026-08-23-pp06-provider-selection-wiring.md`) interacts.** Both touch the connection fields. If PP-6 lands first, this screen renders the provider/model selectors it introduces rather than the three plain text boxes; if this lands first, PP-6 updates both surfaces. Either order works — just do not build the connection section twice.
- PP-2/PP-3 touch the chat transcript, not the header or view switching. No conflict.

## Global Constraints

- Every user-visible string goes through `chat-ui.ts`'s `STRINGS` table with `en` and `he` entries. No hardcoded English.
- The settings view must respect the existing RTL machinery: `setLang` sets `dir` on `.ai-dock` and re-runs `applyStrings()`; anything added here must be reachable by `applyStrings`' `[data-t]` / `[data-t-title]` / `[data-t-placeholder]` sweep, or it will not relocalize on a language switch.
- Colors and spacing come from existing CSS custom properties in `shared/chat-ui/chat-ui.css` — no new raw hex outside a token definition line.
- **The document system message is untrusted user text that reaches the model's system prompt.** Render it with `textContent`/`value`, never `innerHTML`. Cap its length (Task 7 Step 4). It must never be written to a tool result or a log.
- C# 7.3 / .NET Framework 4.8 only in `.cs` — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Rebuild all three bundles and MSBuild all three projects after any `chat-ui.ts`/`chat-ui.css` change.
- Do not regress the existing settings dropdown. Its tested contract is that fields take effect only on Save, not on input (`shared/chat-ui/chat-ui.test.ts`) — this screen follows the same rule.

---

### Task 1: View switching

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

**Interfaces:**
- Produces: internal `setView('chat' | 'settings')`; `ChatUIHandle.openSettings()` for completeness.

- [ ] **Step 1: Markup.** Add a sibling `<div class="ai-settings-view">` between `.ai-chat` and `.ai-composer` in the template at `shared/chat-ui/chat-ui.ts:90-137`. Both views live in the DOM; the view state toggles a class on `.ai-panel`.

- [ ] **Step 2: CSS.** `.ai-settings-view { display: none; flex: 1; overflow-y: auto; padding: 16px; }`, shown via `.ai-panel.settings-open .ai-settings-view { display: block; }`, with `.ai-panel.settings-open .ai-chat`/`.ai-composer` hidden. Reusing the `flex: 1; overflow-y: auto` shape from `.ai-chat` (`:150`) keeps scrolling behavior identical.

- [ ] **Step 3: "More settings" entry point.** Add a button at the bottom of the existing `#settingsPanel` dropdown, below Save. Clicking it closes the dropdown and calls `setView('settings')`.

- [ ] **Step 4: Header swap.** In the settings view the gear button becomes back-to-conversation: swap its glyph and its `data-t-title` (`settings` → `backToChat`), and switch its click handler. Do this by toggling a class and reading the current view in one handler rather than rebinding listeners, which is easier to get wrong.

- [ ] **Step 5:** Hide the "+ new chat" button in the settings view — it acts on the conversation. Keep the collapse button working in both views.

- [ ] **Step 6: Preserve conversation state.** The chat view is hidden, never re-rendered — `assistantBubble` and any in-flight run must survive a round trip to settings and back. If a run finishes while the settings view is open, the transcript updates underneath and is simply there on return. Verify explicitly; do not clear `chatEl`.

- [ ] **Step 7: Strings** — `moreSettings`, `backToChat`, `settingsScreenTitle`, plus section headings `sectionConnection`, `sectionLanguage`, `sectionScope`, `sectionTools`, `sectionDocMessage`.

**Verification:** `cd shared/chat-ui && npx vitest run`; the view switches both ways and the transcript is intact on return.

---

### Task 2: Connection and language sections

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

- [ ] **Step 1:** Render the same fields the dropdown has — base URL, API key, model, skip-TLS, language toggle — in the settings view, as a labelled section.
- [ ] **Step 2: One source of truth.** The two surfaces must not drift. Read both from the same state on open, and have Save in either place write the same state and refresh the other. The simplest correct approach: a single `readSettingsForm(scope)` / `writeSettingsForm(scope)` pair over `[data-field]` selectors scoped to each container.
- [ ] **Step 3:** Keep the API-key input `type="password"` here too.
- [ ] **Step 4:** If PP-6 has landed, render its provider `<select>` and model control instead of the plain text boxes, and honor its `needsBaseUrl` visibility rule.
- [ ] **Step 5:** The language toggle in this view uses the existing `pendingLang` mechanism — language applies on Save, exactly as the dropdown does today (`chat-ui.ts:256-271`).

**Verification:** changing a value in one surface and saving shows the new value in the other.

---

### Task 3: Edit-scope control

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

- [ ] **Step 1:** Render the four modes (`readOnly`/`commentOnly`/`trackChanges`/`fullAutonomy`) with their existing `STRINGS` labels and descriptions — reuse `MODES` (`chat-ui.ts:66`) rather than re-listing them.
- [ ] **Step 2: Scope changes apply immediately**, unlike everything else on this screen. It mirrors the composer's mode menu, which is immediate, and — per the contract — it resets the tool registry, which the user needs to see happen. Everything else on the screen applies on Save.
- [ ] **Step 3: Make the asymmetry visible.** Label the section with a short note that the scope applies immediately and resets tool registration. Without that, a user who changes scope, toggles tools, then hits Save has no model of what just happened.
- [ ] **Step 4:** Selecting a mode here must drive the *same* path the composer's menu drives — update the composer's selected item and label, fire `options.onModeChange(mode)` once, and apply the `accent` class for `trackChanges` (`chat-ui.ts:277-285`). Extract that block into a shared `selectMode(mode, source)` so the two controls cannot diverge.
- [ ] **Step 5:** Conversely, changing the mode from the composer while the settings view is closed must still reset the tool registry — the reset belongs to the mode change, not to this screen.

**Verification:** changing scope from either control updates both and resets the tool list.

---

### Task 4: Tool registry model

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts` (from PP-0)

**Interfaces:**
- Produces: `ChatUIOptions.tools: ToolDisplayEntry[]` and `ChatUIOptions.onToolRegistrationChange(registered: string[]): void`; `ChatUIHandle.setToolScope(available: string[], registered: string[]): void`.

```ts
export interface ToolDisplayEntry {
  name: string                                  // the schema name, e.g. 'apply_commands'
  label: { en: string; he: string }             // short human name
  description: { en: string; he: string }       // one sentence, UI-only
}
```

- [ ] **Step 1: State in the shell, not the UI.** The registry has three inputs — all tools, the scope's available set, the user's overrides — and the shell already owns `toolsForMode()`. Keep the authoritative state there:

```ts
let registeredTools: Set<string> | null = null   // null = "scope default"

function availableForMode(): string[] { /* the existing toolsForMode() name list */ }

function activeTools(): AgentSkill['tools'] {
  const available = availableForMode()
  return config.tools.filter(t => available.includes(t.name)
    && (registeredTools === null || registeredTools.has(t.name)))
}
```

`skill`'s live `get tools()` returns `activeTools()`, so a toggle takes effect on the next turn with no other plumbing.

- [ ] **Step 2: Reset on mode change.** In the shell's `onModeChange`, set `registeredTools = null` **before** notifying the UI, then push the new scope down with `ui.setToolScope(availableForMode(), availableForMode())`.
- [ ] **Step 3: Enforce unregister-within-scope in the model, not only the UI.** `activeTools()` intersects with the scope's available set, so an override naming an out-of-scope tool can never widen the offered set even if the UI is bypassed.
- [ ] **Step 4: Never let the set go empty.** If the user unregisters every tool, the model gets no tools and will answer from thin air. Refuse the last toggle-off and show a short note. (Alternative — allow it as a deliberate "chat only" mode. Pick one; do not leave it undefined.)
- [ ] **Step 5:** Registration is in-memory: no `localStorage`, no bridge message, nothing written on document close. That is the specified lifetime.

**Verification:** unit-level — toggling a tool off removes it from `activeTools()`; a mode change restores the full scope default.

---

### Task 5: Localized tool labels (content work)

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Add a `toolDisplay` map to each app's `startAddIn` config — one entry per tool, with an `en`/`he` label and one-sentence description.
- [ ] **Step 2: Do not reuse the schema `description` verbatim.** Those are written for the model (they enumerate parameter shapes and enum values) and are the wrong register for a settings screen. Write short user-facing sentences: *"Reads the document's structure and a text preview"*, not *"Reads the active Word document's paragraph/word count and a text preview."*
- [ ] **Step 3: Coverage** — Word 7 tools, Excel 10, PowerPoint 23. This is ~40 entries × 2 languages and is most of this task's effort. Budget accordingly.
- [ ] **Step 4: Fail loudly on drift.** A tool present in `config.tools` with no `toolDisplay` entry should fall back to its raw schema name and log a console warning naming it — so a tool added later shows up rather than silently vanishing from the screen.
- [ ] **Step 5:** Group the map next to the tool array in each `entry.ts`, so adding a tool and adding its label are the same edit.
- [ ] **Step 6:** The Hebrew text should be reviewed by someone who reads Hebrew before shipping — the existing `STRINGS` table is the style reference. Flag it for review rather than treating machine translation as done.

**Verification:** every tool in every app renders with a real name and description in both languages.

---

### Task 6: Tool list UI

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

- [ ] **Step 1:** Render one row per tool: a toggle, the localized label, and the localized description beneath it in the muted secondary style the mode-menu items already use (`.ai-mode-menu-item .desc`).
- [ ] **Step 2: Control type.** The request says "radio"; the semantics are per-tool on/off, so implement it as a checkbox or switch styled to match the panel — a true `<input type="radio">` group would make the tools mutually exclusive, which is not the intent. Note the interpretation in a code comment.
- [ ] **Step 3: Out-of-scope rows** render with a `disabled` input and an `.out-of-scope` class (reduced opacity), and carry a `title` hint: *"Not available in {mode} — change the edit scope above to enable."* Put the hint on the row, not the disabled input; disabled inputs do not reliably fire hover tooltips in WebView2.
- [ ] **Step 4:** Keep out-of-scope tools **visible**, not filtered out — the point of the hint is discovering that the tool exists and what would unlock it.
- [ ] **Step 5:** Group rows by scope-availability (available first, then out-of-scope under a subheading) so a Read-only user is not scrolling past 20 greyed rows to reach the 2 live ones. For PowerPoint's 23 tools this is the difference between usable and not.
- [ ] **Step 6:** `setToolScope` re-renders the list. Re-render on language change too — the descriptions are localized, so `applyStrings()` alone will not update dynamically generated content unless each row carries `data-t` keys, which it cannot (the strings come from config, not `STRINGS`). Hook the re-render into `setLang`.
- [ ] **Step 7:** Show a live count in the section heading — *"Tools (6 of 9 registered)"* — so the state is legible without reading every row.

**Verification:** toggles work, out-of-scope rows are inert with a hint, the list relocalizes on a language switch.

---

### Task 7: Document system message — persistence

**Files:**
- Create: `OfficeAi.Shared/DocSettingsStore.cs`
- Modify: `OfficeAi.Shared/PaneHostBase.cs` (from PP-0) or the three `TaskPaneHost.cs`
- Modify: `shared/web-src/app-shell/bridge.ts` (from PP-0)
- Create/Modify: `OfficeAi.Shared.Tests/DocSettingsStoreTests.cs`

- [ ] **Step 1: The store.** Model it on `ChatStore` (`OfficeAi.Shared/ChatStore.cs`) — same `LocalApplicationData\<AppFolder>\` root, same per-document id from `ChatStore.ChatIdForFile`, but a single JSON document per file rather than an append-only log:

```csharp
public struct DocSettings { public string SystemMessage; }

public static class DocSettingsStore
{
    public static DocSettings Load(string appDataFolderName, string chatId)
    public static void Save(string appDataFolderName, string chatId, DocSettings settings)
}
```

Write to `...\DocSettings\<chatId>.json`. Keep it beside `ChatHistory\`, not inside it, so the chat-log loader never sees a non-`.jsonl` file.

- [ ] **Step 2: Corruption tolerance.** `ChatStore.LoadSinceLastDivider` skips malformed lines rather than losing the file (`ChatStore.cs:79-82`). Match that spirit: a malformed or unreadable settings file returns an empty `DocSettings` rather than throwing — this runs during pane startup and an exception there is a dead add-in.
- [ ] **Step 3: Atomic write.** Write to a temp file and move over the target, so a crash mid-write cannot leave a truncated file that silently discards the user's guidelines.
- [ ] **Step 4: Cap the length** at a documented maximum (suggest 8 KB) on the C# side, and mirror the cap as a `maxlength` in the UI. It is prepended to every system prompt in every turn of every conversation, so an unbounded value is a per-turn token cost the user cannot see.
- [ ] **Step 5: Bridge kinds.** Add `load-doc-settings` → `doc-settings-loaded` and `save-doc-settings`, handled next to the existing `load-history`/`append-message` branches. Both use `GetChatId()`, so they inherit the per-document keying and the lazy-COM-resolution rule for free.
- [ ] **Step 6: Tests.** `OfficeAi.Shared.Tests` already covers `ChatStore`; add round-trip, missing-file, corrupt-file, and over-cap cases for `DocSettingsStore`. This is the one part of this feature that *can* be unit-tested — the COM and UI layers cannot.
- [ ] **Step 7: The unsaved-document case** is handled by Task 7b — a system message typed into a never-saved document must not be lost the moment the user saves.

**Verification:** `OfficeAi.Shared.Tests` passes; a message saved, document closed, Office restarted, document reopened — the message is still there.

---

### Task 7b: Re-key provisional chat ids on first save

**Files:**
- Modify: `OfficeAi.Shared/ChatStore.cs`, `OfficeAi.Shared/DocSettingsStore.cs`
- Modify: `OfficeAi.Shared/PaneHostBase.cs` (from PP-0), or the three `TaskPaneHost.cs`
- Modify: `OfficeAi.Shared.Tests/`

**Problem:** `GetChatId()` returns `"unsaved-" + Process.GetCurrentProcess().Id` for a never-saved document and caches it for the session (`WordAiAddIn/TaskPaneHost.cs:49-62`, identically in the other two). Anything written under that provisional id — chat history today, the document system message after Task 7 — is orphaned the moment the user saves, because the next session computes a path hash instead.

**Prerequisite — do not start this before PP-1 Task 5 Step 1.** The provisional id is currently *the same for every unsaved document in the process*: open Word and press New twice, and `Document1` and `Document2` share `unsaved-8412`. That is merely untidy today (two transcripts interleave in one file); it becomes destructive here, because re-keying would carry `Document2`'s conversation and guidelines onto `Document1`'s saved id. PP-1 Task 5 Step 1 appends the window handle, giving each unsaved document its own file and making the rename safe.

**Approach: no save events.** Word has `DocumentBeforeSave` but no after-save event, and during `BeforeSave` the path is still empty — so an event-driven version needs a deferral onto the WinForms message queue, and PowerPoint's after-save story would need confirming besides. None of that is necessary: the only reason a save is invisible is that `_chatId` is cached unconditionally. Make the cache conditional and the two call sites that matter are already there.

- [ ] **Step 1: Make a provisional id re-checkable**

```csharp
private string GetChatId()
{
    // A saved id is final. An "unsaved-" id is provisional: re-check the path
    // every call, so the first use after the user saves migrates history and
    // doc settings onto the real per-file id.
    if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

    if (string.IsNullOrEmpty(DocumentPath)) return _chatId ?? (_chatId = ProvisionalId());

    string saved = ChatStore.ChatIdForFile(DocumentFullName);
    if (_chatId != null)
    {
        ChatStore.Migrate(AppFolder, _chatId, saved);
        DocSettingsStore.Migrate(AppFolder, _chatId, saved);
    }
    return _chatId = saved;
}
```

`DocumentPath`/`DocumentFullName`/`ProvisionalId` are the per-app abstract hooks `PaneHostBase` already needs (PP-0 Task 8 Step 2). The path read is one cheap COM property on operations that are already doing file I/O — no perf concern.

- [ ] **Step 2: Cover save-then-close.** The lazy re-check fires on the next `append-message` / `load-doc-settings`, which covers "save and keep working". A user who saves and immediately closes would never trigger it — so call `GetChatId()` once from the document-close handler PP-1 Task 1 Step 4 already wires for pane disposal. One line, and it closes the gap.

- [ ] **Step 3: `ChatStore.Migrate` — append, never clobber.** The target file may already exist (the user saved over a path they had chatted about before). `ChatStore` is append-only JSONL, so concatenating the provisional file's lines onto the existing one is trivially valid and chronologically correct. Delete the source afterwards. If the source does not exist, no-op silently.

- [ ] **Step 4: `DocSettingsStore.Migrate` — non-empty wins.** A single JSON document cannot be concatenated, so it needs a rule: move the provisional settings over only when their `SystemMessage` is non-empty; otherwise keep whatever the target already had. Never overwrite real guidelines with an empty string.

- [ ] **Step 5: Leave Save As alone.** The Step 1 guard treats a saved id as final, so Save As keeps writing under the original file's id for the rest of the session — today's behavior. Whether a conversation should follow the copy or stay with the original has no obviously right answer, and changing it silently is worse than a documented quirk. Record it as a comment, not a TODO.

- [ ] **Step 6: Failure is non-fatal.** A migration that throws (file locked, permissions) must be caught and swallowed with the provisional id retained — this runs inside pane operations, and an exception here is a dead add-in. Worst case the user keeps the provisional id for the session and retries on the next save.

- [ ] **Step 7: Tests.** Both stores are plain file I/O in `OfficeAi.Shared` and fully unit-testable, unlike the COM layer — which is why this task is cheap to get right. Cover: migrate to a fresh target; migrate onto an existing target (append order for chat, non-empty rule for settings); missing source; empty provisional settings; and a re-run of the migration being a no-op.

**Verification:** `OfficeAi.Shared.Tests` passes. Manually: open a new unsaved document, type guidelines and send a message, save the file, close it, reopen it — both the transcript and the guidelines are there under the saved name, and no `unsaved-*` file remains.

---

### Task 8: Document system message — injection into conversations

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts` (from PP-0)
- Modify: `shared/chat-ui/chat-ui.ts`

- [ ] **Step 1: UI.** A `<textarea>` in the settings view with a `data-t-placeholder` hint explaining what it is for (*"Background and guidelines about this document — included at the start of every new conversation"*), `dir="auto"` so Hebrew guidelines render correctly, and the Task 7 Step 4 `maxlength`.
- [ ] **Step 2: The injection point.** `AgentLoop` already exposes `systemSuffix?(): string`, appended to the system prompt every turn (`shared/web-src/agent-core/loop.ts:69`, applied at `:468`), and no app currently uses it. That is the hook — no agent-core change needed.
- [ ] **Step 3: "New conversations" means frozen at conversation start.** A per-turn read of the live value would retroactively change an in-flight conversation, which the contract excludes. Capture it instead:

```ts
// Frozen when a conversation starts, not read live: editing the document
// guidelines must not retroactively change a conversation already in progress.
let activeDocMessage = ''
function beginConversation(): void { activeDocMessage = savedDocMessage }

new AgentLoop({
  systemSuffix: () => activeDocMessage
    ? '\n\nDocument guidelines from the user:\n' + activeDocMessage
    : '',
  ...
})
```

Call `beginConversation()` on initial load and in `onNewChat`, next to the existing `loop.reset()`.

- [ ] **Step 4: Label the injected block** as user-supplied (the wording above does that). It is untrusted text landing in the system prompt; the model should treat it as the user's standing instructions, not as system policy.
- [ ] **Step 5: Saving does not restart the conversation.** Save persists the new value; it takes effect on the next New chat. Say so directly under the textarea — otherwise a user saves guidelines, sees no change, and saves again.
- [ ] **Step 6:** Do not persist the injected block into `ChatStore` — it is not a conversation message. It reaches the model through the system prompt only.

**Verification:** set a message, start a new chat, ask "what guidelines were you given?" — the model repeats them. Edit mid-conversation and confirm the current conversation is unaffected until New chat.

---

### Task 9: Save semantics

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`

- [ ] **Step 1: One rule, matching the existing dropdown.** Save applies connection settings, language, tool registration, and the document system message together. The scope control is the sole exception (Task 3 Step 2) and is labelled as such.
- [ ] **Step 2:** Extend `onSettingsSave`'s payload with `docSystemMessage: string` and `registeredTools: string[]`, or add a separate `onFullSettingsSave` if the dropdown's narrower payload should stay untouched. Prefer extending — one save path is easier to keep correct than two — and update the existing dropdown call site to pass the unchanged values through.
- [ ] **Step 3: Unsaved-changes guard.** Clicking back with pending edits shows a small inline confirm (*"Discard unsaved changes?"*) rather than silently dropping them. Track a dirty flag on input across the view's fields.
- [ ] **Step 4: Confirmation.** Show a brief inline "Saved" acknowledgement in the view. Do not navigate back automatically on Save — the user may want to keep adjusting, and an auto-return makes the tool list feel like it reset.
- [ ] **Step 5: Never log or echo the API key** in the acknowledgement or anywhere else.

**Verification:** each field type round-trips through Save; back-with-changes prompts; back-after-save does not.

---

### Task 10: Tests

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`

- [ ] **Step 1:** "More settings" switches to the settings view; the gear becomes back; back returns to chat.
- [ ] **Step 2:** Messages rendered before opening settings are still in the DOM after returning (Task 1 Step 6).
- [ ] **Step 3:** The tool list renders one row per supplied tool with the localized label for the current language, and re-localizes on a language switch.
- [ ] **Step 4:** Out-of-scope tools render disabled with the hint; in-scope tools toggle.
- [ ] **Step 5:** `setToolScope` re-renders and clears prior toggles.
- [ ] **Step 6:** Save fires `onSettingsSave` once with the document message and the registered-tool list; field input alone does not fire it (preserving the existing tested contract).
- [ ] **Step 7:** A document system message containing HTML renders as literal text — the XSS guard.
- [ ] **Step 8:** Selecting a mode from the settings view updates the composer's mode label and fires `onModeChange` exactly once.

**Verification:** `cd shared/chat-ui && npx vitest run` — all green.

---

### Task 11: Manual verification matrix

- [ ] Settings dropdown → "More settings" → full screen; gear is now a back button; "+" is hidden; collapse still works.
- [ ] A run in progress when settings opens completes correctly, and its output is visible on return.
- [ ] Connection fields edited here and saved are reflected in the dropdown, and the next request uses them.
- [ ] Language switch on this screen relocalizes the whole screen, including tool names and descriptions, and flips to RTL for Hebrew.
- [ ] Every tool appears with a real localized name and description, in all three apps.
- [ ] Scope change (from this screen and from the composer) resets all toggles to that scope's default, in both directions.
- [ ] Out-of-scope tools are greyed, non-clickable, and show the change-scope hint on hover.
- [ ] Unregistering a tool takes it out of the model's options: unregister `read_blocks`, ask something needing it, confirm the model does not call it.
- [ ] Toggles survive navigating to chat and back, and are reset by a scope change.
- [ ] Toggles are gone after closing and reopening the document (specified lifetime).
- [ ] Document system message: type guidelines, Save, New chat, confirm the model follows them.
- [ ] Same message still present after closing the document, quitting Office, and reopening the file.
- [ ] Editing the message mid-conversation does not change the running conversation; it applies on the next New chat.
- [ ] Two different documents have independent messages and independent tool state.
- [ ] Back with unsaved changes prompts; back after saving does not.
- [ ] The whole screen is usable at the pane's default 420px width and does not force horizontal scrolling.
- [ ] Repeat the core pass in all three apps — PowerPoint's 23-tool list is the layout stress case.
