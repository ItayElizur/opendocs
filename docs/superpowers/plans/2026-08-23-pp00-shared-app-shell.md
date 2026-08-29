# PP-0: Shared Add-in Shell — Stop Triplicating the Host Wiring

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source:** not a PP item — a structural prerequisite identified while writing them. PP-1, PP-2, PP-4, and PP-6 each say "do this identically in three files"; this plan removes that instruction from all of them.

**Goal:** One copy of the add-in host wiring — WebView2 bridge, settings, transport, chat-UI mount, and `AgentLoop` event plumbing — consumed by all three apps, so each `entry.ts` contains only what is genuinely app-specific (its tool definitions, system prompt, and starter prompts).

## Measured duplication (verified, not estimated)

Diffing the three apps with app-specific nouns normalized:

| Surface | Size | Lines that actually differ |
|---|---|---|
| `entry.ts` tail — `mountChatUI` options + `AgentLoop` wiring | ~86 × 3 | **0** (except 3 starter-prompt strings) |
| `entry.ts` head — bridge types, message listener, settings, transport | ~120 × 3 | **17**, all of them Word's selection-tracking block |
| `Ribbon.cs` | 72 × 3 | **2** (the namespace) |
| `ThisAddIn.cs` | 67 × 3 | Excel vs PowerPoint: **2** (namespace); Word: +21 (selection hookup) |
| `TaskPaneHost.cs` | ~104 × 3 | 23–38 (chat-id resolution, the `*Tools.Mode` target) |

≈900–1000 lines of near-identical code across 12 files. `shared/chat-ui/chat-ui.ts` is **already** shared and is not part of this problem — the chat UI proper is fine; the glue around it is what was copied.

**Divergence is already happening**, which is the real argument. Three apps implement editing-mode tool filtering three different ways: Word and Excel use a live `get tools()` getter on the skill (`WordAiAddIn/web-src/entry.ts:264-267`, `ExcelAiAddIn/web-src/entry.ts:259-262`), PowerPoint mutates `skill.tools` from inside `onModeChange` (`PowerPointAiAddIn/web-src/entry.ts:447`). All three happen to work — `AgentLoop.startTurn()` re-reads `skill.tools` each turn — but nobody chose that split; it accumulated. PowerPoint's system prompt also asserts mode-dependent tools in a way its wiring only accidentally satisfies.

**Architecture:** a new `shared/web-src/app-shell/` package exporting one entry point, `startAddIn(config)`. It owns everything currently duplicated; each app passes a small config object. The C# side gets `PaneHostBase` and `RibbonBase` in the existing `OfficeAi.Shared` project.

The selection-context block currently unique to Word moves into the shell unconditionally: Excel and PowerPoint simply never receive a `selection-changed` message today, so the code is inert there, and when either grows selection support it becomes free. That is what turns the head's 17-line divergence into zero.

**Tech Stack:** TypeScript (esbuild-bundled, path aliases already established); C# 7.3 / .NET Framework 4.8 for the host-side extraction.

## Sequencing

- **Tasks 1–6 (TypeScript) go before PP-2, PP-4, and PP-6.** Each of those edits the same regions of the same three files; extracting first turns nine edits into three. The tail is already byte-identical, so this is mostly cut-and-move.
- **Tasks 7–9 (C#) coordinate with PP-1.** PP-1 already rewrites all three `ThisAddIn.cs` and `TaskPaneHost.cs` for the per-window pane registry. If PP-1 has not started, do Tasks 7–9 here and let PP-1 build on the base classes. If PP-1 is in flight, **PP-1 owns the C# extraction** — do Tasks 7–9 as part of it rather than creating a merge conflict across six files.

## Global Constraints

- **Pure refactor.** No behavior change anywhere except Task 6 Step 2 (unifying PowerPoint's mode filtering onto the getter pattern), which is called out explicitly. Any other observable difference is a bug in this plan's execution.
- Do not touch `shared/chat-ui/chat-ui.ts` or `shared/web-src/agent-core` / `ai-provider`. This plan moves *callers*, not the packages they call.
- Keep the shell free of app-specific knowledge — no `if (app === 'word')`. Anything that differs comes through the config object. If something cannot be expressed that way, leave it in the app's `entry.ts` rather than adding a flag.
- C# 7.3 / .NET Framework 4.8 only in `.cs` — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Rebuild all three bundles and MSBuild all three projects after every task. Bundles are gitignored; do not commit them.
- **The esbuild command changes in this plan** (a fourth alias). It appears in the Global Constraints of most `2026-08-23-pp*.md` plans; Task 10 updates them.

---

### Task 1: The bridge module

**Files:**
- Create: `shared/web-src/app-shell/bridge.ts`

**Interfaces:**
- Produces: `initBridge(handlers)`, `requestHistory()`, `persistMessage(role, text)`, `callDotNetTool(name, input)`, `postNewChatDivider()`, `postMode(mode)`, `postCollapse(collapsed)`, `postTlsBypass(enabled)`, and the `ToolCallMessage`/`ToolResultMessage`/`OtherMessage` types.

- [ ] **Step 1:** Move the message-protocol interfaces (`WordAiAddIn/web-src/entry.ts:22-47`) verbatim. They are identical in all three files.
- [ ] **Step 2:** Move `callDotNetTool` (`:92-106`), `requestHistory` (`:84-86`), and `persistMessage` (`:88-90`) verbatim.
- [ ] **Step 3:** Move the `chrome.webview.addEventListener('message', ...)` listener (`:55-82`) into `initBridge(handlers)`, where `handlers` carries the callbacks the listener currently calls inline (`onHistoryLoaded`, `onSelectionChanged`, plus the tool-result resolution it already owns internally).
- [ ] **Step 4:** Include Word's `selection-changed` branch (`:57-61`) unconditionally — it is inert in apps that never receive the message. This is the step that removes the head's only real divergence.
- [ ] **Step 5:** Wrap every raw `chrome.webview.postMessage` in a named function so no caller outside this module constructs a message shape by hand. Grep for `chrome.webview.postMessage` afterwards: only `bridge.ts` should match.

**Verification:** `npx tsc --noEmit` in each app once Task 4-6 wire it up; no behavior to test yet.

---

### Task 2: The settings + transport module

**Files:**
- Create: `shared/web-src/app-shell/settings.ts`

**Interfaces:**
- Produces: `StoredSettings`, `loadSettings()`, `saveSettings()`, `makeTransport()`, `MAX_TOKENS`.

- [ ] **Step 1:** Move `StoredSettings`, `loadSettings`, `saveSettings`, `currentSettings`, `MAX_TOKENS`, and `makeTransport` (`WordAiAddIn/web-src/entry.ts:108-157`) verbatim. All three copies are identical apart from a comment naming the app.
- [ ] **Step 2:** `SETTINGS_STORAGE_KEY` stays `'airchat-settings'`. Each app has its own WebView2 user-data folder, so one key across three apps does not collide — keep the comment at `:100-105` explaining that, since it is the reason this is safe to share.
- [ ] **Step 3:** Export `currentSettings` through a getter rather than as a mutable binding, so `makeTransport`'s request-time read stays correct and no consumer can hold a stale copy.
- [ ] **Step 4:** Leave `MAX_TOKENS` at its current `1024` here. Raising it is PP-4 Task 1 — which becomes a one-line change in one file once this lands, instead of three.

**Verification:** typechecks; no behavior change.

---

### Task 3: `startAddIn` — the bootstrap

**Files:**
- Create: `shared/web-src/app-shell/bootstrap.ts`, `shared/web-src/app-shell/index.ts`

**Interfaces:**
- Produces:

```ts
export interface AddInConfig {
  /** every tool this app implements */
  tools: AgentSkill['tools']
  systemPrompt: string
  skillId: string
  starters: Array<{ en: string; he: string }>
  /** tool names available in Read only mode */
  readOnlyTools: string[]
  /** additionally available in Comment only mode (Word's add_comment; empty elsewhere) */
  commentOnlyExtraTools?: string[]
  /** inject the user's current selection into per-turn context (Word today) */
  useSelectionContext?: boolean
}

export function startAddIn(config: AddInConfig): void
```

- [ ] **Step 1:** Move the `mountChatUI` options block (`WordAiAddIn/web-src/entry.ts:279-320`) into `startAddIn`, sourcing `starters` from the config — the only part that differs between apps.
- [ ] **Step 2:** Move the `AgentLoop` construction and its whole `events` object (`:323-361`) verbatim, plus `currentToolGroup`/`activeSteps` and the trailing `requestHistory()` call (`:363`).
- [ ] **Step 3:** Build the skill inside the shell from the config, using the live-getter pattern Word and Excel already use:

```ts
const skill: AgentSkill = {
  id: config.skillId,
  systemPrompt: config.systemPrompt,
  // Live getter, not a fixed array: AgentLoop.startTurn() re-reads
  // skill.tools every turn, so this recomputes from the current editing mode
  // without rebuilding the skill or touching agent-core.
  get tools() { return toolsForMode() },
  ...(config.useSelectionContext ? { buildContext: () => selectionContext() } : {}),
  executeTool: (call) => callDotNetTool(call.name, call.input),
}
```

- [ ] **Step 4:** Implement `toolsForMode()` once, from `readOnlyTools` + `commentOnlyExtraTools`, replacing three hand-written variants.
- [ ] **Step 5:** Keep `initialSettings`, the on-load `set-tls-bypass` post (`:321`), and every existing comment that explains a non-obvious decision. Those comments are the record of past debugging (the WebView2 profile note, the live-getter note, the TLS-on-load note) — losing them in a move is a real cost.
- [ ] **Step 6:** `index.ts` re-exports the public surface: `startAddIn`, `AddInConfig`, and anything an `entry.ts` still needs directly.

**Verification:** typechecks against all three apps' configs.

---

### Task 4: Wire the alias

**Files:**
- Modify: `WordAiAddIn/tsconfig.json`, `ExcelAiAddIn/tsconfig.json`, `PowerPointAiAddIn/tsconfig.json`
- Modify: the build invocation for each app

- [ ] **Step 1:** Add `"@officeai/app-shell": ["../shared/web-src/app-shell/index.ts"]` to each `tsconfig.json`'s `paths`. The three files are currently byte-identical — keep them that way.
- [ ] **Step 2:** Add `--alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts` to the esbuild command, alongside the three existing aliases.
- [ ] **Step 3:** Find where the esbuild command actually lives for each app (the `.csproj` has no `Exec`/`PreBuild` target, so it is invoked manually or from an external script). Whatever the answer, update every copy — and note the location in this plan as you find it, because several `2026-08-23-pp*.md` plans quote the command in their Global Constraints.
- [ ] **Step 4:** If `shared/chat-ui/vitest.config.ts` needs the alias to resolve shell imports in tests, add it there too.

**Verification:** all three bundles build with the new alias before any `entry.ts` is migrated.

---

### Task 5: Migrate Word and Excel

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Word.** Delete everything now in the shell; keep `ALL_WORD_TOOLS`, the system prompt, and the starters. The file becomes imports + tool array + one `startAddIn({...})` call with `useSelectionContext: true`, `readOnlyTools: ['get_document_context','read_blocks']`, `commentOnlyExtraTools: ['add_comment']`. Expect ~363 → ~130 lines.
- [ ] **Step 2:** Build and **diff the bundle's behavior, not just its size** — run Word, send a message, run a tool, change mode, change language, save settings, collapse the pane, select text. This is the checkpoint that proves the extraction is faithful; the remaining migrations are then mechanical.
- [ ] **Step 3: Excel.** Same, with `readOnlyTools` from `READ_ONLY_TOOL_NAMES` (`ExcelAiAddIn/web-src/entry.ts:229-232`), no `commentOnlyExtraTools` (Excel has no comment tool — keep the existing comment explaining that documented gap), `useSelectionContext: false`. Expect ~351 → ~200 lines.
- [ ] **Step 4:** Re-run Step 2's manual pass in Excel.

**Verification:** both apps behave identically to before, by the Step 2 checklist.

---

### Task 6: Migrate PowerPoint, unifying its mode filtering

**Files:**
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Same migration. `READER_TOOLS`/`MUTATION_TOOLS` (`:144`, `:157`) collapse into one `tools` array plus a `readOnlyTools` name list, so the split stops being structural.
- [ ] **Step 2: The one intentional behavior change.** PowerPoint currently reassigns `powerPointSkill.tools` inside `onModeChange` (`:447`); after this it uses the shared live getter like the other two. Functionally equivalent — `AgentLoop` re-reads `skill.tools` per turn either way — but verify explicitly: set Read only, confirm the mutating tools are no longer offered, set Full autonomy, confirm they return.
- [ ] **Step 3:** PowerPoint has no Comment-only-specific tool; it currently treats Comment only as Read only (`:447`). Preserve that exactly by leaving `commentOnlyExtraTools` unset, and keep a comment saying it is a documented gap rather than an oversight.
- [ ] **Step 4:** Expect ~506 → ~380 lines. The smaller reduction is correct — PowerPoint's 23 tool schemas are genuinely app-specific and are most of the file.
- [ ] **Step 5:** Re-run Task 5 Step 2's manual pass in PowerPoint.

**Verification:** all three apps behave identically to before; `git diff --stat` shows a net deletion of roughly 400 lines.

---

### Task 7: C# — `RibbonBase`

**Files:**
- Create: `OfficeAi.Shared/RibbonBase.cs`
- Modify: `WordAiAddIn/Ribbon.cs`, `ExcelAiAddIn/Ribbon.cs`, `PowerPointAiAddIn/Ribbon.cs`

**Read the Sequencing note above before starting: if PP-1 is in flight, PP-1 owns Tasks 7-9.**

- [ ] **Step 1:** Move all of `Ribbon.cs` — the ribbon XML, `OnRibbonLoad`, `GetLogoImage`, and the `PictureConverter`/`AxHostConverter` helper — into an abstract `RibbonBase`. The three files differ by two lines (the namespace).
- [ ] **Step 2:** The one app-specific hook is `Globals.ThisAddIn.TogglePane()`, since `Globals` is VSTO-generated per project. Make it `protected abstract void TogglePane();`.
- [ ] **Step 3:** Each app's `Ribbon.cs` shrinks to a `[ComVisible(true)] public class Ribbon : RibbonBase` with a one-line `TogglePane` override. Keep `[ComVisible(true)]` on the **concrete** class — COM visibility does not inherit usefully here, and a missing attribute fails at runtime with no ribbon and no error.
- [ ] **Step 4:** `GetLogoImage` reads `web/logo.png` relative to `AppDomain.CurrentDomain.BaseDirectory` — verify that still resolves from the shared assembly's perspective. It should (BaseDirectory is the host process's), but this is exactly the kind of thing that silently breaks in an extraction; confirm the ribbon button still shows its icon in all three apps.

**Verification:** all three build; the Home-tab button appears with its logo and toggles the pane in each app.

---

### Task 8: C# — `PaneHostBase`

**Files:**
- Create: `OfficeAi.Shared/PaneHostBase.cs`
- Modify: all three `TaskPaneHost.cs`

- [ ] **Step 1:** Move into an abstract `PaneHostBase : UserControl`: the `_status` label and `UpdateStatus`, the `_bridge` field, the `RequestPaneWidth` event, and the `OnOtherMessage` branches for `load-history`, `append-message`, `new-chat-divider`, `collapse-pane`, and `expand-pane`.
- [ ] **Step 2:** Two abstract hooks: `protected abstract string GetChatId();` and `protected abstract void SetEditingMode(EditingMode mode);` — those are exactly the parts that differ (the COM type used to resolve the document, and which `*Tools` class receives the mode).
- [ ] **Step 3:** The app-data folder name (`"WordAiAddIn"` etc.) becomes a constructor parameter, not a hardcoded string.
- [ ] **Step 4: Preserve the constructor COM warning verbatim.** `WordAiAddIn/TaskPaneHost.cs:29-40` documents a confirmed repro where touching `Application.ActiveDocument` in the constructor silently kills the VSTO connection. Move that comment to `PaneHostBase`'s constructor, where it now guards three apps instead of one.
- [ ] **Step 5:** Word keeps `OnSelectionChanged` in its own subclass — it is genuinely Word-only until another app grows selection support.
- [ ] **Step 6:** Add `PaneHostBase.cs` and `RibbonBase.cs` to `OfficeAi.Shared.csproj` if that project does not glob its sources.

**Verification:** all three build; panes initialize to `ready`, history loads, mode changes reach the right `*Tools` class, collapse/expand still resizes the native pane.

---

### Task 9: C# — hand off to PP-1

**Files:** none (coordination)

- [ ] **Step 1:** PP-1 (`2026-08-23-pp01-taskpane-per-window.md`) Tasks 1-3 write the same pane registry three times. With `PaneHostBase` in place, note in PP-1 that the registry itself should be a generic in `OfficeAi.Shared` — `PaneRegistry<TWindow>` parameterized by delegates for "get hwnd from window" and "get document from window", since `Word.Window`/`Excel.Window`/`PowerPoint.DocumentWindow` share no interface.
- [ ] **Step 2:** Update PP-1's Tasks 2 and 3 from "port Task 1's registry verbatim" to "instantiate the shared registry", and add the generic's construction to its Task 1.
- [ ] **Step 3:** Do not attempt the generic registry in this plan — it needs PP-1's event-wiring work to have a shape worth generalizing. Extracting it speculatively would guess wrong.

---

### Task 10: Update the plan set

**Files:**
- Modify: `docs/superpowers/plans/2026-08-23-pp02-*.md`, `pp04-*.md`, `pp06-*.md`, `pp01-*.md`, `pp-index.md`

- [ ] **Step 1:** PP-2 Task 3, PP-4 Tasks 1/2/3, and PP-6 Task 2 all say "do this in all three `entry.ts`". Rewrite each as a single edit in `shared/web-src/app-shell/`. PP-4 Task 1 in particular collapses from three files to one line.
- [ ] **Step 2:** PP-1's Tasks 1-4 gain the `PaneHostBase`/`RibbonBase` dependency and the Task 9 registry note.
- [ ] **Step 3:** Update the esbuild command in every plan's Global Constraints to include the fourth alias.
- [ ] **Step 4:** Add PP-0 to `2026-08-23-pp-index.md`'s ordering table as tier 1, ahead of PP-2, and record the dependency in its cross-plan section.
- [ ] **Step 5:** Note in the index that PP-3 is *not* affected — `chat-ui.ts` was already shared, which is worth stating so nobody assumes this plan changes it.

**Verification:** no remaining plan instructs a three-file edit for anything the shell now owns.
