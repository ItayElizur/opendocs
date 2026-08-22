# Add-in UX Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix four user-reported UX defects in the live VSTO add-ins: (1) a native debug status label permanently consumes screen space beneath a redundant title bar, (2) the panel's "minimize"/collapse button does nothing, (3) switching the Settings language between English/Hebrew has no effect anywhere, and (4) the empty "New chat" state never shows the suggested starter prompts designed in the approved mockup. Also adds an automated test proving RTL support actually works end-to-end at the component level: a Hebrew message sent and a Hebrew reply received both render inside the RTL-marked container with the mockup's approved alignment rules applying.

**Architecture:** Items (3) and (4) both live inside the shared `shared/chat-ui/chat-ui.ts` component and are tightly coupled (the user explicitly wants starter suggestions to also switch language) — they're built as one language-then-content sequence in that file, then wired per-app. Item (2) requires a real native-pane width change (not just a CSS illusion), since the CustomTaskPane's actual OS-native width doesn't shrink on its own just because the HTML inside it does — so it adds one new WebMessage kind (`collapse-pane`/`expand-pane`) and a small event hookup from `TaskPaneHost.cs` back up to `ThisAddIn.cs` in each of the three add-ins. Item (1) is the smallest, self-contained change to `OfficeAi.Shared/WebViewBridgeHost.cs` plus a one-line tweak per `TaskPaneHost.cs`.

**Tech Stack:** TypeScript (esbuild-bundled), C# 7.3/.NET Framework 4.8 (VSTO), same conventions as the rest of `officeoffice`.

**Spec:** `shared/chat-ui/mockup.html` is the source of truth for the approved chat UI design (already referenced throughout `docs/superpowers/plans/2026-08-22-office-ai-toolset-port.md`) — this plan implements pieces of that mockup (collapse-to-rail, `STRINGS`/`STARTERS` i18n tables, `setLang`/`applyStrings` pattern) that were previously deferred or left as dead/unwired code.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only in all `.cs` files — no `using`-declarations, no target-typed `new()`, no switch expressions (matches every other file in this project).
- UI chrome colors must use the existing CSS custom properties in `shared/chat-ui/chat-ui.css` (`var(--color-...)` etc.) — no new raw hex values in `.ts`/`.css` except inside a token definition line, per this repo's `CLAUDE.md` theming rules.
- Language switching in `chat-ui.ts` is scoped to the panel's own root element only — never touch `document.documentElement` — matching the existing, already-approved pattern for RTL/theme scoping in this component.
- Do not add file-attachment UI, a multi-session chat browser, or anything else explicitly excluded by the original toolset-port plan's Global Constraints — this plan only touches the four items listed above.
- Rebuild each add-in's esbuild bundle (`npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap`) and re-run `MSBuild -t:Build` after any `shared/chat-ui/chat-ui.ts` or `entry.ts` change, in each of `WordAiAddIn/`, `ExcelAiAddIn/`, `PowerPointAiAddIn/` — a stale bundle silently ships the old behavior. These are gitignored build artifacts; do not commit them.
- No automated tests exist (or are expected) for COM-executor `.cs` methods, per the toolset-port plan's Global Constraints — this plan doesn't touch COM executor logic at all, only `TaskPaneHost.cs`/`ThisAddIn.cs` plumbing and `WebViewBridgeHost.cs`, so this constraint mostly doesn't apply, but keep the same spirit: don't invent COM-level tests.

---

### Task 1: Auto-hide the native debug status label once ready

**Files:**
- Modify: `OfficeAi.Shared/WebViewBridgeHost.cs`
- Modify: `WordAiAddIn/TaskPaneHost.cs`
- Modify: `ExcelAiAddIn/TaskPaneHost.cs`
- Modify: `PowerPointAiAddIn/TaskPaneHost.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new — this is a pure behavior fix, no API changes. `WebViewBridgeHost`'s existing `Action<string> setStatus` callback parameter keeps the same signature.

**Problem:** each `TaskPaneHost` docks a plain WinForms `Label` (`_status`, `Dock = DockStyle.Top, Height = 24`) above the WebView2 control, showing raw debug text like `"WebView2: initializing..."`, `"ready"`, `"Executing tool: X"`. Once init succeeds the label never goes away — it permanently occupies a 24px strip above the WebView2-rendered chat UI (which already has its own polished header), visually duplicating the native CustomTaskPane's own title ("Airchat Office") and wasting vertical space. Fix: hide the label whenever the current status is exactly `"ready"`; show it (with the real text) otherwise — during startup, during a fatal WebView2 init failure, or while a tool is actively executing / just errored.

- [ ] **Step 1: Make `WebViewBridgeHost` revert to "ready" after a successful tool call, and report a distinct string on tool error**

Read `OfficeAi.Shared/WebViewBridgeHost.cs` current content, then modify the tool-call branch inside `OnWebMessageReceived` (currently):
```csharp
var (requestId, toolName, input) = ToolProtocol.ParseToolCall(e.WebMessageAsJson);
_setStatus("Executing tool: " + toolName);
ToolResult result = _executor(toolName, input);
_setStatus("Tool done: " + toolName + (result.IsError ? " (error)" : ""));
if (_webView.CoreWebView2 != null)
{
    _webView.CoreWebView2.PostWebMessageAsJson(ToolProtocol.SerializeToolResult(requestId, result));
}
```
to:
```csharp
var (requestId, toolName, input) = ToolProtocol.ParseToolCall(e.WebMessageAsJson);
_setStatus("Executing tool: " + toolName);
ToolResult result = _executor(toolName, input);
if (_webView.CoreWebView2 != null)
{
    _webView.CoreWebView2.PostWebMessageAsJson(ToolProtocol.SerializeToolResult(requestId, result));
}
_setStatus(result.IsError ? "Tool error: " + toolName : "ready");
```
(A successful tool call now hides the label again immediately; a failed one leaves a diagnostic message visible until the next status change. The in-panel "Running tools..." work-group in `chat-ui.ts` already shows per-tool success/failure to the user — this native label is now purely a startup/fatal-error indicator, not a duplicate of that UI.)

- [ ] **Step 2: Add an `UpdateStatus` method to each `TaskPaneHost.cs` that hides the label when the status is "ready"**

In `WordAiAddIn/TaskPaneHost.cs`, change the constructor's last line (currently):
```csharp
_bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s, OnOtherMessage);
```
to:
```csharp
_bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", UpdateStatus, OnOtherMessage);
```
and add this private method (placed right after the constructor, before `GetChatId()`):
```csharp
private void UpdateStatus(string s)
{
    _status.Text = s;
    _status.Visible = s != "ready";
}
```

Apply the identical change to `ExcelAiAddIn/TaskPaneHost.cs` (constructor line uses `ExcelTools.Execute, "ExcelAiAddIn"`) and `PowerPointAiAddIn/TaskPaneHost.cs` (constructor line uses `PowerPointTools.Execute, "PowerPointAiAddIn"`) — same `UpdateStatus` method body, verbatim, in each file.

- [ ] **Step 3: Build and manually verify in each app**

Run, from `C:/dev/officeoffice`:
```bash
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```
Expected: all three build with 0 warnings/errors. Then manually open each app: the status strip should be invisible once the panel loads (no blank/gray bar above the chat header), should briefly reappear if a tool call takes noticeably long or errors, and should NOT permanently reappear after a normal successful tool call.

- [ ] **Step 4: Commit**

```bash
git add OfficeAi.Shared/WebViewBridgeHost.cs WordAiAddIn/TaskPaneHost.cs ExcelAiAddIn/TaskPaneHost.cs PowerPointAiAddIn/TaskPaneHost.cs
git commit -m "fix: hide native debug status label once ready instead of permanently"
```

---

### Task 2: `chat-ui.ts` — language switching (i18n core)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/chat-ui/chat-ui.test.ts`

**Interfaces:**
- Consumes: nothing new.
- Produces: a module-level `STRINGS` table and `t(key)` lookup that Task 3 (starter pills) also uses. Changes the public `ChatUIOptions`/`ChatUIHandle` shape: removes the unused `title: string` option field (it was accepted but never read — the panel title is now always the localized `panelTitle` string), and replaces `setScopeHint(label: string)` with `setSelectionScope(selection: { hasSelection: boolean; preview: string } | null): void` so the component (not each app's `entry.ts`) owns the localized "Whole document" / "Selection: ..." wording. Task 5 updates `WordAiAddIn/web-src/entry.ts` to call the new method.

**Design:** mirrors `shared/chat-ui/mockup.html`'s already-approved `STRINGS`/`t()`/`applyStrings()`/`setLang()` pattern (see `mockup.html:447-556`), adapted to the real component's DOM structure and TypeScript. Static, persistent elements (panel title, settings labels, mode-menu items, composer placeholder) get `data-t`/`data-t-title`/`data-t-placeholder` attributes and are updated in place by `applyStrings()`. The empty-state block is recreated on demand (mount / `resetToEmpty` / `showHistoric`), so its `data-t="emptyTitle"` element is refreshed automatically every time `emptyStateHtml()` runs — but if the language changes while an empty state is *already* showing, `setLang()` must re-inject it. Language only takes effect when Settings' Save button is pressed (existing `pendingLang` mechanism), matching prior explicit feedback ("language will only change after pressing save").

- [ ] **Step 1: Add the `STRINGS` table, `Lang` type, and `t()`/`applyStrings()` helpers**

In `shared/chat-ui/chat-ui.ts`, after the existing `export type EditingMode = ...` line, add:
```typescript
export type Lang = 'en' | 'he'

const STRINGS: Record<string, Record<Lang, string>> = {
  panelTitle:           { en: 'Airchat Office', he: "איירצ'אט אופיס" },
  inputPlaceholder:     { en: 'Ask Airchat Office to edit this document...', he: 'בקש מ-Airchat Office לערוך את המסמך...' },
  send:                 { en: 'Send', he: 'שלח' },
  newChat:              { en: 'New chat', he: 'שיחה חדשה' },
  settings:             { en: 'Settings', he: 'הגדרות' },
  settingsTitle:        { en: 'Airchat Office Settings', he: 'הגדרות Airchat Office' },
  settingsBaseUrl:      { en: 'API Base URL', he: 'כתובת בסיס API' },
  settingsApiKey:       { en: 'API Key', he: 'מפתח API' },
  settingsModel:        { en: 'Model name', he: 'שם המודל' },
  settingsLanguage:     { en: 'Language', he: 'שפה' },
  save:                 { en: 'Save', he: 'שמור' },
  collapse:             { en: 'Collapse panel', he: 'כווץ חלונית' },
  historySep:           { en: 'Earlier conversation', he: 'שיחה קודמת' },
  scopeWholeDoc:        { en: 'Whole document', he: 'כל המסמך' },
  emptyTitle:           { en: 'What can I help with?', he: 'איך אפשר לעזור?' },
  modeReadOnly:         { en: 'Read only', he: 'קריאה בלבד' },
  modeReadOnlyDesc:     { en: 'AI can only read, never edit', he: 'הבינה יכולה רק לקרוא, לא לערוך' },
  modeCommentOnly:      { en: 'Comment only', he: 'הערות בלבד' },
  modeCommentOnlyDesc:  { en: 'Adds comments, no content edits', he: 'מוסיפה הערות בלבד, ללא עריכת תוכן' },
  modeTrackChanges:     { en: 'Track changes', he: 'מעקב אחר שינויים' },
  modeTrackChangesDesc: { en: 'Edits as reviewable revisions', he: 'עריכות כתיקונים לאישור' },
  modeFullAutonomy:     { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
  modeFullAutonomyDesc: { en: 'Edits applied directly', he: 'עריכות מוחלות ישירות' },
}
```

- [ ] **Step 2: Remove the unused `title` option; replace `setScopeHint` with `setSelectionScope` in the interfaces**

Change `ChatUIOptions` (remove the `title: string` line):
```typescript
export interface ChatUIOptions {
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; lang: Lang }) => void
}
```
Change `ChatUIHandle`'s `setScopeHint` line:
```typescript
  setSelectionScope(selection: { hasSelection: boolean; preview: string } | null): void
```

- [ ] **Step 3: Add `data-t`/`data-t-placeholder` attributes to the static template**

In `mountChatUI`'s template literal, apply these exact text-content changes (attribute additions only, no structural changes):
- `<span>Airchat Office</span>` (panel title span) → `<span data-t="panelTitle">Airchat Office</span>`
- `<h4>Airchat Office Settings</h4>` → `<h4 data-t="settingsTitle">Airchat Office Settings</h4>`
- `<label>API Base URL</label>` → `<label data-t="settingsBaseUrl">API Base URL</label>`
- `<label>API Key</label>` → `<label data-t="settingsApiKey">API Key</label>`
- `<label>Model name</label>` → `<label data-t="settingsModel">Model name</label>`
- `<label>Language</label>` → `<label data-t="settingsLanguage">Language</label>`
- `<button class="ai-btn-primary">Save</button>` → `<button class="ai-btn-primary" data-t="save">Save</button>`
- `<textarea class="ai-textarea" rows="1" placeholder="Ask Airchat Office to edit this document...">` → add `data-t-placeholder="inputPlaceholder"` (keep the existing `placeholder="..."` attribute as-is, as a pre-JS fallback)
- Each of the 4 mode-menu items' inner spans, e.g. `<span>Read only</span><span class="desc">AI can only read, never edit</span>` → `<span data-t="modeReadOnly">Read only</span><span class="desc" data-t="modeReadOnlyDesc">AI can only read, never edit</span>` (and the matching `modeCommentOnly`/`modeCommentOnlyDesc`, `modeTrackChanges`/`modeTrackChangesDesc`, `modeFullAutonomy`/`modeFullAutonomyDesc` keys for the other three items)
- `<div class="ai-header-btn" data-t-title="collapse">` — already has `data-t-title="collapse"`, no change needed (Task 4 adds the matching button element move; this attribute already exists on today's collapse button).

Do NOT add `data-t` to `#scopeHintLabel` (its text is owned by `setSelectionScope`/`refreshScopeHint`, not the static-attribute scan) or to the empty-state title (owned by `emptyStateHtml()`, re-rendered fresh each time — see Step 5).

- [ ] **Step 4: Implement `t()`, `applyStrings()`, `refreshScopeHint()`, `refreshModeLabel()`, and `setLang()`**

Inside `mountChatUI`, after the existing `let assistantBubble` / `let pendingLang` declarations, add:
```typescript
  let currentLang: Lang = 'en'
  let lastSelection: { hasSelection: boolean; preview: string } | null = null

  function t(key: string): string {
    return STRINGS[key]?.[currentLang] ?? key
  }

  function applyStrings(): void {
    root.querySelectorAll<HTMLElement>('[data-t]').forEach((el) => {
      el.textContent = t(el.dataset.t!)
    })
    root.querySelectorAll<HTMLElement>('[data-t-title]').forEach((el) => {
      const s = t(el.dataset.tTitle!)
      el.title = s
      el.setAttribute('aria-label', s)
    })
    root.querySelectorAll<HTMLTextAreaElement>('[data-t-placeholder]').forEach((el) => {
      el.placeholder = t(el.dataset.tPlaceholder!)
    })
  }

  function refreshScopeHint(): void {
    if (lastSelection && lastSelection.hasSelection) {
      const prefix = currentLang === 'he' ? 'בחירה: "' : 'Selection: "'
      scopeHintLabel.textContent = prefix + lastSelection.preview + '..."'
    } else {
      scopeHintLabel.textContent = t('scopeWholeDoc')
    }
  }

  function refreshModeLabel(): void {
    const selected = root.querySelector<HTMLElement>('.ai-mode-menu-item.selected')
    if (selected) modeBtnLabel.textContent = selected.querySelector('span')!.textContent
  }

  function setLang(l: Lang): void {
    currentLang = l
    root.querySelectorAll<HTMLButtonElement>('.ai-lang-toggle button').forEach((b) => {
      b.classList.toggle('active', b.dataset.lang === l)
    })
    applyStrings()
    refreshScopeHint()
    refreshModeLabel()
    const existingEmpty = chatEl.querySelector('.ai-chat-empty')
    if (existingEmpty) {
      existingEmpty.remove()
      chatEl.insertAdjacentHTML('beforeend', emptyStateHtml())
    }
  }
```
(`refreshModeLabel` re-reads the already-updated-by-`applyStrings` `<span data-t="...">` text inside the currently-selected mode item, so it must run AFTER `applyStrings()` in `setLang`, as written above. `dir`/`lang` attribute placement on the actual `.ai-dock` wrapper element is added in Task 4, once that wrapper exists — until then, RTL layout simply doesn't visually flip, which is fine since Task 4 lands before this plan is done and no intermediate release happens.)

- [ ] **Step 5: Wire `setLang` into the Save button, and add `setSelectionScope` to the returned handle**

Replace the existing Save-button listener body (currently calls `options.onSettingsSave(...)` directly) with:
```typescript
  root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.addEventListener('click', () => {
    setLang(pendingLang)
    options.onSettingsSave({
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
      lang: pendingLang,
    })
    settingsPanel.classList.remove('open')
  })
```
Replace the returned handle's `setScopeHint(label) { scopeHintLabel.textContent = label }` method with:
```typescript
    setSelectionScope(selection) {
      lastSelection = selection
      refreshScopeHint()
    },
```
Also call `applyStrings()` once, right after the template is injected and `chatEl`/other consts are resolved (so the panel shows correctly-cased default English strings from `STRINGS` immediately on mount rather than only the raw hardcoded template text) — add `applyStrings()` as the last line before the `let assistantBubble` declarations.

- [ ] **Step 6: Update `chat-ui.test.ts` for the changed API**

In `shared/chat-ui/chat-ui.test.ts`:
- Remove `title: 'Airchat Office'` from the `mountChatUI(root, { ... })` call in `setup()` (the option no longer exists).
- Change the `setScopeHint` test:
```typescript
  it('setSelectionScope updates the hint label text for a live selection, and reverts to Whole document', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, preview: 'Q3 revenue grew...' })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selection: "Q3 revenue grew...."')
    handle.setSelectionScope(null)
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Whole document')
  })
```
(Note the exact expected string includes the literal trailing `..."` per `refreshScopeHint`'s `preview + '..."'` concatenation — adjust the sample `preview` value above if needed so the assertion is exact and not confusing; e.g. use `'Q3 revenue grew'` as the preview so the rendered result reads `Selection: "Q3 revenue grew..."` cleanly.)
- Add a language-switch test:
```typescript
  it('saving settings with a language change updates panel strings via Save, not live', () => {
    const { root, onSettingsSave } = setup()
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    // Not yet applied - Hebrew string should not appear until Save.
    expect(root.querySelector('[data-t="panelTitle"]')!.textContent).toBe('Airchat Office')
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('[data-t="panelTitle"]')!.textContent).toBe("איירצ'אט אופיס")
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ lang: 'he' }))
  })
```

- [ ] **Step 7: Run the tests**

Run: `cd shared/chat-ui && npx vitest run`
Expected: all tests pass, including the two changed/new ones.

- [ ] **Step 8: Commit**

```bash
git add shared/chat-ui/chat-ui.ts shared/chat-ui/chat-ui.test.ts
git commit -m "feat(chat-ui): wire Settings language toggle to actually switch UI strings"
```

---

### Task 3: `chat-ui.ts` — starter suggestion pills in the empty state

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/chat-ui/chat-ui.test.ts`

**Interfaces:**
- Consumes: `t()`/`currentLang` from Task 2 (this task must run after Task 2).
- Produces: a new required `ChatUIOptions.starters: Array<{ en: string; he: string }>` field — each of the three add-ins' `entry.ts` (Task 5/6) must supply exactly 3 app-specific starter prompts, matching `mockup.html`'s per-app `STARTERS[app].empty` sets (`mockup.html:477-505`).

**Design:** matches the mockup's actual (not aspirational) behavior exactly — `renderEmpty()` in the mockup always uses the `.empty` starter set (the `.active` set exists in the mockup's data table but its only would-be use site is dead code that always resolves to `.empty` too — see `mockup.html:665-666` — so there is only one starter set to implement, not two). Clicking a pill populates the composer textarea with that prompt's text (matching `mockup.html`'s `onclick="...value=this.textContent"`); it does not auto-send, so the user can review/edit before sending.

- [ ] **Step 1: Write the failing test**

Add to `shared/chat-ui/chat-ui.test.ts`'s `setup()` helper, add a `starters` option:
```typescript
  const handle = mountChatUI(root, {
    onSend, onModeChange, onSettingsSave, onNewChat,
    starters: [
      { en: 'Summarize this document', he: 'סכם את המסמך' },
      { en: 'Fix grammar issues', he: 'תקן שגיאות דקדוק' },
      { en: 'Improve conciseness', he: 'שפר תמציתיות' },
    ],
  })
```
Add a new test:
```typescript
  it('shows starter pills in the empty state, and clicking one fills the textarea', () => {
    const { root } = setup()
    const pills = root.querySelectorAll<HTMLElement>('.ai-starter')
    expect(pills.length).toBe(3)
    expect(pills[0].textContent).toBe('Summarize this document')
    pills[0].click()
    expect(root.querySelector<HTMLTextAreaElement>('.ai-textarea')!.value).toBe('Summarize this document')
  })
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `cd shared/chat-ui && npx vitest run -t "starter pills"`
Expected: FAIL — `.ai-starters` is currently always empty.

- [ ] **Step 3: Implement starter pill rendering and click-to-fill**

Add `starters: Array<{ en: string; he: string }>` to `ChatUIOptions`:
```typescript
export interface ChatUIOptions {
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; lang: Lang }) => void
  starters: Array<{ en: string; he: string }>
}
```

Replace `emptyStateHtml()`:
```typescript
function emptyStateHtml(options: ChatUIOptions, currentLang: Lang): string {
  const pills = options.starters
    .map((s) => `<div class="ai-starter">${escapeHtml(s[currentLang])}</div>`)
    .join('')
  return `<div class="ai-chat-empty"><div class="ai-chat-empty-title" data-t="emptyTitle">What can I help with?</div><div class="ai-starters">${pills}</div></div>`
}
```
(Now takes `options`/`currentLang` as parameters since it's a module-level function, not a closure over `mountChatUI`'s locals — update every call site inside `mountChatUI` from `emptyStateHtml()` to `emptyStateHtml(options, currentLang)`: the mount-time `chatEl.innerHTML = ...`, `resetToEmpty()`, `showHistoric()`'s `insertAdjacentHTML` call, and `setLang()`'s re-injection call added in Task 2 Step 4.)

Add a single delegated click listener for starter pills, placed near the other `chatEl`-related listeners (after `chatEl` is resolved):
```typescript
  chatEl.addEventListener('click', (e) => {
    const target = (e.target as HTMLElement).closest('.ai-starter')
    if (target) textarea.value = target.textContent || ''
  })
```

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `cd shared/chat-ui && npx vitest run`
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add shared/chat-ui/chat-ui.ts shared/chat-ui/chat-ui.test.ts
git commit -m "feat(chat-ui): render localized starter suggestion pills in the empty state"
```

---

### Task 4: `chat-ui.ts` — working collapse-to-rail UI

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/chat-ui/chat-ui.test.ts`

**Interfaces:**
- Consumes: `setLang`'s existing `dir`/`lang` attribute logic from Task 2 (moves it from `root` to the new `.ai-dock` wrapper element this task introduces).
- Produces: a new `ChatUIOptions.onCollapseChange: (collapsed: boolean) => void` callback. Task 7 wires each app's `entry.ts` to post a `collapse-pane`/`expand-pane` WebMessage from this callback, and Task 6's C# changes actually shrink/restore the native CustomTaskPane's width in response.

**Design:** wraps the existing `.ai-panel` markup inside a `.ai-dock` container with a narrow `.ai-rail` sibling (both CSS classes already exist, unused, in `chat-ui.css` — this task is the first thing that actually renders them). Clicking the header's collapse button (`data-t-title="collapse"`, already present) hides `.ai-panel` and shows the narrow rail via the existing `.ai-dock.collapsed` CSS rule; clicking the rail re-expands. CSS alone only shrinks the *content*, not the actual OS-native docked pane — Task 6/7 make the pane itself narrower so there's no leftover blank space beside the rail.

- [ ] **Step 1: Wrap the template in `.ai-dock`/`.ai-rail`**

Change `mountChatUI`'s `root.innerHTML = ...` template from:
```typescript
  root.innerHTML = `
    <div class="ai-panel">
      ...
    </div>
  `
```
to:
```typescript
  root.innerHTML = `
    <div class="ai-dock">
      <div class="ai-rail" data-t="panelTitle"></div>
      <div class="ai-panel">
        ...
      </div>
    </div>
  `
```
(keep everything inside the existing `<div class="ai-panel">...</div>` exactly as Tasks 2/3 left it — only the outer wrapper and the new `.ai-rail` sibling are added). Note `.ai-rail` reuses the `panelTitle` string (via `data-t`) rather than a separate always-English "AIRCHAT OFFICE" wordmark as in the demo mockup — a deliberate simplification so the rail's label localizes along with everything else, consistent with this plan's Task 2 i18n work, instead of adding a brand-wordmark exception.

- [ ] **Step 2: Wire the collapse/expand click handlers**

Add `onCollapseChange: (collapsed: boolean) => void` to `ChatUIOptions`.

After the existing `const scopeHintLabel = ...` line, add:
```typescript
  const dockEl = root.querySelector<HTMLDivElement>('.ai-dock')!
  const railEl = root.querySelector<HTMLDivElement>('.ai-rail')!
  const collapseBtn = root.querySelector<HTMLButtonElement>('[data-t-title="collapse"]')!

  function setCollapsed(collapsed: boolean): void {
    dockEl.classList.toggle('collapsed', collapsed)
    options.onCollapseChange(collapsed)
  }

  collapseBtn.addEventListener('click', () => setCollapsed(true))
  railEl.addEventListener('click', () => setCollapsed(false))
```

- [ ] **Step 3: Move `dir`/`lang` attribute-setting from `root` to `dockEl` in `setLang`**

In `setLang()` (from Task 2 Step 4), add these two lines (they were not present in Task 2 since `.ai-dock` didn't exist yet):
```typescript
    dockEl.setAttribute('lang', l)
    dockEl.setAttribute('dir', l === 'he' ? 'rtl' : 'ltr')
```
placed as the first two lines inside `setLang(l)`, before `currentLang = l`.

- [ ] **Step 4: Write a test for the collapse behavior**

Add to `chat-ui.test.ts`:
```typescript
  it('clicking collapse hides the panel and shows the rail; clicking the rail re-expands', () => {
    const onCollapseChange = vi.fn()
    const root = document.createElement('div')
    document.body.appendChild(root)
    mountChatUI(root, {
      onSend: vi.fn(), onModeChange: vi.fn(), onSettingsSave: vi.fn(), onNewChat: vi.fn(),
      starters: [], onCollapseChange,
    })
    root.querySelector<HTMLButtonElement>('[data-t-title="collapse"]')!.click()
    expect(root.querySelector('.ai-dock')!.classList.contains('collapsed')).toBe(true)
    expect(onCollapseChange).toHaveBeenCalledWith(true)
    root.querySelector<HTMLElement>('.ai-rail')!.click()
    expect(root.querySelector('.ai-dock')!.classList.contains('collapsed')).toBe(false)
    expect(onCollapseChange).toHaveBeenCalledWith(false)
  })
```
Also add `onCollapseChange: vi.fn()` to the shared `setup()` helper's `mountChatUI` call (Task 3 Step 1 already added a non-empty `starters` array there) so every other existing test keeps compiling under the now-required options.

- [ ] **Step 5: Run the tests**

Run: `cd shared/chat-ui && npx vitest run`
Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add shared/chat-ui/chat-ui.ts shared/chat-ui/chat-ui.test.ts
git commit -m "feat(chat-ui): implement working collapse-to-rail panel UI"
```

---

### Task 5: Wire `WordAiAddIn/web-src/entry.ts` to the new chat-ui.ts APIs

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: `ChatUIOptions.starters`/`onCollapseChange` (Tasks 3/4), `ChatUIHandle.setSelectionScope` (Task 2).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Update the `mountChatUI` call**

In `WordAiAddIn/web-src/entry.ts`, find the `mountChatUI(root, { ... })` call and:
- Remove the `title: 'Airchat Office',` line (option no longer exists).
- Add:
```typescript
  starters: [
    { en: 'Summarize the key points of this document', he: 'סכם את הנקודות העיקריות במסמך' },
    { en: 'Polish the whole document for a more professional tone', he: 'לטש את כל המסמך לטון מקצועי יותר' },
    { en: 'Continue writing from where the document leaves off', he: 'המשך לכתוב מהיכן שהמסמך מסתיים' },
  ],
  onCollapseChange: (collapsed) => {
    chrome.webview.postMessage({ kind: collapsed ? 'collapse-pane' : 'expand-pane' })
  },
```
(These 3 English/Hebrew pairs are copied verbatim from `mockup.html`'s `STARTERS.word.active` set — Word's mockup used its `active`/already-has-content phrasing here rather than the blank-document `empty` set, since a Word document opened in a real add-in almost always already has content, unlike the mockup's from-scratch demo scenario. This is an intentional, small deviation from the generic Task 3 spec's "always use `.empty`" default, appropriate specifically for Word.)

- [ ] **Step 2: Update the `selection-changed` handler to call `setSelectionScope`**

Replace the existing handler body (currently):
```typescript
  if (data.kind === 'selection-changed') {
    latestSelection = data as unknown as typeof latestSelection
    ui.setScopeHint(
      latestSelection.hasSelection ? `Selection: "${latestSelection.preview}..."` : 'Whole document',
    )
  }
```
with:
```typescript
  if (data.kind === 'selection-changed') {
    latestSelection = data as unknown as typeof latestSelection
    ui.setSelectionScope(latestSelection.hasSelection ? { hasSelection: true, preview: latestSelection.preview } : null)
  }
```

- [ ] **Step 3: Typecheck and rebuild**

Run, from `WordAiAddIn/`:
```bash
npx tsc --noEmit
npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap
```
Expected: 0 tsc errors, esbuild succeeds.

- [ ] **Step 4: Commit**

```bash
git add WordAiAddIn/web-src/entry.ts
git commit -m "feat(word): wire starter prompts, localized scope hint, and pane collapse"
```

---

### Task 6: Wire `ExcelAiAddIn`/`PowerPointAiAddIn`'s `entry.ts` to the new chat-ui.ts APIs

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: same as Task 5.
- Produces: nothing new.

Same shape as Task 5, batched into one task since neither app tracks text selection (no `setSelectionScope` call needed — their scope hint stays the default "Whole document"/"כל המסמך" from Task 2's `refreshScopeHint`, which requires no code here at all).

- [ ] **Step 1: Update `ExcelAiAddIn/web-src/entry.ts`'s `mountChatUI` call**

Remove `title: 'Airchat Office',`. Add:
```typescript
  starters: [
    { en: 'Summarize this sheet', he: 'סכם את הגיליון הזה' },
    { en: 'Add a totals row', he: 'הוסף שורת סיכום' },
    { en: 'Check the formulas', he: 'בדוק את הנוסחאות' },
  ],
  onCollapseChange: (collapsed) => {
    chrome.webview.postMessage({ kind: collapsed ? 'collapse-pane' : 'expand-pane' })
  },
```
(copied verbatim from `mockup.html`'s `STARTERS.excel.empty`/`.active` sets, which are identical for Excel.)

- [ ] **Step 2: Update `PowerPointAiAddIn/web-src/entry.ts`'s `mountChatUI` call**

Remove `title: 'Airchat Office',`. Add:
```typescript
  starters: [
    { en: "Improve this slide's title and copy", he: 'שפר את הכותרת והטקסט של השקופית' },
    { en: "Make this slide's bullets more concise", he: 'קצר את התבליטים בשקופית' },
    { en: 'Check the whole deck for typos and fix them', he: 'בדוק שגיאות כתיב בכל המצגת ותקן אותן' },
  ],
  onCollapseChange: (collapsed) => {
    chrome.webview.postMessage({ kind: collapsed ? 'collapse-pane' : 'expand-pane' })
  },
```
(copied verbatim from `mockup.html`'s `STARTERS.powerpoint.active` set — same "already has content" reasoning as Word's Task 5.)

- [ ] **Step 3: Typecheck and rebuild both**

Run, from each of `ExcelAiAddIn/` and `PowerPointAiAddIn/`:
```bash
npx tsc --noEmit
npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap
```
Expected: 0 tsc errors, esbuild succeeds, for both.

- [ ] **Step 4: Commit**

```bash
git add ExcelAiAddIn/web-src/entry.ts PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(excel,powerpoint): wire starter prompts and pane collapse"
```

---

### Task 7: C# pane-resize plumbing for collapse/expand

**Files:**
- Modify: `WordAiAddIn/TaskPaneHost.cs`
- Modify: `WordAiAddIn/ThisAddIn.cs`
- Modify: `ExcelAiAddIn/TaskPaneHost.cs`
- Modify: `ExcelAiAddIn/ThisAddIn.cs`
- Modify: `PowerPointAiAddIn/TaskPaneHost.cs`
- Modify: `PowerPointAiAddIn/ThisAddIn.cs`

**Interfaces:**
- Consumes: the `collapse-pane`/`expand-pane` WebMessage kinds posted by Tasks 5/6's `onCollapseChange` handlers.
- Produces: nothing new — this is the last task in the chain.

**Design:** `TaskPaneHost` doesn't own the `CustomTaskPane` object (`ThisAddIn` does, so it can add it to `this.CustomTaskPanes`) — so resizing it needs an event flowing from `TaskPaneHost` back up to `ThisAddIn`, the reverse direction of Word's existing `WindowSelectionChange` flow (which flows `ThisAddIn` → `TaskPaneHost`). `PANE_WIDTH_EXPANDED = 420` matches the existing `_taskPane.Width = 420;` startup value already hardcoded identically in all three `ThisAddIn.cs` files; `PANE_WIDTH_COLLAPSED = 34` matches `chat-ui.css`'s existing `.ai-dock.collapsed { width: 34px; }` rule so the native pane width and the CSS-collapsed rail width agree exactly (no leftover blank strip or clipped rail).

- [ ] **Step 1: Add a `RequestPaneWidth` event and `collapse-pane`/`expand-pane` handling to each `TaskPaneHost.cs`**

In `WordAiAddIn/TaskPaneHost.cs`, add a public event near the top of the class (after the existing `private string _chatId;` field):
```csharp
        public event Action<int> RequestPaneWidth;
```
(add `using System;` to the file's `using` list if not already present — check the existing `using` block first; `WordAiAddIn/TaskPaneHost.cs` currently starts with `using System.Diagnostics;` etc. without a plain `using System;`, so add it as the first line.)

Add two case labels to the `OnOtherMessage` switch, after the existing `case "set-mode":` block:
```csharp
                case "collapse-pane":
                    RequestPaneWidth?.Invoke(34);
                    break;
                case "expand-pane":
                    RequestPaneWidth?.Invoke(420);
                    break;
```

Apply the identical event declaration and the identical two `case` blocks to `ExcelAiAddIn/TaskPaneHost.cs` and `PowerPointAiAddIn/TaskPaneHost.cs` (same widths, same event name, same `using System;` check).

- [ ] **Step 2: Subscribe to `RequestPaneWidth` in each `ThisAddIn.cs`**

In `WordAiAddIn/ThisAddIn.cs`, add a line in `ThisAddIn_Startup` right after `_taskPane.Visible = true;`:
```csharp
            _taskPaneControl.RequestPaneWidth += width => _taskPane.Width = width;
```
(No corresponding unsubscribe is needed in `ThisAddIn_Shutdown` — `_taskPaneControl` and `_taskPane` are both torn down together at add-in shutdown, unlike the `Application.WindowSelectionChange` event which is owned by the long-lived `Application` object and must be explicitly unhooked.)

Apply the identical line to `ExcelAiAddIn/ThisAddIn.cs` and `PowerPointAiAddIn/ThisAddIn.cs`'s `ThisAddIn_Startup`, in the same position (right after their own `_taskPane.Visible = true;` line).

- [ ] **Step 3: Build all three and manually verify**

Run, from `C:/dev/officeoffice`:
```bash
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```
Expected: all three build with 0 warnings/errors.

Manually verify in each app: clicking the header's collapse chevron shrinks the entire docked pane to a narrow rail (no blank space beside it) showing the vertical "Airchat Office" label; clicking the rail restores the pane to its original 420px width with the full chat UI visible again.

- [ ] **Step 4: Commit**

```bash
git add WordAiAddIn/TaskPaneHost.cs WordAiAddIn/ThisAddIn.cs ExcelAiAddIn/TaskPaneHost.cs ExcelAiAddIn/ThisAddIn.cs PowerPointAiAddIn/TaskPaneHost.cs PowerPointAiAddIn/ThisAddIn.cs
git commit -m "feat: resize the native task pane to a rail on collapse, restore on expand"
```

---

### Task 8: RTL round-trip test — Hebrew message sent, Hebrew reply received, both correctly aligned

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`

**Interfaces:**
- Consumes: `setLang`'s Hebrew switch and the `.ai-dock[dir]` attribute wiring (Task 2 Step 4/5, Task 4 Step 3), `addUserMessage`/`beginAssistantMessage`/`endAssistantMessage` (pre-existing).
- Produces: nothing new — this is a test-only task, run last since it depends on every other task in this plan being complete.

**Scope:** `jsdom` (vitest's test environment) does not compute real visual layout, so this test cannot verify pixel-level alignment — it verifies the two things that actually drive correct RTL rendering in a real WebView2: (1) the `dir="rtl"` attribute lands on the actual ancestor element the CSS's `[dir='rtl'] .ai-msg-user` / `[dir='rtl'] .ai-msg-assistant` selectors key off of (`.ai-dock`, confirmed in Task 4 Step 3), so the browser's real layout engine applies the mockup-approved alignment rules; and (2) a full send-and-receive round trip with actual Hebrew text renders both bubbles' text content correctly (no mangling/truncation of the Hebrew string, which would indicate an encoding bug independent of layout). Real pixel-level visual confirmation is a manual verification step, listed below, to run once in an actual WebView2-hosted add-in — consistent with this project's established pattern of pairing an automated test with one manual check for anything a headless DOM can't fully verify.

- [ ] **Step 1: Write the failing test**

Add to `shared/chat-ui/chat-ui.test.ts`:
```typescript
  it('RTL round trip: switching to Hebrew, sending a Hebrew message, and receiving a Hebrew reply all render inside the RTL-marked container', () => {
    const { root, handle } = setup()

    // Switch to Hebrew via Settings -> Save (language only takes effect on Save).
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('.ai-dock')!.getAttribute('dir')).toBe('rtl')

    // Send a Hebrew user message.
    const hebrewQuestion = 'סכם את המסמך הזה בבקשה'
    handle.addUserMessage(hebrewQuestion)

    // Simulate a streamed Hebrew assistant reply.
    handle.beginAssistantMessage()
    const hebrewAnswer = 'המסמך עוסק בתקציב הרבעוני ובתחזית המכירות.'
    handle.updateAssistantMessage(hebrewAnswer.slice(0, 10))
    handle.endAssistantMessage(hebrewAnswer)

    // Both messages must render intact (no mangling) and inside the
    // RTL-marked ancestor, so the CSS's [dir='rtl'] .ai-msg-user /
    // [dir='rtl'] .ai-msg-assistant alignment rules actually apply to them.
    const userMsg = root.querySelector<HTMLElement>('[dir="rtl"] .ai-msg-user')
    const assistantMsg = root.querySelector<HTMLElement>('[dir="rtl"] .ai-msg-assistant')
    expect(userMsg).not.toBeNull()
    expect(assistantMsg).not.toBeNull()
    expect(userMsg!.textContent).toBe(hebrewQuestion)
    expect(assistantMsg!.textContent).toBe(hebrewAnswer)

    // Switching back to English moves the container back to LTR and the
    // already-rendered Hebrew messages remain readable (dir is a container
    // attribute, not per-message - Unicode bidi still renders Hebrew
    // characters correctly regardless of the container's base direction).
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="en"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('.ai-dock')!.getAttribute('dir')).toBe('ltr')
    expect(root.querySelector('.ai-msg-user')!.textContent).toBe(hebrewQuestion)
  })
```

- [ ] **Step 2: Run the test to verify it fails (before this plan's earlier tasks) or passes (after)**

Run: `cd shared/chat-ui && npx vitest run -t "RTL round trip"`
Expected: PASS, since Tasks 2/4 (already completed earlier in this same plan) provide `dir` attribute wiring and Hebrew string handling. If this task is somehow executed out of order before Tasks 2/4, it will correctly FAIL — do not skip Tasks 2/4 to make this pass artificially.

- [ ] **Step 3: Run the full suite**

Run: `cd shared/chat-ui && npx vitest run`
Expected: all tests pass (this project's full `chat-ui.test.ts` suite, including every test added across Tasks 2-4 and this task).

- [ ] **Step 4: Manual verification (real WebView2, not just jsdom)**

In each of Word/Excel/PowerPoint, after rebuilding: switch Settings to Hebrew and Save, send a Hebrew message, and confirm visually that (a) the user bubble aligns to the left with its rounded-corner "tail" on the left (per `chat-ui.css`'s `[dir='rtl'] .ai-msg-user` rule), (b) the assistant's reply text flows right-to-left and reads correctly, (c) the panel header stays pinned left-title/right-icons (unaffected by RTL, per the existing `.ai-panel-header { direction: ltr }` rule), and (d) the send button and composer mirror correctly (arrow icon flips, scope hint/mode button sit on the visually correct side). This is the pixel-level check jsdom cannot perform.

- [ ] **Step 5: Commit**

```bash
git add shared/chat-ui/chat-ui.test.ts
git commit -m "test(chat-ui): add RTL round-trip coverage for Hebrew send/receive"
```
