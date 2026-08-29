# PP-1: Task Pane in Every Document Window — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-1 (P0).

> **Dependency on PP-0 (added 2026-08-24):** `2026-08-23-pp00-shared-app-shell.md` deliberately **deferred its C# extraction (Tasks 7-9: `RibbonBase`, `PaneHostBase` in `OfficeAi.Shared`) to this plan**, to avoid two agents rewriting the same six files (`*/ThisAddIn.cs`, `*/TaskPaneHost.cs`). **Before Task 1 below, do PP-0 Tasks 7-9 first** if they have not already landed (check `docs/superpowers/plans/STATUS.md`) — build `RibbonBase`/`PaneHostBase` in `OfficeAi.Shared`, then have each app's `Ribbon.cs`/`TaskPaneHost.cs` subclass them. Tasks 1-4 below then implement the per-window pane registry **as an addition to `ThisAddIn.cs`** (which does not move into a base class — `Globals.ThisAddIn` is VSTO-generated per project, so the registry itself can't live in shared code without a generic wrapper; see the note at the end of Task 1). If PP-0 Tasks 7-9 are somehow already done by someone else, adapt the per-window logic below onto the existing `PaneHostBase`/`RibbonBase` rather than recreating them.
>
> **Also fold in:** PP-0 Task 9 asked this plan to consider generalizing the pane registry itself into `OfficeAi.Shared` as `PaneRegistry<TWindow>`, parameterized by delegates for "get hwnd from window" and "get document from window" (since `Word.Window`/`Excel.Window`/`PowerPoint.DocumentWindow` share no common interface). Task 1 below still writes the registry inline per-app first (it needs to exist and be verified in at least one app before a generic shape is worth committing to); if the three copies in Tasks 1-3 turn out identical apart from types, extract the generic at that point rather than guessing its shape up front.

**Goal:** Make the Airchat Office pane reachable from *every* Word/Excel/PowerPoint document window in a running Office instance — not only the window that happened to be active during `ThisAddIn_Startup` — with each window getting its own pane, its own chat history (keyed by that window's document), and its own editing mode.

**Architecture:** All three add-ins hold a single `CustomTaskPane _taskPane` + single `TaskPaneHost _taskPaneControl` created once in `Startup` (`WordAiAddIn/ThisAddIn.cs:11-20`, identically `ExcelAiAddIn/ThisAddIn.cs:8-16`, `PowerPointAiAddIn/ThisAddIn.cs:8-16`). VSTO's `CustomTaskPanes.Add` has a 3-arg overload `Add(control, title, window)` that binds a pane to a specific window; the 2-arg overload used today implicitly binds to whatever window is active at that instant. The fix is a **pane registry**: a `Dictionary<hwnd, PaneEntry>` plus a window-activation hook that lazily creates a pane the first time each window is seen, and disposes it when that window's document closes.

Per-window panes are chosen over reparenting one pane, because `ChatStore.ChatIdForFile` is already per-document (`WordAiAddIn/TaskPaneHost.cs:49-62`) and `TaskPaneHost` owns a WebView2 whose in-page `AgentLoop` history is per-control — reparenting one control would silently mix two documents' conversations into one transcript. Each pane hosting its own WebView2 costs memory; Task 6 caps that with lazy creation (pane created on first *activation* of a window, not on document open) and disposal on close.

Two correctness consequences fall out of having more than one live pane, and are in scope because the feature is broken without them:
- `TaskPaneHost.GetChatId()` reads `Application.ActiveDocument` — with N panes, a background pane would key its history off whatever document is active *now*. It must key off the document it was created for.
- `WordTools.Mode` / `ExcelTools.Mode` / `PowerPointTools.Mode` are `static` — a mode change in one window would silently change every window's mode. Mode becomes per-document.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, VSTO COM interop (`Microsoft.Office.Tools.CustomTaskPane`), matching every other file in this repo.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Never touch COM in a `TaskPaneHost` constructor. The comment at `WordAiAddIn/TaskPaneHost.cs:29-40` documents a confirmed repro: reading `Application.ActiveDocument` inside the constructor (which runs inside `CustomTaskPanes.Add`) kills the whole VSTO connection with no exception, no error entry — just `Connect=False` forever. Pass the owning document **in** as a constructor argument (the caller already holds it as a COM object from an event parameter); store the reference, never dereference it in the constructor.
- Every COM event sink must be wrapped in `try { } catch { }` at its outermost level, following the existing `Application_WindowSelectionChange` pattern (`WordAiAddIn/ThisAddIn.cs:45-56`) — an exception escaping a COM event sink silently disconnects the add-in.
- No automated tests for VSTO/COM host code (existing project convention) — verification is build + the manual matrix in Task 7.
- Do not change the WebView2 message contract (`chrome.webview.postMessage` kinds); this plan is host-side only.
- Keep the pane title `"Airchat Office"` and default width `420` exactly as today.

---

### Task 1: Word — pane registry and per-window creation

**Files:**
- Modify: `WordAiAddIn/ThisAddIn.cs`

**Interfaces:**
- Produces: `private PaneEntry EnsurePaneFor(Word.Window window)` — consumed by Task 4's `TogglePane()` and copied in shape by Tasks 2 and 3.

- [ ] **Step 1: Replace the single-pane fields with a registry**

Replace `private CustomTaskPane _taskPane;` / `private TaskPaneHost _taskPaneControl;` with:

```csharp
private sealed class PaneEntry
{
    public CustomTaskPane Pane;
    public TaskPaneHost Control;
}

// Keyed by the window's Hwnd rather than the Word.Window RCW: reference
// equality on an RCW is not reliable across separate COM calls, while Hwnd is
// a stable int, unique per top-level document window.
private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();
```

- [ ] **Step 2: Extract pane creation into a method taking the target window**

```csharp
private PaneEntry EnsurePaneFor(Word.Window window)
{
    int hwnd = window.Hwnd;
    PaneEntry existing;
    if (_panes.TryGetValue(hwnd, out existing)) return existing;

    TaskPaneHost control = new TaskPaneHost(window.Document, hwnd);
    CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Airchat Office", window);
    pane.Width = 420;
    pane.Visible = true;

    PaneEntry entry = new PaneEntry { Pane = pane, Control = control };
    control.RequestPaneWidth += width => ApplyPaneWidth(pane, width);
    _panes[hwnd] = entry;
    return entry;
}
```

`ApplyPaneWidth(CustomTaskPane pane, int width)` is the existing lambda body from `ThisAddIn.cs:19-38` lifted verbatim into a private method — same docked-position guard, same silent `catch`, now taking the pane as a parameter instead of closing over the field.

- [ ] **Step 3: Wire window activation**

In `ThisAddIn_Startup`, replacing the old inline creation:

```csharp
this.Application.WindowActivate += Application_WindowActivate;
this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
this.Application.WindowSelectionChange += Application_WindowSelectionChange;  // unchanged
EnsurePaneFor(this.Application.ActiveWindow);  // the startup window, as today
```

```csharp
private void Application_WindowActivate(Word.Document doc, Word.Window window)
{
    try { EnsurePaneFor(window); }
    catch { /* pane creation is best-effort; never break the add-in connection */ }
}
```

`WindowActivate` is the single hook covering every path (File > Open, File > New, a file double-clicked while the app runs, and `View > New Window`) — each newly-created window fires it as it becomes active, so `DocumentOpen`/`NewDocument` do **not** also need wiring. Confirm this in Task 7's matrix; if a listed path is observed *not* to fire `WindowActivate`, add `DocumentOpen`/`NewDocument` handlers calling `EnsurePaneFor(doc.ActiveWindow)` as a fallback rather than replacing this hook.

- [ ] **Step 4: Dispose panes when their document closes**

```csharp
private void Application_DocumentBeforeClose(Word.Document doc, ref bool cancel)
{
    try
    {
        var toRemove = new List<int>();
        foreach (Word.Window w in doc.Windows) toRemove.Add(w.Hwnd);
        foreach (int hwnd in toRemove)
        {
            PaneEntry entry;
            if (!_panes.TryGetValue(hwnd, out entry)) continue;
            _panes.Remove(hwnd);
            entry.Pane.Visible = false;
            this.CustomTaskPanes.Remove(entry.Pane);
            entry.Control.Dispose();
        }
    }
    catch { }
}
```

The two-pass shape (collect hwnds, then mutate) avoids mutating `_panes` while enumerating a COM collection that may itself change.

- [ ] **Step 5:** Unsubscribe `WindowActivate` and `DocumentBeforeClose` in `ThisAddIn_Shutdown`, alongside the existing `WindowSelectionChange -=`.

- [ ] **Step 6:** `Application_WindowSelectionChange` currently calls `_taskPaneControl.OnSelectionChanged(selection)`. Route it to the pane owning the selection's window instead:

```csharp
private void Application_WindowSelectionChange(Word.Selection selection)
{
    try
    {
        PaneEntry entry;
        if (_panes.TryGetValue(selection.Document.ActiveWindow.Hwnd, out entry))
            entry.Control.OnSelectionChanged(selection);
    }
    catch { }
}
```

Without this, a selection in window B pushes selected text into window A's chat context — a silent cross-document data leak into the prompt.

**Verification:**
- [ ] `MSBuild WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds.
- [ ] Manual smoke: open Word, File > Open a second document, confirm both windows show the pane. Full matrix deferred to Task 7.

---

### Task 2: Excel — same registry, Excel event surface

**Files:**
- Modify: `ExcelAiAddIn/ThisAddIn.cs`

- [ ] **Step 1:** Port Task 1's registry verbatim, substituting Excel types: key on `Excel.Window.Hwnd`, construct `new TaskPaneHost((Excel.Workbook)window.Parent, hwnd)`.

- [ ] **Step 2:** Wire the events. Excel's `WindowActivate` signature carries both objects:

```csharp
this.Application.WindowActivate += Application_WindowActivate;          // (Excel.Workbook, Excel.Window)
this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose; // (Excel.Workbook, ref bool)
EnsurePaneFor(this.Application.ActiveWindow);
```

Disposal iterates `wb.Windows` exactly as Task 1 Step 4 iterates `doc.Windows`.

- [ ] **Step 3:** Excel has no existing selection-change hookup (`ExcelAiAddIn/ThisAddIn.cs` wires none) — nothing to re-route, skip Task 1 Step 6's equivalent.

**Verification:** `MSBuild ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds.

---

### Task 3: PowerPoint — same registry, PowerPoint event surface

**Files:**
- Modify: `PowerPointAiAddIn/ThisAddIn.cs`

- [ ] **Step 1:** Port Task 1's registry, keyed on `PowerPoint.DocumentWindow.HWND` (PowerPoint's PIA spells it all-caps; confirm at build time), constructing `new TaskPaneHost(window.Presentation, hwnd)`.

- [ ] **Step 2:** Wire `Application.WindowActivate(PowerPoint.Presentation, PowerPoint.DocumentWindow)` and `Application.PresentationClose(PowerPoint.Presentation)` for disposal.

PowerPoint also exposes `AfterPresentationOpen`/`AfterNewPresentation`; as in Task 1, prefer `WindowActivate` alone and add the others only if Task 7 shows a path it misses.

**Verification:** `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds.

---

### Task 4: Ribbon toggle targets the active window's pane

**Files:**
- Modify: `WordAiAddIn/ThisAddIn.cs`, `ExcelAiAddIn/ThisAddIn.cs`, `PowerPointAiAddIn/ThisAddIn.cs`

**Problem:** `TogglePane()` (`WordAiAddIn/ThisAddIn.cs:66-72` plus the two identical copies at `:49` in the other add-ins) flips `_taskPane.Visible` on the one startup-bound pane. From any other window the ribbon button appears to do nothing — the second half of the reported symptom.

- [ ] **Step 1:** Rewrite in all three:

```csharp
public void TogglePane()
{
    try
    {
        PaneEntry entry = EnsurePaneFor(this.Application.ActiveWindow);
        entry.Pane.Visible = !entry.Pane.Visible;
    }
    catch { }
}
```

Routing through `EnsurePaneFor` means the button also *recovers* a window that somehow never got a pane, instead of no-opping.

**Verification:** all three build; toggling from the second window hides/shows only that window's pane.

---

### Task 5: Per-document chat id and per-document editing mode

**Files:**
- Modify: `WordAiAddIn/TaskPaneHost.cs`, `ExcelAiAddIn/TaskPaneHost.cs`, `PowerPointAiAddIn/TaskPaneHost.cs`
- Modify: `WordAiAddIn/WordTools.cs`, `ExcelAiAddIn/ExcelTools.cs`, `PowerPointAiAddIn/PowerPointTools.cs`

**Problem:** with one pane these two globals were harmless; with N panes they are cross-talk bugs.

- [ ] **Step 1: Constructor takes the owning document + hwnd**

`TaskPaneHost` gains `private readonly Word.Document _document;` and `private readonly int _hwnd;`, set from new constructor parameters. `GetChatId()` (`WordAiAddIn/TaskPaneHost.cs:49-62`) uses `_document` instead of `Globals.ThisAddIn.Application.ActiveDocument`. Everything else about the method is unchanged, **except** the unsaved fallback: `"unsaved-" + Process.GetCurrentProcess().Id` is now ambiguous across two unsaved documents in one process, so make it `"unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd`. Keep the lazy-evaluation shape — the constructor still must not dereference `_document`.

- [ ] **Step 2: Mode becomes per-document**

Replace `public static EditingMode Mode = EditingMode.FullAutonomy;` (`WordAiAddIn/WordTools.cs:19`, and the equivalents in the other two) with:

```csharp
private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

public static void SetMode(string docKey, EditingMode mode) { ModeByDoc[docKey] = mode; }

private static EditingMode ModeFor(string docKey)
{
    EditingMode m;
    return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
}
```

`docKey` is the same string `GetChatId()` produces (already unique per document, already computed lazily).

- [ ] **Step 3: Thread the key to `Execute` without changing the shared delegate**

`Execute` reads `Mode` at `WordAiAddIn/WordTools.cs:28-55`. Change its signature to `Execute(string docKey, string name, JsonElement input)` and adapt at the call site with a closure, leaving `OfficeAi.Shared/ToolProtocol.cs`'s `ToolExecutor` delegate and `OfficeAi.Shared.Tests` untouched:

```csharp
_bridge = new WebViewBridgeHost(this, (n, i) => WordTools.Execute(GetChatId(), n, i),
                                "WordAiAddIn", UpdateStatus, OnOtherMessage);
```

- [ ] **Step 4:** In each `TaskPaneHost.OnOtherMessage`'s `"set-mode"` branch, call `WordTools.SetMode(GetChatId(), EditingMode.X)` instead of assigning the static field.

- [ ] **Step 5 (known limitation, document it):** the COM executors resolve their target via `Globals.ThisAddIn.Application.ActiveDocument` / `ActiveWorkbook` / `ActivePresentation` (e.g. `WordAiAddIn/WordTools.cs:98`). A tool call is always initiated by a user in the focused window, so the active document is normally the right one — but a long-running run whose user switches windows mid-run would write into the newly-active document. Add a comment at each `ActiveDoc`-style property recording this, and leave the fix out of scope (it needs per-document COM target resolution across every executor method, which is a separate, much larger change). Note it in Task 7's report.

**Verification:**
- [ ] `OfficeAi.Shared.Tests` still passes unchanged — this task must not alter `ToolProtocol`'s public shape.
- [ ] All three add-ins build.
- [ ] Manual: set window A to Read only and window B to Full autonomy; an edit request in B succeeds while the same request in A is refused.

---

### Task 6: Lazy creation and disposal audit

**Files:**
- Modify: all three `ThisAddIn.cs`; possibly `OfficeAi.Shared/WebViewBridgeHost.cs`

- [ ] **Step 1:** Confirm `EnsurePaneFor` is reachable only from `WindowActivate`, `TogglePane`, and the single startup call — so a document open but never activated pays no WebView2 cost. Add a one-line comment saying so, so a later refactor doesn't move creation to `DocumentOpen`.
- [ ] **Step 2:** Confirm `entry.Control.Dispose()` actually tears down WebView2. Read `OfficeAi.Shared/WebViewBridgeHost.cs`: if it holds a `WebView2` control that is parented into `TaskPaneHost.Controls`, WinForms disposal already covers it; if not, implement `IDisposable` on `WebViewBridgeHost` and call it from an override of `TaskPaneHost.Dispose(bool disposing)`.
- [ ] **Step 3:** Each WebView2 also holds a user-data folder (see `WebViewBridgeHost`'s `userDataFolder`). Confirm N panes in one process share one folder without lock contention; if a second pane fails to initialize with a user-data-folder lock error, give each pane its own subfolder keyed by hwnd and record that in the file's comments.

**Verification:** open and close 10 documents in one session; `msedgewebview2.exe` process count returns to its baseline, and pane #2 onward initializes to `ready` rather than showing a WebView2 init error in the status label.

---

### Task 7: Manual verification matrix

- [ ] For each of Word, Excel, PowerPoint, with the add-in installed:
  1. Launch the app with no document → pane visible in the startup window.
  2. File > Open an existing file → pane visible in the new window.
  3. File > New → pane visible.
  4. Double-click a file in Explorer while the app is already running → pane visible.
  5. `View > New Window` on an existing document (Word/Excel) → pane visible in the clone.
  6. Ribbon "Airchat Office" button in each window hides/shows only that window's pane.
  7. Send a message in window A, then in window B → each transcript stays in its own window; closing and reopening file A restores only A's history.
  8. Set different editing modes in A and B → each is enforced independently.
  9. Select text in window B → the selection context appears only in B's pane.
  10. Close a document → its pane goes away; other windows' panes keep working.
- [ ] For any path where the pane fails to appear, add a temporary `Debug.WriteLine` in `Application_WindowActivate`, read it in DebugView, and record whether the event fired — before adding fallback event hooks.
- [ ] Report the Task 5 Step 5 limitation (active-document tool targeting during window switches) as a known issue for a follow-up item.
