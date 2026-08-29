# FT-2: Selection Context for Excel and PowerPoint

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source:** feature request, 2026-08-23. Sibling to Word's existing selection support.

**Goal:** What the user has selected reaches the model in Excel and PowerPoint the way it already does in Word — the scope-hint pill in the composer reflects it, and each run's per-turn context names it — so "make this bold", "chart this", "fix this slide" resolve to what the user is actually looking at.

## Why this fits — and fits better than it does in Word

Word's version is a workaround. Its tools address paragraphs by index while a text selection is a character range, so the two vocabularies do not line up: `buildContext()` injects up to 24,000 characters of raw selected **content** (`WordAiAddIn/TaskPaneHost.cs:64-77` → `WordAiAddIn/web-src/entry.ts:270-271`), and `apply_commands` needed a `Target.scope: "selection"` escape hatch to act on it.

In the other two apps the selection is already expressed in exactly the vocabulary the tools take:

- **Excel** tools take A1 addresses — `read_range`, `format_range`, `add_chart`, all of them. `Application.SheetSelectionChange(Sh, Target)` hands over `Target.Address`, the same string. Inject `Sheet1!B2:D40` and every existing tool is immediately usable against it.
- **PowerPoint** tools take `slideIndex` + `shapeIndex` and nothing else — `ResolveShape` is two lookups (`PowerPointAiAddIn/PowerPointTools.cs:145-150`). `Application.WindowSelectionChange(Sel)` gives `Sel.Type` (`ppSelectionSlides` / `ppSelectionShapes` / `ppSelectionText` / `ppSelectionNone`), `Sel.SlideRange`, and `Sel.ShapeRange`. "Shape 3 on slide 2 is selected" *is* a tool call's parameters.

**No tool schema changes anywhere in this plan.** That is the main reason it is small.

The payload therefore differs per app by design: Word sends **content**, Excel sends an **address**, PowerPoint sends **indices**. Excel must never send cell values — a column selection is over a million cells, and `read_range` already exists with a 2000-cell cap for exactly this.

## What already exists

The UI half is in the shared component and only Word feeds it: `setSelectionScope` (`shared/chat-ui/chat-ui.ts:66,389`), the `.ai-scope-hint` pill (`:122`), `refreshScopeHint()` (`:196-203`), and a passing test (`chat-ui.test.ts:100`). Excel and PowerPoint simply never call it.

PP-0 moves the `selection-changed` handling into the shared shell unconditionally, so after PP-0 the TypeScript receive-side is already done for all three apps. The remaining work is: emit the event from C#, define each payload, and write the context string.

Excel is partly there through a different door — `get_workbook_context` already reports the selection address (`ExcelAiAddIn/ExcelTools.cs:125`) — but only when the model happens to call that tool, and it never reaches the pill.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (COM event sinks); TypeScript in `shared/web-src/app-shell/` and `shared/chat-ui/`.

## Dependencies

- **PP-0 (`2026-08-23-pp00-shared-app-shell.md`)** — Task 1 Step 4 already moves Word's selection block into the shell. Without it this is three copies of the receive-side.
- **PP-1 (`2026-08-23-pp01-taskpane-per-window.md`)** — with per-window panes the event must reach the pane for the window it happened in. PP-1 Task 1 Step 6 does this for Word; Tasks 2 and 3 here do the same for Excel and PowerPoint.
- Independent of PP-2/3/4 and FT-1.

## Global Constraints

- **Debounce every selection event.** This is the one thing that decides whether the feature feels good. `SheetSelectionChange` fires on *every* arrow-key press and PowerPoint's fires on every shape click during ordinary editing — far more often than Word's text-selection events. Undebounced, this posts a WebMessage across the WebView2 bridge on every keystroke.
- **Never send cell values from Excel.** Address, dimensions, and count only. The model calls `read_range` when it needs data.
- Cap every string that crosses the bridge. Word already caps its text at 24,000 chars (`TaskPaneHost.cs:68`); PowerPoint's shape-text preview needs a much smaller cap.
- Every COM event sink is wrapped in an outer `try { } catch { }`, per the existing pattern (`WordAiAddIn/ThisAddIn.cs:45-56`) — an escaping exception silently disconnects the add-in.
- **PowerPoint's COM collections are 1-based; the tools are 0-based.** Every index that crosses the bridge is converted once, at the C# boundary, and is 0-based from there on. This is the single most likely bug in this plan.
- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- No tool schema changes. If a task seems to need one, stop — the payload is wrong.
- Rebuild bundles + MSBuild after any TypeScript change.

---

### Task 1: Shared debounced dispatch

**Files:**
- Modify: `OfficeAi.Shared/PaneHostBase.cs` (from PP-0)

**Interfaces:**
- Produces: `protected void PostSelection(object payload)` — coalesces bursts and drops duplicates; consumed by Tasks 2, 3, and 4.

- [ ] **Step 1: Debounce on the UI thread.** Use `System.Windows.Forms.Timer`, **not** `System.Timers.Timer` — the WinForms timer ticks on the UI thread, where the WebView2 control lives, so no cross-thread marshaling is needed. `PaneHostBase` is a `UserControl`, so the message pump is already there.

```csharp
private readonly Timer _selectionTimer;   // System.Windows.Forms.Timer
private object _pendingSelection;
private string _lastSignature;

protected void PostSelection(object payload, string signature)
{
    if (signature == _lastSignature) return;   // same selection re-reported; drop
    _pendingSelection = payload;
    _selectionTimer.Stop();
    _selectionTimer.Start();                   // restart: fires ~200ms after the burst ends
}
```

On tick: stop the timer, set `_lastSignature`, and post `_pendingSelection` through the bridge.

- [ ] **Step 2: Pick the interval empirically.** Start at 200ms. Verify by holding an arrow key down in Excel for a few seconds: the pill should settle once, not flicker, and the WebView2 should not visibly work. Record the chosen value and why in a comment.
- [ ] **Step 3: The signature** is a cheap string identifying the selection (`"Sheet1!B2:D40"`, `"slides:2,3"`, `"shapes:1/3"`). It suppresses the common case of an event firing without the selection actually changing.
- [ ] **Step 4: Dispose the timer** in the control's dispose path, alongside PP-1 Task 6's WebView2 teardown — a live timer on a disposed pane throws on tick.
- [ ] **Step 5: Do not suppress during a run.** `buildContext()` snapshots at run start (`shared/web-src/agent-core/loop.ts` calls it once per run), so a selection change mid-run cannot affect the in-flight request. Updating the pill while a run is going is correct and desirable.

**Verification:** all three build; Word's existing selection behavior is unchanged apart from now being debounced.

---

### Task 2: Excel — `SheetSelectionChange`

**Files:**
- Modify: `ExcelAiAddIn/ThisAddIn.cs`, `ExcelAiAddIn/TaskPaneHost.cs`

- [ ] **Step 1: Wire the event** in `ThisAddIn_Startup`, unsubscribing in `Shutdown`:

```csharp
this.Application.SheetSelectionChange += Application_SheetSelectionChange;
```

The handler signature is `(object Sh, Excel.Range Target)`. Route to the pane for the window the selection happened in — with PP-1's registry, resolve from `Application.ActiveWindow`; guard for the case where no pane exists for it yet.

- [ ] **Step 2: Build the payload — extent, not just the address**

A bare A1 address is the least legible thing available. `B2:D40` makes the reader do arithmetic to learn it is 39 rows, and `B1:B1048576` tells them nothing at all except that something went wrong. Report the **extent** — how many rows and columns, and which — with the address alongside it rather than instead of it.

```csharp
Excel.Worksheet sheet = Sh as Excel.Worksheet;
string address = Target.Address[false, false];        // "B2:D40", no $ signs
bool multi = Target.Areas.Count > 1;
long cellCount = Target.CountLarge;                    // NOT Count
int rows = Target.Rows.Count;
int cols = Target.Columns.Count;
string firstRow = Target.Row.ToString();               // 2
string firstCol = ColumnLetter(Target.Column);         // "B"
bool entireColumns = Target.Address[false, false].Contains(":") && Target.Rows.Count == sheet.Rows.Count;
bool entireRows = Target.Columns.Count == sheet.Columns.Count;
```

**Use `CountLarge`, not `Count`.** `Range.Count` is an `int` and a whole-sheet selection is ~17 billion cells, which overflows; `CountLarge` returns a `long` (Excel 2007+). Confirm the member exists in this project's PIA at build time — `ExcelTools.cs:22-26` sets the precedent for what to do if a member is missing.

`ColumnLetter(int)` is a small helper (1 → `A`, 27 → `AA`); write it once here rather than parsing letters back out of the address string.

Post `{ kind: "selection-changed", app: "excel", hasSelection, sheet, address, cellCount, rows, cols, firstRow, firstCol, entireColumns, entireRows, multi, effectiveAddress, effectiveCellCount }` — the last two from Step 2b.

- [ ] **Step 2b: Whole-column and whole-row selections — intersect with `UsedRange`**

Clicking a column header selects 1,048,576 cells. Reported literally, that is both useless to the user and actively misleading to the model, which will either try to read a million cells (and hit `read_range`'s 2000-cell cap) or refuse. Intersect the selection with the sheet's used range and report the effective extent alongside the literal one:

```csharp
Excel.Range effective = Globals.ThisAddIn.Application.Intersect(Target, sheet.UsedRange);
string effectiveAddress = effective == null ? null : effective.Address[false, false];
long effectiveCellCount = effective == null ? 0 : effective.CountLarge;
```

`Intersect` returns `null` when there is no overlap (an empty column) — handle that rather than dereferencing it. Selecting column B in a sheet with 200 rows of data then yields an effective `B1:B200` / 200 cells, which is what both the pill and the model should work from.

Compute this only when `entireColumns || entireRows || cellCount > 10_000` — `UsedRange` is not free, and there is no reason to pay for it on an ordinary drag-selection.

- [ ] **Step 3: Single-cell is not "no selection".** In Excel there is *always* a selection — the cursor is always somewhere. Treat a single cell as a real selection (the user clicked A5 and said "explain this formula") but distinguish it in the context string. Only `hasSelection: false` when the selection is not a `Range` at all (a chart or shape is selected, in which case `Sh`/`Target` may not behave as expected — guard for a null cast).
- [ ] **Step 4: Multi-area selections.** Ctrl-click produces `"B2:D4,F1:F9"`, which single-address tools cannot take. Report the full address in the pill (it is what the user sees), and in the context string state that it is multiple areas and name them — the model can then issue one tool call per area.
- [ ] **Step 5: Optional value preview — recommend not.** A tiny sample (the top-left cell) would help the model decide whether to read. It also silently sends spreadsheet contents to the provider on every click, which address-only avoids entirely. Ship address-only; note the option in a comment and revisit if the model turns out to over-call `read_range`.

**Verification:** build; selecting a range updates the pill; holding an arrow key does not flood the bridge.

---

### Task 3: PowerPoint — `WindowSelectionChange`

**Files:**
- Modify: `PowerPointAiAddIn/ThisAddIn.cs`, `PowerPointAiAddIn/TaskPaneHost.cs`

- [ ] **Step 1: Wire** `this.Application.WindowSelectionChange += ...` with signature `(PowerPoint.Selection Sel)`. Route to the owning window's pane via `Sel.Parent` (a `DocumentWindow`) if reachable, otherwise `Application.ActiveWindow`.

- [ ] **Step 2: Branch on `Sel.Type`** and emit one of three payload shapes, converting every index to 0-based:

```csharp
switch (Sel.Type)
{
    case PowerPoint.PpSelectionType.ppSelectionSlides:
        // Sel.SlideRange -> slide indices, 1-based in COM
        // payload: { kind: "slides", slideIndexes: [1, 2] }
    case PowerPoint.PpSelectionType.ppSelectionShapes:
        // Sel.ShapeRange -> shape indices within Sel.SlideRange[1]
        // payload: { kind: "shapes", slideIndex, shapeIndexes, names, textPreview }
    case PowerPoint.PpSelectionType.ppSelectionText:
        // a text cursor inside one shape
        // payload: { kind: "shapeText", slideIndex, shapeIndex, text }
    case PowerPoint.PpSelectionType.ppSelectionNone:
        // payload: { hasSelection: false }
}
```

- [ ] **Step 3: The 0-based conversion.** `Slides[i]` and `Shapes[i]` are 1-based in COM; `ResolveShape` does `Slides[slideIndex + 1].Shapes[shapeIndex + 1]` (`PowerPointTools.cs:149`). So subtract 1 here, once, and never again. Write that in a comment at the conversion site — an off-by-one makes the model edit the wrong shape, which looks like a model failure rather than a plumbing bug.
- [ ] **Step 4: Shape names.** `shape.Name` (e.g. `"Title 1"`, `"Content Placeholder 2"`) is far more meaningful to a model than a bare index. Include the names alongside the indices; they cost nothing and make the context string readable.
- [ ] **Step 5: Text preview, capped.** For shape selections include a short preview of each shape's text (suggest 80 chars per shape, at most 5 shapes) so the model can tell which is which. `TextFrame.HasText` must be checked first — reading `TextRange.Text` on a shape with no text frame throws.
- [ ] **Step 6: `ppSelectionText` is a sub-selection**, not the whole shape. Send the selected run's text, and note in the context that the user selected text *inside* the shape — the distinction matters for "make this bold" (the run) versus "delete this" (the shape).
- [ ] **Step 7: Multi-slide selections** in the slide sorter are common and useful ("delete these three slides"). Report the whole list.

**Verification:** build; selecting a shape, a slide, and text inside a shape each update the pill correctly, with correct 0-based indices.

---

### Task 4: Context strings in the shell

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`, `shared/web-src/app-shell/bridge.ts` (both from PP-0)

**Interfaces:**
- Produces: a `SelectionContext` discriminated union and `AddInConfig.describeSelection?(sel): string`.

- [ ] **Step 1: One union across the three apps**

```ts
export type SelectionContext =
  | { kind: 'none' }
  | { kind: 'text'; preview: string; fullText: string }                                  // Word
  | { kind: 'range'; sheet: string; address: string; cellCount: number; multi: boolean }  // Excel
  | { kind: 'slides'; slideIndexes: number[] }                                            // PowerPoint
  | { kind: 'shapes'; slideIndex: number; shapeIndexes: number[]; names: string[]; textPreview: string[] }
  | { kind: 'shapeText'; slideIndex: number; shapeIndex: number; text: string }
```

- [ ] **Step 2: `describeSelection` per app**, supplied through the config so the shell stays app-agnostic. It must state the addressing vocabulary explicitly, because that is what makes the selection *actionable* rather than merely informative:

  - Excel, ordinary range: `The user has selected Sheet1!B2:D40 - 39 rows (2-40) x 3 columns (B-D), 117 cells. Use this address directly with the range tools; call read_range if you need the values.`
  - Excel, whole columns: `The user has selected all of columns B-D on Sheet1. Only B1:B200 contains data (200 cells) - use that bounded range, not the full column, which exceeds read_range's 2000-cell cap.` Naming the cap is what stops the model burning a turn on a rejected call.
  - Excel, single cell: `The user has selected the single cell Sheet1!C7.` No dimensions — "1 row x 1 column" is noise.
  - PowerPoint shapes: `The user has selected shape 3 ("Revenue chart") on slide 2. These are 0-based indices in the form the tools take (slideIndex, shapeIndex).`
  - PowerPoint slides: `The user has selected slides 2, 3, 4 (0-based).`
  - PowerPoint shape text: `The user has selected text inside shape 1 ("Title 1") on slide 0: "..." — the selection is a run within that shape, not the whole shape.`
  - Word: unchanged wording (`Content selected by the user:\n...`), so no behavior change there.

- [ ] **Step 3: Keep it in `buildContext`,** which `AgentLoop.run()` calls once per run — the same hook Word uses (`WordAiAddIn/web-src/entry.ts:270-271`). Do **not** use `systemSuffix` (that is FT-1's document guidelines, and it is read every turn — a selection that changes mid-run would then leak into later turns of the same run).
- [ ] **Step 4: Empty selection contributes nothing.** `{ kind: 'none' }` returns `''`, exactly as Word does today, so no per-turn cost when there is no selection.
- [ ] **Step 5:** Enable it in the Excel and PowerPoint configs via PP-0's `useSelectionContext` flag, which currently only Word sets.

**Verification:** typechecks; a run started with a selection carries the right sentence in its first user message.

---

### Task 5: Per-app scope-hint wording

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`

- [ ] **Step 1:** The no-selection label is `scopeWholeDoc` — `'Whole document'` / `'כל המסמך'` (`chat-ui.ts:22`) — shown in all three apps. In Excel it should read *Whole sheet* / *כל הגיליון*, in PowerPoint *Whole deck* / *כל המצגת*. Add both strings and let the mount options pick the key.
- [ ] **Step 2:** The selection label is built inline as `'Selection: "' + preview + '..."'` with a hardcoded Hebrew variant (`:198-199`). That shape assumes text. Add `scopeSelectionPrefix` to `STRINGS` and let the caller pass a ready-made short label instead — `B2:D40 (39×3)`, `Slide 2, shape 3`, or Word's quoted text excerpt.
- [ ] **Step 3:** Widen `setSelectionScope`'s parameter from `{ hasSelection, preview }` to also accept a `label` the shell computes, keeping `preview` for Word's existing quoted form. Update the existing test rather than replacing it.
- [ ] **Step 4:** Do not put the address in a quoted-excerpt format — `Selection: "B2:D40..."` reads like truncated text. The trailing ellipsis belongs only to the text case.

- [ ] **Step 5: Excel label formats.** The pill is inside a 420px panel, so the label has to earn its width. Lead with the extent, keep the address as the qualifier:

| Selection | Label |
|---|---|
| `B2:D40` | `B2:D40 · 39×3` |
| single cell `C7` | `C7` |
| whole column B | `column B · 200 rows with data` |
| whole columns B:D | `columns B–D · 200×3 with data` |
| whole row 5 | `row 5` |
| multi-area | `2 areas · 156 cells` |

`39×3` is rows×columns, in that order, matching how Excel itself reports a drag selection in the status bar — so it is already the notation the user knows. Localize the words (`column`, `rows with data`, `areas`) through `STRINGS`; the address and the `39×3` figure stay as-is in both languages.

- [ ] **Step 6: Fall back to the literal extent when `UsedRange` gives nothing.** An entire empty column has no intersection (Step 2b's `Intersect` returns `null`); label it `column B · empty` rather than showing 1,048,576 rows.

**Verification:** `cd shared/chat-ui && npx vitest run`; the pill reads sensibly in each app in both languages.

---

### Task 6: Tests

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`

- [ ] **Step 1:** `setSelectionScope` with a range label renders `Selection: B2:D40` with no quotes or ellipsis; with Word's text preview it renders the existing quoted form (regression guard on `chat-ui.test.ts:100`).
- [ ] **Step 2:** `null` reverts to the per-app whole-scope label, and the right one for each configured app.
- [ ] **Step 3:** The pill relocalizes on a language switch while a selection is active — `refreshScopeHint()` is already called from `setLang` (`:218`), so this is a guard, not new work.
- [ ] **Step 4:** If the shell's `describeSelection` implementations are exported, unit-test each union member's sentence — particularly that PowerPoint's are 0-based. That is the cheapest possible catch for the off-by-one.

**Verification:** `npx vitest run` — all green.

---

### Task 7: Manual verification matrix

**Excel**
- [ ] Select `B2:D40` → pill reads `Selection: B2:D40 · 39×3`; "make this bold" bolds exactly that range with no clarifying question.
- [ ] Select a single cell with a formula → pill reads `Selection: C7` (no `1×1`); "explain this" explains that cell.
- [ ] Select a whole column with 200 rows of data → pill reads `column B · 200 rows with data`, not a million-row address; no hang, no values sent; `cellCount` is correct (the `CountLarge` check).
- [ ] Ask the model to summarize that whole-column selection → it reads the bounded `B1:B200`, not the full column, and does not hit `read_range`'s 2000-cell cap.
- [ ] Select an entirely empty column → pill reads `column B · empty`; nothing crashes on the null `Intersect`.
- [ ] Select whole columns B:D → `columns B–D · 200×3 with data`.
- [ ] Select a whole row → `row 5`.
- [ ] Select column ZZ (beyond `UsedRange`) → the empty-intersection path, not an exception.
- [ ] Ctrl-click two areas → pill reads `2 areas · 156 cells`; the model handles them as two ranges.
- [ ] Verify `UsedRange` is not computed for an ordinary small drag-selection (Step 2b's threshold) — add a temporary counter if needed.
- [ ] Hold an arrow key across 50 cells → the pill settles once; the panel stays responsive.
- [ ] Select a chart object → no crash, no bogus range.
- [ ] Switch sheets → the pill follows the new sheet's selection.

**PowerPoint**
- [ ] Click a shape → pill names it; "make this text bigger" targets that shape with correct indices.
- [ ] Select three shapes → all reported; a styling request applies to all three.
- [ ] Select slides in the sorter → "delete these" targets the right slides (needs PP-19's `delete_slide`).
- [ ] Select text inside a shape → context distinguishes the run from the shape.
- [ ] Click empty canvas → reverts to `Whole deck`.
- [ ] **Off-by-one check:** with a shape on slide 3 selected, confirm the edit lands on slide 3 and not slide 2 or 4.

**Both / cross-cutting**
- [ ] Word's existing behavior is unchanged.
- [ ] With PP-1's per-window panes, a selection in window B updates only B's pill.
- [ ] Selection changing mid-run does not alter the in-flight request.
- [ ] Hebrew UI: the pill reads correctly and the panel stays RTL.
