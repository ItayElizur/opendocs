# PP-15: Excel Chart Types and Data Rebinding — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-15 (P1) — three related findings.

**Goal:** The full 6-type chart vocabulary is available *and documented* on creation, not just on editing; and an existing chart can be repointed at a new or extended data range instead of having to be deleted and recreated.

**Architecture:** Three defects, all in the gap between `AddChart` and `EditChartExcel`:

1. **`AddChart` ignores the map it should be using.** `ExcelChartTypeMap` (`ExcelAiAddIn/ExcelTools.cs:58-66`) holds six entries — column/bar/line/area/pie/doughnut. `AddChart` (`:643-663`) does not consult it; it hardcodes a three-way ternary at `:656`: `t == "line" ? 4 : t == "pie" ? 5 : 51`. Everything else silently becomes a clustered column.
2. **`edit_chart`'s real breadth is undocumented.** `EditChartExcel` (`:665-726`) *does* use the map (`:672`), so it accepts all six — but the schema only ever documents `add_chart`'s narrower three (`ExcelAiAddIn/web-src/entry.ts:210`), so nothing tells the model bar/area/doughnut exist.
3. **No rebinding.** `EditChartExcel` changes type, title, legend, data labels, series colors, and series names — but never touches `SetSourceData`. A chart cannot follow a growing table.

There is also a fourth, unlisted issue worth folding in: `edit_chart` addresses charts by `chartPath` = the ChartObject's *name* (`:667-669`), and `AddChart` never returns the name of the chart it just created (`:643-663` returns nothing). So the model must guess `"Chart 1"` to edit what it just made. Task 4 fixes that; without it, the rebinding feature is hard to actually use.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Excel` (+ `dynamic` for the chart engine, as the file already does); TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **One chart vocabulary across the repo.** `ExcelChartTypeMap`, `PptChartTypeMap` (`PowerPointAiAddIn/PowerPointTools.cs:430-437`), and any Word map added by PP-9 use the same names for the same xlChartType codes. Note the existing discrepancy: `PptChartTypeMap` maps `"bar"` to 51 (xlColumn**Clustered**), which is wrong — Excel's map has it as 57 (xlBarClustered). Do not "harmonize" Excel down to PowerPoint's bug; PP-21/PP-22 fixes the PowerPoint side.
- No silent fallbacks. An unrecognized `chartType` is an error naming the value and listing valid ones — replacing the current fallback-to-column in both `AddChart` and (implicitly) `EditChartExcel`, where a bad value simply fails the `TryGetValue` and is skipped without comment.
- Existing behavior for `column`/`line`/`pie` must not change.
- No automated tests for COM executor methods (project convention). Verification is build + Task 6's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.
- If PP-5 has landed, express the enums through its `EXCEL_OPS` table.

---

### Task 1: `AddChart` uses the shared map

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Replace the ternary at `:653-658` with a `ExcelChartTypeMap` lookup:

```csharp
int chartTypeCode = 51; // xlColumnClustered
if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
{
    if (!ExcelChartTypeMap.TryGetValue(ct.GetString(), out chartTypeCode))
        throw new ArgumentException("add_chart: unknown chartType '" + ct.GetString() +
                                    "'. Valid: " + string.Join(", ", ExcelChartTypeMap.Keys) + ".");
}
```

- [ ] **Step 2:** Apply the same treatment in `EditChartExcel` (`:672-675`): today a bad `chartType` fails `TryGetValue` and is silently skipped while the tool still reports success. Split the condition so an unrecognized value throws instead of being skipped.
- [ ] **Step 3:** Consider widening the shared map. `columnStacked` (52) and `barStacked` (58) are common asks and `PptChartTypeMap` already has `barStacked`; `scatter` (-4169) and `radar` (-4151) are plausible. Add only types verified to render correctly from a plain rectangular `SetSourceData` range — scatter in particular has different data-shape expectations, so verify before adding it, and leave it out if it needs special handling.
- [ ] **Step 4:** Add a comment above `ExcelChartTypeMap` noting it is now the single source for both add and edit paths, and cross-referencing the PowerPoint/Word maps.

**Verification:** `MSBuild ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug`; `chartType: 'doughnut'` creates a doughnut chart, not a column chart.

---

### Task 2: Rebinding `edit_chart` to a new range

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** In `EditChartExcel`, before the type change (order matters — some type changes reset the plot):

```csharp
if (op.TryGetProperty("dataRange", out var dr) && dr.ValueKind == JsonValueKind.String)
{
    Excel.Range source = Sheet(op).Range[dr.GetString()];
    chart.SetSourceData(source);
}
```

- [ ] **Step 2: `plotBy`.** `SetSourceData` takes an optional `PlotBy` (`xlColumns` = 2 / `xlRows` = 1). Accept an optional `plotBy: 'columns'|'rows'` and pass it through; without it Excel guesses from the range shape, which flips a chart's series/categories unpredictably when a range grows from 2 columns to 3.
- [ ] **Step 3: Ordering.** Rebinding resets series colors and names. Document that `seriesColors`/`seriesData` are applied *after* `dataRange` in the same call, so one call can rebind and restyle — and make sure the code order matches (`dataRange` first, then type, then title/legend/labels, then `seriesColors`/`seriesData` last, which is already the tail of the method at `:704-725`).
- [ ] **Step 4: Cross-sheet ranges.** `Sheet(op)` resolves the *chart's* sheet from the `sheet` property (`:111-119`), but a data range may live on another sheet. Accept an optional `dataSheet` and resolve the range against it. Without this, "chart the data on Sheet2 into the chart on Sheet1" is impossible.

**Verification:** build; adding rows to a table and calling `edit_chart` with the extended `dataRange` updates the chart in place, preserving title and legend.

---

### Task 3: Schema

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** `add_chart.chartType` gets the full enum — replacing `"column"|"line"|"pie" - other values silently become column` at `ExcelAiAddIn/web-src/entry.ts:210`, whose trailing clause is exactly the behavior Task 1 removes. Delete that clause; leaving it would now be a lie in the other direction.
- [ ] **Step 2:** `edit_chart` gets `chartType` (same enum), `dataRange`, `dataSheet`, `plotBy`.
- [ ] **Step 3:** Document `chartPath` properly — it is the ChartObject's *name* (e.g. `"Chart 1"`), obtainable from `read_sheet_features` or from `add_chart`'s return value once Task 4 lands. The current parenthetical "the chart's name" is thin for a required addressing field.
- [ ] **Step 4:** Document the rebinding ordering from Task 2 Step 3.
- [ ] **Step 5:** `add_chart` also accepts `title` but nothing else cosmetic; state that legend/colors/labels are set afterwards via `edit_chart`, so the model doesn't send them to `add_chart` and get a silent drop.

**Verification:** bundle rebuilds; "make this a horizontal bar chart" produces `chartType: 'bar'` on the first attempt.

---

### Task 4: `add_chart` returns the chart's name

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Problem:** `AddChart` returns `void` and `ProposeOperations` reports a bare `add_chart: ok` (`:531`). To edit the chart it just created, the model must guess the name Excel assigned.

- [ ] **Step 1:** Change `AddChart` to return a `string` describing what it made, including `chartObj.Name`.
- [ ] **Step 2:** In `ProposeOperations`, append that string to the result line instead of `ok`. Check whether the batch loop's uniform `lines.AppendLine(kind + ": ok")` shape (`:483-556`) needs a per-op variant; if PP-5 Task 4 or PP-12 Task 3 has reworked this loop, follow whatever reporting shape they established.
- [ ] **Step 3:** Optionally accept a `name` parameter on `add_chart` and set `chartObj.Name` to it, so the model can choose a stable, meaningful id up front. This is the more robust pattern — recommend it in the description.
- [ ] **Step 4:** Do the same for `add_shape` (`:746-773`) and `add_sparkline` (`:728-744`), which have the identical unaddressable-after-creation problem and are edited via the same name-based `visualId`.

**Verification:** create a chart, then edit it in the very next operation of the same batch using the returned name.

---

### Task 5: Series-name and category sourcing

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Context:** `edit_chart`'s `seriesData` only sets `series.Name` (`:707-725`) — a literal string. Excel's normal idiom is a cell reference (`=Sheet1!$B$1`) so the name follows the data.

- [ ] **Step 1:** Accept either form: a bare string sets the literal name; a string starting with `=` is passed through as a formula reference. `series.Name` accepts both, so this is a documentation-and-passthrough change, not new logic.
- [ ] **Step 2:** Document it in the schema.
- [ ] **Step 3:** Consider accepting `categoriesRange` on `edit_chart` to set `chart.SeriesCollection(1).XValues`. Useful when the header row is not adjacent to the data. Include only if it verifies cleanly; it is the least important item in this plan.

**Verification:** a series named `=Sheet1!$B$1` shows the cell's text in the legend and updates when the cell is edited.

---

### Task 6: Manual verification matrix

- [ ] `add_chart` with each of the six types produces that type in real Excel.
- [ ] `add_chart` with `chartType: 'nonsense'` → specific error listing valid types; no chart created.
- [ ] `edit_chart` changes an existing column chart to bar, area, and doughnut in turn.
- [ ] `edit_chart` with `dataRange` extended by 5 rows → the chart shows the new rows, keeps its title and legend.
- [ ] `edit_chart` with `dataRange` + `plotBy: 'rows'` → series and categories swap as expected.
- [ ] `edit_chart` with `dataSheet` pointing at another sheet → chart on Sheet1 plots Sheet2's data.
- [ ] `add_chart` returns a name; the next operation in the same batch edits that chart by name successfully.
- [ ] `add_chart` with an explicit `name` → that name is usable immediately.
- [ ] `seriesData` with a literal name and with an `=` reference both work.
- [ ] Natural language: "make this a horizontal bar chart" → a real bar chart. "Update the chart to include the new rows" → rebinding, not a duplicate chart.
