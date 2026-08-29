# PP-9: Word Chart Parity — Categories, Named Multi-Series, Chart Type

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-9 (P1).

**Goal:** Bring Word's `edit_chart` up to what this repo's own PowerPoint `add_chart` already does correctly — labeled categories, named multi-series, and a real chart-type choice — so an AI-authored chart in Word stops being generic unlabeled column bars.

**Architecture:** Word's chart object model is the same shared Office chart engine PowerPoint uses: a `Shape` with `HasChart`, whose data lives in an embedded Excel workbook reached via `chart.ChartData.Workbook`. `PowerPointTools.AddChartPpt` (`PowerPointAiAddIn/PowerPointTools.cs:437-503`) already implements the full pattern correctly, including the part that is easy to get wrong: opening the embedded workbook, writing the grid, `SetSourceData(sheet.UsedRange)`, then `Close(SaveChanges: true)` inside a `finally` with an explicit `Marshal.ReleaseComObject` so no hidden Excel host process leaks.

Word's `EditChart` (`WordAiAddIn/WordTools.cs:132-169`) does none of that — it sets `series.Values = values` directly on series 1 of a hardcoded `xlColumnClustered` chart. Direct `Values` assignment works for a single unlabeled series and is exactly why categories and series names are unreachable.

The plan therefore **ports PowerPoint's proven implementation into Word** rather than inventing a Word-specific approach. Two Word-specific decisions:

1. **Chart identity.** `EditChart` finds "the first shape with `HasChart`" (`WordTools.cs:143-150`) and creates one if none exists — so a document with two charts can only ever address the first. Task 3 adds explicit addressing while keeping the create-or-edit convenience.
2. **Insertion position.** `Shapes.AddChart2(-1, type, 0, 0, 300, 200)` places a floating shape at document origin. Task 4 anchors it at a paragraph index instead, which is what "put a chart after the intro" requires.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, VSTO COM interop against `Microsoft.Office.Interop.Word` plus `dynamic` for the shared chart engine (the existing file already documents why `dynamic` is used here — `WordTools.cs:128-131`).

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Follow `AddChartPpt`'s COM-release discipline exactly. The embedded chart workbook **must** be closed and released in a `finally`; a leaked one leaves an invisible Excel process alive for the rest of the Word session.
- Backward compatibility: the current schema is `{ title, values }`, both required (`WordAiAddIn/web-src/entry.ts:193-205`). A call in that old shape must keep working — `values` without `categories` produces a single unnamed series, as today.
- No automated tests for COM executor methods (project convention). Verification is build + the manual matrix in Task 6.
- Rebuild the bundle + MSBuild after any `entry.ts` change (command in `2026-08-23-pp02-tool-steps-chronological-order.md`'s Global Constraints).
- If PP-5 (gateway schemas) has landed, declare this tool's enums through its tables rather than by hand.

---

### Task 1: Chart-type map, mirroring PowerPoint's

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Produces: `private static readonly Dictionary<string, int> WordChartTypeMap` — consumed by Tasks 2 and 3.

- [ ] **Step 1:** Add the map next to the other statics near the top of the class, using the same xlChartType codes as `PptChartTypeMap` (`PowerPointAiAddIn/PowerPointTools.cs:430-437`) and `ExcelChartTypeMap` (`ExcelAiAddIn/ExcelTools.cs:58-66`), so all three apps speak one vocabulary:

```csharp
// Same xlChartType codes as ExcelChartTypeMap / PptChartTypeMap - one chart
// vocabulary across all three add-ins.
private static readonly Dictionary<string, int> WordChartTypeMap = new Dictionary<string, int>
{
    ["column"] = 51,        // xlColumnClustered
    ["columnStacked"] = 52, // xlColumnStacked
    ["bar"] = 57,           // xlBarClustered
    ["line"] = 4,           // xlLine
    ["area"] = 1,           // xlArea
    ["pie"] = 5,            // xlPie
    ["doughnut"] = -4120,   // xlDoughnut
};
```

- [ ] **Step 2:** Note the deliberate difference from `PptChartTypeMap`, which maps the key `"bar"` to 51 (xlColumn**Clustered**) — a naming bug in that file, not a model to copy. Word's map uses the correct code for each name; PP-21/PP-22 covers the PowerPoint side. Add a comment recording the discrepancy so the two are not "harmonized" in the wrong direction later.

- [ ] **Step 3:** Unknown chart type is an **error**, not a silent fallback to column — the whole class of bugs in PP-15/PP-21/PP-22. Throw `ArgumentException` naming the value and listing valid ones.

**Verification:** `MSBuild WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds.

---

### Task 2: Port the embedded-workbook data writer

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Produces: `private static void WriteChartData(dynamic chart, List<string> categories, JsonElement seriesArray)` — consumed by Task 3.

- [ ] **Step 1:** Port `AddChartPpt`'s data-writing block (`PowerPointAiAddIn/PowerPointTools.cs:454-495`) verbatim into this helper: series names into row 1 starting at column 2, categories into column 1 starting at row 2, values filling the grid, then `chart.SetSourceData(sheet.UsedRange)` — all inside the `try`/`finally` that closes and releases `dataWorkbook`.
- [ ] **Step 2:** Handle the no-categories case: when `categories` is empty, write `1..N` as categories so `UsedRange` still forms a rectangle. Without this, `SetSourceData` on a one-column range produces a chart with no plotted series.
- [ ] **Step 3:** Handle the legacy single-`values` case by normalizing it up front — Task 3 converts `{ values: [...] }` into a one-element `series` array with no name before calling this helper, so the helper itself only ever handles one shape.
- [ ] **Step 4:** Validate that every series' `values` array has the same length as `categories` (or as each other) and throw a specific error if not. A ragged grid produces a silently wrong chart, which is worse than a rejected call.

**Verification:** build succeeds.

---

### Task 3: Rewrite `EditChart`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [ ] **Step 1: New schema**

```ts
{
  name: 'edit_chart',
  description:
    'Creates or edits a native Word chart. Supply categories + one or more named series for a labeled chart. ' +
    'chartIndex addresses an existing chart (0-based, in document shape order); omit it to edit the first chart, ' +
    'or pass create:true to always add a new one.',
  inputSchema: {
    type: 'object',
    properties: {
      title: { type: 'string' },
      chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'line', 'area', 'pie', 'doughnut'] },
      categories: { type: 'array', items: { type: 'string' } },
      series: {
        type: 'array',
        items: {
          type: 'object',
          properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } },
          required: ['values'],
        },
      },
      values: { type: 'array', items: { type: 'number' }, description: 'Legacy single-series shorthand; prefer series.' },
      chartIndex: { type: 'number' },
      create: { type: 'boolean' },
      afterBlockIndex: { type: 'number', description: '0-based paragraph index to anchor a NEW chart after; -1 = start.' },
    },
    required: [],
  },
}
```

Note `required` drops to empty: with create-or-edit semantics, a call that only changes a title is legitimate.

- [ ] **Step 2: Resolve the target chart**

Replace the "first shape with `HasChart`" scan (`WordTools.cs:143-150`) with a helper that collects **all** chart shapes into a list, then selects by `chartIndex` (0-based), defaulting to index 0. If `create` is true, or no chart exists, add one — via Task 4's anchored insertion.

- [ ] **Step 3: Apply**

Order matters: set `chart.ChartType` *before* writing data (some type changes reset series formatting), then call `WriteChartData`, then set the title. Normalize legacy `values` into `series` first, per Task 2 Step 3.

- [ ] **Step 4: Report accurately**

The output string already names created-vs-updated (`WordTools.cs:164`). Extend it with the chart index, type, series count, and category count, so the transcript (and PP-3's output view) shows what was actually built.

- [ ] **Step 5: Update the system prompt** at `WordAiAddIn/web-src/entry.ts:250-257`, which currently says the assistant can "create or edit a native Word chart" — mention categories and multi-series so the capability is discoverable in prose too.

**Verification:** build + bundle; manual chart creation with two named series and five categories renders with a legend and axis labels.

---

### Task 4: Anchored insertion

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

- [ ] **Step 1:** When creating a chart with `afterBlockIndex` present, resolve that paragraph via `ActiveDoc.Paragraphs[i + 1]` (the same 0-based convention `ResolveTargetParagraphs` uses, `WordTools.cs:307-377`), collapse a range to its end, and use the `Shapes.AddChart2` overload that takes an `Anchor` range — or, if an inline chart is wanted, `InlineShapes.AddChart2` at that range. Prefer **inline**: an inline chart flows with the text, which is what "add a chart after this paragraph" means to a user, whereas a floating shape at (0,0) overlaps the text.
- [ ] **Step 2:** `-1` means "start of document", matching `insertToc`/`moveBlocks`' existing convention (`WordAiAddIn/web-src/entry.ts:234-236`).
- [ ] **Step 3:** With no `afterBlockIndex`, keep today's behavior (floating shape) so existing calls do not change position. Document the difference in the schema description.
- [ ] **Step 4:** If Task 3's chart enumeration walks `doc.Shapes`, extend it to walk `doc.InlineShapes` as well — otherwise a chart this task inserts inline becomes unaddressable by `chartIndex`. Define the ordering as: inline shapes in document order first, then floating shapes, and state that in the schema description so indices are predictable.

**Verification:** manual — a chart inserted after paragraph 3 appears between paragraphs 3 and 4 and moves with the text when paragraphs above it are edited.

---

### Task 5: Shared-code note

**Files:** none modified (assessment)

- [ ] **Step 1:** After Task 2, `WriteChartData` in `WordTools.cs` and the block in `AddChartPpt` are near-identical. Assess whether to extract into `OfficeAi.Shared` — it would need to take a `dynamic chart` and no app-specific types, so it is feasible.
- [ ] **Step 2:** **Recommendation: don't, yet.** `OfficeAi.Shared` currently holds only app-agnostic plumbing (`ChatStore`, `ToolProtocol`, `WebViewBridgeHost`); introducing `dynamic` COM chart logic there sets a precedent for moving Office-specific code into it. Add a cross-reference comment in both files instead, and revisit if a third copy appears (Excel's `AddChart` is the likely third — see PP-15).

---

### Task 6: Manual verification matrix

- [ ] Legacy call `{ title: 'X', values: [1,2,3] }` → still creates a single-series column chart titled X (no regression).
- [ ] `{ chartType: 'bar', categories: ['Q1','Q2','Q3'], series: [{name:'2025', values:[1,2,3]}, {name:'2026', values:[4,5,6]}] }` → a bar chart with three labeled categories, two named series, and a legend showing both names.
- [ ] `{ chartType: 'pie', categories: [...], series: [{values:[...]}]}` → a pie chart with labeled slices.
- [ ] `{ chartType: 'nonsense' }` → a specific error listing valid types; no chart created or modified.
- [ ] Ragged series (`categories` length 3, `values` length 2) → specific error; no partial chart.
- [ ] Two charts in one document: `chartIndex: 1` edits the second one, `chartIndex: 0` the first.
- [ ] `afterBlockIndex: 2` → chart appears after the third paragraph, inline.
- [ ] After creating five charts in one session, Task Manager shows no orphaned `EXCEL.EXE` process (the COM-release check).
