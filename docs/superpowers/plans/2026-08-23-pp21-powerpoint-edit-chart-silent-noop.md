# PP-21: PowerPoint `edit_chart` — No More Silent No-Ops

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-21 (P1).

**Goal:** `edit_chart`'s `chartType` and `legendPos` have documented `enum` values matching what the handler accepts, and an out-of-range value is rejected with an error the model can react to — instead of the current "does nothing (or the wrong thing) and reports success."

**Architecture:** Two defects in `EditChartPpt` (`PowerPointAiAddIn/PowerPointTools.cs:505-548`), both ending in the same unconditional `return new ToolResult { Output = "Chart updated." ... }` at `:547`.

1. **`chartType`** (`:510-513`) is applied only when `PptChartTypeMap.TryGetValue` succeeds. A typo or an unlisted name fails the lookup, the whole `if` is skipped, and the tool still reports "Chart updated." "Change this to a bar chart" can silently do nothing.
2. **`legendPos`** (`:519-531`) is schema-typed as a bare `{ type: 'string' }` (`PowerPointAiAddIn/web-src/entry.ts:356`) with no valid-value guidance, while the handler expects exactly `"none"`/`"r"`/`"t"`/`"l"`. The dispatch at `:529` is `pos == "r" ? -4152 : pos == "t" ? -4160 : pos == "l" ? -4131 : -4107` — a terminal `else` meaning **any** unmatched value lands on bottom. So the natural phrasings a model will actually send — `"right"`, `"bottom"`, `"top"` — silently move the legend to the bottom. `"right"` in particular produces the exact opposite of the request in three cases out of four.

There is a third issue behind #1 worth fixing at the same time: `PptChartTypeMap` (`:430-437`) maps `"bar"` to **51 = xlColumnClustered**, not 57 = xlBarClustered. Excel's map has this right (`ExcelAiAddIn/ExcelTools.cs:58-66`). So even a *successful* `chartType: 'bar'` produces a column chart — a silent wrong result that no amount of enum work would catch.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.PowerPoint` + `dynamic` for the shared chart engine; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **Never report success for a parameter that was not applied.** Every task here serves that one rule.
- One chart vocabulary across the repo: names and codes match `ExcelChartTypeMap` (`ExcelAiAddIn/ExcelTools.cs:58-66`) and any Word map from PP-9. Where the two currently disagree, Excel is correct.
- Task 1 Step 2 changes the observable behavior of an existing valid input (`chartType: 'bar'` starts producing an actual bar chart). That is a bug fix, not a regression — call it out in the commit message.
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: Fix and enumerate `chartType`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Error on an unrecognized value.** Split the compound condition at `:510` so the lookup failure is distinguishable from an absent parameter:

```csharp
if (input.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
{
    int typeCode;
    if (!PptChartTypeMap.TryGetValue(ct.GetString(), out typeCode))
        throw new ArgumentException("edit_chart: unknown chartType '" + ct.GetString() +
                                    "'. Valid: " + string.Join(", ", PptChartTypeMap.Keys) + ".");
    chart.ChartType = typeCode;
}
```

- [ ] **Step 2: Fix the `bar` mapping.** `PptChartTypeMap["bar"] = 51` is xlColumnClustered. Correct it to 57 (xlBarClustered) and add `"column" = 51` / `"columnStacked" = 52` so both orientations are addressable. Check whether `"barStacked" = 52` (`:432`) has the same problem — 52 is xlColumnStacked, so it should be 58 (xlBarStacked). Verify each code against the xlChartType enumeration before changing it; getting this wrong replaces one silent wrong result with another.
- [ ] **Step 3: Back-compat note.** Anything relying on `"bar"` meaning a column chart changes behavior. That is the point — but grep the repo (system prompts, docs, starter prompts) for hardcoded `"bar"` usages first, and check `AddChartPpt`'s fallback at `:441`, which defaults to 51 and is described as "bar" in PP-22.
- [ ] **Step 4: Schema** — `chartType` gets a real `enum` from the corrected map, replacing `{ type: 'string' }` at `entry.ts:355`.
- [ ] **Step 5:** Apply the same enum to `add_chart`'s `kind` — same map, same vocabulary. PP-22 owns that field; if it has landed, verify rather than redo, and make sure both fields ended up with the *same* list.

**Verification:** `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug`; `chartType: 'bar'` produces horizontal bars; `chartType: 'nonsense'` errors specifically.

---

### Task 2: Fix and enumerate `legendPos`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Accept the natural names, keep the short ones.** The handler's `"r"`/`"t"`/`"l"` are a genoffice-ism; a model will say `"right"`. Support both through one map:

```csharp
private static readonly Dictionary<string, int> PptLegendPositions =
    new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
{
    ["right"] = -4152, ["r"] = -4152,     // xlLegendPositionRight
    ["top"] = -4160, ["t"] = -4160,       // xlLegendPositionTop
    ["left"] = -4131, ["l"] = -4131,      // xlLegendPositionLeft
    ["bottom"] = -4107, ["b"] = -4107,    // xlLegendPositionBottom
    ["corner"] = -4161,                   // xlLegendPositionCorner
};
```

Verify `xlLegendPositionCorner`'s code before including it; drop it if unverified rather than guessing.

- [ ] **Step 2: Replace the ternary chain** at `:529` with a `TryGetValue` + throw. The terminal `else → bottom` is the entire defect; nothing may replace it with another default.
- [ ] **Step 3:** Keep `"none"` handled separately as it is now (`:522-525`) — it sets `HasLegend = false` rather than a position — but include it in the schema enum and in the error message's valid list.
- [ ] **Step 4: Schema** — `legendPos` gets `enum: ['none','right','top','left','bottom','corner','r','t','l','b']`. Consider listing only the long names plus `none` to keep the enum readable while the handler still accepts the short aliases for back-compat; note that choice in the description.
- [ ] **Step 5:** Align with Excel, whose `edit_chart` already uses `"none"|"right"|"top"|"left"|"bottom"` (`ExcelAiAddIn/web-src/entry.ts:211`) and whose handler does the same terminal-else-to-bottom trick (`ExcelTools.cs:686-695`). **Fix Excel's the same way in this plan** — it is the identical five-line defect and splitting it across two plans guarantees one gets forgotten.

**Verification:** build; `legendPos: 'right'` puts the legend on the right in both apps; an unknown value errors specifically.

---

### Task 3: Audit `edit_chart`'s remaining parameters

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`

- [ ] **Step 1: `dataLabels`** (`:533-540`) reads `dl.ValueKind == JsonValueKind.True`, so any non-boolean (including the string `"value"`, which Excel's equivalent accepts) is treated as `false` — silently turning labels off when asked to turn them on. Either accept Excel's `'none'|'value'|'percent'` vocabulary, or reject non-booleans explicitly. Prefer matching Excel.
- [ ] **Step 2: `gridlines`** (`:541-544`) has the same boolean coercion, and `chart.Axes(2)` throws for chart types with no value axis (pie, doughnut). Catch that and return a specific message instead of a raw COM error.
- [ ] **Step 3: `title`** (`:514-518`) is fine.
- [ ] **Step 4: Report what applied.** Replace the flat `"Chart updated."` (`:547`) with a list of the properties actually changed. This is the general guard against the whole class: even a parameter that silently no-ops in future becomes visible in the transcript.
- [ ] **Step 5:** `ResolveShape(input)` (`:507`) throws an unhelpful error when the target shape is not a chart. Check `shape.HasChart` first and report `"edit_chart: shape <n> on slide <m> is not a chart."`.

**Verification:** build; each parameter either works or errors clearly; the result names what changed.

---

### Task 4: Cross-app chart-vocabulary reconciliation

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1:** After Tasks 1-2, tabulate the chart-type names and codes across `ExcelChartTypeMap`, `PptChartTypeMap`, and any Word map from PP-9. They should agree exactly.
- [ ] **Step 2:** Record the table in `docs/ai-tool-surface.md` as the canonical vocabulary, so the next person adding a chart type in any app has one place to look.
- [ ] **Step 3:** Note the `bar` code fix and its date — it changes existing behavior and someone will eventually ask why.

---

### Task 5: Manual verification matrix

- [ ] `edit_chart {chartType: 'bar'}` → horizontal bars (previously a column chart).
- [ ] `edit_chart {chartType: 'column'}` → vertical columns.
- [ ] Every other type in the corrected map renders as its name says.
- [ ] `edit_chart {chartType: 'nonsense'}` → specific error listing valid types; chart unchanged.
- [ ] `edit_chart {legendPos: 'right'}` → legend on the right (previously bottom).
- [ ] `legendPos` with each of `left`, `top`, `bottom`, `none` → correct.
- [ ] `legendPos: 'r'` → still right (back-compat).
- [ ] `legendPos: 'nonsense'` → specific error; legend unchanged.
- [ ] The same two legend checks in **Excel**'s `edit_chart` (Task 2 Step 5).
- [ ] `dataLabels: 'value'` → labels shown, not silently hidden.
- [ ] `gridlines: true` on a pie chart → clear message, not a raw COM error.
- [ ] `edit_chart` targeting a non-chart shape → specific error.
- [ ] The result text names every property actually applied.
- [ ] Natural language: "change this to a bar chart" and "move the legend to the right" both do what they say — the two failures the source item names.
