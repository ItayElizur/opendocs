# PP-18: `set_page_setup`, `delete_table`, `add_sparkline` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-18 (P2) — three small, independent findings.

**Goal:** Each of these three operations produces the outcome its name and schema promise, or says clearly why it cannot. Each is a plausible-but-wrong result from a reasonable request today, with no error to signal it.

**Architecture:** Three unrelated defects, grouped only by size. They can be done in any order and by different people.

1. **`set_page_setup` silently discards `scale`.** `SetPageSetup` (`ExcelAiAddIn/ExcelTools.cs:852-865`) applies `scale` → `setup.Zoom` at `:857-858`, then the `fitToWidth` branch at `:859-863` unconditionally sets `setup.Zoom = false` before setting `FitToPagesWide`. Excel treats Zoom and FitToPages as mutually exclusive (the code comment at `:861` says so) — but the schema (`ExcelAiAddIn/web-src/entry.ts:206`) lists `scale?, fitToWidth?, fitToHeight?` side by side with nothing marking them exclusive, so nothing stops the model combining them and silently losing one. Worse, `fitToHeight` at `:864-865` sets `FitToPagesTall` **without** clearing Zoom — so `{scale, fitToHeight}` produces yet a third behavior. There is also a real bug in the `scale` path: `setup.Zoom` accepts an int or `false`; assigning `(int)scale.GetDouble()` without range-checking 10..400 throws an opaque COM error.
2. **`delete_table` keeps the data.** `DeleteTable` (`:1085-1088`) calls `.Unlist()` — which converts the ListObject back to a plain range, keeping every cell value and all formatting. Its own comment says so. But the operation is named `delete_table` and the schema says nothing, so "delete the table" leaves the data sitting there looking deleted-but-not.
3. **`add_sparkline` defaults its target onto its own source.** `AddSparkline` (`:728-744`) defaults `targetCell` to `dataRange` (`:731`) — the sparkline draws inside the very cells holding its source data, overlaying the numbers. Undocumented.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Excel`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **Prefer an explicit error over a plausible guess.** Each of these is a case where the code currently guesses and the guess is defensible in isolation; the harm is that it is invisible.
- Do not silently change what an existing well-formed call does, except where a task explicitly says the current behavior is wrong (Task 1's conflict case, which today has no single defined behavior anyway).
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after `entry.ts` changes.
- If PP-5 has landed, express schema changes through its `EXCEL_OPS` table.

---

### Task 1: `set_page_setup` — reject the conflicting combination

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Detect and reject the conflict** at the top of `SetPageSetup`, before touching COM:

```csharp
bool hasScale = op.TryGetProperty("scale", out var scaleEl) && scaleEl.ValueKind == JsonValueKind.Number;
bool hasFit = (op.TryGetProperty("fitToWidth", out var ftwEl) && ftwEl.ValueKind == JsonValueKind.Number)
           || (op.TryGetProperty("fitToHeight", out var fthEl) && fthEl.ValueKind == JsonValueKind.Number);
if (hasScale && hasFit)
    throw new ArgumentException(
        "set_page_setup: 'scale' and 'fitToWidth'/'fitToHeight' are mutually exclusive in Excel " +
        "(matching its own Page Setup UI). Pass either scale, or one/both fit values - not both.");
```

Rejecting rather than picking a winner is right here: there is no defensible choice, and the model can retry immediately with an unambiguous call.

- [ ] **Step 2: Fix the `fitToHeight`-only path.** `FitToPagesTall` also requires `Zoom = false`. Restructure so `Zoom = false` is set once when *either* fit value is present, then apply whichever were given:

```csharp
if (hasFit)
{
    setup.Zoom = false;
    if (ftwEl.ValueKind == JsonValueKind.Number) setup.FitToPagesWide = (int)ftwEl.GetDouble();
    if (fthEl.ValueKind == JsonValueKind.Number) setup.FitToPagesTall = (int)fthEl.GetDouble();
}
else if (hasScale) { /* validated scale */ }
```

- [ ] **Step 3: Validate `scale`** against Excel's 10..400 range with a specific error.
- [ ] **Step 4: Support "fit to N pages wide, unlimited tall".** Excel expresses that as `FitToPagesTall = false`. Accept `fitToHeight: 0` (or `null`) to mean unlimited and document it — it is the single most common real-world page-setup request ("fit on one page wide") and today it is unexpressible.
- [ ] **Step 5: Schema** — mark the exclusivity in the description explicitly, and document the `scale` range and the `fitToHeight: 0` convention. If the schema is structural (PP-5), express the exclusivity with `oneOf` between a scale branch and a fit branch.

**Verification:** build; `{scale: 80}` scales; `{fitToWidth: 1, fitToHeight: 0}` fits one page wide, unlimited tall; `{scale: 80, fitToWidth: 1}` errors specifically; `{scale: 900}` errors specifically.

---

### Task 2: `delete_table` — make the name match the behavior

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Decide.** Three options: (a) rename the operation, (b) add a parameter, (c) document only.

  **Take (b)**, with (c)'s documentation attached. Renaming breaks saved conversations and any prompt that learned the current name; documentation alone leaves "delete the table and its data" impossible. A parameter makes both behaviors reachable and forces the schema to state which is the default.

- [ ] **Step 2: Implement**

```csharp
private static void DeleteTable(JsonElement op)
{
    Excel.ListObject table = ResolveTable(op);
    bool deleteData = op.TryGetProperty("deleteData", out var dd) && dd.ValueKind == JsonValueKind.True;
    if (deleteData)
    {
        Excel.Range body = table.Range;   // includes the header row
        table.Unlist();
        body.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
    }
    else
    {
        table.Unlist();   // default, unchanged: converts to a plain range, keeps data + formatting
    }
}
```

Capture the range **before** `Unlist()` — the `ListObject` reference is invalid afterwards.

- [ ] **Step 3: Shift direction.** `xlShiftUp` is the safer default (a table usually occupies full-width rows), but it moves everything below. Accept an optional `shift: 'up'|'left'|'none'` where `'none'` uses `ClearContents()` instead of `Delete()`, leaving the cells empty in place. Default `'up'`.
- [ ] **Step 4: Schema** — spell out the default in the description: *"Converts the table back to a plain range, keeping all data and formatting (this is Excel's Unlist). Pass deleteData:true to also remove the cells."*
- [ ] **Step 5: Report accurately.** Return a string saying which happened, so the transcript (and PP-3's output view) shows "converted to range, data kept" rather than a bare `ok`.

**Verification:** build; default keeps the data; `deleteData: true` removes it; both report accurately.

---

### Task 3: `add_sparkline` — require or explain the target

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Make `targetCell` required.** It is the only field with no sane default — the current default (`dataRange`, `:731`) is actively wrong, drawing the sparkline over its own numbers. Throw a specific error when absent, naming the usual choice (the cell immediately right of, or below, the data).
- [ ] **Step 2 (alternative if a default is wanted):** derive one — the cell immediately to the right of `dataRange` for a row-shaped range, immediately below for a column-shaped one — and **report the chosen cell** in the result. Prefer Step 1: a wrong-but-silent default is exactly the failure mode this whole product plan is about, and a derived default is still a guess. Only take Step 2 if a product owner wants the call to never fail.
- [ ] **Step 3: Validate shape.** A sparkline group's target must have the same number of cells as the data has rows/columns (one sparkline per data row, or one for a single range). Detect the mismatch and throw before the COM call, which otherwise fails opaquely.
- [ ] **Step 4:** `type` (`:732`) falls back to `line` for any unrecognized value (`:734-736`) — the same silent-fallback pattern. Add the `enum` (`line`/`column`/`stacked`) to the schema and error on anything else.
- [ ] **Step 5: Return the group's identity** if one is available, so the sparkline can be edited or removed later; if the PIA exposes no stable id, say so in the description rather than leaving the model to guess.
- [ ] **Step 6: Schema** — document `targetCell` as required, the shape rule, and the `type` enum.

**Verification:** build; a sparkline lands in the requested cell; omitting `targetCell` errors specifically; a shape mismatch errors specifically.

---

### Task 4: Sweep for sibling instances

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs` (as found)

- [ ] **Step 1:** All three defects here are instances of two patterns: *silent fallback for an unrecognized value*, and *a default that is defensible in code but wrong in practice*. Grep `ExcelTools.cs` for `TryGetProperty(...) ? ... : "<literal>"` and for ternary chains ending in a bare fallback, and list every instance.
- [ ] **Step 2:** For each, classify: harmless (a genuinely sane default, e.g. `hasHeader` defaulting true), or a PP-18-class defect. Fix the latter the same way; record the former in a comment so the next audit does not re-flag it.
- [ ] **Step 3:** Do **not** fix instances already owned by another plan — `add_chart.chartType` is PP-15, `add_shape.shapeType` is PP-16, the conditional-formatting fallbacks are PP-14. Note the overlap and move on.
- [ ] **Step 4:** Add the resulting list to `docs/ai-tool-surface.md` so it is a verified baseline rather than a one-off sweep.

**Verification:** the list exists and every entry is either fixed, deliberately kept, or assigned to a named plan.

---

### Task 5: Manual verification matrix

- [ ] `set_page_setup {scale: 80}` alone → 80% scaling in Page Setup.
- [ ] `{fitToWidth: 1}` alone → fits one page wide.
- [ ] `{fitToHeight: 2}` alone → fits two pages tall (this is the path that was silently broken).
- [ ] `{fitToWidth: 1, fitToHeight: 0}` → one page wide, unlimited tall.
- [ ] `{scale: 80, fitToWidth: 1}` → specific error; page setup unchanged.
- [ ] `{scale: 5}` and `{scale: 900}` → specific range errors.
- [ ] `delete_table` default → table gone from the Name Box, data still present and readable.
- [ ] `delete_table {deleteData: true}` → data gone, rows shifted up.
- [ ] `delete_table {deleteData: true, shift: 'none'}` → cells emptied in place.
- [ ] `add_sparkline` without `targetCell` → specific error, nothing drawn.
- [ ] `add_sparkline` with a valid `targetCell` → sparkline in that cell, source data still readable.
- [ ] `add_sparkline` with a mismatched target shape → specific error.
- [ ] `add_sparkline {type: 'nonsense'}` → specific error listing valid types.
