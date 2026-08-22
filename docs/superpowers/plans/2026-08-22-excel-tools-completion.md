# Excel Tools Completion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the gap between `ExcelAiAddIn/ExcelTools.cs` (currently 3 read tools + `propose_operations` with 9 of 65 real operation kinds) and genoffice's real `apps/sheets` tool surface (per `C:\Dev\genoffice\docs\ai-tool-surface.md` and `apps/sheets/src/domain/workbook-dsl.ts`) by adding 6 standalone tools (`read_formats`, `read_sheet_features`, `find_cells`, `select_range`, `trace_precedents`, `trace_dependents`) and 56 new `propose_operations` kinds.

**Architecture:** Every operation maps to a genuinely native Excel COM/Interop call — this is not an approximation layer. The two structurally hardest pieces are `add_pivot` (a real, live, refreshable `PivotTable` object via `PivotCaches.Create`/`CreatePivotTable`, which is actually *simpler* than genoffice's own dual-path "bake now, write real OOXML parts at save" implementation, since COM's PivotTable is inherently the real thing) and `edit_chart` (data changes go through the chart's embedded-workbook `ChartData.Workbook`, a COM-automation-inside-COM-automation pattern that must release its Excel Application reference carefully to avoid an orphaned hidden process). A handful of operations need one-time unit-conversion or lookup-table work called out explicitly below (`set_col_width`'s pixel→character-width conversion, `add_shape`'s OOXML-preset-name→`MsoAutoShapeType` table, `set_page_setup`'s margin-preset point values) — these are documented as deliberate, reasonable approximations, not silent bugs.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (VSTO COM Interop against `Microsoft.Office.Interop.Excel`), matching every other file in `ExcelAiAddIn/`.

**Spec:** Parameter shapes and semantics below are read directly from genoffice's real source (`apps/sheets/src/domain/workbook-dsl.ts` and its executor in `apps/sheets/src/renderer/ai/tools.ts`).

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Do not modify the 3 existing standalone tools (`get_workbook_context`, `read_range`, `read_cells`) or the 9 existing `propose_operations` kinds (`set_cell`, `set_formula`, `set_range`, `format_range`, `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`, `add_chart`) — only add new `case` branches and new private helper methods.
- Every new tool/operation respects the existing editing-mode gate in `ExcelTools.Execute` — do not bypass it. New mutating operations are NOT added to `AlwaysAllowedTools`; new read-only tools (`read_formats`, `read_sheet_features`, `find_cells`, `select_range`, `trace_precedents`, `trace_dependents`) ARE added to `AlwaysAllowedTools`, matching the existing 3 read tools' treatment.
- `genoffice`'s `load_guide` tool is deliberately NOT ported — it's a prompt-management convenience for genoffice's much larger 65-op system prompt; with this plan's tool descriptions written directly and concretely (not lazily loaded), it isn't needed. Revisit only if the system prompt grows large enough to cause real context-budget problems.
- Any COM object explicitly created for a "dip into another app" operation (`edit_chart`'s embedded chart workbook) must be released deterministically (`Marshal.ReleaseComObject` or closing/quitting the sub-object) before the tool returns — a leaked reference orphans a hidden Excel/embedded-object process that outlives the visible session.
- `Application.DisplayAlerts` must be set to `false` before any operation that would otherwise show a native confirmation dialog (`delete_sheet`), and restored to `true` in a `finally` block immediately after — never leave it permanently disabled.
- Rebuild the esbuild bundle and re-run MSBuild after any `entry.ts` change (Task 15): `npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap`, then `MSBuild ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug`.
- No automated tests for COM-executor methods — verification is build + manual interactive testing in real Excel, same convention as every prior task in this project.

---

### Task 1: Simple standalone tools — `select_range`, `clear_cell`, `clear_range`, `read_formats`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: existing `Sheet(JsonElement)` helper.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Add `select_range` as a standalone tool**

Add `"select_range"` to `AlwaysAllowedTools`, add a `case "select_range": return SelectRange(input);` to `Execute`'s switch, and:
```csharp
private static ToolResult SelectRange(JsonElement input)
{
    string address = input.GetProperty("address").GetString();
    Excel.Worksheet sheet = Sheet(input);
    sheet.Activate();
    Excel.Range range = sheet.Range[address];
    range.Select();
    return new ToolResult { Output = "Selected " + address + " on " + sheet.Name + ".", Summary = "select_range" };
}
```

- [ ] **Step 2: Add `clear_cell`/`clear_range` as `propose_operations` kinds**

Add to `ProposeOperations`'s switch:
```csharp
                        case "clear_cell":
                            Sheet(op).Range[op.GetProperty("address").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_range":
                            Sheet(op).Range[op.GetProperty("range").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```

- [ ] **Step 3: Add `read_formats` as a standalone tool**

Add `"read_formats"` to `AlwaysAllowedTools`, add `case "read_formats": return ReadFormats(input);`, and:
```csharp
private static ToolResult ReadFormats(JsonElement input)
{
    string address = input.GetProperty("address").GetString();
    Excel.Range range = Sheet(input).Range[address];
    if (range.Cells.Count > 200)
    {
        return new ToolResult { Output = "Range exceeds 200-cell cap.", IsError = true, Summary = "read_formats" };
    }
    var sb = new System.Text.StringBuilder();
    foreach (Excel.Range cell in range.Cells)
    {
        bool bold = (bool)(cell.Font.Bold ?? false);
        bool italic = (bool)(cell.Font.Italic ?? false);
        bool underline = !(cell.Font.Underline is bool underlineOff) || underlineOff == false ? cell.Font.Underline.ToString() != "-4142" : false;
        string numberFormat = cell.NumberFormat as string;
        bool hasDefaultFormat = !bold && !italic && (numberFormat == "General" || numberFormat == null);
        if (hasDefaultFormat) continue; // only explicitly-formatted cells, matches genoffice
        sb.AppendLine($"{cell.Address[false, false]}: bold={bold}, italic={italic}, numberFormat={numberFormat}");
    }
    return new ToolResult { Output = sb.ToString(), Summary = "read_formats" };
}
```
(Underline detection via raw Interop is fiddly since `Font.Underline` is a `XlUnderlineStyle` boxed as object, not a bool — the check above treats anything other than `xlUnderlineStyleNone` (-4142) as "has underline"; simplify further during manual verification if the boxing behaves differently than expected in practice.)

- [ ] **Step 4: Build and manually verify**

Run from `ExcelAiAddIn/`: MSBuild command from Global Constraints. Expected: 0 errors.

Manually verify: `select_range` with `{"address":"B2:C3"}` navigates/selects that range in the open workbook; `clear_cell`/`clear_range` via `propose_operations` empty their target(s) without touching formatting; `read_formats` on a range with one bold cell and one plain cell reports only the bold one.

- [ ] **Step 5: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add select_range, clear_cell, clear_range, read_formats"
```

---

### Task 2: Layout operations — `sort_range`, `merge_cells`, `unmerge_cells`, `set_row_height`, `set_col_width`, `set_rows_hidden`, `set_cols_hidden`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Add the 7 `propose_operations` kinds**

Add to `ProposeOperations`'s switch:
```csharp
                        case "sort_range": SortRange(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "merge_cells":
                            Sheet(op).Range[op.GetProperty("range").GetString()].Merge();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "unmerge_cells":
                            Sheet(op).Range[op.GetProperty("range").GetString()].UnMerge();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_row_height": SetRowHeight(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_col_width": SetColWidth(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_rows_hidden": SetRowsHidden(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_cols_hidden": SetColsHidden(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```
and the helper methods:
```csharp
private static void SortRange(JsonElement op)
{
    string range = op.GetProperty("range").GetString();
    string byColumn = op.GetProperty("byColumn").GetString();
    string order = op.GetProperty("order").GetString();
    bool hasHeader = op.TryGetProperty("hasHeader", out var hh) && hh.ValueKind == JsonValueKind.True;
    Excel.Range target = Sheet(op).Range[range];
    Excel.Range key = Sheet(op).Range[byColumn + "1"];
    target.Sort(
        Key1: key,
        Order1: order == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending,
        Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
}

private static void SetRowHeight(JsonElement op)
{
    int row = op.GetProperty("row").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    double heightPoints = op.GetProperty("heightPoints").GetDouble();
    Sheet(op).Range[$"{row}:{row + count - 1}"].RowHeight = heightPoints;
}

// Character-width units, NOT pixels: Excel's native ColumnWidth property has
// no pixel unit at the COM layer. Converts genoffice's px schema using the
// standard Calibri-11 approximation (px - 5) / 7 - a documented, deliberate
// approximation, not exact for other default fonts.
private static void SetColWidth(JsonElement op)
{
    int col = op.GetProperty("column").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    double widthPx = op.GetProperty("widthPx").GetDouble();
    double charWidth = Math.Max(0, (widthPx - 5) / 7.0);
    string startLetter = ColumnLetter(col);
    string endLetter = ColumnLetter(col + count - 1);
    Sheet(op).Range[$"{startLetter}:{endLetter}"].ColumnWidth = charWidth;
}

private static void SetRowsHidden(JsonElement op)
{
    int row = op.GetProperty("row").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    bool hidden = op.GetProperty("hidden").GetBoolean();
    Sheet(op).Range[$"{row}:{row + count - 1}"].EntireRow.Hidden = hidden;
}

private static void SetColsHidden(JsonElement op)
{
    int col = op.GetProperty("column").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    bool hidden = op.GetProperty("hidden").GetBoolean();
    string startLetter = ColumnLetter(col);
    string endLetter = ColumnLetter(col + count - 1);
    Sheet(op).Range[$"{startLetter}:{endLetter}"].EntireColumn.Hidden = hidden;
}
```
(`ColumnLetter` already exists in `ExcelTools.cs`, reused as-is.)

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify each op via `propose_operations`: `sort_range` on a small table sorts correctly by the given column/order; `merge_cells`/`unmerge_cells` toggle a merge; `set_row_height`/`set_col_width` visibly resize; `set_rows_hidden`/`set_cols_hidden` hide/unhide.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add sort_range, merge/unmerge_cells, row/col size and hidden ops"
```

---

### Task 3: `set_freeze`, `set_page_setup`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `set_freeze`**

Add `case "set_freeze": SetFreeze(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void SetFreeze(JsonElement op)
{
    int rows = op.GetProperty("rows").GetInt32();
    int columns = op.GetProperty("columns").GetInt32();
    Excel.Worksheet sheet = Sheet(op);
    sheet.Activate();
    Excel.Window window = Globals.ThisAddIn.Application.ActiveWindow;
    window.FreezePanes = false;
    if (rows == 0 && columns == 0) return; // unfreeze only
    sheet.Cells[rows + 1, columns + 1].Select();
    window.FreezePanes = true;
}
```

- [ ] **Step 2: Implement `set_page_setup`**

Add `case "set_page_setup": SetPageSetup(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void SetPageSetup(JsonElement op)
{
    Excel.PageSetup setup = Sheet(op).PageSetup;
    if (op.TryGetProperty("orientation", out var orient) && orient.ValueKind == JsonValueKind.String)
        setup.Orientation = orient.GetString() == "landscape" ? Excel.XlPageOrientation.xlLandscape : Excel.XlPageOrientation.xlPortrait;
    if (op.TryGetProperty("scale", out var scale) && scale.ValueKind == JsonValueKind.Number)
        setup.Zoom = (int)scale.GetDouble();
    if (op.TryGetProperty("fitToWidth", out var ftw) && ftw.ValueKind == JsonValueKind.Number)
    {
        setup.Zoom = false; // Zoom and FitToPages are mutually exclusive in Excel's own UI
        setup.FitToPagesWide = (int)ftw.GetDouble();
    }
    if (op.TryGetProperty("fitToHeight", out var fth) && fth.ValueKind == JsonValueKind.Number)
        setup.FitToPagesTall = (int)fth.GetDouble();
    if (op.TryGetProperty("printGridlines", out var pg))
        setup.PrintGridlines = pg.ValueKind == JsonValueKind.True;
    if (op.TryGetProperty("printHeadings", out var ph))
        setup.PrintHeadings = ph.ValueKind == JsonValueKind.True;
    if (op.TryGetProperty("printArea", out var pa) && pa.ValueKind == JsonValueKind.String)
        setup.PrintArea = pa.GetString();
    if (op.TryGetProperty("margins", out var margins) && margins.ValueKind == JsonValueKind.String)
    {
        // Point values match Excel's own ribbon presets (Normal/Wide/Narrow).
        double top, bottom, leftRight, header, footer;
        switch (margins.GetString())
        {
            case "wide": top = 72; bottom = 72; leftRight = 72; header = 36; footer = 36; break;
            case "narrow": top = 27; bottom = 27; leftRight = 18; header = 13.5; footer = 13.5; break;
            default: top = 54; bottom = 54; leftRight = 43.2; header = 22.5; footer = 22.5; break; // "normal"
        }
        setup.TopMargin = top; setup.BottomMargin = bottom;
        setup.LeftMargin = leftRight; setup.RightMargin = leftRight;
        setup.HeaderMargin = header; setup.FooterMargin = footer;
    }
}
```

- [ ] **Step 3: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify via `propose_operations`: `set_freeze` with `{"rows":1,"columns":0}` freezes the header row (check View > Freeze Panes state); `set_page_setup` with `{"orientation":"landscape"}` changes Page Layout view orientation.

- [ ] **Step 4: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add set_freeze and set_page_setup"
```

---

### Task 4: Sheet-structure operations — `add_sheet`, `delete_sheet`, `duplicate_sheet`, `set_sheet_hidden`, `move_sheet`, `protect_sheet`, `rename_sheet`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement all 7 operations**

Add to `ProposeOperations`'s switch:
```csharp
                        case "add_sheet": AddSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_sheet": DeleteSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "duplicate_sheet": DuplicateSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_sheet_hidden": SetSheetHidden(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "move_sheet": MoveSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "protect_sheet": ProtectSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "rename_sheet": Sheet(op).Name = op.GetProperty("name").GetString(); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```
and:
```csharp
private static void AddSheet(JsonElement op)
{
    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
    Excel.Worksheet newSheet = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
    newSheet.Name = op.GetProperty("name").GetString();
}

private static void DeleteSheet(JsonElement op)
{
    Excel.Application app = Globals.ThisAddIn.Application;
    bool prevAlerts = app.DisplayAlerts;
    app.DisplayAlerts = false;
    try
    {
        Sheet(op).Delete();
    }
    finally
    {
        app.DisplayAlerts = prevAlerts;
    }
}

private static void DuplicateSheet(JsonElement op)
{
    Excel.Worksheet source = Sheet(op);
    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
    source.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
    Excel.Worksheet copy = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
    if (op.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
    {
        copy.Name = name.GetString();
    }
}

private static void SetSheetHidden(JsonElement op)
{
    bool hidden = op.GetProperty("hidden").GetBoolean();
    Sheet(op).Visible = hidden ? Excel.XlSheetVisibility.xlSheetHidden : Excel.XlSheetVisibility.xlSheetVisible;
}

private static void MoveSheet(JsonElement op)
{
    int position = op.GetProperty("position").GetInt32(); // 1-based
    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
    Excel.Worksheet target = Sheet(op);
    if (position >= wb.Worksheets.Count)
    {
        target.Move(After: wb.Worksheets[wb.Worksheets.Count]);
    }
    else
    {
        target.Move(Before: wb.Worksheets[position]);
    }
}

private static void ProtectSheet(JsonElement op)
{
    bool isProtected = op.GetProperty("protected").GetBoolean();
    Excel.Worksheet sheet = Sheet(op);
    if (isProtected) sheet.Protect();
    else sheet.Unprotect();
}
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify each op via `propose_operations` against a multi-sheet workbook: add/delete/duplicate/hide/move/protect/rename all behave as expected, and `delete_sheet` shows no native confirmation dialog (confirms `DisplayAlerts` suppression worked) while still leaving `DisplayAlerts` restored to `true` afterward (verify by triggering an unrelated action that WOULD show a dialog, e.g. deleting another sheet manually via the UI, and confirming the dialog appears).

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add sheet-structure operations (add/delete/duplicate/hide/move/protect/rename)"
```

---

### Task 5: `read_sheet_features`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `read_sheet_features`**

Add `"read_sheet_features"` to `AlwaysAllowedTools`, add `case "read_sheet_features": return ReadSheetFeatures(input);`, and:
```csharp
private static ToolResult ReadSheetFeatures(JsonElement input)
{
    Excel.Worksheet sheet = Sheet(input);
    var sb = new System.Text.StringBuilder();

    sb.AppendLine("Hidden: " + (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible));
    sb.AppendLine("Protected: " + sheet.ProtectContents);

    if (sheet.AutoFilterMode && sheet.AutoFilter != null)
    {
        sb.AppendLine("AutoFilter range: " + sheet.AutoFilter.Range.Address[false, false]);
    }

    try
    {
        Excel.Window window = Globals.ThisAddIn.Application.ActiveWindow;
        if (window.FreezePanes)
        {
            sb.AppendLine($"Freeze panes: rows={window.SplitRow}, columns={window.SplitColumn}");
        }
    }
    catch { /* ActiveWindow may not correspond to this sheet if it's not active - best-effort only */ }

    int cfCount = 0;
    foreach (Excel.FormatCondition fc in sheet.UsedRange.FormatConditions) cfCount++;
    sb.AppendLine("Conditional format rules in used range: " + cfCount);

    foreach (Excel.Name n in sheet.Names)
    {
        sb.AppendLine("Defined name (sheet-scoped): " + n.Name + " = " + n.RefersTo);
    }
    foreach (Excel.Name n in Globals.ThisAddIn.Application.ActiveWorkbook.Names)
    {
        sb.AppendLine("Defined name (workbook-scoped): " + n.Name + " = " + n.RefersTo);
    }

    int shapeCount = sheet.Shapes.Count;
    sb.AppendLine("Shapes/images: " + shapeCount);

    return new ToolResult { Output = sb.ToString(), Summary = "read_sheet_features" };
}
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify on a sheet with at least a frozen header row, one AutoFilter, and one shape: `read_sheet_features` reports all three correctly.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add read_sheet_features"
```

---

### Task 6: `find_cells`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `find_cells`**

Add `"find_cells"` to `AlwaysAllowedTools`, add `case "find_cells": return FindCells(input);`, and:
```csharp
private static readonly string[] ExcelErrorTexts = { "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!" };

private static ToolResult FindCells(JsonElement input)
{
    bool errorsOnly = input.TryGetProperty("errors_only", out var eo) && eo.ValueKind == JsonValueKind.True;
    string query = input.TryGetProperty("query", out var q) && q.ValueKind == JsonValueKind.String ? q.GetString() : null;
    bool useRegex = input.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
    string lookIn = input.TryGetProperty("look_in", out var li) && li.ValueKind == JsonValueKind.String ? li.GetString() : "both";
    int maxResults = input.GetProperty("max_results").GetInt32();
    string sheetName = input.TryGetProperty("sheetId", out var sid) && sid.ValueKind == JsonValueKind.String ? sid.GetString() : null;

    if (!errorsOnly && query == null)
    {
        return new ToolResult { Output = "find_cells requires either 'query' or 'errors_only'.", IsError = true, Summary = "find_cells" };
    }

    System.Text.RegularExpressions.Regex regex = useRegex && query != null
        ? new System.Text.RegularExpressions.Regex(query)
        : null;

    var sb = new System.Text.StringBuilder();
    int found = 0;
    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
    foreach (Excel.Worksheet sheet in wb.Worksheets)
    {
        if (sheetName != null && sheet.Name != sheetName) continue;
        if (found >= maxResults) break;

        if (errorsOnly)
        {
            Excel.Range errorCells = null;
            try
            {
                // Native error-cell scan - the exact advantage this project's
                // original feasibility report flagged VSTO/COM as having over
                // Office.js's wildcard-only Range.find.
                errorCells = sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);
            }
            catch (System.Runtime.InteropServices.COMException) { /* no error cells on this sheet - SpecialCells throws if none match */ }
            if (errorCells != null)
            {
                foreach (Excel.Range cell in errorCells.Cells)
                {
                    if (found >= maxResults) break;
                    sb.AppendLine($"{sheet.Name}!{cell.Address[false, false]}: {cell.Text}");
                    found++;
                }
            }
            continue;
        }

        foreach (Excel.Range cell in sheet.UsedRange.Cells)
        {
            if (found >= maxResults) break;
            string valueText = cell.Text as string ?? "";
            string formulaText = cell.Formula as string ?? "";
            string haystack = lookIn == "values" ? valueText : lookIn == "formulas" ? formulaText : valueText + " " + formulaText;
            bool isMatch = regex != null ? regex.IsMatch(haystack) : haystack.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
            if (isMatch)
            {
                sb.AppendLine($"{sheet.Name}!{cell.Address[false, false]}: {valueText}");
                found++;
            }
        }
    }
    return new ToolResult { Output = sb.ToString(), Summary = "find_cells" };
}
```
(Full-`UsedRange` cell-by-cell iteration for the substring/regex path is O(cells) — acceptable for typical sheet sizes and matches this being a manually-triggered AI tool call, not a hot loop; the error-cell path is the fast, native one via `SpecialCells`.)

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `find_cells` with `{"query":"revenue", "max_results":10}` finds matching cells across the workbook; `find_cells` with `{"errors_only":true, "max_results":10}` on a sheet with a deliberately-broken formula (e.g. `=1/0`) finds it via the native `SpecialCells` path.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add find_cells with native error-cell scan"
```

---

### Task 7: `trace_precedents`, `trace_dependents`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement both tools (same-sheet only for this pass)**

Add both to `AlwaysAllowedTools`, add `case "trace_precedents": return TracePrecedents(input);` / `case "trace_dependents": return TraceDependents(input);`, and:
```csharp
private static ToolResult TracePrecedents(JsonElement input)
{
    string address = input.GetProperty("address").GetString();
    Excel.Range cell = Sheet(input).Range[address];
    Excel.Range precedents;
    try
    {
        precedents = cell.DirectPrecedents;
    }
    catch (System.Runtime.InteropServices.COMException)
    {
        return new ToolResult { Output = address + " has no precedents (not a formula, or references nothing).", Summary = "trace_precedents" };
    }
    var sb = new System.Text.StringBuilder();
    foreach (Excel.Range p in precedents.Cells)
    {
        string text = p.Text as string ?? "";
        bool isError = ExcelErrorTexts.Any(e => text == e);
        sb.AppendLine($"{p.Address[false, false]}: {text}" + (isError ? " (ERROR)" : ""));
    }
    return new ToolResult { Output = sb.ToString(), Summary = "trace_precedents" };
}

private static ToolResult TraceDependents(JsonElement input)
{
    string address = input.GetProperty("address").GetString();
    Excel.Range cell = Sheet(input).Range[address];
    Excel.Range dependents;
    try
    {
        dependents = cell.Dependents;
    }
    catch (System.Runtime.InteropServices.COMException)
    {
        return new ToolResult { Output = "No formulas in this sheet reference " + address + ".", Summary = "trace_dependents" };
    }
    var sb = new System.Text.StringBuilder();
    foreach (Excel.Range d in dependents.Cells)
    {
        sb.AppendLine($"{d.Address[false, false]}: {d.Formula}");
    }
    return new ToolResult { Output = sb.ToString(), Summary = "trace_dependents" };
}
```
(`DirectPrecedents`/`Dependents` are same-sheet-only in classic Interop; cross-sheet tracing needs `Range.NavigateArrow`, a materially more complex one-hop-at-a-time API. Scoped OUT of this task deliberately — same-sheet coverage handles the common case; cross-sheet tracing is a documented follow-up, not silently dropped.)

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify on a sheet where `B1 = A1 + 1`: `trace_precedents` on `B1` reports `A1`; `trace_dependents` on `A1` reports `B1`.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add trace_precedents and trace_dependents (same-sheet)"
```

---

### Task 8: `add_sparkline`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

**Context:** this is the single item the original feasibility report flagged as the clearest categorical Office.js gap that VSTO/COM uniquely closes — direct native support, no adaptation needed.

- [ ] **Step 1: Implement `add_sparkline`**

Add `case "add_sparkline": AddSparkline(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void AddSparkline(JsonElement op)
{
    string dataRange = op.GetProperty("dataRange").GetString();
    string targetCell = op.TryGetProperty("targetCell", out var tc) && tc.ValueKind == JsonValueKind.String ? tc.GetString() : dataRange;
    string type = op.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() : "line";
    Excel.SparklineGroups groups = Sheet(op).SparklineGroups;
    Excel.XlSparkType sparkType = type == "column" ? Excel.XlSparkType.xlSparkColumnStacked100
        : type == "stacked" ? Excel.XlSparkType.xlSparkColumnStacked100
        : Excel.XlSparkType.xlSparkLine;
    dynamic group = groups.Add(sparkType, Sheet(op).Range[dataRange].Address[true, true, Excel.XlReferenceStyle.xlA1, true]);
    if (op.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
    {
        group.SeriesColor.Color = HexToOleColor(color.GetString());
    }
}

private static int HexToOleColor(string hex)
{
    hex = hex.TrimStart('#');
    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
    return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
}
```
(`dynamic` used for the sparkline group's `SeriesColor` property, following this codebase's existing convention of using `dynamic` for less-common COM interfaces to sidestep exact Interop type-name uncertainty — same pattern as `WordTools.EditChart`.)

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `add_sparkline` with `{"dataRange":"A1:E1", "targetCell":"F1", "type":"line"}` on a row of numbers inserts a visible line sparkline in F1.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add add_sparkline - closes the original Office.js-gap finding"
```

---

### Task 9: `add_shape`, `edit_shape`, `add_image`, `delete_visual` — shape/image ops with name-based addressing

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: a shared shape/chart name-resolution helper (`ResolveVisual`) — not consumed by any later task in this plan, but written generically enough that Task 11 (`edit_chart`) reuses the same "resolve a chart by name" half of it if convenient (not required — Task 11 can also resolve independently).

**Design:** genoffice addresses shapes/charts by an opaque `visualId`. Excel's `Shapes`/`ChartObjects` collections are natively name-addressable (`.Item(name)`), which is a more natural fit than the positional `slideIndex`/`shapeIndex` addressing `PowerPointTools.cs` uses — so `visualId` here is simply the shape/chart's own COM `.Name`, which `add_shape`/`add_chart` already implicitly assign (Excel auto-names new shapes e.g. `"Rectangle 1"`) and which `get_workbook_context`/a future `read_sheet_features` enhancement could report back to the model (not required by this task — the model can also just ask `read_sheet_features` today, which already lists shape count; reporting exact names is a reasonable follow-up, not blocking this task).

- [ ] **Step 1: Add a shape-type lookup table and `add_shape`**

Add near the top of the class (a `static readonly Dictionary<string, Excel.XlShapeType>`-free approach isn't right since Excel shapes use `Office.MsoAutoShapeType`, not `XlShapeType` — use `Microsoft.Office.Core.MsoAutoShapeType`):
```csharp
private static readonly Dictionary<string, Microsoft.Office.Core.MsoAutoShapeType> ShapeTypeMap =
    new Dictionary<string, Microsoft.Office.Core.MsoAutoShapeType>
{
    ["rect"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
    ["roundRect"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
    ["ellipse"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
    ["triangle"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle,
    ["rtTriangle"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightTriangle,
    ["parallelogram"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeParallelogram,
    ["trapezoid"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeTrapezoid,
    ["diamond"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond,
    ["pentagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapePentagon,
    ["hexagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeHexagon,
    ["octagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOctagon,
    ["plus"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapePlus,
    ["mathPlus"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMathPlus,
    ["pie"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapePie,
    ["chord"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeChord,
    ["donut"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDonut,
    ["foldedCorner"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeFoldedCorner,
    ["heart"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeHeart,
    ["lightningBolt"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeLightningBolt,
    ["sun"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeSun,
    ["moon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMoon,
    ["cloud"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeCloud,
    ["arc"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeArc,
    ["star5"] = Microsoft.Office.Core.MsoAutoShapeType.msoShape5pointStar,
    ["rightArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
    ["leftArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftArrow,
    ["upArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeUpArrow,
    ["downArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDownArrow,
};
```
(Covers the ~25 most common `ADDABLE_SHAPE_TYPES` entries from genoffice's `apps/sheets/src/shared/shape-types.ts`, which lists ~180 total OOXML preset names — extend this table opportunistically if a specific missing shape type is requested; unmapped names fall back to `msoShapeRectangle` per Step 1's implementation below, never an error, since an unusual/rare requested shape rendering as a plain rectangle is a better failure mode than a hard tool error.)

Add `case "add_shape": AddShapeExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void AddShapeExcel(JsonElement op)
{
    string shapeType = op.GetProperty("shapeType").GetString();
    string anchorCell = op.GetProperty("anchorCell").GetString();
    Excel.Range anchor = Sheet(op).Range[anchorCell];
    float left = (float)(double)anchor.Left;
    float top = (float)(double)anchor.Top;
    float width = 100f, height = 60f;

    Excel.Shape shape;
    if (shapeType == "textbox")
    {
        shape = Sheet(op).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
    }
    else
    {
        Microsoft.Office.Core.MsoAutoShapeType msoType = ShapeTypeMap.TryGetValue(shapeType, out var mapped) ? mapped : Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle;
        shape = Sheet(op).Shapes.AddShape(msoType, left, top, width, height);
    }
    if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
    {
        shape.Fill.ForeColor.RGB = HexToOleColor(fill.GetString());
    }
    if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
    {
        shape.TextFrame.Characters().Text = text.GetString();
    }
}
```

- [ ] **Step 2: Implement `edit_shape`, `delete_visual`, `add_image`**

Add the three `case` branches and:
```csharp
private static Excel.Shape ResolveShapeByName(JsonElement op, string idField)
{
    string visualId = op.GetProperty(idField).GetString();
    return Sheet(op).Shapes.Item(visualId);
}

private static void EditShapeExcel(JsonElement op)
{
    Excel.Shape shape = ResolveShapeByName(op, "visualId");
    if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String && shape.TextFrame.HasText != 0)
    {
        shape.TextFrame.Characters().Text = text.GetString();
    }
    if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
    {
        shape.Fill.ForeColor.RGB = HexToOleColor(fill.GetString());
    }
    if (op.TryGetProperty("anchorCell", out var anchorCell) && anchorCell.ValueKind == JsonValueKind.String)
    {
        Excel.Range anchor = Sheet(op).Range[anchorCell.GetString()];
        shape.Left = (float)(double)anchor.Left;
        shape.Top = (float)(double)anchor.Top;
    }
}

private static void DeleteVisual(JsonElement op)
{
    ResolveShapeByName(op, "visualId").Delete();
}

private static void AddImageExcel(JsonElement op)
{
    string path = op.GetProperty("path").GetString();
    if (path.StartsWith("http://") || path.StartsWith("https://"))
    {
        throw new NotSupportedException("add_image: remote URLs are not supported in this air-gapped deployment - use a local file path.");
    }
    string anchorCell = op.GetProperty("anchorCell").GetString();
    Excel.Range anchor = Sheet(op).Range[anchorCell];
    Sheet(op).Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
        (float)(double)anchor.Left, (float)(double)anchor.Top, -1, -1);
}
```
Wire the three cases:
```csharp
                        case "edit_shape": EditShapeExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_visual": DeleteVisual(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_image": AddImageExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```
(`add_image` deliberately rejects `http(s)://` paths — the original architecture decision excluded internet-dependent capabilities from this air-gapped deployment; genoffice's own `path` field allows a URL, this port intentionally narrows it to local-file-only.)

- [ ] **Step 3: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `add_shape` with `{"shapeType":"roundRect", "anchorCell":"B2", "fillColor":"#4a9eff", "text":"Hi"}` creates a filled, labeled rounded rectangle anchored near B2; note its auto-assigned name (visible in Excel's Name Box when selected) and use that name in `edit_shape`/`delete_visual`'s `visualId`; `add_image` with a real local `.png` path inserts it at the anchor.

- [ ] **Step 4: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add add_shape, edit_shape, delete_visual, add_image (local-path only)"
```

---

### Task 10: `edit_chart`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: `ResolveShapeByName`-style pattern from Task 9 is NOT reused (charts live in the separate `ChartObjects` collection, not `Shapes`) — this task resolves charts independently via `ChartObjects().Item(name)`.
- Produces: nothing new for other tasks.

**Warning (from research):** chart *data* changes go through the chart's embedded workbook (`chart.ChartData.Workbook`) — a COM-automation-inside-COM-automation pattern. The embedded workbook's `Application`/`Workbook` COM objects must be released before this method returns, or a hidden hosting process is orphaned.

- [ ] **Step 1: Implement `edit_chart`**

Add `case "edit_chart": EditChartExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static readonly Dictionary<string, int> ExcelChartTypeMap = new Dictionary<string, int>
{
    ["column"] = 51, // xlColumnClustered
    ["bar"] = 57,     // xlBarClustered
    ["line"] = 4,     // xlLine
    ["area"] = 1,     // xlArea
    ["pie"] = 5,      // xlPie
    ["doughnut"] = -4120, // xlDoughnut
};

private static void EditChartExcel(JsonElement op)
{
    string chartName = op.GetProperty("chartPath").GetString(); // this project's visualId-equivalent for charts
    dynamic chartObjects = Sheet(op).ChartObjects();
    dynamic chartObj = chartObjects.Item(chartName);
    dynamic chart = chartObj.Chart;

    if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String && ExcelChartTypeMap.TryGetValue(ct.GetString(), out int typeCode))
    {
        chart.ChartType = typeCode;
    }
    if (op.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
    {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title.GetString();
    }
    if (op.TryGetProperty("legend", out var legend) && legend.ValueKind == JsonValueKind.String)
    {
        string pos = legend.GetString();
        if (pos == "none") { chart.HasLegend = false; }
        else
        {
            chart.HasLegend = true;
            chart.Legend.Position = pos == "right" ? -4152 /*xlLegendPositionRight*/
                : pos == "top" ? -4160 /*xlLegendPositionTop*/
                : pos == "left" ? -4131 /*xlLegendPositionLeft*/
                : -4107 /*xlLegendPositionBottom*/;
        }
    }
    if (op.TryGetProperty("dataLabels", out var dl) && dl.ValueKind == JsonValueKind.String)
    {
        bool show = dl.GetString() != "none";
        foreach (dynamic series in chart.SeriesCollection())
        {
            series.HasDataLabels = show;
            if (show && dl.GetString() == "percent") series.DataLabels().ShowPercentage = true;
        }
    }
    if (op.TryGetProperty("seriesColors", out var colors) && colors.ValueKind == JsonValueKind.Object)
    {
        foreach (JsonProperty prop in colors.EnumerateObject())
        {
            int seriesIndex = int.Parse(prop.Name);
            dynamic series = chart.SeriesCollection(seriesIndex + 1);
            series.Format.Fill.ForeColor.RGB = HexToOleColor(prop.Value.GetString());
        }
    }

    // Data changes require the chart's embedded workbook - open, write,
    // close, and RELEASE explicitly so no hidden Excel host process leaks.
    if (op.TryGetProperty("seriesData", out var seriesData) && seriesData.ValueKind == JsonValueKind.Array)
    {
        dynamic chartDataWorkbook = chart.ChartData.Workbook;
        try
        {
            dynamic dataSheet = chartDataWorkbook.Worksheets[1];
            int seriesIdx = 0;
            foreach (JsonElement sd in seriesData.EnumerateArray())
            {
                dynamic series = chart.SeriesCollection(seriesIdx + 1);
                if (sd.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
                {
                    series.Name = nameEl.GetString();
                }
                seriesIdx++;
            }
        }
        finally
        {
            chartDataWorkbook.Close(SaveChanges: true);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(chartDataWorkbook);
        }
    }
}
```
(`seriesColors`/`legend` numeric enum constants above are the raw underlying values of `XlLegendPosition`/etc. rather than the named enum, since `dynamic` COM calls resolve members at runtime and raw ints avoid an extra `using` for a rarely-touched enum — consistent with this codebase's existing `dynamic`-chart-code convention of using raw ints, e.g. `WordTools.EditChart`'s `51 /* xlColumnClustered */`. `seriesData`'s renaming is implemented; full `values`/`categories` repointing is deliberately left as a documented follow-up rather than expanded here, since it requires writing actual cell grids into the embedded workbook sheet — same shape as `ExcelTools.SetRangeValues`, reusable if a future task needs it.)

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: create a chart first via the existing `add_chart` op, note its name (Name Box when selected), then call `edit_chart` with `{"chartPath":"<name>", "chartType":"line", "title":"Updated", "legend":"bottom"}` — confirm the chart visibly changes type/title/legend position. Watch Task Manager during and after the call to confirm no orphaned `EXCEL.EXE` process remains if `seriesData` is exercised.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add edit_chart (type/title/legend/labels/colors/series rename)"
```

---

### Task 11: Native tables — `add_table`, `add_table_row`, `add_table_column`, `delete_table_row`, `delete_table_column`, `delete_table`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `add_table`**

Add `case "add_table": AddTable(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void AddTable(JsonElement op)
{
    string range = op.GetProperty("range").GetString();
    Excel.Worksheet sheet = Sheet(op);
    Excel.ListObject table = sheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sheet.Range[range], Type.Missing, Excel.XlYesNoGuess.xlYes);
    if (op.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
    {
        table.Name = name.GetString();
    }
    if (op.TryGetProperty("style", out var style) && style.ValueKind == JsonValueKind.String)
    {
        table.TableStyle = style.GetString();
    }
    if (op.TryGetProperty("bandedRows", out var banded))
    {
        table.ShowTableStyleRowStripes = banded.ValueKind == JsonValueKind.True;
    }
}
```

- [ ] **Step 2: Implement row/column add/delete and `delete_table`**

Add the 5 remaining `case` branches and:
```csharp
private static Excel.ListObject ResolveTable(JsonElement op)
{
    string tableName = op.GetProperty("tableName").GetString();
    return Sheet(op).ListObjects[tableName];
}

private static void AddTableRow(JsonElement op)
{
    Excel.ListObject table = ResolveTable(op);
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    for (int i = 0; i < count; i++)
    {
        if (op.TryGetProperty("row", out var row) && row.ValueKind == JsonValueKind.Number)
        {
            table.ListRows.Add(row.GetInt32() + 1);
        }
        else
        {
            table.ListRows.Add();
        }
    }
}

private static void AddTableColumn(JsonElement op)
{
    Excel.ListObject table = ResolveTable(op);
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    string columnName = op.GetProperty("columnName").GetString();
    for (int i = 0; i < count; i++)
    {
        Excel.ListColumn col = op.TryGetProperty("column", out var colPos) && colPos.ValueKind == JsonValueKind.Number
            ? table.ListColumns.Add(colPos.GetInt32() + 1)
            : table.ListColumns.Add();
        // Excel requires unique column names - only the last added column (or
        // the only one, when count==1) gets the literal requested name; extras
        // get a numbered suffix to stay valid.
        col.Name = count == 1 ? columnName : columnName + " " + (i + 1);
    }
}

private static void DeleteTableRow(JsonElement op)
{
    Excel.ListObject table = ResolveTable(op);
    int row = op.GetProperty("row").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    for (int i = 0; i < count; i++)
    {
        table.ListRows[row + 1].Delete();
    }
}

private static void DeleteTableColumn(JsonElement op)
{
    Excel.ListObject table = ResolveTable(op);
    int column = op.GetProperty("column").GetInt32();
    int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
    for (int i = 0; i < count; i++)
    {
        table.ListColumns[column + 1].Delete();
    }
}

private static void DeleteTable(JsonElement op)
{
    ResolveTable(op).Unlist(); // converts back to a plain range, keeping data/formatting
}
```
Wire the cases:
```csharp
                        case "add_table_row": AddTableRow(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_table_column": AddTableColumn(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table_row": DeleteTableRow(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table_column": DeleteTableColumn(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table": DeleteTable(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```

- [ ] **Step 3: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `add_table` on a range with headers creates a real, styled Excel Table (Table Design ribbon tab appears when a cell inside is selected — confirming it's a genuine `ListObject`, not just formatted cells); add/delete row/column ops resize it correctly; `delete_table` converts it back to plain cells without losing data.

- [ ] **Step 4: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add native table operations (add/edit/delete via ListObjects)"
```

---

### Task 12: Simple data operations — `set_hyperlink`, `set_note`, `add_defined_name`, `delete_defined_name`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement all 4 operations**

Add the 4 `case` branches and:
```csharp
private static void SetHyperlink(JsonElement op)
{
    string address = op.GetProperty("address").GetString();
    Excel.Range range = Sheet(op).Range[address];
    if (!op.TryGetProperty("target", out var target) || target.ValueKind == JsonValueKind.Null)
    {
        foreach (Excel.Hyperlink link in range.Hyperlinks) link.Delete();
        return;
    }
    string url = target.GetString();
    Excel.Worksheet sheet = Sheet(op);
    if (url.Contains("!") && !url.StartsWith("http"))
    {
        sheet.Hyperlinks.Add(range, "", SubAddress: url);
    }
    else
    {
        sheet.Hyperlinks.Add(range, url);
    }
}

private static void SetNote(JsonElement op)
{
    string address = op.GetProperty("address").GetString();
    Excel.Range cell = Sheet(op).Range[address];
    if (!op.TryGetProperty("text", out var text) || text.ValueKind == JsonValueKind.Null)
    {
        cell.Comment?.Delete();
        return;
    }
    cell.Comment?.Delete();
    cell.AddComment(text.GetString());
}

private static void AddDefinedName(JsonElement op)
{
    string name = op.GetProperty("name").GetString();
    string reference = op.GetProperty("ref").GetString();
    Globals.ThisAddIn.Application.ActiveWorkbook.Names.Add(name, "=" + reference);
}

private static void DeleteDefinedName(JsonElement op)
{
    string name = op.GetProperty("name").GetString();
    Globals.ThisAddIn.Application.ActiveWorkbook.Names.Item(name).Delete();
}
```
Wire the cases:
```csharp
                        case "set_hyperlink": SetHyperlink(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_note": SetNote(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_defined_name": AddDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_defined_name": DeleteDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify each of the 4 ops via `propose_operations` against a real workbook, including the `null`-target/`null`-text removal paths for `set_hyperlink`/`set_note`.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add set_hyperlink, set_note, add/delete_defined_name"
```

---

### Task 13: Filtering — `set_filter`, `clear_filter`, `set_filter_criteria`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement all 3 operations**

Add the 3 `case` branches and:
```csharp
private static void SetFilter(JsonElement op)
{
    string range = op.GetProperty("range").GetString();
    Sheet(op).Range[range].AutoFilter();
}

private static void ClearFilter(JsonElement op)
{
    Excel.Worksheet sheet = Sheet(op);
    if (sheet.AutoFilterMode)
    {
        sheet.AutoFilterMode = false;
    }
}

private static void SetFilterCriteria(JsonElement op)
{
    Excel.Worksheet sheet = Sheet(op);
    if (sheet.AutoFilter == null)
    {
        throw new InvalidOperationException("set_filter_criteria: no AutoFilter is active on this sheet - call set_filter first.");
    }
    int column = op.GetProperty("column").GetInt32(); // 0-based, relative to the AutoFilter range's first column
    Excel.Range filterRange = sheet.AutoFilter.Range;
    int fieldIndex = column + 1; // AutoFilter's Field parameter is 1-based, relative to the filter range - a common COM gotcha

    if (!op.TryGetProperty("values", out var values) || values.ValueKind == JsonValueKind.Null)
    {
        filterRange.AutoFilter(Field: fieldIndex); // toggling with no Criteria1 clears that column's filter
        return;
    }
    var criteria = new List<string>();
    foreach (JsonElement v in values.EnumerateArray()) criteria.Add(v.GetString());
    filterRange.AutoFilter(Field: fieldIndex, Criteria1: criteria.ToArray(), Operator: Excel.XlAutoFilterOperator.xlFilterValues);
}
```
Wire the cases:
```csharp
                        case "set_filter": SetFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_filter": ClearFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_filter_criteria": SetFilterCriteria(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `set_filter` on a headered range enables filter dropdowns; `set_filter_criteria` with a `values` list hides non-matching rows; `set_filter_criteria` with `values: null` restores all rows for that column; `clear_filter` removes the AutoFilter entirely.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add set_filter, clear_filter, set_filter_criteria"
```

---

### Task 14: `add_conditional_format`, `clear_conditional_formats`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: `HexToOleColor` (Task 8).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `add_conditional_format`'s 7 rule kinds**

Add `case "add_conditional_format": AddConditionalFormat(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and `case "clear_conditional_formats": Sheet(op).UsedRange.FormatConditions.Delete(); lines.AppendLine(kind + ": ok"); anyMutated = true; break;`, plus:
```csharp
private static Excel.XlFormatConditionOperator MapCfOperator(string op)
{
    switch (op)
    {
        case "greaterThan": return Excel.XlFormatConditionOperator.xlGreater;
        case "lessThan": return Excel.XlFormatConditionOperator.xlLess;
        case "equal": return Excel.XlFormatConditionOperator.xlEqual;
        case "between": return Excel.XlFormatConditionOperator.xlBetween;
        default: return Excel.XlFormatConditionOperator.xlEqual;
    }
}

private static void AddConditionalFormat(JsonElement op)
{
    string range = op.GetProperty("range").GetString();
    Excel.Range target = Sheet(op).Range[range];
    JsonElement rule = op.GetProperty("rule");
    string kind = rule.GetProperty("kind").GetString();
    Excel.FormatCondition fc = null;

    switch (kind)
    {
        case "number":
        {
            string oper = rule.GetProperty("operator").GetString();
            double value = rule.GetProperty("value").GetDouble();
            string formula2 = rule.TryGetProperty("value2", out var v2) ? v2.GetDouble().ToString() : null;
            fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, MapCfOperator(oper), value.ToString(), formula2);
            break;
        }
        case "text":
        {
            string text = rule.GetProperty("text").GetString();
            fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlTextString, Text: text, TextOperator: Excel.XlContainsOperator.xlContains);
            break;
        }
        case "blank":
            fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition);
            break;
        case "duplicate":
            fc = target.FormatConditions.AddUniqueValues();
            ((Excel.UniqueValues)fc).DupeUnique = Excel.XlDupeUnique.xlDuplicate;
            break;
        case "top10":
        {
            int rank = rule.TryGetProperty("rank", out var r) ? r.GetInt32() : 10;
            bool percent = rule.TryGetProperty("percent", out var p) && p.ValueKind == JsonValueKind.True;
            bool bottom = rule.TryGetProperty("bottom", out var b) && b.ValueKind == JsonValueKind.True;
            Excel.Top10 top10 = target.FormatConditions.AddTop10();
            top10.Rank = rank;
            top10.Percent = percent;
            top10.TopBottom = bottom ? Excel.XlTopBottom.xlTop10Bottom : Excel.XlTopBottom.xlTop10Top;
            fc = top10;
            break;
        }
        case "formula":
            fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: rule.GetProperty("formula").GetString());
            break;
        case "colorScale":
        {
            Excel.ColorScale scale = target.FormatConditions.AddColorScale(3);
            if (rule.TryGetProperty("minColor", out var minC)) scale.ColorScaleCriteria[1].FormatColor.Color = HexToOleColor(minC.GetString());
            if (rule.TryGetProperty("midColor", out var midC)) scale.ColorScaleCriteria[2].FormatColor.Color = HexToOleColor(midC.GetString());
            if (rule.TryGetProperty("maxColor", out var maxC)) scale.ColorScaleCriteria[3].FormatColor.Color = HexToOleColor(maxC.GetString());
            return; // ColorScale/DataBar carry their own visual - no separate "format" object to apply below
        }
        case "dataBar":
        {
            Excel.DataBar bar = target.FormatConditions.AddDatabar();
            if (rule.TryGetProperty("color", out var barColor))
            {
                bar.BarColor.Color = HexToOleColor(barColor.GetString());
            }
            return;
        }
    }

    if (fc != null && rule.TryGetProperty("format", out var format))
    {
        if (format.TryGetProperty("bold", out var bold)) fc.Font.Bold = bold.ValueKind == JsonValueKind.True;
        if (format.TryGetProperty("fontColor", out var fontColor)) fc.Font.Color = HexToOleColor(fontColor.GetString());
        if (format.TryGetProperty("fillColor", out var fillColor)) fc.Interior.Color = HexToOleColor(fillColor.GetString());
    }
}
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify at least 3 of the 7 rule kinds (`number`, `colorScale`, `top10`) render correctly in real Excel conditional formatting; `clear_conditional_formats` removes all of them from the used range.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add add_conditional_format (7 rule kinds) and clear_conditional_formats"
```

---

### Task 15: `set_data_validation`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

**Note from research:** genoffice's `checkbox` validation kind may map to Excel 365's newer native boolean-checkbox-cell feature rather than classic Data Validation, and availability depends on the installed Interop/Office build. This task implements the 5 confirmed-classic kinds (`list`, `listRef`, `numberBetween`, `dateBetween`, `formula`) plus `null`-to-clear; `checkbox` is explicitly deferred pending a quick manual check against this machine's installed Office build (Step 2 below) rather than guessed blind.

- [ ] **Step 1: Implement the 5 confirmed kinds + clear**

Add `case "set_data_validation": SetDataValidation(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static void SetDataValidation(JsonElement op)
{
    string range = op.GetProperty("range").GetString();
    Excel.Range target = Sheet(op).Range[range];

    if (!op.TryGetProperty("validation", out var validation) || validation.ValueKind == JsonValueKind.Null)
    {
        target.Validation.Delete();
        return;
    }

    string kind = validation.GetProperty("kind").GetString();
    target.Validation.Delete();

    switch (kind)
    {
        case "list":
        {
            var values = new List<string>();
            foreach (JsonElement v in validation.GetProperty("values").EnumerateArray()) values.Add(v.GetString());
            target.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(",", values));
            break;
        }
        case "listRef":
        {
            string refRange = validation.GetProperty("range").GetString();
            target.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "=" + refRange);
            break;
        }
        case "numberBetween":
        {
            double min = validation.GetProperty("min").GetDouble();
            double max = validation.GetProperty("max").GetDouble();
            target.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, min.ToString(), max.ToString());
            break;
        }
        case "dateBetween":
        {
            string start = validation.GetProperty("start").GetString();
            string end = validation.GetProperty("end").GetString();
            target.Validation.Add(Excel.XlDVType.xlValidateDate, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, start, end);
            break;
        }
        case "formula":
        {
            string formula = validation.GetProperty("formula").GetString();
            target.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, formula);
            break;
        }
        case "checkbox":
            throw new NotSupportedException("set_data_validation: 'checkbox' kind is not yet implemented pending manual verification against this machine's installed Office build - see plan Task 15 Step 2.");
        default:
            throw new ArgumentException("set_data_validation: unknown validation kind '" + kind + "'.");
    }
}
```

- [ ] **Step 2: Manually check `checkbox` support and report the finding**

In a real Excel session on this machine, check whether `Microsoft.Office.Interop.Excel.Range` exposes an `InsertCheckbox`-style member, or whether `Validation.Add` accepts a boolean-checkbox-producing type in the installed Interop version (Object Browser in Visual Studio, or trial-and-error against the actual PIA). Record the finding as a comment above the `case "checkbox":` line (either implement it if a real API exists, following the same pattern as the other 5 kinds, or leave the `NotSupportedException` in place with a comment citing what was checked and why it's unavailable on this Office build).

- [ ] **Step 3: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify `list`, `numberBetween`, and `dateBetween` produce real, working Data Validation dropdowns/restrictions in Excel's own Data > Data Validation dialog; `null` validation removes it.

- [ ] **Step 4: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add set_data_validation (list/listRef/numberBetween/dateBetween/formula)"
```

---

### Task 16: `add_pivot`, `refresh_pivot`

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new — this is the last operational task in the plan.

**Context:** the largest single COM sequence in this plan, but every piece is a well-documented native Excel pattern (this is what a VBA macro recorder produces for "Insert Pivot Table"). COM's `PivotTable` is inherently the real, live, refreshable pivot object — genuinely simpler than genoffice's own dual-path "bake computed values now, write real OOXML pivot parts at save" implementation.

- [ ] **Step 1: Implement `add_pivot`**

Add `case "add_pivot": AddPivot(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;` and:
```csharp
private static Excel.XlConsolidationFunction MapPivotAgg(string agg)
{
    switch (agg)
    {
        case "count": return Excel.XlConsolidationFunction.xlCount;
        case "average": return Excel.XlConsolidationFunction.xlAverage;
        case "max": return Excel.XlConsolidationFunction.xlMax;
        case "min": return Excel.XlConsolidationFunction.xlMin;
        default: return Excel.XlConsolidationFunction.xlSum;
    }
}

private static void AddPivot(JsonElement op)
{
    string sourceRange = op.GetProperty("sourceRange").GetString();
    string targetCell = op.GetProperty("targetCell").GetString();
    Excel.Worksheet sourceSheet = Sheet(op);
    Excel.Worksheet targetSheet = op.TryGetProperty("targetSheetId", out var tsid) && tsid.ValueKind == JsonValueKind.String
        ? (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[tsid.GetString()]
        : sourceSheet;

    Excel.PivotCache cache = Globals.ThisAddIn.Application.ActiveWorkbook.PivotCaches().Create(
        Excel.XlPivotTableSourceType.xlDatabase, sourceSheet.Range[sourceRange]);
    Excel.PivotTable pivot = cache.CreatePivotTable(targetSheet.Range[targetCell],
        op.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String ? name.GetString() : "PivotTable1");

    void AddField(string fieldName, Excel.XlPivotFieldOrientation orientation)
    {
        Excel.PivotField field = (Excel.PivotField)pivot.PivotFields(fieldName);
        field.Orientation = orientation;
    }

    if (op.TryGetProperty("rowFields", out var rowFields))
    {
        if (rowFields.ValueKind == JsonValueKind.Array)
            foreach (JsonElement f in rowFields.EnumerateArray()) AddField(f.GetString(), Excel.XlPivotFieldOrientation.xlRowField);
        else
            AddField(rowFields.GetString(), Excel.XlPivotFieldOrientation.xlRowField);
    }
    if (op.TryGetProperty("columnField", out var colField) && colField.ValueKind == JsonValueKind.String)
    {
        AddField(colField.GetString(), Excel.XlPivotFieldOrientation.xlColumnField);
    }
    if (op.TryGetProperty("pageFields", out var pageFields) && pageFields.ValueKind == JsonValueKind.Array)
    {
        foreach (JsonElement f in pageFields.EnumerateArray()) AddField(f.GetString(), Excel.XlPivotFieldOrientation.xlPageField);
    }
    foreach (JsonElement v in op.GetProperty("values").EnumerateArray())
    {
        string fieldName = v.GetProperty("field").GetString();
        string agg = v.TryGetProperty("agg", out var a) ? a.GetString() : "sum";
        if (v.TryGetProperty("formula", out var formula) && formula.ValueKind == JsonValueKind.String)
        {
            pivot.CalculatedFields().Add(fieldName, formula.GetString());
        }
        Excel.PivotField dataField = (Excel.PivotField)pivot.PivotFields(fieldName);
        dataField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
        dataField.Function = MapPivotAgg(agg);
        if (v.TryGetProperty("numFmt", out var numFmt) && numFmt.ValueKind == JsonValueKind.String)
        {
            dataField.NumberFormat = numFmt.GetString();
        }
    }
}

private static void RefreshPivot(JsonElement op)
{
    foreach (Excel.PivotTable pivot in Sheet(op).PivotTables())
    {
        pivot.PivotCache().Refresh();
    }
}
```
(Date/range groupings and label/value pivot filters from the full genoffice schema are deliberately NOT included in this first pass — `PivotField.LabelRange`+`.Group` and `.PivotFilters.Add` are real, documented APIs but add meaningful additional surface; the core "build a working pivot with rows/columns/pages/aggregated values" path above already delivers a genuinely native, refreshable pivot table, which is the highest-value 80% of this tool. Extend in a follow-up task if grouping/filtering is specifically needed.)

Wire `refresh_pivot`:
```csharp
                        case "refresh_pivot": RefreshPivot(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```

- [ ] **Step 2: Build and manually verify**

Run the MSBuild command. Expected: 0 errors.

Manually verify: `add_pivot` with a real data range (headers + a few data rows), one row field, one value field (`{"field":"Revenue","agg":"sum"}`) produces a genuine, native Excel PivotTable (PivotTable Analyze/Design ribbon tabs appear when a cell inside is selected, and it behaves identically to one built via Insert > PivotTable — drag a field in the PivotTable Fields pane and confirm it updates live). `refresh_pivot` after manually editing a source cell updates the pivot's aggregated values.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/ExcelTools.cs
git commit -m "feat(excel): add add_pivot (rows/columns/pages/values/calculated fields) and refresh_pivot"
```
