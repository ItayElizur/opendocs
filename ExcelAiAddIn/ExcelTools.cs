using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static class ExcelTools
    {
        // Task 11 (per-document since PP-1): see WordTools.cs's identical
        // pattern for the rationale - keyed by TaskPaneHost.GetChatId().
        private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

        public static void SetMode(string docKey, EditingMode mode)
        {
            ModeByDoc[docKey] = mode;
        }

        private static EditingMode ModeFor(string docKey)
        {
            EditingMode m;
            return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
        }

        private static readonly string[] ExcelErrorTexts = { "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!" };

        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "get_workbook_context", "read_range", "read_cells", "select_range", "read_formats", "read_sheet_features", "find_cells", "trace_precedents", "trace_dependents",
        };

        // Shape-name lookup now lives in OfficeAi.Shared.ShapeTypes (Phase 0) -
        // union of this map and PowerPoint's near-identical copy. PP-16:
        // mirrors EXCEL_SHAPE_TYPES / add_shape's shapeType enum in
        // ExcelAiAddIn/web-src/entry.ts exactly, plus the separately-handled
        // "textbox". Edit both together.

        // PP-15: the single chart-type vocabulary source for BOTH add_chart and
        // edit_chart (AddChart previously ignored this map entirely). Cross-
        // referenced against PowerPointAiAddIn/PowerPointTools.cs's
        // PptChartTypeMap and any Word chart map - those must use the same
        // names for the same xlChartType codes. Note PptChartTypeMap's "bar"
        // was independently wrong (51/xlColumnClustered instead of 57/
        // xlBarClustered) - fixed on that side (PP-21/22), not by changing
        // Excel's (already-correct) codes here.
        private static readonly Dictionary<string, int> ExcelChartTypeMap = new Dictionary<string, int>
        {
            ["column"] = 51,        // xlColumnClustered
            ["columnStacked"] = 52, // xlColumnStacked
            ["bar"] = 57,           // xlBarClustered
            ["barStacked"] = 58,    // xlBarStacked
            ["line"] = 4,           // xlLine
            ["area"] = 1,           // xlArea
            ["pie"] = 5,            // xlPie
            ["doughnut"] = -4120,   // xlDoughnut
        };

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            try
            {
                EditingMode mode = ModeFor(docKey);
                // Excel has no add_comment-equivalent tool yet, so Comment Only
                // mode allows no mutating tools at all (documented gap - see
                // Task 16 brief). Track Changes mode currently behaves the
                // same as Full Autonomy for gating purposes: Excel's
                // track-changes equivalent (Workbook.HighlightChangesOnScreen /
                // shared-workbook change tracking) is more limited than
                // Word's TrackRevisions and is out of scope for this task, so
                // there is deliberately no COM call wired up for it here.
                bool isMutating = !AlwaysAllowedTools.Contains(name);
                if (mode == EditingMode.ReadOnly && isMutating)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
                }
                if (mode == EditingMode.CommentOnly && isMutating)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Comment Only.", IsError = true, Summary = name };
                }

                switch (name)
                {
                    case "get_workbook_context": return GetWorkbookContext();
                    case "read_range": return ReadRange(input);
                    case "read_cells": return ReadCells(input);
                    case "select_range": return SelectRange(input);
                    case "read_formats": return ReadFormats(input);
                    case "read_sheet_features": return ReadSheetFeatures(input);
                    case "find_cells": return FindCells(input);
                    case "trace_precedents": return TracePrecedents(input);
                    case "trace_dependents": return TraceDependents(input);
                    case "propose_operations": return ProposeOperations(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        // Known limitation (PP-1 Task 5 Step 5): resolves the ACTIVE workbook/
        // sheet right now, not necessarily the one whose pane initiated this
        // tool call - see WordTools.cs's ActiveDoc for the identical
        // rationale and the same out-of-scope decision.
        private static Excel.Worksheet Sheet(JsonElement input)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            if (input.ValueKind == JsonValueKind.Object && input.TryGetProperty("sheet", out var s) && s.ValueKind == JsonValueKind.String)
            {
                return (Excel.Worksheet)app.ActiveWorkbook.Sheets[s.GetString()];
            }
            return (Excel.Worksheet)app.ActiveSheet;
        }

        private static ToolResult GetWorkbookContext()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string usedRange = sheet.UsedRange.Address[false, false];
            string selection = ((Excel.Range)Globals.ThisAddIn.Application.Selection).Address[false, false];

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Active sheet: {sheet.Name}");
            sb.AppendLine($"UsedRange: {usedRange}");
            sb.AppendLine($"Selection: {selection}");
            sb.AppendLine("Sheets:");
            // Mirrors get_deck_context's per-slide preview pattern (one line
            // per sheet) so the model gets the whole workbook's shape in one
            // call instead of needing a round-trip per sheet.
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                sb.AppendLine($"- {s.Name}: {s.UsedRange.Address[false, false]}");
            }

            return new ToolResult { Output = sb.ToString(), Summary = "get_workbook_context" };
        }

        private static ToolResult ReadRange(JsonElement input)
        {
            string address = input.GetProperty("address").GetString();
            Excel.Range range = Sheet(input).Range[address];
            if (range.Cells.Count > 2000)
            {
                return new ToolResult { Output = "Range exceeds 2000-cell cap.", IsError = true, Summary = "read_range" };
            }
            object[,] values = range.Value2 as object[,];
            var sb = new System.Text.StringBuilder();
            if (values == null)
            {
                sb.Append(range.Value2 ?? "");
            }
            else
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    var cells = Enumerable.Range(1, values.GetLength(1)).Select(c => values[r, c]?.ToString() ?? "");
                    sb.AppendLine(string.Join("\t", cells));
                }
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_range" };
        }

        private static ToolResult ReadCells(JsonElement input)
        {
            Excel.Worksheet sheet = Sheet(input);
            var sb = new System.Text.StringBuilder();
            foreach (JsonElement addr in input.GetProperty("addresses").EnumerateArray())
            {
                string a = addr.GetString();
                object value = ((Excel.Range)sheet.Range[a]).Value2;
                sb.AppendLine($"{a}: {value}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_cells" };
        }

        private static ToolResult SelectRange(JsonElement input)
        {
            string address = input.GetProperty("address").GetString();
            Excel.Worksheet sheet = Sheet(input);
            sheet.Activate();
            Excel.Range range = sheet.Range[address];
            range.Select();
            return new ToolResult { Output = "Selected " + address + " on " + sheet.Name + ".", Summary = "select_range" };
        }

        // PP-13 Task 3: measured ~1.1s for a 200-cell read with the widened
        // per-cell property set (12 COM reads/cell vs. the previous 4) on this
        // dev machine - comfortably under the ~2s budget the plan flagged, so
        // the 200-cell cap is kept as-is rather than lowered or split into a
        // properties?:string[] filter.
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
                object underlineRaw = cell.Font.Underline;
                int underlineCode = underlineRaw != null ? Convert.ToInt32(underlineRaw) : -4142;
                bool underline = underlineCode != -4142; // -4142 == xlUnderlineStyleNone
                string numberFormat = cell.NumberFormat as string;
                bool strikethrough = (bool)(cell.Font.Strikethrough ?? false);
                string fontName = cell.Font.Name as string;
                double fontSize = cell.Font.Size is double sz ? sz : 0;
                Excel.XlHAlign hAlign = (Excel.XlHAlign)(int)cell.HorizontalAlignment;
                Excel.XlVAlign vAlign = (Excel.XlVAlign)(int)cell.VerticalAlignment;
                bool wrapText = (bool)(cell.WrapText ?? false);
                int rotation = cell.Orientation is int rot ? rot : 0;
                int indent = cell.IndentLevel is int ind ? ind : 0;
                bool hasBorder = false;
                foreach (Excel.XlBordersIndex idx in BorderEdgeMap.Values)
                {
                    if (cell.Borders[idx].LineStyle != Excel.XlLineStyle.xlLineStyleNone) { hasBorder = true; break; }
                }

                // Widened to match everything format_range can now set (PP-13) -
                // a cell that is only e.g. centered, or only bordered, must not
                // be filtered out as "unformatted", or a read-modify-write cycle
                // silently drops that property.
                bool hasDefaultFormat = !bold && !italic && !underline && !strikethrough
                    && (numberFormat == "General" || numberFormat == null)
                    && hAlign == Excel.XlHAlign.xlHAlignGeneral && vAlign == Excel.XlVAlign.xlVAlignBottom
                    && !wrapText && rotation == 0 && indent == 0 && !hasBorder;
                if (hasDefaultFormat) continue; // only explicitly-formatted cells, matches genoffice

                sb.AppendLine($"{cell.Address[false, false]}: bold={bold}, italic={italic}, underline={underline}, " +
                    $"strikethrough={strikethrough}, fontName={fontName}, fontSize={fontSize}, numberFormat={numberFormat}, " +
                    $"horizontalAlignment={HAlignName(hAlign)}, verticalAlignment={VAlignName(vAlign)}, wrapText={wrapText}, " +
                    $"textRotation={rotation}, indent={indent}, hasBorder={hasBorder}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_formats" };
        }

        private static string HAlignName(Excel.XlHAlign v)
        {
            foreach (var kv in HAlignMap) if (kv.Value == v) return kv.Key;
            return v.ToString();
        }

        private static string VAlignName(Excel.XlVAlign v)
        {
            foreach (var kv in VAlignMap) if (kv.Value == v) return kv.Key;
            return v.ToString();
        }

        // Shared by find_cells and propose_operations' find_replace: which
        // sheets a call should touch. sheetId names one specific sheet;
        // otherwise allSheets picks every sheet in the workbook (Ctrl+F's
        // "Within: Workbook") or, by default, just the active sheet (Ctrl+F's
        // default "Within: Sheet") - a narrower default than this tool used
        // to have (previously: omitting sheetId meant the WHOLE workbook).
        private static List<Excel.Worksheet> ResolveSheetsToSearch(Excel.Workbook wb, string sheetId, bool allSheets)
        {
            var sheets = new List<Excel.Worksheet>();
            if (sheetId != null)
            {
                sheets.Add((Excel.Worksheet)wb.Sheets[sheetId]);
            }
            else if (allSheets)
            {
                foreach (Excel.Worksheet s in wb.Worksheets) sheets.Add(s);
            }
            else
            {
                sheets.Add((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            }
            return sheets;
        }

        // Excel's own native Find/FindNext - the same engine behind Ctrl+F -
        // over one sheet for one LookIn mode, instead of reading .Text on
        // every single cell in the range regardless of match count. Wraps
        // around like Ctrl+F itself; detected and stopped via the address of
        // the first hit rather than scanning the whole range unconditionally.
        private static int NativeFindInSheet(Excel.Worksheet sheet, string query, Excel.XlFindLookIn lookIn, int limit,
            HashSet<string> seenAddresses, System.Text.StringBuilder sb)
        {
            if (limit <= 0) return 0;
            Excel.Range usedRange = sheet.UsedRange;
            Excel.Range current = usedRange.Find(
                What: query, LookIn: lookIn, LookAt: Excel.XlLookAt.xlPart,
                SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext,
                MatchCase: false);

            Excel.Range first = null;
            int count = 0;
            while (current != null)
            {
                string cellAddress = current.Address[false, false];
                if (first == null) first = current;
                else if (cellAddress == first.Address[false, false]) break; // wrapped back to the first hit

                string key = sheet.Name + "!" + cellAddress;
                if (seenAddresses.Add(key))
                {
                    sb.AppendLine($"{key}: {current.Text}");
                    count++;
                    if (count >= limit) break;
                }
                current = usedRange.FindNext(current);
            }
            return count;
        }

        private static ToolResult FindCells(JsonElement input)
        {
            bool errorsOnly = input.TryGetProperty("errors_only", out var eo) && eo.ValueKind == JsonValueKind.True;
            string query = input.TryGetProperty("query", out var q) && q.ValueKind == JsonValueKind.String ? q.GetString() : null;
            bool useRegex = input.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
            string lookIn = input.TryGetProperty("look_in", out var li) && li.ValueKind == JsonValueKind.String ? li.GetString() : "both";
            int maxResults = input.GetProperty("max_results").GetInt32();
            string sheetId = input.TryGetProperty("sheetId", out var sid) && sid.ValueKind == JsonValueKind.String ? sid.GetString() : null;
            bool allSheets = input.TryGetProperty("allSheets", out var allEl) && allEl.ValueKind == JsonValueKind.True;

            if (!errorsOnly && query == null)
            {
                return new ToolResult { Output = "find_cells requires either 'query' or 'errors_only'.", IsError = true, Summary = "find_cells" };
            }

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<Excel.Worksheet> sheets = ResolveSheetsToSearch(wb, sheetId, allSheets);
            var sb = new System.Text.StringBuilder();
            int found = 0;

            if (errorsOnly)
            {
                foreach (Excel.Worksheet sheet in sheets)
                {
                    if (found >= maxResults) break;
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
                }
                return new ToolResult { Output = sb.ToString(), Summary = "find_cells" };
            }

            if (useRegex)
            {
                var regex = new System.Text.RegularExpressions.Regex(query);
                foreach (Excel.Worksheet sheet in sheets)
                {
                    if (found >= maxResults) break;
                    foreach (Excel.Range cell in sheet.UsedRange.Cells)
                    {
                        if (found >= maxResults) break;
                        string valueText = cell.Text as string ?? "";
                        string formulaText = cell.Formula as string ?? "";
                        string haystack = lookIn == "values" ? valueText : lookIn == "formulas" ? formulaText : valueText + " " + formulaText;
                        if (regex.IsMatch(haystack))
                        {
                            sb.AppendLine($"{sheet.Name}!{cell.Address[false, false]}: {valueText}");
                            found++;
                        }
                    }
                }
                return new ToolResult { Output = sb.ToString(), Summary = "find_cells" };
            }

            // Plain substring - Excel's own native Find/FindNext, not a
            // per-cell scan. xlValues/xlFormulas are separate native passes
            // (Excel's Find only searches one LookIn mode per call); "both"
            // runs both and de-dupes by address so a cell matching in either
            // is reported once, matching the old per-cell "both" semantics.
            var seenAddresses = new HashSet<string>();
            foreach (Excel.Worksheet sheet in sheets)
            {
                if (found >= maxResults) break;
                if (lookIn == "values" || lookIn == "both")
                    found += NativeFindInSheet(sheet, query, Excel.XlFindLookIn.xlValues, maxResults - found, seenAddresses, sb);
                if (found < maxResults && (lookIn == "formulas" || lookIn == "both"))
                    found += NativeFindInSheet(sheet, query, Excel.XlFindLookIn.xlFormulas, maxResults - found, seenAddresses, sb);
            }
            return new ToolResult { Output = sb.ToString(), Summary = "find_cells" };
        }

        // The write-side counterpart of find_cells (shares its sheetId/
        // allSheets scoping - active sheet only by default, matching Ctrl+H's
        // "Within: Sheet") - only replaces within literal cell VALUES
        // (cell.Value2 is a string), never formulas or numbers, so a formula
        // is never corrupted by a text substitution.
        private static int FindReplaceExcel(JsonElement op)
        {
            string find = op.GetProperty("find").GetString();
            string replace = op.GetProperty("replace").GetString();
            bool useRegex = op.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
            bool matchCase = op.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            string sheetId = op.TryGetProperty("sheetId", out var sid) && sid.ValueKind == JsonValueKind.String ? sid.GetString() : null;
            bool allSheets = op.TryGetProperty("allSheets", out var allEl) && allEl.ValueKind == JsonValueKind.True;

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<Excel.Worksheet> sheets = ResolveSheetsToSearch(wb, sheetId, allSheets);
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int replaced = 0;

            if (useRegex)
            {
                var regex = new System.Text.RegularExpressions.Regex(find, matchCase
                    ? System.Text.RegularExpressions.RegexOptions.None
                    : System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                foreach (Excel.Worksheet sheet in sheets)
                {
                    foreach (Excel.Range cell in sheet.UsedRange.Cells)
                    {
                        if (!(cell.Value2 is string text)) continue; // numbers/dates/formulas/blank - not a text replace target
                        if (!regex.IsMatch(text)) continue;
                        cell.Value2 = regex.Replace(text, replace);
                        replaced++;
                    }
                }
                return replaced;
            }

            foreach (Excel.Worksheet sheet in sheets)
            {
                replaced += NativeFindReplaceInSheet(sheet, find, replace, matchCase, comparison);
            }
            return replaced;
        }

        // Locates matches via Excel's own native Find/FindNext (the engine
        // behind Ctrl+F/Ctrl+H) instead of reading .Text on every cell in the
        // range, then replaces only the matched cell's literal text VALUE
        // directly. current.Value2 is only ever a string for a literal text
        // cell, so a numeric/date/formula cell that merely DISPLAYS a match
        // (native Find matched its formatted text) is correctly skipped
        // here, same safety scope as before.
        private static int NativeFindReplaceInSheet(Excel.Worksheet sheet, string find, string replace, bool matchCase, StringComparison comparison)
        {
            Excel.Range usedRange = sheet.UsedRange;
            Excel.Range current = usedRange.Find(
                What: find, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart,
                SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext,
                MatchCase: matchCase);

            Excel.Range first = null;
            int replaced = 0;
            while (current != null)
            {
                string address = current.Address[false, false];
                if (first == null) first = current;
                else if (address == first.Address[false, false]) break; // wrapped back to the first hit

                // Captured before mutating current's cell, so FindNext's own
                // position tracking is never asked to reason about a cell
                // whose content just changed underneath it.
                Excel.Range next = usedRange.FindNext(current);

                if (current.Value2 is string text && text.IndexOf(find, comparison) >= 0)
                {
                    current.Value2 = TextUtil.ReplaceAllOccurrences(text, find, replace, comparison);
                    replaced++;
                }

                current = next;
            }
            return replaced;
        }

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

            sb.AppendLine("Detected tables (contiguous data blocks):");
            List<string> tables = FindDataTables(sheet);
            if (tables.Count == 0)
            {
                sb.AppendLine("(none - sheet is empty)");
            }
            else
            {
                foreach (string region in tables) sb.AppendLine("- " + region);
            }

            return new ToolResult { Output = sb.ToString(), Summary = "read_sheet_features" };
        }

        // Splits the sheet's used range into separate contiguous data blocks -
        // BFS flood-fill over non-empty cells, 4-directional, no gap tolerance
        // (a single fully-blank row or column separates two tables). Ported
        // from nlputils' ExcelTableParser.find_data_tables, simplified for COM:
        // reads Value2 in one bulk call instead of per-cell (a per-cell COM
        // round-trip would make flood-fill prohibitively slow), and doesn't
        // special-case merged cells - Value2 already reports a merge's
        // non-anchor cells as empty, which is enough for the common case
        // (merge's text lives in its top-left cell) but means a merge wider
        // than the flood fill's reach can still fragment a table. A sheet with
        // scattered blank cells inside one logical table will similarly
        // over-fragment into several regions - acceptable for now since the
        // goal is finding stacked/side-by-side tables, not validating layout.
        private static List<string> FindDataTables(Excel.Worksheet sheet)
        {
            var regions = new List<string>();
            Excel.Range used = sheet.UsedRange;
            int rowCount = used.Rows.Count;
            int colCount = used.Columns.Count;
            if (rowCount == 0 || colCount == 0) return regions;

            object valueObj = used.Value2;
            bool[,] hasContent = new bool[rowCount, colCount];
            if (rowCount == 1 && colCount == 1)
            {
                hasContent[0, 0] = valueObj != null;
            }
            else
            {
                object[,] values = (object[,])valueObj;
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        hasContent[r - 1, c - 1] = values[r, c] != null;
                    }
                }
            }

            int startRow = used.Row;
            int startCol = used.Column;
            bool[,] visited = new bool[rowCount, colCount];
            int[] dr = { 0, 0, 1, -1 };
            int[] dc = { 1, -1, 0, 0 };

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    if (!hasContent[r, c] || visited[r, c]) continue;

                    int minR = r, maxR = r, minC = c, maxC = c;
                    var queue = new Queue<KeyValuePair<int, int>>();
                    queue.Enqueue(new KeyValuePair<int, int>(r, c));
                    visited[r, c] = true;
                    while (queue.Count > 0)
                    {
                        KeyValuePair<int, int> cur = queue.Dequeue();
                        int cr = cur.Key, cc = cur.Value;
                        minR = Math.Min(minR, cr); maxR = Math.Max(maxR, cr);
                        minC = Math.Min(minC, cc); maxC = Math.Max(maxC, cc);
                        for (int d = 0; d < 4; d++)
                        {
                            int nr = cr + dr[d], nc = cc + dc[d];
                            if (nr < 0 || nr >= rowCount || nc < 0 || nc >= colCount) continue;
                            if (visited[nr, nc] || !hasContent[nr, nc]) continue;
                            visited[nr, nc] = true;
                            queue.Enqueue(new KeyValuePair<int, int>(nr, nc));
                        }
                    }

                    Excel.Range topLeft = (Excel.Range)sheet.Cells[startRow + minR, startCol + minC];
                    Excel.Range bottomRight = (Excel.Range)sheet.Cells[startRow + maxR, startCol + maxC];
                    string address = sheet.Range[topLeft, bottomRight].Address[false, false];
                    int rows = maxR - minR + 1;
                    int cols = maxC - minC + 1;
                    regions.Add($"{address} ({rows}x{cols})");
                }
            }

            return regions;
        }

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

        private static Excel.XlFormatConditionOperator MapCfOperator(string op)
        {
            switch (op)
            {
                case "greaterThan": return Excel.XlFormatConditionOperator.xlGreater;
                case "lessThan": return Excel.XlFormatConditionOperator.xlLess;
                case "equal": return Excel.XlFormatConditionOperator.xlEqual;
                case "notEqual": return Excel.XlFormatConditionOperator.xlNotEqual;
                case "greaterEqual": return Excel.XlFormatConditionOperator.xlGreaterEqual;
                case "lessEqual": return Excel.XlFormatConditionOperator.xlLessEqual;
                case "between": return Excel.XlFormatConditionOperator.xlBetween;
                case "notBetween": return Excel.XlFormatConditionOperator.xlNotBetween;
                default:
                    throw new ArgumentException("add_conditional_format: unknown operator '" + op +
                        "'. Valid: greaterThan, lessThan, equal, notEqual, greaterEqual, lessEqual, between, notBetween.");
            }
        }

        private static double GetCfNumber(JsonElement rule, string field)
        {
            JsonElement v = rule.GetProperty(field);
            if (v.ValueKind == JsonValueKind.Number) return v.GetDouble();
            if (v.ValueKind == JsonValueKind.String)
            {
                double parsed;
                if (double.TryParse(v.GetString(), out parsed)) return parsed;
            }
            throw new ArgumentException("add_conditional_format: '" + field + "' must be a number (or a numeric string).");
        }

        private static Excel.XlContainsOperator MapCfTextMatch(string match)
        {
            switch (match)
            {
                case "contains": return Excel.XlContainsOperator.xlContains;
                case "notContains": return Excel.XlContainsOperator.xlDoesNotContain;
                case "beginsWith": return Excel.XlContainsOperator.xlBeginsWith;
                case "endsWith": return Excel.XlContainsOperator.xlEndsWith;
                default:
                    throw new ArgumentException("add_conditional_format: unknown text match '" + match +
                        "'. Valid: contains, notContains, beginsWith, endsWith.");
            }
        }

        // Returns a short description of what was created, for the batch result
        // line (PP-14 Task 5 Step 4).
        private static string AddConditionalFormat(JsonElement op)
        {
            JsonElement rangeEl;
            if (!op.TryGetProperty("range", out rangeEl) || rangeEl.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_conditional_format: missing required field \"range\".");
            string range = rangeEl.GetString();
            Excel.Range target;
            try
            {
                target = Sheet(op).Range[range];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new ArgumentException("add_conditional_format: '" + range + "' is not a valid range address.");
            }

            JsonElement rule;
            if (!op.TryGetProperty("rule", out rule) || rule.ValueKind != JsonValueKind.Object)
                throw new ArgumentException("add_conditional_format: missing required field \"rule\".");
            JsonElement kindEl;
            if (!rule.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_conditional_format: rule is missing a string \"kind\" field.");
            string kind = kindEl.GetString();
            Excel.FormatCondition fc = null;
            string detail = "";

            switch (kind)
            {
                case "number":
                {
                    string oper = rule.GetProperty("operator").GetString();
                    Excel.XlFormatConditionOperator mappedOp = MapCfOperator(oper); // throws on unknown
                    double value = GetCfNumber(rule, "value");
                    bool needsSecond = mappedOp == Excel.XlFormatConditionOperator.xlBetween || mappedOp == Excel.XlFormatConditionOperator.xlNotBetween;
                    string formula2 = null;
                    if (needsSecond)
                    {
                        if (!rule.TryGetProperty("value2", out _))
                            throw new ArgumentException("add_conditional_format: 'value2' is required when operator is '" + oper + "'.");
                        formula2 = GetCfNumber(rule, "value2").ToString();
                    }
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, mappedOp, value.ToString(), formula2);
                    detail = "number " + oper + " " + value + (formula2 != null ? ".." + formula2 : "");
                    break;
                }
                case "text":
                {
                    string text = rule.GetProperty("text").GetString();
                    string matchName = rule.TryGetProperty("match", out var matchEl) && matchEl.ValueKind == JsonValueKind.String ? matchEl.GetString() : "contains";
                    Excel.XlContainsOperator matchOp = MapCfTextMatch(matchName); // throws on unknown
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlTextString, String: text, TextOperator: matchOp);
                    detail = "text " + matchName + " \"" + text + "\"";
                    break;
                }
                case "blank":
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition);
                    detail = "blank";
                    break;
                case "duplicate":
                {
                    string modeName = rule.TryGetProperty("mode", out var modeEl) && modeEl.ValueKind == JsonValueKind.String ? modeEl.GetString() : "duplicate";
                    Excel.XlDupeUnique dupeUnique;
                    if (modeName == "duplicate") dupeUnique = Excel.XlDupeUnique.xlDuplicate;
                    else if (modeName == "unique") dupeUnique = Excel.XlDupeUnique.xlUnique;
                    else throw new ArgumentException("add_conditional_format: unknown mode '" + modeName + "'. Valid: duplicate, unique.");
                    fc = target.FormatConditions.AddUniqueValues();
                    ((Excel.UniqueValues)fc).DupeUnique = dupeUnique;
                    detail = modeName;
                    break;
                }
                case "top10":
                {
                    int rank = rule.TryGetProperty("rank", out var r) ? r.GetInt32() : 10;
                    bool percent = rule.TryGetProperty("percent", out var p) && p.ValueKind == JsonValueKind.True;
                    bool bottom = rule.TryGetProperty("bottom", out var b) && b.ValueKind == JsonValueKind.True;
                    Excel.Top10 top10 = target.FormatConditions.AddTop10();
                    top10.Rank = rank;
                    top10.Percent = percent;
                    top10.TopBottom = bottom ? Excel.XlTopBottom.xlTop10Bottom : Excel.XlTopBottom.xlTop10Top;
                    if (rule.TryGetProperty("format", out var top10Format))
                    {
                        if (top10Format.TryGetProperty("bold", out var bold)) top10.Font.Bold = bold.ValueKind == JsonValueKind.True;
                        if (top10Format.TryGetProperty("fontColor", out var fontColor)) top10.Font.Color = ColorUtil.HexToOle(fontColor.GetString());
                        if (top10Format.TryGetProperty("fillColor", out var fillColor)) top10.Interior.Color = ColorUtil.HexToOle(fillColor.GetString());
                    }
                    return "top10 range=" + range + " kind=top10 rank=" + rank + (percent ? "%" : "") + (bottom ? " bottom" : " top");
                    // Top10 doesn't implement FormatCondition in this PIA (confirmed via reflection) - format applied directly above, mirroring colorScale/dataBar's early-return pattern
                }
                case "formula":
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: rule.GetProperty("formula").GetString());
                    detail = "formula";
                    break;
                case "colorScale":
                {
                    Excel.ColorScale scale = target.FormatConditions.AddColorScale(3);
                    if (rule.TryGetProperty("minColor", out var minC)) scale.ColorScaleCriteria[1].FormatColor.Color = ColorUtil.HexToOle(minC.GetString());
                    if (rule.TryGetProperty("midColor", out var midC)) scale.ColorScaleCriteria[2].FormatColor.Color = ColorUtil.HexToOle(midC.GetString());
                    if (rule.TryGetProperty("maxColor", out var maxC)) scale.ColorScaleCriteria[3].FormatColor.Color = ColorUtil.HexToOle(maxC.GetString());
                    return "colorScale range=" + range; // ColorScale/DataBar carry their own visual - no separate "format" object to apply below
                }
                case "dataBar":
                {
                    Excel.Databar bar = target.FormatConditions.AddDatabar();
                    if (rule.TryGetProperty("color", out var barColor))
                    {
                        bar.BarColor.Color = ColorUtil.HexToOle(barColor.GetString());
                    }
                    return "dataBar range=" + range;
                }
                default:
                    throw new ArgumentException("add_conditional_format: unknown rule kind '" + kind +
                        "'. Valid: number, text, blank, duplicate, top10, formula, colorScale, dataBar.");
            }

            if (fc != null && rule.TryGetProperty("format", out var format))
            {
                if (format.TryGetProperty("bold", out var bold)) fc.Font.Bold = bold.ValueKind == JsonValueKind.True;
                if (format.TryGetProperty("fontColor", out var fontColor)) fc.Font.Color = ColorUtil.HexToOle(fontColor.GetString());
                if (format.TryGetProperty("fillColor", out var fillColor)) fc.Interior.Color = ColorUtil.HexToOle(fillColor.GetString());
            }
            return kind + " range=" + range + " (" + detail + ")";
        }

        // PP-5: mirrors EXCEL_OPS's `required` arrays in
        // ExcelAiAddIn/web-src/entry.ts exactly (minus "kind" itself, which is
        // validated separately in ProposeOperations before this table is
        // consulted) - the two must be edited together. Kinds needing no
        // field beyond "kind" (set_page_setup, delete_sheet, duplicate_sheet,
        // refresh_pivot, clear_filter, clear_conditional_formats) have no
        // entry here, matching WordTools.cs's RequiredFields convention. This
        // is the actual guarantee: the TS schema is documentation the model
        // reads, not a validator that runs (not every provider enforces
        // oneOf/const, and Excel's grouped-variant collapsed branch carries
        // no per-kind structure at all) - this precheck is what turns a
        // missing field into a specific, per-operation error instead of a
        // raw COM/NullReference exception.
        private static readonly Dictionary<string, string[]> RequiredFields = new Dictionary<string, string[]>
        {
            ["set_cell"] = new[] { "address", "value" },
            ["set_formula"] = new[] { "address", "formula" },
            ["set_range"] = new[] { "address", "values" },
            ["clear_cell"] = new[] { "address" },
            ["clear_range"] = new[] { "range" },
            ["find_replace"] = new[] { "find", "replace" },
            ["format_range"] = new[] { "address" },
            ["sort_range"] = new[] { "range", "byColumn", "order" },
            ["merge_cells"] = new[] { "range" },
            ["unmerge_cells"] = new[] { "range" },
            ["set_row_height"] = new[] { "row", "heightPoints" },
            ["set_col_width"] = new[] { "column", "widthPx" },
            ["set_rows_hidden"] = new[] { "row", "hidden" },
            ["set_cols_hidden"] = new[] { "column", "hidden" },
            ["set_freeze"] = new[] { "rows", "columns" },
            ["insert_rows"] = new[] { "startRow", "count" },
            ["delete_rows"] = new[] { "startRow", "count" },
            ["insert_cols"] = new[] { "startCol", "count" },
            ["delete_cols"] = new[] { "startCol", "count" },
            ["add_sheet"] = new[] { "name" },
            ["set_sheet_hidden"] = new[] { "hidden" },
            ["move_sheet"] = new[] { "position" },
            ["protect_sheet"] = new[] { "protected" },
            ["rename_sheet"] = new[] { "name" },
            ["add_chart"] = new[] { "dataRange" },
            ["edit_chart"] = new[] { "chartPath" },
            ["delete_visual"] = new[] { "visualId" },
            ["add_sparkline"] = new[] { "dataRange", "targetCell" },
            ["add_shape"] = new[] { "shapeType", "anchorCell" },
            ["edit_shape"] = new[] { "visualId" },
            ["add_image"] = new[] { "path", "anchorCell" },
            ["add_table"] = new[] { "range" },
            ["add_table_row"] = new[] { "tableName" },
            ["add_table_column"] = new[] { "tableName", "columnName" },
            ["delete_table_row"] = new[] { "tableName", "row" },
            ["delete_table_column"] = new[] { "tableName", "column" },
            ["delete_table"] = new[] { "tableName" },
            ["add_pivot"] = new[] { "sourceRange", "targetCell", "values" },
            ["set_hyperlink"] = new[] { "address" },
            ["set_note"] = new[] { "address" },
            ["add_defined_name"] = new[] { "name", "ref" },
            ["delete_defined_name"] = new[] { "name" },
            ["set_filter"] = new[] { "range" },
            ["set_filter_criteria"] = new[] { "column" },
            ["add_conditional_format"] = new[] { "range", "rule" },
            ["set_data_validation"] = new[] { "range" },
        };

        private static ToolResult ProposeOperations(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            bool anyError = false;
            foreach (JsonElement op in input.GetProperty("operations").EnumerateArray())
            {
                string kind = null;
                try
                {
                    JsonElement kindEl;
                    if (!op.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("Operation is missing a string \"kind\" field.");
                    kind = kindEl.GetString();
                    ToolArgs.ValidateRequired(kind, op, RequiredFields, "Operation");
                    switch (kind)
                    {
                        case "set_cell":
                            Sheet(op).Range[op.GetProperty("address").GetString()].Value2 = JsonValueToObject(op.GetProperty("value"));
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_formula":
                            Sheet(op).Range[op.GetProperty("address").GetString()].Formula = op.GetProperty("formula").GetString();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_range":
                            SetRangeValues(op);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "format_range":
                        {
                            string formatNote = FormatRange(op);
                            lines.AppendLine(kind + ": ok" + (formatNote != null ? " (" + formatNote + ")" : ""));
                            anyMutated = true; break;
                        }
                        case "clear_cell":
                            Sheet(op).Range[op.GetProperty("address").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_range":
                            Sheet(op).Range[op.GetProperty("range").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "find_replace":
                        {
                            int replaced = FindReplaceExcel(op);
                            lines.AppendLine(kind + ": " + replaced + " cell(s) changed");
                            if (replaced > 0) anyMutated = true;
                            break;
                        }
                        case "insert_rows": InsertDeleteRows(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_rows": InsertDeleteRows(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "insert_cols": InsertDeleteCols(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_cols": InsertDeleteCols(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_chart":
                        {
                            string chartDetail = AddChart(op);
                            lines.AppendLine(kind + ": ok (" + chartDetail + ")");
                            anyMutated = true; break;
                        }
                        case "add_sparkline":
                        {
                            string sparklineDetail = AddSparkline(op);
                            lines.AppendLine(kind + ": ok (" + sparklineDetail + ")");
                            anyMutated = true; break;
                        }
                        case "add_shape":
                        {
                            string shapeName = AddShapeExcel(op);
                            lines.AppendLine(kind + ": ok (name=" + shapeName + ")");
                            anyMutated = true; break;
                        }
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
                        case "set_freeze": SetFreeze(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_page_setup": SetPageSetup(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_sheet": AddSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_sheet": DeleteSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "duplicate_sheet": DuplicateSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_sheet_hidden": SetSheetHidden(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "move_sheet": MoveSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "protect_sheet": ProtectSheet(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "rename_sheet": Sheet(op).Name = op.GetProperty("name").GetString(); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "edit_shape": EditShapeExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "edit_chart": EditChartExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_visual": DeleteVisual(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_image": AddImageExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_table": AddTable(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_table_row": AddTableRow(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_table_column": AddTableColumn(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table_row": DeleteTableRow(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table_column": DeleteTableColumn(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_table":
                        {
                            string deleteTableDetail = DeleteTable(op);
                            lines.AppendLine(kind + ": " + deleteTableDetail);
                            anyMutated = true; break;
                        }
                        case "set_hyperlink": SetHyperlink(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_note": SetNote(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_defined_name": AddDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_defined_name": DeleteDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_filter": SetFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_filter": ClearFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_filter_criteria": SetFilterCriteria(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_conditional_format":
                        {
                            string cfDetail = AddConditionalFormat(op);
                            lines.AppendLine(kind + ": ok - " + cfDetail);
                            anyMutated = true; break;
                        }
                        case "clear_conditional_formats": Sheet(op).UsedRange.FormatConditions.Delete(); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_data_validation": SetDataValidation(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_pivot": AddPivot(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "refresh_pivot": RefreshPivot(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        default:
                            lines.AppendLine(kind + ": unknown operation kind"); anyError = true; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine((kind ?? "(unknown kind)") + ": ERROR - " + ex.Message); anyError = true;
                }
            }
            return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "propose_operations" };
        }

        private static object JsonValueToObject(JsonElement v)
        {
            switch (v.ValueKind)
            {
                case JsonValueKind.String: return v.GetString();
                case JsonValueKind.Number: return v.GetDouble();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                default: return null;
            }
        }

        private static void SetRangeValues(JsonElement op)
        {
            string address = op.GetProperty("address").GetString();
            JsonElement rows = op.GetProperty("values");
            int rowCount = rows.GetArrayLength();
            int colCount = rows[0].GetArrayLength();
            object[,] grid = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                JsonElement row = rows[r];
                for (int c = 0; c < colCount; c++) grid[r, c] = JsonValueToObject(row[c]);
            }
            Excel.Range topLeft = Sheet(op).Range[address];
            topLeft.Resize[rowCount, colCount].Value2 = grid;
        }

        // PP-13: horizontal/vertical alignment name -> XlHAlign/XlVAlign, mirroring
        // this file's existing ExcelChartTypeMap pattern (ShapeTypes moved to
        // OfficeAi.Shared in Phase 0).
        private static readonly Dictionary<string, Excel.XlHAlign> HAlignMap = new Dictionary<string, Excel.XlHAlign>
        {
            ["general"] = Excel.XlHAlign.xlHAlignGeneral,
            ["left"] = Excel.XlHAlign.xlHAlignLeft,
            ["center"] = Excel.XlHAlign.xlHAlignCenter,
            ["right"] = Excel.XlHAlign.xlHAlignRight,
            ["fill"] = Excel.XlHAlign.xlHAlignFill,
            ["justify"] = Excel.XlHAlign.xlHAlignJustify,
            ["centerAcrossSelection"] = Excel.XlHAlign.xlHAlignCenterAcrossSelection,
            ["distributed"] = Excel.XlHAlign.xlHAlignDistributed,
        };

        private static readonly Dictionary<string, Excel.XlVAlign> VAlignMap = new Dictionary<string, Excel.XlVAlign>
        {
            ["top"] = Excel.XlVAlign.xlVAlignTop,
            ["center"] = Excel.XlVAlign.xlVAlignCenter,
            ["bottom"] = Excel.XlVAlign.xlVAlignBottom,
            ["justify"] = Excel.XlVAlign.xlVAlignJustify,
            ["distributed"] = Excel.XlVAlign.xlVAlignDistributed,
        };

        // PP-13: border edge name -> XlBordersIndex, and style name -> (LineStyle, Weight).
        private static readonly Dictionary<string, Excel.XlBordersIndex> BorderEdgeMap = new Dictionary<string, Excel.XlBordersIndex>
        {
            ["left"] = Excel.XlBordersIndex.xlEdgeLeft,
            ["top"] = Excel.XlBordersIndex.xlEdgeTop,
            ["bottom"] = Excel.XlBordersIndex.xlEdgeBottom,
            ["right"] = Excel.XlBordersIndex.xlEdgeRight,
            ["insideHorizontal"] = Excel.XlBordersIndex.xlInsideHorizontal,
            ["insideVertical"] = Excel.XlBordersIndex.xlInsideVertical,
            ["diagonalDown"] = Excel.XlBordersIndex.xlDiagonalDown,
            ["diagonalUp"] = Excel.XlBordersIndex.xlDiagonalUp,
        };

        private static readonly string[] OutlineEdges = { "left", "top", "bottom", "right" };
        private static readonly string[] AllEdges = { "left", "top", "bottom", "right", "insideHorizontal", "insideVertical" };

        // Returns a note to append to the batch result line (e.g. borders skipped
        // on a single cell), or null when there is nothing to report.
        private static string FormatRange(JsonElement op)
        {
            string note = null;
            Excel.Range range = Sheet(op).Range[op.GetProperty("address").GetString()];
            if (op.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean();
            if (op.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean();
            if (op.TryGetProperty("numberFormat", out var nf)) range.NumberFormat = nf.GetString();
            if (op.TryGetProperty("fillColor", out var fc) && fc.ValueKind == JsonValueKind.String)
            {
                range.Interior.Color = ColorUtil.HexToOle(fc.GetString());
            }

            if (op.TryGetProperty("fontName", out var fn) && fn.ValueKind == JsonValueKind.String) range.Font.Name = fn.GetString();
            if (op.TryGetProperty("fontSize", out var fs) && fs.ValueKind == JsonValueKind.Number) range.Font.Size = fs.GetDouble();
            if (op.TryGetProperty("fontColor", out var fcol) && fcol.ValueKind == JsonValueKind.String) range.Font.Color = ColorUtil.HexToOle(fcol.GetString());
            if (op.TryGetProperty("strikethrough", out var st)) range.Font.Strikethrough = st.ValueKind == JsonValueKind.True;

            if (op.TryGetProperty("underline", out var underline))
            {
                if (underline.ValueKind == JsonValueKind.True || underline.ValueKind == JsonValueKind.False)
                {
                    range.Font.Underline = underline.ValueKind == JsonValueKind.True
                        ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                        : Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                }
                else
                {
                    string u = underline.GetString();
                    switch (u)
                    {
                        case "none": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone; break;
                        case "single": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle; break;
                        case "double": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleDouble; break;
                        case "singleAccounting": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting; break;
                        case "doubleAccounting": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting; break;
                        default:
                            throw new ArgumentException("format_range: unknown underline '" + u + "'. Valid: none, single, double, singleAccounting, doubleAccounting (or a boolean).");
                    }
                }
            }

            if (op.TryGetProperty("horizontalAlignment", out var hAlign) && hAlign.ValueKind == JsonValueKind.String)
            {
                Excel.XlHAlign mapped;
                if (!HAlignMap.TryGetValue(hAlign.GetString(), out mapped))
                    throw new ArgumentException("format_range: unknown horizontalAlignment '" + hAlign.GetString() + "'. Valid: " + string.Join(", ", HAlignMap.Keys) + ".");
                range.HorizontalAlignment = mapped;
            }
            if (op.TryGetProperty("verticalAlignment", out var vAlign) && vAlign.ValueKind == JsonValueKind.String)
            {
                Excel.XlVAlign mapped;
                if (!VAlignMap.TryGetValue(vAlign.GetString(), out mapped))
                    throw new ArgumentException("format_range: unknown verticalAlignment '" + vAlign.GetString() + "'. Valid: " + string.Join(", ", VAlignMap.Keys) + ".");
                range.VerticalAlignment = mapped;
            }

            if (op.TryGetProperty("wrapText", out var wt)) range.WrapText = wt.ValueKind == JsonValueKind.True;
            if (op.TryGetProperty("textRotation", out var tr) && tr.ValueKind == JsonValueKind.Number)
            {
                int deg = tr.GetInt32();
                if (deg < -90 || deg > 90)
                    throw new ArgumentOutOfRangeException("textRotation", "textRotation must be between -90 and 90 degrees.");
                range.Orientation = deg;
            }
            if (op.TryGetProperty("indent", out var ind) && ind.ValueKind == JsonValueKind.Number)
            {
                int lvl = ind.GetInt32();
                if (lvl < 0 || lvl > 15)
                    throw new ArgumentOutOfRangeException("indent", "indent must be between 0 and 15.");
                range.IndentLevel = lvl;
            }

            if (op.TryGetProperty("borders", out var borders) && borders.ValueKind == JsonValueKind.Object)
            {
                note = ApplyBorders(range, borders);
            }
            return note;
        }

        // Returns a note when interior-border edges were silently skipped on a
        // single-cell range (Excel throws for insideHorizontal/insideVertical
        // there), or null otherwise.
        private static string ApplyBorders(Excel.Range range, JsonElement borders)
        {
            var edges = new List<string>();
            if (borders.TryGetProperty("preset", out var preset) && preset.ValueKind == JsonValueKind.String)
            {
                switch (preset.GetString())
                {
                    case "none":
                        foreach (Excel.XlBordersIndex idx in BorderEdgeMap.Values) range.Borders[idx].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                        return null; // clearing is the whole request; edges/style below don't apply to "none"
                    case "outline": edges.AddRange(OutlineEdges); break;
                    case "all": edges.AddRange(AllEdges); break;
                    case "thick-outline": edges.AddRange(OutlineEdges); break;
                    default:
                        throw new ArgumentException("format_range: unknown borders.preset '" + preset.GetString() + "'. Valid: none, outline, all, thick-outline.");
                }
            }
            if (borders.TryGetProperty("edges", out var edgesEl) && edgesEl.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement e in edgesEl.EnumerateArray())
                {
                    string edge = e.GetString();
                    if (!BorderEdgeMap.ContainsKey(edge))
                        throw new ArgumentException("format_range: unknown borders.edges value '" + edge + "'. Valid: " + string.Join(", ", BorderEdgeMap.Keys) + ".");
                    if (!edges.Contains(edge)) edges.Add(edge);
                }
            }
            if (edges.Count == 0) edges.AddRange(OutlineEdges); // borders object given with no preset/edges - sane default

            string styleName = borders.TryGetProperty("style", out var styleEl) && styleEl.ValueKind == JsonValueKind.String
                ? styleEl.GetString()
                : (preset.ValueKind == JsonValueKind.String && preset.GetString() == "thick-outline" ? "thick" : "thin");
            Excel.XlLineStyle lineStyle;
            Excel.XlBorderWeight weight;
            switch (styleName)
            {
                case "thin": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlThin; break;
                case "medium": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlMedium; break;
                case "thick": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlThick; break;
                case "double": lineStyle = Excel.XlLineStyle.xlDouble; weight = Excel.XlBorderWeight.xlThick; break;
                case "dotted": lineStyle = Excel.XlLineStyle.xlDot; weight = Excel.XlBorderWeight.xlThin; break;
                case "dashed": lineStyle = Excel.XlLineStyle.xlDash; weight = Excel.XlBorderWeight.xlThin; break;
                case "none": lineStyle = Excel.XlLineStyle.xlLineStyleNone; weight = Excel.XlBorderWeight.xlThin; break;
                default:
                    throw new ArgumentException("format_range: unknown borders.style '" + styleName + "'. Valid: thin, medium, thick, double, dotted, dashed, none.");
            }

            int? oleColor = null;
            if (borders.TryGetProperty("color", out var colorEl) && colorEl.ValueKind == JsonValueKind.String)
                oleColor = ColorUtil.HexToOle(colorEl.GetString());

            bool singleCell = range.Cells.Count == 1;
            bool skippedInterior = false;
            foreach (string edge in edges)
            {
                if (singleCell && (edge == "insideHorizontal" || edge == "insideVertical"))
                {
                    skippedInterior = true;
                    continue; // Excel throws for interior borders on a single cell - silently skip, noted to the caller below
                }
                Excel.Border border = range.Borders[BorderEdgeMap[edge]];
                border.LineStyle = lineStyle;
                if (lineStyle != Excel.XlLineStyle.xlLineStyleNone)
                {
                    border.Weight = weight;
                    if (oleColor.HasValue) border.Color = oleColor.Value;
                }
            }
            return skippedInterior
                ? "insideHorizontal/insideVertical skipped (single-cell range has no interior edges)"
                : null;
        }

        private static void InsertDeleteRows(JsonElement op, bool insert)
        {
            int startRow = op.GetProperty("startRow").GetInt32();
            int count = op.GetProperty("count").GetInt32();
            Excel.Range rows = Sheet(op).Range[$"{startRow}:{startRow + count - 1}"];
            if (insert) rows.EntireRow.Insert(); else rows.EntireRow.Delete();
        }

        private static void InsertDeleteCols(JsonElement op, bool insert)
        {
            int startCol = op.GetProperty("startCol").GetInt32();
            int count = op.GetProperty("count").GetInt32();
            string startLetter = TextUtil.ColumnLetter(startCol);
            string endLetter = TextUtil.ColumnLetter(startCol + count - 1);
            Excel.Range cols = Sheet(op).Range[$"{startLetter}:{endLetter}"];
            if (insert) cols.EntireColumn.Insert(); else cols.EntireColumn.Delete();
        }

        // Returns a description including the created chart's name (PP-15
        // Task 4), so a follow-up edit_chart in the same batch can address it
        // without guessing Excel's auto-assigned "Chart 1"-style name.
        // Excel's SetSourceData auto-detects category (x-axis) labels only when
        // the leftmost column/top row of the bound range is text - if it's
        // numeric, Excel can't tell it apart from another value series and the
        // chart falls back to a plain 1,2,3... index. This forces every
        // series' XValues to an explicit range so the model can put a numeric
        // column (dates, ids, years) on the x-axis on purpose.
        private static void ApplyChartCategoryRange(dynamic chart, Excel.Range categories)
        {
            dynamic seriesCollection = chart.SeriesCollection();
            int count = seriesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                seriesCollection.Item(i).XValues = categories;
            }
        }

        private static string AddChart(JsonElement op)
        {
            Excel.Worksheet sheet = Sheet(op);
            string dataRange = op.GetProperty("dataRange").GetString();
            dynamic chartObjects = sheet.ChartObjects();
            dynamic chartObj = chartObjects.Add(100, 20, 400, 250);
            dynamic chart = chartObj.Chart;
            chart.SetSourceData(sheet.Range[dataRange]);
            if (op.TryGetProperty("categoryRange", out var catEl) && catEl.ValueKind == JsonValueKind.String)
            {
                ApplyChartCategoryRange(chart, sheet.Range[catEl.GetString()]);
            }
            int chartTypeCode = 51; // xlColumnClustered
            if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                if (!ExcelChartTypeMap.TryGetValue(ct.GetString(), out chartTypeCode))
                    throw new ArgumentException("add_chart: unknown chartType '" + ct.GetString() +
                                                "'. Valid: " + string.Join(", ", ExcelChartTypeMap.Keys) + ".");
            }
            chart.ChartType = chartTypeCode;
            if (op.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            if (op.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
            {
                chartObj.Name = nameEl.GetString();
            }
            return "name=" + (string)chartObj.Name;
        }

        private static void EditChartExcel(JsonElement op)
        {
            string chartName = op.GetProperty("chartPath").GetString(); // this project's visualId-equivalent for charts
            dynamic chartObjects = Sheet(op).ChartObjects();
            dynamic chartObj = chartObjects.Item(chartName);
            dynamic chart = chartObj.Chart;

            // PP-15 Task 2: rebinding, done FIRST - some chart-type changes
            // reset the plot, so this must happen before ChartType below.
            // dataSheet lets the chart's own sheet differ from the data's
            // sheet (e.g. chart on Sheet1, data on Sheet2).
            if (op.TryGetProperty("dataRange", out var dr) && dr.ValueKind == JsonValueKind.String)
            {
                Excel.Worksheet dataSheet = op.TryGetProperty("dataSheet", out var dsEl) && dsEl.ValueKind == JsonValueKind.String
                    ? (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[dsEl.GetString()]
                    : Sheet(op);
                Excel.Range source = dataSheet.Range[dr.GetString()];
                if (op.TryGetProperty("plotBy", out var pb) && pb.ValueKind == JsonValueKind.String)
                {
                    Excel.XlRowCol plotBy;
                    if (pb.GetString() == "rows") plotBy = Excel.XlRowCol.xlRows;
                    else if (pb.GetString() == "columns") plotBy = Excel.XlRowCol.xlColumns;
                    else throw new ArgumentException("edit_chart: unknown plotBy '" + pb.GetString() + "'. Valid: rows, columns.");
                    chart.SetSourceData(source, plotBy);
                }
                else
                {
                    chart.SetSourceData(source); // Excel infers orientation from the range shape when omitted
                }
            }

            if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                int typeCode;
                if (!ExcelChartTypeMap.TryGetValue(ct.GetString(), out typeCode))
                    throw new ArgumentException("edit_chart: unknown chartType '" + ct.GetString() +
                                                "'. Valid: " + string.Join(", ", ExcelChartTypeMap.Keys) + ".");
                chart.ChartType = typeCode;
            }

            // Independent of dataRange/plotBy above - fixes the x-axis to an
            // explicit column/row even when the chart isn't otherwise being
            // rebound (see ApplyChartCategoryRange).
            if (op.TryGetProperty("categoryRange", out var catEl) && catEl.ValueKind == JsonValueKind.String)
            {
                Excel.Worksheet categorySheet = op.TryGetProperty("dataSheet", out var catDsEl) && catDsEl.ValueKind == JsonValueKind.String
                    ? (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[catDsEl.GetString()]
                    : Sheet(op);
                ApplyChartCategoryRange(chart, categorySheet.Range[catEl.GetString()]);
            }
            if (op.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            if (op.TryGetProperty("legend", out var legend) && legend.ValueKind == JsonValueKind.String)
            {
                // PP-21 Task 2 Step 5: was a terminal-else-to-bottom - any
                // unmatched value (a model could plausibly send anything not
                // in this exact set) silently moved the legend to the bottom
                // instead of erroring, the identical defect PP-21 fixes on
                // the PowerPoint side.
                string pos = legend.GetString();
                if (pos == "none") { chart.HasLegend = false; }
                else
                {
                    int position;
                    switch (pos)
                    {
                        case "right": position = -4152; break; // xlLegendPositionRight
                        case "top": position = -4160; break;   // xlLegendPositionTop
                        case "left": position = -4131; break;  // xlLegendPositionLeft
                        case "bottom": position = -4107; break; // xlLegendPositionBottom
                        default:
                            throw new ArgumentException("edit_chart: unknown legend '" + pos + "'. Valid: none, right, top, left, bottom.");
                    }
                    chart.HasLegend = true;
                    chart.Legend.Position = position;
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
                    series.Format.Fill.ForeColor.RGB = ColorUtil.HexToOle(prop.Value.GetString());
                }
            }

            if (op.TryGetProperty("seriesData", out var seriesData) && seriesData.ValueKind == JsonValueKind.Array)
            {
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
        }

        // Returns a description including the target address (PP-18 Task 3
        // Step 5) - this PIA's SparklineGroup exposes no separate stable id
        // beyond the cells it occupies, so the target address IS the
        // addressing handle for a later edit (there is no delete_sparkline/
        // edit_sparkline operation yet to consume it, but it is at least
        // visible in the transcript for the user/model to reason about).
        private static string AddSparkline(JsonElement op)
        {
            string dataRange = op.GetProperty("dataRange").GetString();
            // Required, not defaulted to dataRange: that default drew the
            // sparkline over its own source numbers, an actively-wrong result
            // that was silent before this fix.
            if (!op.TryGetProperty("targetCell", out var tc) || tc.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_sparkline: 'targetCell' is required (the cell immediately right of, or below, dataRange is the usual choice - it must not overlap dataRange).");
            string targetCell = tc.GetString();

            string type = "line";
            if (op.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String)
            {
                type = t.GetString();
                if (type != "line" && type != "column" && type != "stacked")
                    throw new ArgumentException("add_sparkline: unknown type '" + type + "'. Valid: line, column, stacked.");
            }

            Excel.Worksheet sheet = Sheet(op);
            Excel.Range dataRangeObj = sheet.Range[dataRange];
            Excel.Range targetRangeObj = sheet.Range[targetCell];

            // Shape validation: one sparkline per data row (multi-cell target
            // matching row count) or a single sparkline for the whole range
            // (single-cell target). Anything else fails opaquely in the COM
            // call below without this check.
            int dataRows = dataRangeObj.Rows.Count;
            int targetCells = targetRangeObj.Cells.Count;
            bool overlaps = string.Equals(
                dataRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, false],
                targetRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, false],
                StringComparison.OrdinalIgnoreCase);
            if (overlaps)
                throw new ArgumentException("add_sparkline: targetCell must not be the same as dataRange - the sparkline would draw over its own source data.");
            if (targetCells != 1 && targetCells != dataRows)
                throw new ArgumentException("add_sparkline: targetCell has " + targetCells + " cell(s) but dataRange has " +
                    dataRows + " row(s) - targetCell must be a single cell (one sparkline for the whole range) or match the row count (one sparkline per row).");

            dynamic groups = targetRangeObj.SparklineGroups;
            Excel.XlSparkType sparkType = type == "column" ? Excel.XlSparkType.xlSparkColumn
                : type == "stacked" ? Excel.XlSparkType.xlSparkColumnStacked100
                : Excel.XlSparkType.xlSparkLine;
            dynamic group = groups.Add(sparkType, dataRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, true]);
            if (op.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
            {
                group.SeriesColor.Color = ColorUtil.HexToOle(color.GetString());
            }
            return "target=" + targetCell;
        }

        // Returns the created shape's name (PP-16 Task 3), so a follow-up
        // edit_shape/delete_visual in the same batch can address it without
        // guessing Excel's auto-generated name.
        private static string AddShapeExcel(JsonElement op)
        {
            string shapeType = op.GetProperty("shapeType").GetString();
            string anchorCell = op.GetProperty("anchorCell").GetString();
            Excel.Range anchor = Sheet(op).Range[anchorCell];
            float left = (float)(double)anchor.Left;
            float top = (float)(double)anchor.Top;
            float width = 100f, height = 60f;

            Excel.Shape shape;
            if (string.Equals(shapeType, "textbox", StringComparison.OrdinalIgnoreCase))
            {
                shape = Sheet(op).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            }
            else
            {
                int msoTypeInt;
                if (!ShapeTypes.ByName.TryGetValue(shapeType, out msoTypeInt))
                    throw new ArgumentException("add_shape: unknown shapeType '" + shapeType + "'. Valid: textbox, " +
                                                string.Join(", ", ShapeTypes.ByName.Keys) + ".");
                shape = Sheet(op).Shapes.AddShape((Microsoft.Office.Core.MsoAutoShapeType)msoTypeInt, left, top, width, height);
            }
            if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
            {
                shape.Fill.ForeColor.RGB = ColorUtil.HexToOle(fill.GetString());
            }
            if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
            {
                shape.TextFrame.Characters().Text = text.GetString();
            }
            if (op.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
            {
                shape.Name = nameEl.GetString();
            }
            return shape.Name;
        }

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
            string startLetter = TextUtil.ColumnLetter(col);
            string endLetter = TextUtil.ColumnLetter(col + count - 1);
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
            string startLetter = TextUtil.ColumnLetter(col);
            string endLetter = TextUtil.ColumnLetter(col + count - 1);
            Sheet(op).Range[$"{startLetter}:{endLetter}"].EntireColumn.Hidden = hidden;
        }

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

        private static void SetPageSetup(JsonElement op)
        {
            Excel.PageSetup setup = Sheet(op).PageSetup;
            if (op.TryGetProperty("orientation", out var orient) && orient.ValueKind == JsonValueKind.String)
                setup.Orientation = orient.GetString() == "landscape" ? Excel.XlPageOrientation.xlLandscape : Excel.XlPageOrientation.xlPortrait;

            bool hasScale = op.TryGetProperty("scale", out var scaleEl) && scaleEl.ValueKind == JsonValueKind.Number;
            bool hasFitWidth = op.TryGetProperty("fitToWidth", out var ftwEl) && ftwEl.ValueKind == JsonValueKind.Number;
            bool hasFitHeight = op.TryGetProperty("fitToHeight", out var fthEl) && fthEl.ValueKind == JsonValueKind.Number;
            bool hasFit = hasFitWidth || hasFitHeight;
            if (hasScale && hasFit)
                throw new ArgumentException(
                    "set_page_setup: 'scale' and 'fitToWidth'/'fitToHeight' are mutually exclusive in Excel " +
                    "(matching its own Page Setup UI). Pass either scale, or one/both fit values - not both.");

            if (hasFit)
            {
                setup.Zoom = false; // Zoom and FitToPages are mutually exclusive in Excel's own UI
                if (hasFitWidth) setup.FitToPagesWide = (int)ftwEl.GetDouble();
                // fitToHeight: 0 means "unlimited tall" - Excel expresses that as
                // FitToPagesTall = false, the single most common real-world
                // page-setup request ("fit on one page wide") that was otherwise
                // unexpressible.
                if (hasFitHeight)
                {
                    double fth = fthEl.GetDouble();
                    if (fth == 0) setup.FitToPagesTall = false;
                    else setup.FitToPagesTall = (int)fth;
                }
            }
            else if (hasScale)
            {
                int scaleVal = (int)scaleEl.GetDouble();
                if (scaleVal < 10 || scaleVal > 400)
                    throw new ArgumentOutOfRangeException("scale", "set_page_setup: scale must be between 10 and 400.");
                setup.Zoom = scaleVal;
            }

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
            int position = op.GetProperty("position").GetInt32(); // 1-based, desired final position
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet target = Sheet(op);
            int count = wb.Worksheets.Count;
            int currentIndex = target.Index; // 1-based
            int clamped = Math.Max(1, Math.Min(position, count));
            if (clamped == currentIndex) return; // already there
            if (clamped > currentIndex)
            {
                int afterShift = clamped + 1;
                if (afterShift > count) target.Move(After: wb.Worksheets[count]);
                else target.Move(Before: wb.Worksheets[afterShift]);
            }
            else
            {
                target.Move(Before: wb.Worksheets[clamped]);
            }
        }

        private static void ProtectSheet(JsonElement op)
        {
            bool isProtected = op.GetProperty("protected").GetBoolean();
            Excel.Worksheet sheet = Sheet(op);
            if (isProtected) sheet.Protect();
            else sheet.Unprotect();
        }

        private static Excel.Shape ResolveShapeByName(JsonElement op, string idField)
        {
            string visualId = op.GetProperty(idField).GetString();
            return Sheet(op).Shapes.Item(visualId);
        }

        private static void EditShapeExcel(JsonElement op)
        {
            Excel.Shape shape = ResolveShapeByName(op, "visualId");
            try
            {
                if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
                {
                    shape.TextFrame.Characters().Text = text.GetString();
                }
            }
            catch (System.Runtime.InteropServices.COMException) { /* shape doesn't support a text frame */ }
            if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
            {
                shape.Fill.ForeColor.RGB = ColorUtil.HexToOle(fill.GetString());
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

        // Returns a description of what happened (PP-18 Task 2 Step 5), since
        // "delete_table" defaulting to keeping the data is easy to miss.
        private static string DeleteTable(JsonElement op)
        {
            Excel.ListObject table = ResolveTable(op);
            bool deleteData = op.TryGetProperty("deleteData", out var dd) && dd.ValueKind == JsonValueKind.True;
            string shift = op.TryGetProperty("shift", out var sh) && sh.ValueKind == JsonValueKind.String ? sh.GetString() : "up";

            if (!deleteData)
            {
                table.Unlist(); // converts back to a plain range, keeping data/formatting
                return "converted to a plain range - data and formatting kept (pass deleteData:true to remove the cells too)";
            }

            if (shift != "up" && shift != "left" && shift != "none")
                throw new ArgumentException("delete_table: unknown shift '" + shift + "'. Valid: up, left, none."); // validated BEFORE Unlist() below, so a bad value leaves the table untouched

            Excel.Range body = table.Range; // capture before Unlist() - the ListObject reference is invalid after
            table.Unlist();
            switch (shift)
            {
                case "none":
                    body.ClearContents();
                    return "converted to a plain range, cells cleared in place";
                case "left":
                    body.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    return "converted to a plain range, cells removed (shifted left)";
                default:
                    body.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    return "converted to a plain range, cells removed (shifted up)";
            }
        }

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

        // PP-17: Excel rejects a name colliding with a cell address, starting
        // with a digit, or containing a space - with an unhelpful COM error.
        // Pre-check so the model gets a specific reason instead.
        private static readonly System.Text.RegularExpressions.Regex CellAddressLike =
            new System.Text.RegularExpressions.Regex(@"^\$?[A-Za-z]{1,3}\$?[0-9]+$");

        private static void ValidateDefinedNameSyntax(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("add_defined_name: name cannot be empty.");
            if (char.IsDigit(name[0]))
                throw new ArgumentException("add_defined_name: name '" + name + "' cannot start with a digit.");
            if (name.IndexOf(' ') >= 0)
                throw new ArgumentException("add_defined_name: name '" + name + "' cannot contain spaces.");
            if (CellAddressLike.IsMatch(name))
                throw new ArgumentException("add_defined_name: name '" + name + "' looks like a cell address, which Excel does not allow as a defined name.");
        }

        private static bool DefinedNameExists(Excel.Names names, string name)
        {
            foreach (Excel.Name n in names)
            {
                if (string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static void AddDefinedName(JsonElement op)
        {
            string name = op.GetProperty("name").GetString();
            string reference = op.GetProperty("ref").GetString();
            bool sheetScoped = op.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String && sc.GetString() == "sheet";
            bool overwrite = op.TryGetProperty("overwrite", out var ow) && ow.ValueKind == JsonValueKind.True;

            ValidateDefinedNameSyntax(name);

            Excel.Worksheet targetSheet = Sheet(op); // honors the existing optional "sheet" property either way
            string refersTo = reference.StartsWith("=") ? reference : "=" + reference;
            // Qualify an unqualified reference to the target sheet - otherwise a
            // workbook-scoped name resolves against whichever sheet is active
            // at evaluation time, a latent wrong-answer bug for both scopes.
            if (refersTo.IndexOf('!') < 0)
            {
                string sheetName = targetSheet.Name;
                string quotedSheet = sheetName.IndexOf(' ') >= 0 ? "'" + sheetName + "'" : sheetName;
                refersTo = "=" + quotedSheet + "!" + refersTo.Substring(1);
            }

            Excel.Names names = sheetScoped ? targetSheet.Names : Globals.ThisAddIn.Application.ActiveWorkbook.Names;
            if (!overwrite && DefinedNameExists(names, name))
                throw new ArgumentException("add_defined_name: a " + (sheetScoped ? "sheet" : "workbook") +
                    "-scoped name '" + name + "' already exists. Pass overwrite:true to replace it.");

            names.Add(name, refersTo);
        }

        private static void DeleteDefinedName(JsonElement op)
        {
            string name = op.GetProperty("name").GetString();
            bool sheetScoped = op.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String && sc.GetString() == "sheet";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Names names = sheetScoped ? Sheet(op).Names : wb.Names;

            if (!DefinedNameExists(names, name))
            {
                // Point the model at the other scope if the name exists there -
                // turns a dead end into a self-correcting next turn.
                string searchedScope = sheetScoped ? "sheet" : "workbook";
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (!sheetScoped && DefinedNameExists(sheet.Names, name))
                    {
                        throw new ArgumentException("delete_defined_name: no workbook-scoped name '" + name +
                            "' found; a sheet-scoped name with this name exists on '" + sheet.Name +
                            "' - pass scope:'sheet' and sheet:'" + sheet.Name + "'.");
                    }
                }
                if (sheetScoped && DefinedNameExists(wb.Names, name))
                {
                    throw new ArgumentException("delete_defined_name: no sheet-scoped name '" + name +
                        "' found on this sheet; a workbook-scoped name with this name exists - omit scope to target it.");
                }
                throw new ArgumentException("delete_defined_name: no " + searchedScope + "-scoped name '" + name + "' found.");
            }

            names.Item(name).Delete();
        }

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
                // XlDVType enum verification (via reflection against this machine's
                // Microsoft.Office.Interop.Excel PIA) shows 8 total validation kinds:
                // xlValidateInputOnly, xlValidateWholeNumber, xlValidateDecimal, xlValidateList,
                // xlValidateDate, xlValidateTime, xlValidateTextLength, xlValidateCustom. None of
                // these map to boolean-checkbox cells. The assembly does define CheckBox and
                // CheckBoxes types, but they are form controls (accessed via Shapes.AddFormControl),
                // not Data Validation options. Thus, Excel's native checkbox-cell feature (if it
                // exists in newer Office 365 builds) is not accessible through the Validation API
                // in this Interop version.
                case "checkbox":
                    throw new NotSupportedException("set_data_validation: 'checkbox' kind is not supported in this version of Excel Interop - CheckBox is a form control, not a Data Validation type.");
                default:
                    throw new ArgumentException("set_data_validation: unknown validation kind '" + kind + "'.");
            }
        }

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
    }
}
