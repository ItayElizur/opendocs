using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    public static class ExcelTools
    {
        public static EditingMode Mode = EditingMode.FullAutonomy;

        private static readonly string[] ExcelErrorTexts = { "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!" };

        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "get_workbook_context", "read_range", "read_cells", "select_range", "read_formats", "read_sheet_features", "find_cells", "trace_precedents", "trace_dependents",
        };

        // Note: the brief's source table also lists "plus"/"mathPlus" (MsoAutoShapeType.msoShapePlus/
        // msoShapeMathPlus), but those enum members do not exist in this project's referenced
        // Microsoft.Office.Core PIA (confirmed via CS0117 compile failure) - omitted; requests for
        // either shape type fall back to msoShapeRectangle per the table's existing fallback behavior.
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

        private static readonly Dictionary<string, int> ExcelChartTypeMap = new Dictionary<string, int>
        {
            ["column"] = 51, // xlColumnClustered
            ["bar"] = 57,     // xlBarClustered
            ["line"] = 4,     // xlLine
            ["area"] = 1,     // xlArea
            ["pie"] = 5,      // xlPie
            ["doughnut"] = -4120, // xlDoughnut
        };

        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                // Excel has no add_comment-equivalent tool yet, so Comment Only
                // mode allows no mutating tools at all (documented gap - see
                // Task 16 brief). Track Changes mode currently behaves the
                // same as Full Autonomy for gating purposes: Excel's
                // track-changes equivalent (Workbook.HighlightChangesOnScreen /
                // shared-workbook change tracking) is more limited than
                // Word's TrackRevisions and is out of scope for this task, so
                // there is deliberately no COM call wired up for it here.
                bool isMutating = !AlwaysAllowedTools.Contains(name);
                if (Mode == EditingMode.ReadOnly && isMutating)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
                }
                if (Mode == EditingMode.CommentOnly && isMutating)
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
            return new ToolResult { Output = $"Sheet: {sheet.Name}\nUsedRange: {usedRange}\nSelection: {selection}", Summary = "get_workbook_context" };
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
                bool underline = underlineRaw != null && Convert.ToInt32(underlineRaw) != -4142; // -4142 == xlUnderlineStyleNone
                string numberFormat = cell.NumberFormat as string;
                bool hasDefaultFormat = !bold && !italic && !underline && (numberFormat == "General" || numberFormat == null);
                if (hasDefaultFormat) continue; // only explicitly-formatted cells, matches genoffice
                sb.AppendLine($"{cell.Address[false, false]}: bold={bold}, italic={italic}, underline={underline}, numberFormat={numberFormat}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_formats" };
        }

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
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlTextString, String: text, TextOperator: Excel.XlContainsOperator.xlContains);
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
                    if (rule.TryGetProperty("format", out var top10Format))
                    {
                        if (top10Format.TryGetProperty("bold", out var bold)) top10.Font.Bold = bold.ValueKind == JsonValueKind.True;
                        if (top10Format.TryGetProperty("fontColor", out var fontColor)) top10.Font.Color = HexToOleColor(fontColor.GetString());
                        if (top10Format.TryGetProperty("fillColor", out var fillColor)) top10.Interior.Color = HexToOleColor(fillColor.GetString());
                    }
                    return; // Top10 doesn't implement FormatCondition in this PIA (confirmed via reflection) - format applied directly above, mirroring colorScale/dataBar's early-return pattern
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
                    Excel.Databar bar = target.FormatConditions.AddDatabar();
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

        private static ToolResult ProposeOperations(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            bool anyError = false;
            foreach (JsonElement op in input.GetProperty("operations").EnumerateArray())
            {
                string kind = op.GetProperty("kind").GetString();
                try
                {
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
                            FormatRange(op);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_cell":
                            Sheet(op).Range[op.GetProperty("address").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_range":
                            Sheet(op).Range[op.GetProperty("range").GetString()].ClearContents();
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "insert_rows": InsertDeleteRows(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_rows": InsertDeleteRows(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "insert_cols": InsertDeleteCols(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_cols": InsertDeleteCols(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_chart": AddChart(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_sparkline": AddSparkline(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_shape": AddShapeExcel(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
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
                        case "delete_table": DeleteTable(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_hyperlink": SetHyperlink(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_note": SetNote(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_defined_name": AddDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "delete_defined_name": DeleteDefinedName(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_filter": SetFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "clear_filter": ClearFilter(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_filter_criteria": SetFilterCriteria(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "add_conditional_format": AddConditionalFormat(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
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
                    lines.AppendLine(kind + ": ERROR - " + ex.Message); anyError = true;
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

        private static void FormatRange(JsonElement op)
        {
            Excel.Range range = Sheet(op).Range[op.GetProperty("address").GetString()];
            if (op.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean();
            if (op.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean();
            if (op.TryGetProperty("numberFormat", out var nf)) range.NumberFormat = nf.GetString();
            if (op.TryGetProperty("fillColor", out var fc))
            {
                string hex = fc.GetString().TrimStart('#');
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
            }
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
            string startLetter = ColumnLetter(startCol);
            string endLetter = ColumnLetter(startCol + count - 1);
            Excel.Range cols = Sheet(op).Range[$"{startLetter}:{endLetter}"];
            if (insert) cols.EntireColumn.Insert(); else cols.EntireColumn.Delete();
        }

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }

        private static void AddChart(JsonElement op)
        {
            Excel.Worksheet sheet = Sheet(op);
            string dataRange = op.GetProperty("dataRange").GetString();
            dynamic chartObjects = sheet.ChartObjects();
            dynamic chartObj = chartObjects.Add(100, 20, 400, 250);
            dynamic chart = chartObj.Chart;
            chart.SetSourceData(sheet.Range[dataRange]);
            int chartTypeCode = 51; // xlColumnClustered
            if (op.TryGetProperty("chartType", out var ct))
            {
                string t = ct.GetString();
                chartTypeCode = t == "line" ? 4 : t == "pie" ? 5 : 51;
            }
            chart.ChartType = chartTypeCode;
            if (op.TryGetProperty("title", out var title))
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
        }

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

        private static void AddSparkline(JsonElement op)
        {
            string dataRange = op.GetProperty("dataRange").GetString();
            string targetCell = op.TryGetProperty("targetCell", out var tc) && tc.ValueKind == JsonValueKind.String ? tc.GetString() : dataRange;
            string type = op.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() : "line";
            Excel.Worksheet sheet = Sheet(op);
            dynamic targetRange = sheet.Range[targetCell];
            dynamic groups = targetRange.SparklineGroups;
            Excel.XlSparkType sparkType = type == "column" ? Excel.XlSparkType.xlSparkColumn
                : type == "stacked" ? Excel.XlSparkType.xlSparkColumnStacked100
                : Excel.XlSparkType.xlSparkLine;
            dynamic group = groups.Add(sparkType, sheet.Range[dataRange].Address[true, true, Excel.XlReferenceStyle.xlA1, true]);
            if (op.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
            {
                group.SeriesColor.Color = HexToOleColor(color.GetString());
            }
        }

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

        private static int HexToOleColor(string hex)
        {
            hex = hex.TrimStart('#');
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
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

        private static void DeleteTable(JsonElement op)
        {
            ResolveTable(op).Unlist(); // converts back to a plain range, keeping data/formatting
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
