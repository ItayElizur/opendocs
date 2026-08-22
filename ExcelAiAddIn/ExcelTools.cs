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

        private static void AddSparkline(JsonElement op)
        {
            string dataRange = op.GetProperty("dataRange").GetString();
            string targetCell = op.TryGetProperty("targetCell", out var tc) && tc.ValueKind == JsonValueKind.String ? tc.GetString() : dataRange;
            string type = op.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() : "line";
            dynamic sheet = Sheet(op);
            dynamic groups = sheet.SparklineGroups;
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
    }
}
