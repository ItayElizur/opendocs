using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
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

    }
}

