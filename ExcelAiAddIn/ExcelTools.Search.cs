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
        // Which sheets a call touches: sheetId for one sheet, allSheets for the whole workbook (Ctrl+F's
        // "Within: Workbook"), otherwise the active sheet only (Ctrl+F's default) - see .md for the
        // default-scope history.
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

        // Uses Excel's native Find/FindNext (the Ctrl+F engine) instead of scanning every cell's .Text;
        // wraps like Ctrl+F and stops via the first hit's address.
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
                        // Native error-cell scan - the VSTO/COM advantage over Office.js's wildcard-only
                        // Range.find that this project's feasibility report flagged.
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

            // Native Find/FindNext, not a per-cell scan. xlValues/xlFormulas are separate passes (Excel's
            // Find only searches one LookIn mode per call); "both" runs both and de-dupes by address to
            // match the old per-cell "both" semantics.
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

        // Write-side counterpart of find_cells (same sheetId/allSheets scoping, Ctrl+H's "Within: Sheet"
        // default). Only replaces literal string VALUES - never formulas/numbers - so a formula can't be
        // corrupted by a text substitution.
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

        // Uses Find/FindNext instead of scanning every cell's .Text; only replaces when Value2 is itself a
        // string, so a numeric/date/formula cell that merely displays a match (via its formatted text) is
        // skipped.
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

                // Captured before mutating current's cell - FindNext's own position tracking must not
                // reason about a cell that just changed underneath it.
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

    }
}

