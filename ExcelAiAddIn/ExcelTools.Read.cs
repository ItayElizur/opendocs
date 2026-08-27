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

    }
}

