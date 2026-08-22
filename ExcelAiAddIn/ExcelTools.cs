using System;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static class ExcelTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                switch (name)
                {
                    case "get_workbook_context": return GetWorkbookContext();
                    case "read_range": return ReadRange(input);
                    case "read_cells": return ReadCells(input);
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
    }
}
