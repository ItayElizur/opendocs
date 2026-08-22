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
    }
}
