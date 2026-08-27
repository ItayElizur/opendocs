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

