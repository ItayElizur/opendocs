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

        // Single source for the "what can I send?" answer in the
        // unknown-kind error. Mirrors ProposeOperations' switch - edit both
        // together.
        private static readonly string[] KnownOperationKinds =
        {
            "set_cell", "set_formula", "set_range", "clear_cell", "clear_range", "find_replace",
            "format_range", "sort_range", "merge_cells", "unmerge_cells",
            "set_row_height", "set_col_width", "set_rows_hidden", "set_cols_hidden", "set_freeze",
            "insert_rows", "delete_rows", "insert_cols", "delete_cols", "set_page_setup",
            "add_sheet", "delete_sheet", "duplicate_sheet", "set_sheet_hidden", "move_sheet",
            "protect_sheet", "rename_sheet",
            "add_chart", "edit_chart", "delete_visual", "add_sparkline", "add_shape", "edit_shape",
            "add_image", "add_table", "add_table_row", "add_table_column", "delete_table_row",
            "delete_table_column", "delete_table", "add_pivot", "refresh_pivot",
            "set_hyperlink", "set_note", "add_defined_name", "delete_defined_name",
            "set_filter", "clear_filter", "set_filter_criteria",
            "add_conditional_format", "clear_conditional_formats", "set_data_validation",
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
                            Sheet(op).Range[op.GetProperty("address").GetString()].Value2 = JsonUtil.JsonValueToObject(op.GetProperty("value"));
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
                            // List what IS valid - a bare "unknown operation
                            // kind" is a dead end the model cannot correct
                            // itself from. RequiredFields is not the full set
                            // (kinds with no required fields are absent), so
                            // this reads the switch's own vocabulary.
                            lines.AppendLine(kind + ": unknown operation kind. Valid kinds: " +
                                             string.Join(", ", KnownOperationKinds) + ".");
                            anyError = true; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine((kind ?? "(unknown kind)") + ": ERROR - " + ex.Message); anyError = true;
                }
            }
            return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "propose_operations" };
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
                for (int c = 0; c < colCount; c++) grid[r, c] = JsonUtil.JsonValueToObject(row[c]);
            }
            Excel.Range topLeft = Sheet(op).Range[address];
            topLeft.Resize[rowCount, colCount].Value2 = grid;
        }

    }
}

