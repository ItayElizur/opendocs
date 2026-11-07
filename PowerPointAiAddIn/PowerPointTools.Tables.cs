using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // kind selects the operation: create, set_cell, the four row/column structure edits, or style.
        private static ToolResult EditTable(JsonElement input)
        {
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "create": return AddTable(input);
                case "set_cell": return EditTableCell(input);
                case "insert-row":
                case "delete-row":
                case "insert-col":
                case "delete-col": return EditTableStructure(input);
                case "style": return EditTableStyle(input);
                default:
                    throw new ArgumentException("edit_table: unknown kind '" + kind + "'. Valid: create, set_cell, insert-row, delete-row, insert-col, delete-col, style.");
            }
        }

        private static ToolResult AddTable(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int rows = input.GetProperty("rows").GetInt32();
            int cols = input.GetProperty("cols").GetInt32();
            ReadRectDefaulted(input, 200f, out float left, out float top, out float width, out float height);

            PowerPoint.Slide slide = ResolveSlide(input);
            PowerPoint.Shape tableShape = slide.Shapes.AddTable(rows, cols, left, top, width, height);
            if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
            {
                int r = 0;
                foreach (JsonElement rowEl in cells.EnumerateArray())
                {
                    int c = 0;
                    foreach (JsonElement cellEl in rowEl.EnumerateArray())
                    {
                        string cellText = cellEl.GetString();
                        PowerPoint.TextRange cellRange = tableShape.Table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange;
                        SetText(cellRange, cellText);
                        c++;
                    }
                    r++;
                }
            }
            int newShapeIndex = slide.Shapes.Count - 1;
            string named = ApplyOptionalName(tableShape, input);
            return new ToolResult { Output = "Table added at shapeIndex " + newShapeIndex + (named != null ? " (\"" + named + "\")" : "") + ".", Mutated = true };
        }

        private static PowerPoint.Table ResolveTable(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            if (shape.HasTable != Microsoft.Office.Core.MsoTriState.msoTrue)
                throw new InvalidOperationException("shape " + input.GetProperty("shapeIndex").GetInt32() +
                                                    " on slide " + input.GetProperty("slideIndex").GetInt32() +
                                                    " is not a table. Call read_slide to find the table's shapeIndex.");
            return shape.Table;
        }

        private static ToolResult EditTableCell(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            int row = input.GetProperty("row").GetInt32();
            int col = input.GetProperty("col").GetInt32();
            string text = input.GetProperty("paragraphs").GetString();
            PowerPoint.TextRange range = table.Cell(row + 1, col + 1).Shape.TextFrame.TextRange;
            SetText(range, text);
            return new ToolResult { Output = "Cell updated.", Mutated = true };
        }

        private static ToolResult EditTableStructure(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            string kind = input.GetProperty("kind").GetString();
            int index = input.GetProperty("index").GetInt32();
            bool before = ToolArgs.OptionalFlag(input, "before");
            // index always addresses an EXISTING row/column (0-based); before/
            // after decides which side of it the new one goes on - so the valid
            // range is the same for insert and delete. Un-validated, an
            // out-of-range index threw a raw, unhelpful COM error.
            switch (kind)
            {
                case "insert-row":
                    if (index < 0 || index >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for insert-row.");
                    table.Rows.Add(before ? index + 1 : index + 2);
                    break;
                case "delete-row":
                    if (index < 0 || index >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for delete-row.");
                    table.Rows[index + 1].Delete();
                    break;
                case "insert-col":
                    if (index < 0 || index >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for insert-col.");
                    table.Columns.Add(before ? index + 1 : index + 2);
                    break;
                case "delete-col":
                    if (index < 0 || index >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for delete-col.");
                    table.Columns[index + 1].Delete();
                    break;
                default:
                    return new ToolResult { Output = "Unknown structure kind: " + kind, IsError = true };
            }
            // Deleting/inserting a row or column shifts every later row/column's
            // index - the same trap manage_slide's delete action has. Callers doing more
            // than one structural edit in a run should re-read the table between
            // calls; documented in the schema description too.
            return new ToolResult { Output = "Table structure updated.", Mutated = true };
        }

        private static ToolResult EditTableStyle(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            if (input.TryGetProperty("firstRow", out var firstRow))
            {
                table.FirstRow = firstRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("bandRow", out var bandRow))
            {
                table.HorizBanding = bandRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("shadingColor", out var shading) && shading.ValueKind == JsonValueKind.String)
            {
                int color = ColorUtil.HexToOle(shading.GetString());
                foreach (PowerPoint.Row row in table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        cell.Shape.Fill.ForeColor.RGB = color;
                    }
                }
            }
            if (input.TryGetProperty("borderColor", out _) || input.TryGetProperty("borderWidthPt", out _) || input.TryGetProperty("borderPreset", out _))
            {
                string preset = input.TryGetProperty("borderPreset", out var bp) && bp.ValueKind == JsonValueKind.String ? bp.GetString() : "all";
                if (preset != "all" && preset != "none" && preset != "outline")
                    throw new ArgumentException("edit_table: unknown borderPreset '" + preset + "'. Valid: all, none, outline.");
                bool visible = preset != "none";
                float weight = ToolArgs.OptionalFloat(input, "borderWidthPt") ?? 1f;
                int color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String ? ColorUtil.HexToOle(bc.GetString()) : ColorUtil.HexToOle("#000000");
                PowerPoint.PpBorderType[] sides = { PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom, PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight };
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;
                int rIdx = 0;
                foreach (PowerPoint.Row row in table.Rows)
                {
                    int cIdx = 0;
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        foreach (PowerPoint.PpBorderType side in sides)
                        {
                            // "outline" only draws the table's outer perimeter -
                            // suppress every interior edge per-cell rather than
                            // per-table, reusing the existing per-cell loop.
                            bool sideVisible = visible;
                            if (visible && preset == "outline")
                            {
                                sideVisible = (side == PowerPoint.PpBorderType.ppBorderTop && rIdx == 0)
                                           || (side == PowerPoint.PpBorderType.ppBorderBottom && rIdx == rowCount - 1)
                                           || (side == PowerPoint.PpBorderType.ppBorderLeft && cIdx == 0)
                                           || (side == PowerPoint.PpBorderType.ppBorderRight && cIdx == colCount - 1);
                            }
                            PowerPoint.LineFormat border = cell.Borders[side];
                            border.Visible = Tri(sideVisible);
                            if (sideVisible)
                            {
                                border.Weight = weight;
                                border.ForeColor.RGB = color;
                            }
                        }
                        cIdx++;
                    }
                    rIdx++;
                }
            }
            return new ToolResult { Output = "Table style updated.", Mutated = true };
        }

    }
}

