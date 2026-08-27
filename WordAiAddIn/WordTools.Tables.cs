using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        // PP-23 Task 1: 0-based at the tool boundary, matching every other
        // index in this file; Document.Tables is 1-based in COM. Document
        // order (no inline-vs-floating split needed - Word tables are always
        // flow content, unlike charts/SmartArt which can float).
        private static Word.Table ResolveTable(JsonElement input)
        {
            // Bug found via DebugLog from a real repro (2026-08-24): this
            // used GetProperty (required), throwing KeyNotFoundException
            // whenever tableIndex was omitted - directly contradicting this
            // tool's own documented "omit to target the first table"
            // behavior (and ReadTable's own correct TryGetProperty pattern,
            // right below in this same file). Fixed to match.
            int tableIndex = input.TryGetProperty("tableIndex", out var ti) && ti.ValueKind == JsonValueKind.Number ? ti.GetInt32() : 0;
            Word.Tables tables = ActiveDoc.Tables;
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentOutOfRangeException("tableIndex",
                    "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
            return tables[tableIndex + 1];
        }

        private static ToolResult AddTable(JsonElement input)
        {
            int rows = input.GetProperty("rows").GetInt32();
            int cols = input.GetProperty("cols").GetInt32();
            if (rows < 1 || cols < 1)
                throw new ArgumentException("add_table: rows and cols must each be at least 1.");

            int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : (int?)null;
            Word.Range at = afterBlockIndex.HasValue ? RangeAfterBlock(afterBlockIndex.Value) : EndOfDocumentRange();

            Word.Table table = ActiveDoc.Tables.Add(at, rows, cols);

            if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
            {
                int r = 0;
                foreach (JsonElement rowEl in cells.EnumerateArray())
                {
                    if (r >= rows) break; // ignore extra rows beyond the declared size rather than throwing mid-write
                    int c = 0;
                    foreach (JsonElement cellEl in rowEl.EnumerateArray())
                    {
                        if (c >= cols) break;
                        table.Cell(r + 1, c + 1).Range.Text = cellEl.GetString();
                        c++;
                    }
                    r++;
                }
            }

            int newIndex = ActiveDoc.Tables.Count - 1; // Tables.Add appends; stable immediately after the call
            return new ToolResult
            {
                Output = "Table added at index " + newIndex + " (" + rows + " rows x " + cols + " cols).",
                Mutated = true,
                Summary = "add_table",
            };
        }

        private static ToolResult EditTable(JsonElement input)
        {
            Word.Table table = ResolveTable(input);
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "set_cell":
                {
                    int row = input.GetProperty("row").GetInt32();
                    int col = input.GetProperty("col").GetInt32();
                    if (row < 0 || row >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (table.Rows.Count - 1) + ".");
                    if (col < 0 || col >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (table.Columns.Count - 1) + ".");
                    table.Cell(row + 1, col + 1).Range.Text = input.GetProperty("text").GetString();
                    return new ToolResult { Output = "Cell [" + row + "," + col + "] updated.", Mutated = true, Summary = "edit_table" };
                }
                case "insert_row":
                case "delete_row":
                case "insert_col":
                case "delete_col":
                {
                    // Same index-always-existing, before/after-picks-side
                    // convention as PowerPoint's edit_table_structure - index
                    // always addresses an EXISTING row/column (0-based).
                    int index = input.GetProperty("index").GetInt32();
                    bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
                    if (kind == "insert_row" || kind == "delete_row")
                    {
                        if (index < 0 || index >= table.Rows.Count)
                            throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for " + kind + ".");
                        if (kind == "insert_row") table.Rows.Add(table.Rows[before ? index + 1 : Math.Min(index + 2, table.Rows.Count + 1)]);
                        else table.Rows[index + 1].Delete();
                    }
                    else
                    {
                        if (index < 0 || index >= table.Columns.Count)
                            throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for " + kind + ".");
                        if (kind == "insert_col") table.Columns.Add(table.Columns[before ? index + 1 : Math.Min(index + 2, table.Columns.Count + 1)]);
                        else table.Columns[index + 1].Delete();
                    }
                    return new ToolResult
                    {
                        Output = kind + " applied at index " + index + ". Row/column indices after this point have shifted - re-read the table before another structural edit in the same run.",
                        Mutated = true,
                        Summary = "edit_table",
                    };
                }
                case "set_style":
                {
                    if (input.TryGetProperty("styleName", out var styleEl) && styleEl.ValueKind == JsonValueKind.String)
                    {
                        try { table.set_Style(styleEl.GetString()); }
                        catch (Exception ex) { throw new ArgumentException("edit_table: '" + styleEl.GetString() + "' is not a valid table style name in this document/template. " + ex.Message); }
                    }
                    if (input.TryGetProperty("headerRow", out var hdr))
                        table.ApplyStyleHeadingRows = hdr.ValueKind == JsonValueKind.True;
                    if (input.TryGetProperty("bandedRows", out var band))
                        table.ApplyStyleRowBands = band.ValueKind == JsonValueKind.True;
                    // Post-hoc fix (2026-08-24, user-reported): this branch had
                    // no border support at all - "borders" is a real field on
                    // updateParagraphStyle elsewhere in this file, and the model
                    // (reasonably, given that precedent) called edit_table with
                    // the same field name expecting the same effect. Since
                    // set_style never checked for it, the call silently did
                    // nothing - a real gap, not a user error. table.Borders
                    // mirrors the Word.Border collection updateParagraphStyle
                    // already uses for paragraph borders, applied here at the
                    // whole-table level (outside + inside edges).
                    if (input.TryGetProperty("borders", out var bordersEl))
                    {
                        bool on = bordersEl.ValueKind == JsonValueKind.True;
                        Word.WdColor color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String
                            ? (Word.WdColor)ColorUtil.HexToOle(bc.GetString())
                            : (Word.WdColor)ColorUtil.HexToOle("#000000");
                        // Post-hoc fix (2026-08-24, user-reported): table.Borders
                        // is not just the 6 grid sides - it also includes
                        // wdBorderDiagonalDown/wdBorderDiagonalUp (the rare
                        // cell-split diagonal lines), so the blind foreach over
                        // the whole collection turned those on too, producing
                        // crisscrossing diagonals across every cell. Enumerate
                        // only the real table grid sides explicitly.
                        Word.WdBorderType[] sides =
                        {
                            Word.WdBorderType.wdBorderTop, Word.WdBorderType.wdBorderLeft,
                            Word.WdBorderType.wdBorderBottom, Word.WdBorderType.wdBorderRight,
                            Word.WdBorderType.wdBorderHorizontal, Word.WdBorderType.wdBorderVertical,
                        };
                        foreach (Word.WdBorderType side in sides)
                        {
                            Word.Border border = table.Borders[side];
                            border.LineStyle = on ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleNone;
                            if (on) border.Color = color;
                        }
                    }
                    return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table" };
                }
                case "set_shading":
                {
                    // Post-hoc addition (2026-08-24, user-requested): fills
                    // cell background color at cell/row/col/whole-table
                    // scope. Word.Cell.Shading.BackgroundPatternColor is the
                    // same property/pattern updateParagraphStyle's
                    // shadingFill already uses on paragraphs elsewhere in
                    // this file - applied per-cell here since Word tables
                    // have no single "shade this row" API, only per-cell
                    // shading (matches PowerPoint's own EditTableStyle,
                    // which does the identical per-cell loop for its
                    // shadingColor field).
                    string scope = input.GetProperty("scope").GetString();
                    Word.WdColor color = (Word.WdColor)ColorUtil.HexToOle(input.GetProperty("color").GetString());
                    int rowCount = table.Rows.Count, colCount = table.Columns.Count;
                    switch (scope)
                    {
                        case "cell":
                        {
                            int row = input.GetProperty("row").GetInt32();
                            int col = input.GetProperty("col").GetInt32();
                            if (row < 0 || row >= rowCount)
                                throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (rowCount - 1) + ".");
                            if (col < 0 || col >= colCount)
                                throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (colCount - 1) + ".");
                            table.Cell(row + 1, col + 1).Shading.BackgroundPatternColor = color;
                            break;
                        }
                        case "row":
                        {
                            int row = input.GetProperty("row").GetInt32();
                            if (row < 0 || row >= rowCount)
                                throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (rowCount - 1) + ".");
                            for (int c = 1; c <= colCount; c++)
                            {
                                // A merged cell's non-anchor positions throw on
                                // direct Cell(r,c) access - skip those rather
                                // than failing the whole row, same tolerance
                                // ReadTable already applies.
                                try { table.Cell(row + 1, c).Shading.BackgroundPatternColor = color; } catch { }
                            }
                            break;
                        }
                        case "col":
                        {
                            int col = input.GetProperty("col").GetInt32();
                            if (col < 0 || col >= colCount)
                                throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (colCount - 1) + ".");
                            for (int r = 1; r <= rowCount; r++)
                            {
                                try { table.Cell(r, col + 1).Shading.BackgroundPatternColor = color; } catch { }
                            }
                            break;
                        }
                        case "table":
                        {
                            for (int r = 1; r <= rowCount; r++)
                                for (int c = 1; c <= colCount; c++)
                                {
                                    try { table.Cell(r, c).Shading.BackgroundPatternColor = color; } catch { }
                                }
                            break;
                        }
                        default:
                            throw new ArgumentException("edit_table: unknown scope '" + scope + "' for set_shading. Valid: cell, row, col, table.");
                    }
                    return new ToolResult { Output = "Table shading applied (" + scope + ").", Mutated = true, Summary = "edit_table" };
                }
                default:
                    throw new ArgumentException("edit_table: unknown kind '" + kind + "'. Valid: set_cell, insert_row, delete_row, insert_col, delete_col, set_style, set_shading.");
            }
        }

        private static ToolResult ReadTable(JsonElement input)
        {
            Word.Tables tables = ActiveDoc.Tables;
            if (tables.Count == 0)
                return new ToolResult { Output = "No tables in this document.", Summary = "read_table" };

            int tableIndex = input.TryGetProperty("tableIndex", out var ti) && ti.ValueKind == JsonValueKind.Number ? ti.GetInt32() : 0;
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentOutOfRangeException("tableIndex", "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
            Word.Table table = tables[tableIndex + 1];

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Table " + tableIndex + " of " + tables.Count + " (" + table.Rows.Count + " rows x " + table.Columns.Count + " cols):");
            for (int r = 0; r < table.Rows.Count; r++)
            {
                var cellsOut = new List<string>();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    // A merged cell can make Cell(r,c) throw for the cells it no
                    // longer owns - report a placeholder rather than failing the
                    // whole read over one merged region.
                    try { cellsOut.Add(table.Cell(r + 1, c + 1).Range.Text.TrimEnd('\r', '\a')); }
                    catch { cellsOut.Add("(merged)"); }
                }
                sb.AppendLine("[" + r + "] " + string.Join(" | ", cellsOut));
            }
            return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_table" };
        }

        // PP-23 Task 4: ported from PowerPointTools.SmartArtLayouts.ByName /
        // ResolveSmartArtLayout verbatim - same seven keys, same
        // two-distinct-errors design (unknown key vs. valid-key-but-not-in-
        // this-install's-gallery). SmartArt is the Office-shared object
        // model, not PowerPoint-specific - Application.SmartArtLayouts
        // resolves identically against this add-in's own ThisAddIn.

    }
}

