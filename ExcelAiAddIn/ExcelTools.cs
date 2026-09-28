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
        // Task 11 (per-document since PP-1): see WordTools.cs's identical
        // pattern for the rationale - keyed by TaskPaneHost.GetChatId().
        private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

        public static void SetMode(string docKey, EditingMode mode)
        {
            ModeByDoc[docKey] = mode;
        }

        private static EditingMode ModeFor(string docKey)
        {
            EditingMode m;
            return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
        }

        private static readonly string[] ExcelErrorTexts = { "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!" };

        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "get_workbook_context", "read_range", "read_cells", "select_range", "read_formats", "read_sheet_features", "find_cells", "trace_precedents", "trace_dependents",
        };

        // Shape-name lookup now lives in OfficeAi.Shared.ShapeTypes (Phase 0) -
        // union of this map and PowerPoint's near-identical copy. PP-16:
        // mirrors EXCEL_SHAPE_TYPES / add_shape's shapeType enum in
        // ExcelAiAddIn/web-src/entry.ts exactly, plus the separately-handled
        // "textbox". Edit both together.

        // PP-15: chart-type vocabulary for BOTH add_chart and edit_chart now
        // lives in OfficeAi.Shared.ChartTypes, shared with Word and PowerPoint.
        // The cross-referencing this comment used to describe by hand (and the
        // PptChartTypeMap "bar" bug it records) is what motivated sharing it.

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            try
            {
                EditingMode mode = ModeFor(docKey);
                // Excel has no add_comment-equivalent tool yet, so Comment Only
                // mode allows no mutating tools at all (documented gap - see
                // Task 16 brief). Track Changes mode currently behaves the
                // same as Full Autonomy for gating purposes: Excel's
                // track-changes equivalent (Workbook.HighlightChangesOnScreen /
                // shared-workbook change tracking) is more limited than
                // Word's TrackRevisions and is out of scope for this task, so
                // there is deliberately no COM call wired up for it here.
                bool isMutating = !AlwaysAllowedTools.Contains(name);
                if (mode == EditingMode.ReadOnly && isMutating)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
                }
                if (mode == EditingMode.CommentOnly && isMutating)
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

        // Known limitation (PP-1 Task 5 Step 5): resolves the ACTIVE workbook/
        // sheet right now, not necessarily the one whose pane initiated this
        // tool call - see WordTools.cs's ActiveDoc for the identical
        // rationale and the same out-of-scope decision.
        private static Excel.Worksheet Sheet(JsonElement input)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            if (input.ValueKind == JsonValueKind.Object && input.TryGetProperty("sheet", out var s) && s.ValueKind == JsonValueKind.String)
            {
                return (Excel.Worksheet)app.ActiveWorkbook.Sheets[s.GetString()];
            }
            return (Excel.Worksheet)app.ActiveSheet;
        }

    }
}

