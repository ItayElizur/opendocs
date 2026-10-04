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
        // Per-document editing mode, keyed by TaskPaneHost.GetChatId() - see WordTools.cs for the identical pattern.
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

        // Shape names: OfficeAi.Shared.ShapeTypes; mirrors entry.ts's shapeType enum (PP-16) - edit both together.
        // Chart types: OfficeAi.Shared.ChartTypes, shared with Word and PowerPoint (PP-15).

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            try
            {
                EditingMode mode = ModeFor(docKey);
                // Comment Only blocks all mutating tools (no add_comment-equivalent tool yet); Track Changes
                // currently behaves like Full Autonomy - see ExcelTools.cs.md for why.
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

        // Known limitation (PP-1 Task 5 Step 5): resolves the ACTIVE workbook/sheet, not necessarily the one
        // whose pane initiated this call - see WordTools.cs's ActiveDoc for the identical rationale.
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

