using System;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        // undo_last_action / redo_last_action.
        //
        // IMPORTANT correction vs. the initial plan for this feature, found by
        // actually reflecting against the exact 15.0.0.0 PIA this project
        // references (ExcelAiAddIn.csproj) rather than assuming: Excel's
        // Application.Undo() is real, but it is `void` - there is no bool
        // signal telling the caller whether anything was actually undone.
        // Worse, Excel's object model has NO Application.Redo() method AT
        // ALL (confirmed absent by reflection) - Excel has never exposed one;
        // Ctrl+Y is UI-only. So a plain COM-method implementation can't give
        // an honest "nothing to undo/redo" message, and can't implement redo
        // at all.
        //
        // Both tools instead dispatch through Application.CommandBars'
        // ExecuteMso(idMso)/GetEnabledMso(idMso) - the generic Office
        // ribbon-command API (Microsoft.Office.Core.CommandBars, from the
        // "Office" PIA reference already in this project) that the real
        // Undo/Redo buttons and Ctrl+Z/Ctrl+Y dispatch through internally.
        // GetEnabledMso is checked FIRST, so the result is honest instead of
        // always claiming success - important since Excel is well known to
        // clear its undo stack the instant automation touches the object
        // model, making this inherently best-effort on Excel specifically.
        private static ToolResult UndoLastAction()
        {
            return ExecuteRibbonUndoRedo("Undo", "Undid the last action.",
                "Nothing to undo (Excel's undo history may have been cleared by automation - this is best-effort on Excel).",
                "undo_last_action");
        }

        private static ToolResult RedoLastAction()
        {
            return ExecuteRibbonUndoRedo("Redo", "Redid the last undone action.",
                "Nothing to redo (Excel's undo history may have been cleared by automation - this is best-effort on Excel).",
                "redo_last_action");
        }

        private static ToolResult ExecuteRibbonUndoRedo(string idMso, string successMessage, string nothingMessage, string toolName)
        {
            try
            {
                Microsoft.Office.Core.CommandBars bars = Globals.ThisAddIn.Application.CommandBars;
                if (!bars.GetEnabledMso(idMso))
                {
                    return new ToolResult { Output = nothingMessage, Summary = toolName };
                }
                bars.ExecuteMso(idMso);
                return new ToolResult { Output = successMessage, Mutated = true, Summary = toolName };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Output = "Excel could not " + (idMso == "Undo" ? "undo" : "redo") + ": " + ex.Message,
                    IsError = true,
                    Summary = toolName,
                };
            }
        }
    }
}
