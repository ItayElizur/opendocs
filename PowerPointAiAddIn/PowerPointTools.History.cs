using System;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // undo_last_action / redo_last_action. PowerPoint has no Application.Undo()/Redo(),
        // so these dispatch through the ribbon's CommandBars.ExecuteMso("Undo"/"Redo"),
        // pre-checked with GetEnabledMso so "nothing to undo/redo" is reported honestly.
        // Full rationale and history: PowerPointTools.History.cs.md.
        private static ToolResult UndoLastAction()
        {
            return ExecuteRibbonUndoRedo("Undo", "Undid the last action.", "Nothing to undo.", "undo_last_action");
        }

        private static ToolResult RedoLastAction()
        {
            return ExecuteRibbonUndoRedo("Redo", "Redid the last undone action.", "Nothing to redo.", "redo_last_action");
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
                    Output = "PowerPoint could not " + (idMso == "Undo" ? "undo" : "redo") + ": " + ex.Message,
                    IsError = true,
                    Summary = toolName,
                };
            }
        }
    }
}
