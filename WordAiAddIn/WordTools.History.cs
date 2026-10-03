using System;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // Calls Document.Undo() directly - Application has no such method, see WordTools.History.cs.md.
        private static ToolResult UndoLastAction()
        {
            try
            {
                Word.Document doc = ActiveDoc;
                bool didUndo = doc.Undo();
                return new ToolResult
                {
                    Output = didUndo ? "Undid the last action." : "Nothing to undo.",
                    Mutated = didUndo,
                    Summary = "undo_last_action",
                };
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = "Could not undo: " + ex.Message, IsError = true, Summary = "undo_last_action" };
            }
        }

        // Calls Document.Redo() directly - Application has no such method, see WordTools.History.cs.md.
        private static ToolResult RedoLastAction()
        {
            try
            {
                Word.Document doc = ActiveDoc;
                bool didRedo = doc.Redo();
                return new ToolResult
                {
                    Output = didRedo ? "Redid the last undone action." : "Nothing to redo.",
                    Mutated = didRedo,
                    Summary = "redo_last_action",
                };
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = "Could not redo: " + ex.Message, IsError = true, Summary = "redo_last_action" };
            }
        }
    }
}
