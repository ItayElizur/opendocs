using System;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // save: in-place Save(), or the modal Save As dialog when saveAs is set or the document
        // has never been saved. See WordTools.File.cs.md.
        private static ToolResult SaveDocument(JsonElement input)
        {
            bool saveAs = input.ValueKind == JsonValueKind.Object
                && input.TryGetProperty("saveAs", out var e) && e.ValueKind == JsonValueKind.True;
            Word.Document doc = ActiveDoc;

            if (!saveAs && !string.IsNullOrEmpty(doc.Path))
            {
                doc.Save();
                return new ToolResult { Output = "Saved " + doc.FullName + ".", Summary = "save" };
            }

            // Dialog.Show() returns -1 when the user confirmed (Word has already saved) and 0 on cancel.
            int result = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogFileSaveAs].Show();
            if (result == 0)
            {
                return new ToolResult { Output = "The user cancelled the Save As dialog; nothing was saved.", Summary = "save" };
            }
            return new ToolResult { Output = "Saved as " + ActiveDoc.FullName + ".", Summary = "save" };
        }
    }
}
