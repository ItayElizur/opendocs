using System;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // save: the flow lives in OfficeAi.Shared.SaveTool; this is Word's two steps. See WordTools.File.cs.md.
        private static ToolResult SaveDocument(JsonElement input)
        {
            Word.Document doc = ActiveDoc;
            return SaveTool.Run(input, doc.Path,
                () => { doc.Save(); return doc.FullName; },
                () =>
                {
                    // Dialog.Show() returns -1 when the user confirmed (Word has already saved) and 0 on cancel.
                    int result = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogFileSaveAs].Show();
                    return result == 0 ? null : ActiveDoc.FullName;
                });
        }
    }
}
