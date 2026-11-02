using System;
using System.Text.Json;
using OfficeAi.Shared;
using Office = Microsoft.Office.Core;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // save: the flow lives in OfficeAi.Shared.SaveTool. See PowerPointTools.File.cs.md.
        private static ToolResult SavePresentation(JsonElement input)
        {
            var pres = ActivePresentation;
            return SaveTool.Run(input, pres.Path, () => pres.Save(), ShowSaveAsDialog, () => ActivePresentation.FullName);
        }

        // PowerPoint's Application.Dialogs is an untyped object with no built-in dialog list (Word and Excel have
        // WdWordDialog / XlBuiltInDialog), so use FileDialog. Show() only picks the file (-1 confirmed, 0 cancelled);
        // Execute() performs the actual SaveAs with the chosen name and format.
        private static bool ShowSaveAsDialog()
        {
            Office.FileDialog dlg = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogSaveAs];
            if (dlg.Show() == 0) return false;
            dlg.Execute();
            return true;
        }
    }
}
