using System;
using System.Text.Json;
using OfficeAi.Shared;
using Office = Microsoft.Office.Core;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // save: the flow lives in OfficeAi.Shared.SaveTool; this is PowerPoint's two steps. See PowerPointTools.File.cs.md.
        private static ToolResult SavePresentation(JsonElement input)
        {
            var pres = ActivePresentation;
            return SaveTool.Run(input, pres.Path,
                () => { pres.Save(); return pres.FullName; },
                () =>
                {
                    // PowerPoint has no Dialogs collection. FileDialog.Show() only picks the file (-1 confirmed,
                    // 0 cancelled); Execute() performs the actual SaveAs with the chosen name and format.
                    Office.FileDialog dlg = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogSaveAs];
                    if (dlg.Show() == 0) return null;
                    dlg.Execute();
                    return ActivePresentation.FullName;
                });
        }
    }
}
