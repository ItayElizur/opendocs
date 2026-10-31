using System;
using System.Text.Json;
using OfficeAi.Shared;
using Office = Microsoft.Office.Core;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // save: in-place Save(), or the Save As file dialog when saveAs is set or the presentation
        // has never been saved. See PowerPointTools.File.cs.md.
        private static ToolResult SavePresentation(JsonElement input)
        {
            bool saveAs = input.ValueKind == JsonValueKind.Object
                && input.TryGetProperty("saveAs", out var e) && e.ValueKind == JsonValueKind.True;
            var pres = ActivePresentation;

            if (!saveAs && !string.IsNullOrEmpty(pres.Path))
            {
                pres.Save();
                return new ToolResult { Output = "Saved " + pres.FullName + ".", Summary = "save" };
            }

            // PowerPoint has no Dialogs collection. FileDialog.Show() only picks the file (-1 confirmed,
            // 0 cancelled); Execute() performs the actual SaveAs with the chosen name and format.
            Office.FileDialog dlg = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogSaveAs];
            if (dlg.Show() == 0)
            {
                return new ToolResult { Output = "The user cancelled the Save As dialog; nothing was saved.", Summary = "save" };
            }
            dlg.Execute();
            return new ToolResult { Output = "Saved as " + ActivePresentation.FullName + ".", Summary = "save" };
        }
    }
}
