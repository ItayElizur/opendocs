using System;
using System.Text.Json;
using OfficeAi.Shared;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        // save: in-place Save(), or the modal Save As dialog when saveAs is set or the workbook
        // has never been saved. See ExcelTools.File.cs.md.
        private static ToolResult SaveWorkbook(JsonElement input)
        {
            bool saveAs = input.ValueKind == JsonValueKind.Object
                && input.TryGetProperty("saveAs", out var e) && e.ValueKind == JsonValueKind.True;
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;

            if (!saveAs && !string.IsNullOrEmpty(wb.Path))
            {
                wb.Save();
                return new ToolResult { Output = "Saved " + wb.FullName + ".", Summary = "save" };
            }

            // Dialog.Show() returns true when the user confirmed (Excel has already saved) and false on cancel.
            bool confirmed = app.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            if (!confirmed)
            {
                return new ToolResult { Output = "The user cancelled the Save As dialog; nothing was saved.", Summary = "save" };
            }
            return new ToolResult { Output = "Saved as " + app.ActiveWorkbook.FullName + ".", Summary = "save" };
        }
    }
}
