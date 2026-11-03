using System;
using System.Text.Json;
using OfficeAi.Shared;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        // save: the flow lives in OfficeAi.Shared.SaveTool. See ExcelTools.File.cs.md.
        private static ToolResult SaveWorkbook(JsonElement input)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            return SaveTool.Run(input, wb.Path, () => wb.Save(), ShowSaveAsDialog, () => Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
        }

        // Dialog.Show() returns true when the user confirmed (Excel has already saved) and false on cancel.
        private static bool ShowSaveAsDialog()
        {
            // Sent from the chat pane the pointer can already be invisible; fix that before the dialog opens.
            HiddenPointer.Reset();
            return Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
        }
    }
}
