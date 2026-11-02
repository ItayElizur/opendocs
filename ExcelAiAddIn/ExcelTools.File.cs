using System;
using System.Text.Json;
using OfficeAi.Shared;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        // save: the flow lives in OfficeAi.Shared.SaveTool; this is Excel's three steps. See ExcelTools.File.cs.md.
        private static ToolResult SaveWorkbook(JsonElement input)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;
            return SaveTool.Run(input, wb.Path,
                () => wb.Save(),
                () =>
                {
                    // Sent from the chat pane the pointer can already be invisible; fix that before the dialog opens.
                    HiddenPointer.Reset();
                    // Dialog.Show() returns true when the user confirmed (Excel has already saved) and false on cancel.
                    return app.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
                },
                () => app.ActiveWorkbook.FullName);
        }
    }
}
