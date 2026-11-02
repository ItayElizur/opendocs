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
            bool saveAs = ToolArgs.IsTrue(input, "saveAs");
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;

            if (!saveAs && !string.IsNullOrEmpty(wb.Path))
            {
                wb.Save();
                return new ToolResult { Output = "Saved " + wb.FullName + ".", Summary = "save" };
            }

            // Not Dialogs[xlDialogSaveAs]: with that legacy dialog the mouse pointer vanished when the command was sent
            // while the pointer was over the chat pane. GetSaveAsFilename only picks a name; SaveAs does the save.
            // See ExcelTools.File.cs.md.
            object picked = app.GetSaveAsFilename(wb.Name, SaveAsFilter, 1, "Save As", Type.Missing);
            string path = picked as string;
            if (string.IsNullOrEmpty(path))
            {
                return new ToolResult { Output = "The user cancelled the Save As dialog; nothing was saved.", Summary = "save" };
            }
            wb.SaveAs(path, FileFormatFor(path), Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return new ToolResult { Output = "Saved as " + app.ActiveWorkbook.FullName + ".", Summary = "save" };
        }

        private const string SaveAsFilter =
            "Excel Workbook (*.xlsx),*.xlsx,Excel Macro-Enabled Workbook (*.xlsm),*.xlsm,Excel Binary Workbook (*.xlsb),*.xlsb,"
            + "Excel 97-2003 Workbook (*.xls),*.xls,CSV (*.csv),*.csv";

        // SaveAs needs the format spelled out or it keeps the workbook's current one regardless of the typed extension.
        private static Excel.XlFileFormat FileFormatFor(string path)
        {
            switch (System.IO.Path.GetExtension(path).ToLowerInvariant())
            {
                case ".xlsm": return Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                case ".xlsb": return Excel.XlFileFormat.xlExcel12;
                case ".xls": return Excel.XlFileFormat.xlExcel8;
                case ".csv": return Excel.XlFileFormat.xlCSV;
                default: return Excel.XlFileFormat.xlOpenXMLWorkbook;
            }
        }
    }
}
