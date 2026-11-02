using System;
using System.Runtime.InteropServices;
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
            ResetHiddenPointer();
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

        // When the command is sent with the mouse over the chat pane, the pointer is already invisible on Excel's UI
        // thread by the time Save As opens (null cursor, sometimes also a negative ShowCursor count), and stays that way
        // until it re-enters the pane. Reset both just before the dialog. See ExcelTools.File.cs.md.
        [StructLayout(LayoutKind.Sequential)]
        private struct CursorInfo { public int cbSize; public int flags; public IntPtr hCursor; public int x; public int y; }
        [DllImport("user32.dll")] private static extern bool GetCursorInfo(ref CursorInfo ci);
        [DllImport("user32.dll")] private static extern int ShowCursor(bool show);
        [DllImport("user32.dll")] private static extern IntPtr SetCursor(IntPtr cursor);
        [DllImport("user32.dll")] private static extern IntPtr LoadCursor(IntPtr instance, int name);

        private static int ShowCursorCount()
        {
            ShowCursor(true); // the API only reports the counter when changing it, so bump and undo
            return ShowCursor(false);
        }

        private static void ResetHiddenPointer()
        {
            try
            {
                var ci = new CursorInfo { cbSize = Marshal.SizeOf(typeof(CursorInfo)) };
                bool ok = GetCursorInfo(ref ci);
                int count = ShowCursorCount();
                int steps = 0;
                while (count < 0 && steps < 20) { ShowCursor(true); count = ShowCursorCount(); steps++; }
                bool nullCursor = ok && ci.hCursor == IntPtr.Zero;
                if (nullCursor) SetCursor(LoadCursor(IntPtr.Zero, 32512 /* IDC_ARROW */));
                if (steps > 0 || nullCursor)
                    DebugLog.Write("save: reset hidden pointer (showCursor steps=" + steps + ", null cursor replaced=" + nullCursor + ")");
            }
            catch (Exception ex)
            {
                DebugLog.Write("save: pointer reset failed: " + ex.Message);
            }
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
