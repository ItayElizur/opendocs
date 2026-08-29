using System.Diagnostics;
using System.Text.Json;
using OfficeAi.Shared;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public partial class TaskPaneHost : PaneHostBase
    {
        private readonly Excel.Workbook _workbook;
        private readonly int _hwnd;
        private string _chatId;

        // Deliberately does NOT dereference _workbook here (no .Path/.FullName
        // read) - see WordAiAddIn/TaskPaneHost.cs's identical comment for the
        // confirmed repro of why an eager read at construction time silently
        // kills the whole add-in connection.
        public TaskPaneHost(Excel.Workbook workbook, int hwnd) : base("ExcelAiAddIn")
        {
            _workbook = workbook;
            _hwnd = hwnd;
        }

        protected override ToolResult ExecuteTool(string name, JsonElement input)
        {
            return ExcelTools.Execute(GetChatId(), name, input);
        }

        protected override string GetChatId()
        {
            // A saved id is final - never re-checked again. An "unsaved-" id
            // is provisional: re-check the workbook's Path on every call, so
            // the first use after the user saves migrates chat history and
            // doc settings onto the real per-file id (FT-1 Task 7b). The Path
            // read is one cheap COM property on operations (load-history,
            // append-message, etc.) that are already doing file I/O.
            if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

            if (string.IsNullOrEmpty(_workbook.Path))
            {
                // An unsaved workbook has no on-disk Path; Workbook.FullName
                // falls back to its temp Name (e.g. "Book1") in that case,
                // which is not a stable key across sessions - and with
                // multiple panes now possible in one process, "unsaved-<pid>"
                // alone would collide across two different unsaved workbooks,
                // so the window handle is folded in too.
                return _chatId ?? (_chatId = "unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd);
            }

            string saved = ChatStore.ChatIdForFile(_workbook.FullName);
            if (_chatId != null)
            {
                ChatStore.Migrate("ExcelAiAddIn", _chatId, saved);
                DocSettingsStore.Migrate("ExcelAiAddIn", _chatId, saved);
            }
            // Save As after this point does NOT re-key - see WordAiAddIn/
            // TaskPaneHost.cs's identical comment for the rationale.
            return _chatId = saved;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            ExcelTools.SetMode(GetChatId(), mode);
        }

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }

        // FT-2 Task 2: called from ThisAddIn's SheetSelectionChange handler,
        // routed here via the active window's hwnd. Debounced through
        // PaneHostBase.PostSelection (Task 1) - SheetSelectionChange fires on
        // every arrow-key press and would flood the WebView2 bridge otherwise.
        public void OnSelectionChanged(Excel.Worksheet sheet, Excel.Range target)
        {
            // Task 2 Step 3: a chart/shape selection is not a Range at all -
            // Sh/Target may not behave as a normal range then. Guard rather
            // than let a COM exception escape a COM event sink.
            if (sheet == null || target == null)
            {
                PostSelection(new { kind = "selection-changed", app = "excel", hasSelection = false }, "excel:none");
                return;
            }

            string address = target.Address[false, false];
            int areaCount = target.Areas.Count;
            bool multi = areaCount > 1;
            long cellCount = target.CountLarge; // NOT Count - a whole-sheet selection (~17B cells) overflows Int32
            int rows = target.Rows.Count;
            int cols = target.Columns.Count;
            int firstRow = target.Row;
            string firstCol = ColumnLetter(target.Column);
            // Selecting a whole column spans every row (and vice versa) - this
            // is how a "column B selected" click is distinguished from an
            // ordinary drag-selection that merely happens to be tall.
            bool entireColumns = target.Rows.Count == sheet.Rows.Count;
            bool entireRows = target.Columns.Count == sheet.Columns.Count;

            // Task 2b: report the effective (UsedRange-intersected) extent
            // alongside the literal one for whole-column/row selections and
            // any large selection - a bare "B1:B1048576" is both useless to
            // show the user and something the model would try to read in
            // full, hitting read_range's 2000-cell cap. Only pay for
            // UsedRange when it can matter; an ordinary drag-selection never
            // needs it.
            string effectiveAddress = null;
            long effectiveCellCount = 0;
            int effectiveRows = 0;
            int effectiveCols = 0;
            if (entireColumns || entireRows || cellCount > 10000)
            {
                Excel.Range effective = Globals.ThisAddIn.Application.Intersect(target, sheet.UsedRange);
                if (effective != null)
                {
                    effectiveAddress = effective.Address[false, false];
                    effectiveCellCount = effective.CountLarge;
                    effectiveRows = effective.Rows.Count;
                    effectiveCols = effective.Columns.Count;
                }
            }

            string signature = "excel:" + sheet.Name + "!" + address;
            PostSelection(new
            {
                kind = "selection-changed",
                app = "excel",
                hasSelection = true,
                sheet = sheet.Name,
                address,
                cellCount,
                rows,
                cols,
                firstRow,
                firstCol,
                entireColumns,
                entireRows,
                multi,
                areaCount,
                effectiveAddress,
                effectiveCellCount,
                effectiveRows,
                effectiveCols,
            }, signature);
        }
    }
}
