using System.Diagnostics;
using System.Text.Json;
using System.Threading.Tasks;
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
        // read) - see TaskPaneHost.cs.md for the confirmed repro (kept in sync
        // with WordAiAddIn/TaskPaneHost.cs.md's identical note).
        public TaskPaneHost(Excel.Workbook workbook, int hwnd) : base("ExcelAiAddIn")
        {
            _workbook = workbook;
            _hwnd = hwnd;
        }

        protected override Task<ToolResult> ExecuteTool(string name, JsonElement input)
        {
            return Task.FromResult(ExcelTools.Execute(GetChatId(), name, input));
        }

        protected override string GetChatId()
        {
            // A saved id is final; an "unsaved-" id re-checks the workbook's Path on every call so saving
            // migrates chat history/doc settings onto the real per-file id (FT-1 Task 7b) - see
            // TaskPaneHost.cs.md (kept in sync with WordAiAddIn/TaskPaneHost.cs.md's identical rationale).
            if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

            if (string.IsNullOrEmpty(_workbook.Path))
            {
                // Window handle folded in: an unsaved workbook's FullName falls back to a temp Name (e.g.
                // "Book1"), and with multiple panes possible in one process, "unsaved-<pid>" alone could
                // collide across two different unsaved workbooks - see TaskPaneHost.cs.md.
                return _chatId ?? (_chatId = "unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd);
            }

            string saved = ChatStore.ChatIdForFile(_workbook.FullName);
            if (_chatId != null)
            {
                ChatStore.Migrate("ExcelAiAddIn", _chatId, saved);
                DocSettingsStore.Migrate("ExcelAiAddIn", _chatId, saved);
            }
            // Save As after this point does NOT re-key. See TaskPaneHost.cs.md.
            return _chatId = saved;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            ExcelTools.SetMode(GetChatId(), mode);
        }

        protected override int GetOfficeUiLanguageId()
        {
            return Globals.ThisAddIn.GetOfficeUiLanguageId();
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

        // Called from ThisAddIn's SheetSelectionChange handler via the active window's hwnd; debounced
        // through PaneHostBase.PostSelection since SheetSelectionChange fires on every arrow-key press.
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

            // Reports the UsedRange-intersected extent alongside the literal one for whole-column/row or
            // large selections - a bare "B1:B1048576" exceeds read_range's cap and isn't useful to show.
            // See TaskPaneHost.cs.md.
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
