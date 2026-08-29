using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public partial class ThisAddIn
    {
        private sealed class PaneEntry
        {
            public CustomTaskPane Pane;
            public TaskPaneHost Control;
        }

        // Keyed by the window's Hwnd rather than the Excel.Window RCW: reference
        // equality on an RCW is not reliable across separate COM calls, while
        // Hwnd is a stable int, unique per top-level document window.
        private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();

        // Guards against reentrancy into EnsurePaneFor for the same hwnd -
        // confirmed repro (single document open): CustomTaskPanes.Add can
        // pump the Windows message queue internally, which lets a nested
        // WindowActivate for the window already being set up reenter this
        // method before the outer call has returned and written
        // _panes[hwnd]. Without this guard that constructs a SECOND
        // TaskPaneHost/WebViewBridgeHost for the one window, and both race to
        // create a CoreWebView2Environment against the identical user-data
        // folder - which WebView2 rejects with "the group or resource is not
        // in the correct state" (HRESULT 0x8007139F).
        private readonly HashSet<int> _paneCreationInProgress = new HashSet<int>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;

            // The startup window, as today - every subsequently-opened window
            // gets its own pane via Application_WindowActivate below. Guarded
            // (unlike every other EnsurePaneFor call site, all of which are
            // already wrapped) because Excel can start on its own "Start
            // Screen" template chooser rather than a real workbook - a state
            // ActiveWindow may not represent as a normal, fully-formed
            // Excel.Window (confirmed repro: this call, unguarded, left the
            // add-in showing a blank/gray pane). If that happens here, no
            // pane is created for the Start Screen at all - the first real
            // workbook (Ctrl+N, File > Open, etc.) still gets a working pane
            // via Application_WindowActivate regardless.
            try
            {
                Excel.Window active = this.Application.ActiveWindow;
                if (active != null) EnsurePaneFor(active);
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WindowActivate -= Application_WindowActivate;
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            this.Application.SheetSelectionChange -= Application_SheetSelectionChange;
        }

        // Lazy: only reachable from WindowActivate, TogglePane, and the single
        // startup call above - a workbook that is open but whose window has
        // never been activated pays no WebView2 cost.
        private PaneEntry EnsurePaneFor(Excel.Window window)
        {
            int hwnd = window.Hwnd;
            PaneEntry existing;
            if (_panes.TryGetValue(hwnd, out existing)) return existing;

            // HashSet<T>.Add returns false if hwnd was already present - a
            // reentrant call for the same window bails out here instead of
            // constructing a second pane. See _paneCreationInProgress's
            // declaration for why this is needed.
            if (!_paneCreationInProgress.Add(hwnd)) return null;
            try
            {
                TaskPaneHost control = new TaskPaneHost((Excel.Workbook)window.Parent, hwnd);
                CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Airchat Office", window);
                pane.Width = 420;
                pane.Visible = true;

                PaneEntry entry = new PaneEntry { Pane = pane, Control = control };
                control.RequestPaneWidth += width => ApplyPaneWidth(pane, width);
                _panes[hwnd] = entry;
                return entry;
            }
            finally
            {
                _paneCreationInProgress.Remove(hwnd);
            }
        }

        private static void ApplyPaneWidth(CustomTaskPane pane, int width)
        {
            try
            {
                if (pane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft ||
                    pane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight)
                {
                    pane.Width = width;
                }
            }
            catch
            {
                // Resizing is best-effort - never let a transient Office
                // COM exception (e.g. pane docked top/bottom) propagate
                // out and permanently reveal the debug status label via
                // WebViewBridgeHost's generic error-status path.
            }
        }

        // WindowActivate is the single hook covering every path that produces
        // a window needing a pane: File > Open, File > New, a file
        // double-clicked while the app runs, and View > New Window - each
        // newly-created window fires this as it becomes active.
        private void Application_WindowActivate(Excel.Workbook wb, Excel.Window window)
        {
            try { EnsurePaneFor(window); }
            catch { /* pane creation is best-effort; never break the add-in connection */ }
        }

        // Two-pass shape (collect hwnds, then mutate) avoids mutating _panes
        // while enumerating a COM collection that may itself change.
        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            try
            {
                var toRemove = new List<int>();
                foreach (Excel.Window w in wb.Windows) toRemove.Add(w.Hwnd);
                foreach (int hwnd in toRemove)
                {
                    PaneEntry entry;
                    if (!_panes.TryGetValue(hwnd, out entry)) continue;
                    // FT-1 Task 7b Step 2: one last GetChatId() check before
                    // the pane goes away - covers "save, then immediately
                    // close" (no separate after-save event exists to hook).
                    entry.Control.FlushChatIdMigration();
                    _panes.Remove(hwnd);
                    entry.Pane.Visible = false;
                    this.CustomTaskPanes.Remove(entry.Pane);
                    entry.Control.Dispose();
                }
            }
            catch { }
        }

        // FT-2 Task 2: routed to the pane owning the active window.
        // SheetSelectionChange hands over only the sheet and range, not a
        // window/document reference, so ActiveWindow is how this resolves
        // which pane it belongs to - the same pattern WordAiAddIn's
        // Application_WindowSelectionChange uses.
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            try
            {
                Excel.Window window = this.Application.ActiveWindow;
                if (window == null) return;
                PaneEntry entry;
                if (_panes.TryGetValue(window.Hwnd, out entry))
                {
                    entry.Control.OnSelectionChanged(Sh as Excel.Worksheet, Target);
                }
            }
            catch
            {
                // Selection-change notifications are best-effort; never let
                // one crash out of a COM event sink and kill the add-in
                // connection.
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        // Toggles the ACTIVE window's pane - routing through EnsurePaneFor
        // means the button also recovers a window that somehow never got a
        // pane, instead of no-opping.
        public void TogglePane()
        {
            try
            {
                PaneEntry entry = EnsurePaneFor(this.Application.ActiveWindow);
                if (entry != null) entry.Pane.Visible = !entry.Pane.Visible;
            }
            catch { }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
