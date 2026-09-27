using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using OfficeAi.Shared;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAiAddIn
{
    public partial class ThisAddIn
    {
        private sealed class PaneEntry
        {
            public CustomTaskPane Pane;
            public TaskPaneHost Control;
        }

        // Keyed by the window's HWND (PowerPoint's PIA spells the property
        // all-caps, unlike Word/Excel's Hwnd) rather than the DocumentWindow
        // RCW: reference equality on an RCW is not reliable across separate
        // COM calls, while HWND is a stable int, unique per top-level
        // document window.
        private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();

        // Guards against reentrancy into EnsurePaneFor for the same hwnd -
        // confirmed repro (single presentation open): CustomTaskPanes.Add can
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
            this.Application.PresentationClose += Application_PresentationClose;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // The startup window, as today - every subsequently-opened window
            // gets its own pane via Application_WindowActivate below. Guarded
            // (unlike every other EnsurePaneFor call site, all of which are
            // already wrapped) because PowerPoint can start on its own "Start
            // Screen" template chooser rather than a real presentation - a
            // state ActiveWindow may not represent as a normal, fully-formed
            // PowerPoint.DocumentWindow (confirmed repro: this call,
            // unguarded, left the add-in showing a blank/gray pane). If that
            // happens here, no pane is created for the Start Screen at all -
            // the first real presentation (Ctrl+N, File > Open, etc.) still
            // gets a working pane via Application_WindowActivate regardless.
            try
            {
                PowerPoint.DocumentWindow active = this.Application.ActiveWindow;
                if (active != null) EnsurePaneFor(active);
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WindowActivate -= Application_WindowActivate;
            this.Application.PresentationClose -= Application_PresentationClose;
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        // The one real COM call for Office's UI display language in this
        // app - Ribbon.cs and TaskPaneHost.cs each need their own copy of
        // this (their base classes' GetOfficeUiLanguageId hooks are
        // abstract, since neither shared assembly can see this app's own
        // Globals class), but delegate here rather than re-issuing the COM
        // call themselves, so there is exactly one place per app that can
        // fail and exactly one place that guards against it. A theme-
        // detection bug must never break pane creation (OfficeTheme.cs's own
        // stated posture) - same reasoning applies here: if
        // LanguageSettings throws (an unusual COM/host state), degrade to
        // the code that already means "not Hebrew" rather than letting the
        // ribbon render a blank label or the "load-language" bridge message
        // die silently with no reply ever sent (that one-shot message has no
        // retry - see PaneHostBase's "load-language" case).
        public int GetOfficeUiLanguageId()
        {
            try
            {
                return this.Application.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI];
            }
            catch
            {
                return 0;
            }
        }

        // Read once per pane creation - the CustomTaskPane's native title bar
        // is a third UI surface, separate from the ribbon and the WebView2
        // content, so it needs its own call site even though all three now
        // share GetOfficeUiLanguageId().
        private string PaneTitle()
        {
            return OfficeLanguage.ResolveBrandName(GetOfficeUiLanguageId());
        }

        // Lazy: only reachable from WindowActivate, TogglePane, and the single
        // startup call above - a presentation that is open but whose window
        // has never been activated pays no WebView2 cost.
        private PaneEntry EnsurePaneFor(PowerPoint.DocumentWindow window)
        {
            int hwnd = window.HWND;
            PaneEntry existing;
            if (_panes.TryGetValue(hwnd, out existing)) return existing;

            // HashSet<T>.Add returns false if hwnd was already present - a
            // reentrant call for the same window bails out here instead of
            // constructing a second pane. See _paneCreationInProgress's
            // declaration for why this is needed.
            if (!_paneCreationInProgress.Add(hwnd)) return null;
            try
            {
                TaskPaneHost control = new TaskPaneHost(window.Presentation, hwnd);
                CustomTaskPane pane = this.CustomTaskPanes.Add(control, PaneTitle(), window);
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
        // double-clicked while the app runs.
        private void Application_WindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow window)
        {
            try { EnsurePaneFor(window); }
            catch { /* pane creation is best-effort; never break the add-in connection */ }
        }

        // Two-pass shape (collect hwnds, then mutate) avoids mutating _panes
        // while enumerating a COM collection that may itself change.
        private void Application_PresentationClose(PowerPoint.Presentation pres)
        {
            try
            {
                var toRemove = new List<int>();
                foreach (PowerPoint.DocumentWindow w in pres.Windows) toRemove.Add(w.HWND);
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

        // FT-2 Task 3: routed to the pane owning the window the selection
        // happened in. Sel.Parent normally resolves to the owning
        // DocumentWindow directly; ActiveWindow is the fallback for any
        // selection shape where that cast does not hold.
        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                PowerPoint.DocumentWindow window = null;
                try { window = Sel.Parent as PowerPoint.DocumentWindow; }
                catch { /* fall through to ActiveWindow below */ }
                if (window == null) window = this.Application.ActiveWindow;
                if (window == null) return;

                PaneEntry entry;
                if (_panes.TryGetValue(window.HWND, out entry))
                {
                    entry.Control.OnSelectionChanged(Sel);
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
