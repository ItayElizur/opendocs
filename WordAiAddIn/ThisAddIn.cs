using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public partial class ThisAddIn
    {
        private sealed class PaneEntry
        {
            public CustomTaskPane Pane;
            public TaskPaneHost Control;
        }

        // Keyed by the window's Hwnd rather than the Word.Window RCW: reference
        // equality on an RCW is not reliable across separate COM calls, while
        // Hwnd is a stable int, unique per top-level document window.
        private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();

        // Guards against reentrancy into EnsurePaneFor for the same hwnd - prevents
        // a second TaskPaneHost/WebViewBridgeHost racing to open a WebView2
        // environment on the same user-data folder. See ThisAddIn.cs.md.
        private readonly HashSet<int> _paneCreationInProgress = new HashSet<int>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // Guarded (unlike other EnsurePaneFor call sites) because Word can start
            // on its "Start Screen" template chooser, where ActiveWindow may not be
            // a normal, fully-formed Word.Window. See ThisAddIn.cs.md.
            try
            {
                Word.Window active = this.Application.ActiveWindow;
                if (active != null) EnsurePaneFor(active);
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WindowActivate -= Application_WindowActivate;
            this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        // Single COM call site for the UI language id; Ribbon.cs/TaskPaneHost.cs
        // delegate here instead of re-issuing the call. Degrades to 0 ("not Hebrew")
        // on failure rather than risk breaking pane creation. See ThisAddIn.cs.md.
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

        // The task pane's native title bar is a UI surface separate from the
        // ribbon/WebView2 content, so it needs its own call site.
        private string PaneTitle()
        {
            return OfficeLanguage.ResolveBrandName(GetOfficeUiLanguageId());
        }

        // Lazy: only reachable from WindowActivate, TogglePane, and the single
        // startup call above - a document whose window is never activated pays
        // no WebView2 cost.
        private PaneEntry EnsurePaneFor(Word.Window window)
        {
            int hwnd = window.Hwnd;
            PaneEntry existing;
            if (_panes.TryGetValue(hwnd, out existing)) return existing;

            // Add returns false if hwnd is already present, so a reentrant call for
            // the same window bails out here (see _paneCreationInProgress above).
            if (!_paneCreationInProgress.Add(hwnd)) return null;
            try
            {
                TaskPaneHost control = new TaskPaneHost(window.Document, hwnd);
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
                // Best-effort - a transient COM exception (e.g. pane docked top/bottom)
                // must not propagate and reveal the debug status label.
            }
        }

        // Single hook covering every path that produces a window needing a pane:
        // File > Open, File > New, a double-clicked file, View > New Window.
        private void Application_WindowActivate(Word.Document doc, Word.Window window)
        {
            try { EnsurePaneFor(window); }
            catch { /* pane creation is best-effort; never break the add-in connection */ }
        }

        // Two-pass shape (collect hwnds, then mutate) avoids mutating _panes
        // while enumerating a COM collection that may itself change.
        private void Application_DocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            try
            {
                var toRemove = new List<int>();
                foreach (Word.Window w in doc.Windows) toRemove.Add(w.Hwnd);
                foreach (int hwnd in toRemove)
                {
                    PaneEntry entry;
                    if (!_panes.TryGetValue(hwnd, out entry)) continue;
                    // One last GetChatId() check before the pane goes away - covers
                    // "save, then immediately close" (no after-save event to hook).
                    entry.Control.FlushChatIdMigration();
                    _panes.Remove(hwnd);
                    entry.Pane.Visible = false;
                    this.CustomTaskPanes.Remove(entry.Pane);
                    entry.Control.Dispose();
                }
            }
            catch { }
        }

        // Routed to the pane owning the selection's window - without this, a
        // selection in window B would push selected text into window A's
        // chat context, a silent cross-document data leak into the prompt.
        private void Application_WindowSelectionChange(Word.Selection selection)
        {
            try
            {
                PaneEntry entry;
                if (_panes.TryGetValue(selection.Document.ActiveWindow.Hwnd, out entry))
                {
                    entry.Control.OnSelectionChanged(selection);
                }
            }
            catch (Exception ex)
            {
                // Best-effort - must never crash the COM event sink. Logged (not
                // silently dropped) since a throw here could fully explain
                // "only one OnSelectionChanged logged despite several clicks". See ThisAddIn.cs.md.
                OfficeAi.Shared.DebugLog.WriteException("Application_WindowSelectionChange", ex);
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        // Toggles the active window's pane; routing through EnsurePaneFor also
        // recovers a window that somehow never got a pane, instead of no-opping.
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
