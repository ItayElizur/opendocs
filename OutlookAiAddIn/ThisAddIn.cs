using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public partial class ThisAddIn
    {
        private sealed class PaneEntry
        {
            public Outlook.Explorer Explorer;
            public CustomTaskPane Pane;
            public TaskPaneHost Control;
            // Kept so the COM event sinks can be detached on Close/Shutdown -
            // the parameterless Outlook events (Close/SelectionChange/Activate)
            // carry no sender, so each handler is a closure over its explorer.
            public Outlook.ExplorerEvents_10_CloseEventHandler OnClose;
            public Outlook.ExplorerEvents_10_SelectionChangeEventHandler OnSelectionChange;
        }

        // Outlook's Explorer has no Hwnd (unlike Word/Excel/PowerPoint windows)
        // and RCW reference-equality is unreliable across COM calls, so panes
        // are keyed by the COM identity pointer (IUnknown). A lookup miss falls
        // back to Application.ActiveExplorer().
        private readonly Dictionary<IntPtr, PaneEntry> _panes = new Dictionary<IntPtr, PaneEntry>();

        // Same reentrancy hazard the other hosts document: CustomTaskPanes.Add
        // pumps the message queue, which can re-enter EnsurePaneFor for the same
        // explorer (via its Activate hook) before _panes is written.
        private readonly HashSet<IntPtr> _paneCreationInProgress = new HashSet<IntPtr>();

        // Held in a field so the NewExplorer event sink is not garbage collected.
        private Outlook.Explorers _explorers;

        // Recovery hooks for explorers that exist at startup / arrive via
        // NewExplorer but whose pane is created lazily on first Activate.
        private readonly Dictionary<IntPtr, Outlook.ExplorerEvents_10_ActivateEventHandler> _pendingActivate =
            new Dictionary<IntPtr, Outlook.ExplorerEvents_10_ActivateEventHandler>();

        private static IntPtr IdOf(object comObject)
        {
            IntPtr p = Marshal.GetIUnknownForObject(comObject);
            Marshal.Release(p);
            return p;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Deliberately NO pane creation here. Outlook disables add-ins whose
            // median startup exceeds ~1s over 5 launches, and building a
            // WebView2 CustomTaskPane is well over that budget. Panes are
            // created on an explorer's first Activate (or the ribbon button).
            _explorers = this.Application.Explorers;
            _explorers.NewExplorer += Explorers_NewExplorer;

            try
            {
                foreach (Outlook.Explorer ex in _explorers) ArmActivate(ex);
            }
            catch { /* best-effort - a not-yet-ready explorer still recovers on its own Activate */ }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { if (_explorers != null) _explorers.NewExplorer -= Explorers_NewExplorer; }
            catch { }

            var keys = new List<IntPtr>(_panes.Keys);
            foreach (IntPtr key in keys) RemovePane(key);
            _panes.Clear();
            _pendingActivate.Clear();
            _explorers = null;
        }

        // Subscribes only the Activate recovery hook. The full per-pane event
        // wiring happens in EnsurePaneFor once the pane actually exists.
        private void ArmActivate(Outlook.Explorer explorer)
        {
            IntPtr key = IdOf(explorer);
            if (_panes.ContainsKey(key) || _pendingActivate.ContainsKey(key)) return;

            Outlook.ExplorerEvents_10_ActivateEventHandler handler = null;
            handler = () =>
            {
                try { EnsurePaneFor(explorer); }
                catch { }
            };
            _pendingActivate[key] = handler;
            ((Outlook.ExplorerEvents_10_Event)explorer).Activate += handler;
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                // A brand-new explorer may not be fully initialized yet; if the
                // immediate attempt throws, ArmActivate still recovers it on its
                // first real Activate.
                ArmActivate(explorer);
            }
            catch { }
        }

        private PaneEntry EnsurePaneFor(Outlook.Explorer explorer)
        {
            IntPtr key = IdOf(explorer);

            PaneEntry existing;
            if (_panes.TryGetValue(key, out existing)) return existing;

            if (!_paneCreationInProgress.Add(key)) return null;
            try
            {
                var control = new TaskPaneHost();
                CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Airchat Office", explorer);
                pane.Width = 420;
                pane.Visible = true;

                var entry = new PaneEntry { Explorer = explorer, Pane = pane, Control = control };
                control.RequestPaneWidth += width => ApplyPaneWidth(pane, width);

                var events = (Outlook.ExplorerEvents_10_Event)explorer;

                entry.OnSelectionChange = () =>
                {
                    try
                    {
                        PaneEntry live;
                        if (_panes.TryGetValue(key, out live))
                            live.Control.OnSelectionChanged(explorer.Selection);
                    }
                    catch (Exception ex) { DebugLog.WriteException("Explorer.SelectionChange", ex); }
                };
                entry.OnClose = () =>
                {
                    try { RemovePane(key); } catch { }
                };
                events.SelectionChange += entry.OnSelectionChange;
                events.Close += entry.OnClose;

                _panes[key] = entry;

                // The Activate recovery hook has done its job for this explorer.
                Outlook.ExplorerEvents_10_ActivateEventHandler pending;
                if (_pendingActivate.TryGetValue(key, out pending))
                {
                    try { ((Outlook.ExplorerEvents_10_Event)explorer).Activate -= pending; }
                    catch { }
                    _pendingActivate.Remove(key);
                }

                return entry;
            }
            finally
            {
                _paneCreationInProgress.Remove(key);
            }
        }

        private void RemovePane(IntPtr key)
        {
            PaneEntry entry;
            if (!_panes.TryGetValue(key, out entry)) return;
            _panes.Remove(key);

            try
            {
                var events = (Outlook.ExplorerEvents_10_Event)entry.Explorer;
                if (entry.OnSelectionChange != null) events.SelectionChange -= entry.OnSelectionChange;
                if (entry.OnClose != null) events.Close -= entry.OnClose;
            }
            catch { }

            try
            {
                entry.Pane.Visible = false;
                this.CustomTaskPanes.Remove(entry.Pane);
            }
            catch { }

            try { entry.Control.Dispose(); } catch { }
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
            catch { /* best-effort, mirrors the other hosts */ }
        }

        // Invoked by Ribbon.cs. Creates the active explorer's pane on demand
        // (so the button also recovers an explorer that never got one) then
        // toggles visibility.
        public void TogglePane()
        {
            try
            {
                Outlook.Explorer active = this.Application.ActiveExplorer();
                if (active == null) return;
                PaneEntry entry = EnsurePaneFor(active);
                if (entry != null) entry.Pane.Visible = !entry.Pane.Visible;
            }
            catch { }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
