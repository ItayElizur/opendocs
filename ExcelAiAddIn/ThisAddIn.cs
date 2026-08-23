using System;
using Microsoft.Office.Tools;

namespace ExcelAiAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _taskPane;
        private TaskPaneHost _taskPaneControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _taskPaneControl = new TaskPaneHost();
            _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "Airchat Office");
            _taskPane.Width = 420;
            _taskPane.Visible = true;

            _taskPaneControl.RequestPaneWidth += width =>
            {
                try
                {
                    if (_taskPane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft ||
                        _taskPane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight)
                    {
                        _taskPane.Width = width;
                    }
                }
                catch
                {
                    // Resizing is best-effort - never let a transient Office
                    // COM exception (e.g. pane docked top/bottom) propagate
                    // out and permanently reveal the debug status label via
                    // WebViewBridgeHost's generic error-status path.
                }
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        // Reopens the pane after the user closes it via its native "x" (or
        // hides it), toggled from the Home-tab ribbon button.
        public void TogglePane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
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
