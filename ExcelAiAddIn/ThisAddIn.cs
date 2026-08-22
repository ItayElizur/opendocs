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
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
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
