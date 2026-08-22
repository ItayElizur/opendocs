using System;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
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

            _taskPaneControl.RequestPaneWidth += width => _taskPane.Width = width;

            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        private void Application_WindowSelectionChange(Word.Selection selection)
        {
            try
            {
                _taskPaneControl.OnSelectionChanged(selection);
            }
            catch
            {
                // Selection-change notifications are best-effort; never let one
                // crash out of a COM event sink and kill the add-in connection.
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
