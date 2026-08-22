using System.Windows.Forms;
using OfficeAi.Shared;

namespace WordAiAddIn
{
    public partial class TaskPaneHost : UserControl
    {
        private readonly Label _status;
        private readonly WebViewBridgeHost _bridge;

        public TaskPaneHost()
        {
            _status = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "WebView2: initializing...",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            };
            Controls.Add(_status);

            _bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s);
        }
    }
}
