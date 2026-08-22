using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public partial class TaskPaneHost : UserControl
    {
        private readonly Label _status;
        private readonly WebViewBridgeHost _bridge;
        private readonly string _chatId;

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

            Word.Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            // An unsaved document has no on-disk Path; ActiveDocument.FullName
            // falls back to its temp Name (e.g. "Document1") in that case,
            // which is not a stable key across sessions, so use a per-process
            // fallback id instead (mirrors genoffice's tempChatId concept).
            _chatId = string.IsNullOrEmpty(activeDoc.Path)
                ? "unsaved-" + Process.GetCurrentProcess().Id
                : ChatStore.ChatIdForFile(activeDoc.FullName);

            _bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s, OnOtherMessage);
        }

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider("WordAiAddIn", _chatId);
                    _bridge.PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage("WordAiAddIn", _chatId, role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider("WordAiAddIn", _chatId);
                    break;
            }
        }
    }
}
