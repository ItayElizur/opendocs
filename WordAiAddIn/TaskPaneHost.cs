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
        private string _chatId;

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

            // Deliberately NOT computed here: accessing
            // Globals.ThisAddIn.Application.ActiveDocument eagerly, at the
            // exact moment this constructor runs inside ThisAddIn_Startup
            // (i.e. before CustomTaskPanes.Add() has even returned), hits a
            // COM timing issue in Word's own startup sequence and silently
            // kills the whole add-in connection (VSTO never connects it - no
            // exception, no resiliency-disabled entry, just Connect=False
            // forever). Confirmed by direct repro. _chatId is instead
            // computed lazily on first actual use, in GetChatId() below, by
            // which point the task pane is visible and the user has
            // triggered a message - so ActiveDocument is guaranteed settled.
            _bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s, OnOtherMessage);
        }

        private string GetChatId()
        {
            if (_chatId != null) return _chatId;

            Word.Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            // An unsaved document has no on-disk Path; ActiveDocument.FullName
            // falls back to its temp Name (e.g. "Document1") in that case,
            // which is not a stable key across sessions, so use a per-process
            // fallback id instead (mirrors genoffice's tempChatId concept).
            _chatId = string.IsNullOrEmpty(activeDoc.Path)
                ? "unsaved-" + Process.GetCurrentProcess().Id
                : ChatStore.ChatIdForFile(activeDoc.FullName);
            return _chatId;
        }

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider("WordAiAddIn", GetChatId());
                    _bridge.PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage("WordAiAddIn", GetChatId(), role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider("WordAiAddIn", GetChatId());
                    break;
                case "set-mode":
                    string modeStr = root.GetProperty("mode").GetString();
                    switch (modeStr)
                    {
                        case "readOnly": WordTools.Mode = EditingMode.ReadOnly; break;
                        case "commentOnly": WordTools.Mode = EditingMode.CommentOnly; break;
                        case "trackChanges": WordTools.Mode = EditingMode.TrackChanges; break;
                        case "fullAutonomy": WordTools.Mode = EditingMode.FullAutonomy; break;
                    }
                    break;
            }
        }
    }
}
