using System;
using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;
using OfficeAi.Shared;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAiAddIn
{
    public partial class TaskPaneHost : UserControl
    {
        private readonly Label _status;
        private readonly WebViewBridgeHost _bridge;
        private string _chatId;

        public event Action<int> RequestPaneWidth;

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
            // Globals.ThisAddIn.Application.ActiveWorkbook eagerly, at the
            // exact moment this constructor runs inside ThisAddIn_Startup
            // (i.e. before CustomTaskPanes.Add() has even returned), hits a
            // COM timing issue in Excel's own startup sequence and silently
            // kills the whole add-in connection (VSTO never connects it - no
            // exception, no resiliency-disabled entry, just Connect=False
            // forever). Confirmed by direct repro (see WordAiAddIn's
            // TaskPaneHost.cs). _chatId is instead computed lazily on first
            // actual use, in GetChatId() below, by which point the task pane
            // is visible and the user has triggered a message - so
            // ActiveWorkbook is guaranteed settled.
            _bridge = new WebViewBridgeHost(this, ExcelTools.Execute, "ExcelAiAddIn", UpdateStatus, OnOtherMessage);
        }

        private void UpdateStatus(string s)
        {
            _status.Text = s;
            _status.Visible = s != "ready";
        }

        private string GetChatId()
        {
            if (_chatId != null) return _chatId;

            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            // An unsaved workbook has no on-disk Path; ActiveWorkbook.FullName
            // falls back to its temp Name (e.g. "Book1") in that case, which
            // is not a stable key across sessions, so use a per-process
            // fallback id instead (mirrors genoffice's tempChatId concept).
            _chatId = string.IsNullOrEmpty(activeWorkbook.Path)
                ? "unsaved-" + Process.GetCurrentProcess().Id
                : ChatStore.ChatIdForFile(activeWorkbook.FullName);
            return _chatId;
        }

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider("ExcelAiAddIn", GetChatId());
                    _bridge.PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage("ExcelAiAddIn", GetChatId(), role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider("ExcelAiAddIn", GetChatId());
                    break;
                case "set-mode":
                    string mode = root.GetProperty("mode").GetString();
                    switch (mode)
                    {
                        case "readOnly": ExcelTools.Mode = EditingMode.ReadOnly; break;
                        case "commentOnly": ExcelTools.Mode = EditingMode.CommentOnly; break;
                        case "trackChanges": ExcelTools.Mode = EditingMode.TrackChanges; break;
                        case "fullAutonomy": ExcelTools.Mode = EditingMode.FullAutonomy; break;
                    }
                    break;
                case "collapse-pane":
                    RequestPaneWidth?.Invoke(34);
                    break;
                case "expand-pane":
                    RequestPaneWidth?.Invoke(420);
                    break;
            }
        }
    }
}
