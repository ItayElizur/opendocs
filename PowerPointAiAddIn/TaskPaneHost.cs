using System;
using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;
using OfficeAi.Shared;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAiAddIn
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
            // Globals.ThisAddIn.Application.ActivePresentation eagerly, at the
            // exact moment this constructor runs inside ThisAddIn_Startup
            // (i.e. before CustomTaskPanes.Add() has even returned), hits a
            // COM timing issue in PowerPoint's own startup sequence and silently
            // kills the whole add-in connection (VSTO never connects it - no
            // exception, no resiliency-disabled entry, just Connect=False
            // forever). Confirmed by direct repro (Word). _chatId is instead
            // computed lazily on first actual use, in GetChatId() below, by
            // which point the task pane is visible and the user has
            // triggered a message - so ActivePresentation is guaranteed settled.
            _bridge = new WebViewBridgeHost(this, PowerPointTools.Execute, "PowerPointAiAddIn", UpdateStatus, OnOtherMessage);
        }

        private void UpdateStatus(string s)
        {
            _status.Text = s;
            _status.Visible = s != "ready";
        }

        private string GetChatId()
        {
            if (_chatId != null) return _chatId;

            PowerPoint.Presentation activePresentation = Globals.ThisAddIn.Application.ActivePresentation;
            // An unsaved presentation has no on-disk Path; ActivePresentation.FullName
            // falls back to its temp Name (e.g. "Presentation1") in that case,
            // which is not a stable key across sessions, so use a per-process
            // fallback id instead (mirrors genoffice's tempChatId concept).
            _chatId = string.IsNullOrEmpty(activePresentation.Path)
                ? "unsaved-" + Process.GetCurrentProcess().Id
                : ChatStore.ChatIdForFile(activePresentation.FullName);
            return _chatId;
        }

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider("PowerPointAiAddIn", GetChatId());
                    _bridge.PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage("PowerPointAiAddIn", GetChatId(), role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider("PowerPointAiAddIn", GetChatId());
                    break;
                case "set-mode":
                    string modeStr = root.GetProperty("mode").GetString();
                    switch (modeStr)
                    {
                        case "readOnly": PowerPointTools.Mode = EditingMode.ReadOnly; break;
                        case "commentOnly": PowerPointTools.Mode = EditingMode.CommentOnly; break;
                        case "trackChanges": PowerPointTools.Mode = EditingMode.TrackChanges; break;
                        case "fullAutonomy": PowerPointTools.Mode = EditingMode.FullAutonomy; break;
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
