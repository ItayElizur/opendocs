using System;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAi.Shared
{
    // Shared editing-mode enum dispatched through PaneHostBase.SetEditingMode;
    // each app's *Tools.cs references this instead of declaring its own. See
    // PaneHostBase.cs.md.
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    // Shared base for Word/Excel/PowerPoint's TaskPaneHost.cs - the status
    // label, the WebView2 bridge, and the app-agnostic OnOtherMessage
    // branches live once, here; subclasses supply only the COM document type
    // and app-data folder name. See PaneHostBase.cs.md.
    public abstract class PaneHostBase : UserControl
    {
        private readonly Label _status;
        private readonly string _appDataFolderName;
        private readonly WebViewBridgeHost _bridge;

        // Shared debounced selection dispatch; 200ms chosen empirically. See
        // PaneHostBase.cs.md.
        private readonly Timer _selectionTimer;
        private object _pendingSelection;
        private string _pendingSelectionSignature;
        private string _lastSelectionSignature;

        public event Action<int> RequestPaneWidth;

        protected PaneHostBase(string appDataFolderName)
        {
            _appDataFolderName = appDataFolderName;
            _status = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "WebView2: initializing...",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            };
            Controls.Add(_status);

            _selectionTimer = new Timer { Interval = 200 };
            _selectionTimer.Tick += OnSelectionTimerTick;

            // Never dereference the owning document/workbook/presentation at
            // construction time (no .Path/.FullName access) - doing so
            // silently kills the VSTO connection (no exception, just
            // Connect=False forever). Subclasses must resolve it lazily. See
            // PaneHostBase.cs.md.
            _bridge = new WebViewBridgeHost(this, ExecuteTool, appDataFolderName, UpdateStatus, OnOtherMessage);
        }

        private void UpdateStatus(string s)
        {
            _status.Text = s;
            _status.Visible = s != "ready";
        }

        protected void PostMessage(object payload)
        {
            _bridge.PostMessage(payload);
        }

        // Coalesces bursts of selection-change events and drops exact repeats
        // (by `signature`, e.g. "Sheet1!B2:D40") before even restarting the
        // timer. See PaneHostBase.cs.md.
        protected void PostSelection(object payload, string signature)
        {
            if (signature == _lastSelectionSignature) return;
            _pendingSelection = payload;
            _pendingSelectionSignature = signature;
            _selectionTimer.Stop();
            _selectionTimer.Start();
        }

        private void OnSelectionTimerTick(object sender, EventArgs e)
        {
            _selectionTimer.Stop();
            _lastSelectionSignature = _pendingSelectionSignature;
            PostMessage(_pendingSelection);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _selectionTimer.Tick -= OnSelectionTimerTick;
                _selectionTimer.Stop();
                _selectionTimer.Dispose();
            }
            base.Dispose(disposing);
        }

        // Resolves and executes a tool call against this pane's own document -
        // each subclass closes over its own GetChatId() to thread the
        // per-document key through to e.g. WordTools.Execute(docKey, name, input)
        // without changing the shared ToolExecutor delegate's signature.
        protected abstract Task<ToolResult> ExecuteTool(string name, JsonElement input);

        // The per-document chat-history/mode key, lazily computed and cached
        // by each subclass on first use (never in the constructor - see
        // above). Transparently migrates a still-provisional id onto the real
        // one once the document is saved. See PaneHostBase.cs.md.
        protected abstract string GetChatId();

        // Forces one last GetChatId() check before the pane is disposed, so a
        // save-then-close sequence still migrates the chat id. ThisAddIn
        // can't call the protected GetChatId() directly. See PaneHostBase.cs.md.
        public void FlushChatIdMigration()
        {
            try { GetChatId(); }
            catch { /* best-effort; never let this block pane teardown */ }
        }

        // Routes a mode change to this app's *Tools class, keyed by GetChatId()
        // so the mode is per-document rather than shared across every window.
        protected abstract void SetEditingMode(EditingMode mode);

        // Office's UI display language id - abstract because this shared
        // assembly has no access to any app's own VSTO-generated Globals
        // class. See OfficeLanguage.cs for how the LCID maps to a UI
        // language, and PaneHostBase.cs.md for the exact property path.
        protected abstract int GetOfficeUiLanguageId();

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider(_appDataFolderName, GetChatId());
                    PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage(_appDataFolderName, GetChatId(), role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider(_appDataFolderName, GetChatId());
                    break;
                case "set-mode":
                    string modeStr = root.GetProperty("mode").GetString();
                    switch (modeStr)
                    {
                        case "readOnly": SetEditingMode(EditingMode.ReadOnly); break;
                        case "commentOnly": SetEditingMode(EditingMode.CommentOnly); break;
                        case "trackChanges": SetEditingMode(EditingMode.TrackChanges); break;
                        case "fullAutonomy": SetEditingMode(EditingMode.FullAutonomy); break;
                    }
                    break;
                case "collapse-pane":
                    RequestPaneWidth?.Invoke(34);
                    break;
                case "expand-pane":
                    RequestPaneWidth?.Invoke(420);
                    break;
                // The document system message; both branches use GetChatId()
                // so they inherit the per-document keying and lazy-COM-
                // resolution rule for free, same as chat history above.
                case "load-doc-settings":
                    DocSettings settings = DocSettingsStore.Load(_appDataFolderName, GetChatId());
                    PostMessage(new { kind = "doc-settings-loaded", systemMessage = settings.SystemMessage });
                    break;
                case "save-doc-settings":
                    string systemMessage = root.TryGetProperty("systemMessage", out var sm) && sm.ValueKind == JsonValueKind.String
                        ? sm.GetString()
                        : "";
                    DocSettingsStore.Save(_appDataFolderName, GetChatId(), new DocSettings { SystemMessage = systemMessage });
                    break;
                // One-shot: Office's theme is read once when the pane boots,
                // never re-checked afterward (by design - see OfficeTheme.cs).
                case "load-theme":
                    PostMessage(new { kind = "office-theme", theme = OfficeTheme.ReadEffectiveTheme() });
                    break;
                // Same one-shot posture as load-theme above, for Office's UI
                // display language instead of its theme.
                case "load-language":
                    PostMessage(new { kind = "office-language", language = OfficeLanguage.ResolveUiLanguage(GetOfficeUiLanguageId()) });
                    break;
            }
        }
    }
}
