using System;
using System.Text.Json;
using System.Windows.Forms;

namespace OfficeAi.Shared
{
    // Previously declared identically as three separate app-namespaced enums
    // (WordAiAddIn.EditingMode, ExcelAiAddIn.EditingMode, PowerPointAiAddIn.EditingMode)
    // with the same four members. PaneHostBase's SetEditingMode hook needs one
    // shared type to dispatch through, so this replaces all three - each app's
    // *Tools.cs now references this instead of declaring its own.
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    // Shared across Word/Excel/PowerPoint's TaskPaneHost.cs - the three copies
    // differed only in the COM type used to resolve the owning document (the
    // abstract hooks below) and app-data folder name (now a constructor
    // parameter). Everything else - the status label, the WebView2 bridge,
    // and the OnOtherMessage branches that don't need app-specific
    // resolution - lives once, here.
    public abstract class PaneHostBase : UserControl
    {
        private readonly Label _status;
        private readonly string _appDataFolderName;
        private readonly WebViewBridgeHost _bridge;

        // FT-2 Task 1: shared debounced selection dispatch. A WinForms Timer
        // (not System.Timers.Timer) ticks on the UI thread, where this
        // UserControl and its WebView2 already live - no cross-thread
        // marshaling needed. 200ms was chosen empirically: holding an arrow
        // key down in Excel settles the pill once, after the burst ends,
        // rather than flickering on every keystroke.
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

            // This base constructor never touches the owning document/
            // workbook/presentation - it only takes the app-data folder name.
            // Each subclass's own constructor stores its COM document
            // reference WITHOUT dereferencing it (no .Path/.FullName access
            // there); see GetChatId()'s doc comment on each subclass for the
            // confirmed repro of why an eager read silently kills the whole
            // VSTO connection (no exception, no error - just Connect=False
            // forever) when done at construction time instead of lazily.
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

        // FT-2 Task 1: coalesces bursts of selection-change events and drops
        // exact repeats. `signature` is a cheap string identifying the
        // selection (e.g. "Sheet1!B2:D40", "slides:2,3") - when it matches the
        // last one actually posted, the event is dropped outright rather than
        // even restarting the timer, so an event storm that never changes the
        // selection (e.g. re-entrant COM notifications) costs nothing.
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
        protected abstract ToolResult ExecuteTool(string name, JsonElement input);

        // The per-document chat-history/mode key. Lazily computed and cached
        // by each subclass on first actual use (never in the constructor -
        // see the constructor comment above). A subclass's override re-checks
        // a still-provisional ("unsaved-...") id on every call and migrates
        // ChatStore/DocSettingsStore onto the real id the moment the document
        // is saved (FT-1 Task 7b) - callers never need to know this happens.
        protected abstract string GetChatId();

        // FT-1 Task 7b Step 2: called from each app's document-close handler
        // (ThisAddIn.cs, alongside pane disposal) to force one last GetChatId()
        // check before the pane goes away - covers "save, then immediately
        // close" without needing a save-then-close-specific hook. ThisAddIn
        // cannot call the protected GetChatId() directly (different class,
        // not a subclass), hence this public wrapper.
        public void FlushChatIdMigration()
        {
            try { GetChatId(); }
            catch { /* best-effort; never let this block pane teardown */ }
        }

        // Routes a mode change to this app's *Tools class, keyed by GetChatId()
        // so the mode is per-document rather than shared across every window.
        protected abstract void SetEditingMode(EditingMode mode);

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
                // FT-1 Task 7: the document system message. Both branches use
                // GetChatId(), so they inherit the per-document keying and the
                // lazy-COM-resolution rule for free, same as chat history above.
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
            }
        }
    }
}
