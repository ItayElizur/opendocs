using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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

        // Alt-tabbing away and back makes Office give focus to the document, not the pane, so the cursor in the chat
        // box is lost. When the page reports it lost focus, this timer watches whether the app was left and comes
        // back, and if so gives the pane its focus again (see FocusReturn and PaneHostBase.cs.md).
        private readonly Timer _focusReturnTimer;
        private bool _focusWasAway;

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        private const uint GA_ROOT = 2;

        public event Action<int> RequestPaneWidth;

        protected PaneHostBase(string appDataFolderName)
        {
            _appDataFolderName = appDataFolderName;
            _disabledModes = ModePolicy.DisabledFor(appDataFolderName);
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

            _focusReturnTimer = new Timer { Interval = 250 };
            _focusReturnTimer.Tick += OnFocusReturnTick;

            // Never dereference the owning document/workbook/presentation at
            // construction time (no .Path/.FullName access) - doing so
            // silently kills the VSTO connection (no exception, just
            // Connect=False forever). Subclasses must resolve it lazily. See
            // PaneHostBase.cs.md.
            _bridge = new WebViewBridgeHost(this, ExecuteToolUnderPolicy, appDataFolderName, UpdateStatus, OnOtherMessage);
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

        private void StartFocusReturnWatch()
        {
            if (_focusReturnTimer.Enabled) return;
            _focusWasAway = false;
            _focusReturnTimer.Start();
        }

        private void OnFocusReturnTick(object sender, EventArgs e)
        {
            if (IsDisposed || !IsHandleCreated) { _focusReturnTimer.Stop(); return; }
            IntPtr root = GetAncestor(Handle, GA_ROOT);
            bool inFront = root != IntPtr.Zero && GetForegroundWindow() == root;
            switch (FocusReturn.Next(inFront, ref _focusWasAway))
            {
                case FocusReturnStep.Restore:
                    _focusReturnTimer.Stop();
                    _bridge.WebView.Focus();
                    break;
                case FocusReturnStep.Stop:
                    _focusReturnTimer.Stop();
                    break;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _focusReturnTimer.Tick -= OnFocusReturnTick;
                _focusReturnTimer.Stop();
                _focusReturnTimer.Dispose();
                _selectionTimer.Tick -= OnSelectionTimerTick;
                _selectionTimer.Stop();
                _selectionTimer.Dispose();
            }
            base.Dispose(disposing);
        }

        // The mode the page last reported; null until its first "set-mode".
        private EditingMode? _mode;

        // This app's modes switched off by the administrator's policy.json (wire names).
        private readonly HashSet<string> _disabledModes;

        // Administrator policy (ModePolicy): no tool runs while the pane is in a mode policy.json disables, or before the
        // page has said which mode it is in. Checked here once for all four apps, ahead of each app's own tier gate.
        private Task<ToolResult> ExecuteToolUnderPolicy(string name, JsonElement input)
        {
            if (ModePolicy.Blocks(_disabledModes, _mode))
            {
                return Task.FromResult(new ToolResult { Output = "Blocked: this editing mode is disabled by your administrator.", IsError = true, Summary = name });
            }
            // Re-store this pane's mode under the current chat id before every call: the id changes on a new document's
            // first save (the app would then fall back to Full autonomy), and Outlook panes share one per-mailbox mode.
            if (_mode != null) SetEditingMode(_mode.Value);
            return ExecuteTool(name, input);
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

        // Extra text for the model's system prompt, asked for once when the page starts. Outlook adds the user's name and
        // addresses and the mailbox's work week; the other apps add nothing.
        protected virtual Task<string> GetPromptContextAsync() => Task.FromResult<string>(null);

        // async void, so nothing may escape: an exception here would take down the Office process. The pane can close
        // while Outlook waits on Exchange, so the post sits inside the try as well.
        private async void SendPromptContext()
        {
            try
            {
                string text = null;
                try { text = await GetPromptContextAsync(); }
                catch (Exception ex) { DebugLog.WriteException("GetPromptContextAsync", ex); }
                PostMessage(new { kind = "prompt-context", text = text ?? "" });
            }
            catch (Exception ex) { DebugLog.WriteException("SendPromptContext", ex); }
        }

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
                        case "readOnly": _mode = EditingMode.ReadOnly; break;
                        case "commentOnly": _mode = EditingMode.CommentOnly; break;
                        case "trackChanges": _mode = EditingMode.TrackChanges; break;
                        case "fullAutonomy": _mode = EditingMode.FullAutonomy; break;
                    }
                    // A mode the policy switches off is never handed to the app, so it cannot leak to another pane's calls.
                    if (_mode != null && !ModePolicy.Blocks(_disabledModes, _mode)) SetEditingMode(_mode.Value);
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
                // The page lost keyboard focus (to another program, or to the document): see StartFocusReturnWatch.
                case "pane-blur":
                    StartFocusReturnWatch();
                    break;
                case "load-doc-settings":
                    DocSettings settings = DocSettingsStore.Load(_appDataFolderName, GetChatId());
                    PostMessage(new { kind = "doc-settings-loaded", systemMessage = settings.SystemMessage });
                    break;
                // Provider API keys (provider id -> key); stored encrypted, see SecretStore.
                case "save-secrets":
                    var secrets = new Dictionary<string, string>();
                    if (root.TryGetProperty("secrets", out var sec) && sec.ValueKind == JsonValueKind.Object)
                        foreach (JsonProperty p in sec.EnumerateObject())
                            if (p.Value.ValueKind == JsonValueKind.String) secrets[p.Name] = p.Value.GetString();
                    SecretStore.Save(_appDataFolderName, secrets);
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
                // Answered when the lookup finishes (Outlook asks Exchange); the page uses it from then on.
                case "load-prompt-context":
                    SendPromptContext();
                    break;
            }
        }
    }
}
