using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace OfficeAi.Shared
{
    // Hosts a WebView2 control docked into `host`, loads that app's web/ folder,
    // and bridges JSON WebMessages to/from it. "tool-call" messages are handled
    // here directly (the one thing every app needs identically); anything else
    // (set-mode, selection queries, chat persistence requests) is routed to
    // onOtherMessage, since its meaning is app-specific and this class stays
    // app-agnostic.
    public class WebViewBridgeHost
    {
        private readonly WebView2 _webView;
        private readonly ToolExecutor _executor;
        private readonly Action<string> _setStatus;
        private readonly OtherMessageHandler _onOtherMessage;

        // Off by default. Only settable via the explicit "set-tls-bypass"
        // WebMessage the Settings panel's checkbox sends - never enabled
        // silently. Intended for testing against an internal/air-gapped LLM
        // gateway using a self-signed or otherwise untrusted certificate;
        // the correct fix for a real deployment is trusting the real
        // certificate in Windows' certificate store instead.
        private bool _skipTlsVerify;

        // One CoreWebView2Environment per app-data-folder name, shared across
        // every pane created in this process (keyed defensively; in practice
        // one VSTO add-in's AppDomain only ever passes its own fixed name).
        // Before PP-1 there was only ever one TaskPaneHost/WebViewBridgeHost
        // per process, so CreateAsync was only ever called once - never
        // exercised. PP-1 creates one pane per open document window, so
        // opening (or Office internally re-activating) more than one window
        // can call this constructor again while the first pane's environment
        // is still initializing. Calling CoreWebView2Environment.CreateAsync
        // a second time for the SAME user-data-folder before the first call
        // has finished is exactly what WebView2 rejects with "the group or
        // resource is not in the correct state" (HRESULT 0x8007139F,
        // confirmed repro) - sharing one environment (the officially
        // documented pattern for multiple WebView2 controls against one
        // profile) eliminates the race by construction rather than trying to
        // serialize around it. Caching the in-flight Task (not just the
        // eventual result) matters: a second pane's constructor runs
        // synchronously up to this dictionary check/insert, before any
        // `await`, so it sees and awaits the SAME in-flight task instead of
        // starting a second CreateAsync. Only ever touched from the single
        // STA UI thread all Office COM callbacks run on, so no lock is needed.
        private static readonly Dictionary<string, Task<CoreWebView2Environment>> _environments =
            new Dictionary<string, Task<CoreWebView2Environment>>();

        private static Task<CoreWebView2Environment> GetOrCreateEnvironment(string appDataFolderName)
        {
            Task<CoreWebView2Environment> existing;
            if (_environments.TryGetValue(appDataFolderName, out existing)) return existing;

            string userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                appDataFolderName, "WebView2");
            Task<CoreWebView2Environment> created = CoreWebView2Environment.CreateAsync(null, userDataFolder);
            _environments[appDataFolderName] = created;
            return created;
        }

        public WebView2 WebView => _webView;

        public WebViewBridgeHost(
            Control host,
            ToolExecutor executor,
            string appDataFolderName,
            Action<string> setStatus,
            OtherMessageHandler onOtherMessage = null)
        {
            _executor = executor;
            _setStatus = setStatus ?? (_ => { });
            _onOtherMessage = onOtherMessage;

            _webView = new WebView2 { Dock = DockStyle.Fill };
            host.Controls.Add(_webView);

            InitializeAsync(appDataFolderName);
        }

        private async void InitializeAsync(string appDataFolderName)
        {
            try
            {
                CoreWebView2Environment environment = await GetOrCreateEnvironment(appDataFolderName);
                await _webView.EnsureCoreWebView2Async(environment);

                string webRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "web");
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    webRoot,
                    CoreWebView2HostResourceAccessKind.Allow);

                // Post-hoc fix (2026-08-24, user-reported a CSS fix not
                // taking effect after rebuild + close/reopen Word): the
                // CoreWebView2Environment here uses a PERSISTENT
                // userDataFolder (see GetOrCreateEnvironment below), so its
                // HTTP disk cache survives across Word restarts even though
                // the C# DLLs themselves reload correctly - closing and
                // reopening Word recreates the WebView2 CONTROL, but not its
                // cache. bundle.js/bundle.css/index.html are referenced with
                // no cache-busting query string, so a rebuilt bundle can be
                // silently served stale from cache indefinitely. Force every
                // request to this virtual host to bypass cache and
                // revalidate, matching what a browser's hard-refresh does.
                _webView.CoreWebView2.AddWebResourceRequestedFilter("http://appassets.local/*", CoreWebView2WebResourceContext.All);
                _webView.CoreWebView2.WebResourceRequested += (sender, args) =>
                {
                    args.Request.Headers.SetHeader("Cache-Control", "no-cache, no-store, must-revalidate");
                    args.Request.Headers.SetHeader("Pragma", "no-cache");
                };

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.ServerCertificateErrorDetected += OnServerCertificateErrorDetected;

                // http, not https: appassets.local serves local files only (no
                // real network transport), so the scheme is a free choice - and
                // an insecure page fetching an insecure resource is not mixed
                // content, unlike a secure page doing the same. Without this, a
                // user-configured remote LLM endpoint reachable only over plain
                // HTTP gets silently blocked by Chromium as mixed content
                // before the request ever leaves the process (confirmed via
                // WebView2 DevTools: "Mixed Content: ... was loaded over
                // HTTPS, but requested an insecure resource"). Fetching an
                // HTTPS endpoint from this now-insecure origin is unaffected -
                // mixed content is one-directional. The trade-off is
                // crypto.randomUUID() (secure-context-gated) - see
                // agent-core/id.ts's randomId() fallback, used everywhere this
                // codebase previously called crypto.randomUUID() directly.
                _webView.Source = new Uri("http://appassets.local/index.html");
                _setStatus("ready");
            }
            catch (Exception ex)
            {
                _setStatus("WebView2 init failed: " + ex.Message);
            }
        }

        public void PostMessage(object payload)
        {
            if (_webView.CoreWebView2 == null) return;
            _webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(payload));
        }

        private void OnServerCertificateErrorDetected(object sender, CoreWebView2ServerCertificateErrorDetectedEventArgs e)
        {
            e.Action = _skipTlsVerify
                ? CoreWebView2ServerCertificateErrorAction.AlwaysAllow
                : CoreWebView2ServerCertificateErrorAction.Default;
        }

        // async void: this is an event handler bound to CoreWebView2.WebMessageReceived,
        // raised on the UI thread. The tool-call branch awaits the (now Task-returning)
        // executor so a slow tool - e.g. Outlook's search_contacts, which does its EWS
        // call off the UI thread - no longer blocks the message pump. The outer try/catch
        // is mandatory: an exception escaping an async void crashes the process. The
        // continuation after `await _executor(...)` is NOT guaranteed to resume on the UI
        // thread (an Office host thread may carry no WinForms SynchronizationContext, so a
        // tool that awaited Task.Run resumes on a threadpool thread) - so the result is
        // posted back through PostToolResult, which marshals onto the control's thread.
        private async void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            // Read the payload into a local before the first await - the event args can
            // be invalidated once the handler yields.
            string json = e.WebMessageAsJson;
            try
            {
                string kind;
                using (JsonDocument doc = JsonDocument.Parse(json))
                {
                    kind = doc.RootElement.GetProperty("kind").GetString();
                    if (kind != "tool-call")
                    {
                        if (kind == "set-tls-bypass")
                        {
                            _skipTlsVerify = doc.RootElement.TryGetProperty("enabled", out var enabled) && enabled.GetBoolean();
                        }
                        else
                        {
                            _onOtherMessage?.Invoke(kind, doc.RootElement.Clone());
                        }
                        return;
                    }
                }

                var (requestId, name, input) = ToolProtocol.ParseToolCall(json);
                // Tool execution no longer flashes a status-bar message - the
                // chat UI's own "Running N tools..." work-group (chat-ui.ts)
                // already shows this inline, so the top-of-pane label stayed
                // reserved for real problems (init/message-handling errors).
                ToolResult result;
                try
                {
                    result = await _executor(name, input);
                }
                catch (Exception ex)
                {
                    // No _setStatus here - the continuation may be on a threadpool
                    // thread and touching the status Label off-thread would throw.
                    // A tool error surfaces in the chat UI from the IsError result.
                    result = new ToolResult { Output = ex.Message, IsError = true, Summary = name };
                }

                PostToolResult(requestId, result);
            }
            catch (Exception ex)
            {
                // Only reachable from the synchronous pre-await section (JSON parse,
                // ParseToolCall, _onOtherMessage) - all on the UI thread - so _setStatus
                // is safe here. PostToolResult swallows its own failures.
                _setStatus("message handling error: " + ex.Message);
            }
        }

        // Marshals the tool-result post back onto the WebView2 control's own thread and
        // never throws - touching the control (or CoreWebView2) from a threadpool
        // continuation would otherwise throw, and a re-throw from here would escape the
        // async void handler and take the host process down.
        private void PostToolResult(string requestId, ToolResult result)
        {
            try
            {
                if (_webView.IsDisposed || !_webView.IsHandleCreated) return;
                if (_webView.InvokeRequired)
                    _webView.BeginInvoke((Action)(() => RawPostToolResult(requestId, result)));
                else
                    RawPostToolResult(requestId, result);
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("WebViewBridgeHost.PostToolResult", ex);
            }
        }

        private void RawPostToolResult(string requestId, ToolResult result)
        {
            if (_webView.CoreWebView2 != null)
                _webView.CoreWebView2.PostWebMessageAsJson(ToolProtocol.SerializeToolResult(requestId, result));
        }
    }
}
