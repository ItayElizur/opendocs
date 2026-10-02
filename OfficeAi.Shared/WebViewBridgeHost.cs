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

        // Off by default, settable only via the explicit "set-tls-bypass"
        // WebMessage - never enabled silently. For testing against an
        // internal LLM gateway with a self-signed cert. See WebViewBridgeHost.cs.md.
        private bool _skipTlsVerify;

        // One CoreWebView2Environment per app-data-folder name, shared across
        // every pane in this process and cached as the in-flight Task (not
        // just the eventual result) - a second pane's constructor can run
        // synchronously into this dictionary check before the first pane's
        // environment finishes creating, so it must await the SAME task
        // rather than start a second CreateAsync, which WebView2 rejects for
        // the same user-data-folder (HRESULT 0x8007139F). UI-thread only; no
        // lock needed. See WebViewBridgeHost.cs.md.
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

                // The persistent userDataFolder (see GetOrCreateEnvironment)
                // means the HTTP disk cache survives across Word restarts
                // even though the DLLs reload correctly, so a rebuilt bundle
                // can be served stale indefinitely. Force every request to
                // this virtual host to bypass cache and revalidate, matching
                // a browser hard-refresh. See WebViewBridgeHost.cs.md.
                _webView.CoreWebView2.AddWebResourceRequestedFilter("http://appassets.local/*", CoreWebView2WebResourceContext.All);
                _webView.CoreWebView2.WebResourceRequested += (sender, args) =>
                {
                    args.Request.Headers.SetHeader("Cache-Control", "no-cache, no-store, must-revalidate");
                    args.Request.Headers.SetHeader("Pragma", "no-cache");
                };

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.ServerCertificateErrorDetected += OnServerCertificateErrorDetected;

                // http, not https: appassets.local serves local files only, so
                // the scheme is a free choice, and it avoids Chromium blocking
                // a user-configured plain-HTTP LLM endpoint as mixed content
                // (fetching HTTPS from this insecure origin is unaffected -
                // mixed content is one-directional). See WebViewBridgeHost.cs.md
                // for the trade-off this forces (crypto.randomUUID()).
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

        // async void: an event handler on CoreWebView2.WebMessageReceived (UI
        // thread). The outer try/catch is mandatory - an exception escaping
        // async void crashes the process. The continuation after
        // `await _executor(...)` is not guaranteed to resume on the UI thread,
        // so the result is posted back through PostToolResult, which marshals
        // onto the control's thread. See WebViewBridgeHost.cs.md.
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
                // No status-bar flash here - the chat UI's own "Running N
                // tools..." work group already shows this inline; the
                // top-of-pane label stays reserved for real problems.
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
