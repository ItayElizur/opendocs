using System;
using System.IO;
using System.Text.Json;
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
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    appDataFolderName, "WebView2");
                CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(environment);

                string webRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "web");
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    webRoot,
                    CoreWebView2HostResourceAccessKind.Allow);

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                _webView.Source = new Uri("https://appassets.local/index.html");
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

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                using (JsonDocument doc = JsonDocument.Parse(e.WebMessageAsJson))
                {
                    JsonElement root = doc.RootElement;
                    string kind = root.GetProperty("kind").GetString();
                    if (kind == "tool-call")
                    {
                        var (requestId, toolName, input) = ToolProtocol.ParseToolCall(e.WebMessageAsJson);
                        _setStatus("Executing tool: " + toolName);
                        ToolResult result = _executor(toolName, input);
                        if (_webView.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.PostWebMessageAsJson(ToolProtocol.SerializeToolResult(requestId, result));
                        }
                        _setStatus(result.IsError ? "Tool error: " + toolName : "ready");
                    }
                    else
                    {
                        _onOtherMessage?.Invoke(kind, root.Clone());
                    }
                }
            }
            catch (Exception ex)
            {
                _setStatus("message handling error: " + ex.Message);
            }
        }
    }
}
