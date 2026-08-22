using System;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordAiAddIn
{
    // Hosts a WebView2 control inside a VSTO CustomTaskPane. Spike 1 proved the
    // WebView2 <-> .NET message bridge; spike 3 uses that same bridge to route
    // tool calls from the WebView2-hosted AgentLoop into real COM calls against
    // the live Word document (WordTools.Execute), then posts the result back.
    public partial class TaskPaneHost : UserControl
    {
        private readonly WebView2 _webView;
        private readonly Label _status;

        public TaskPaneHost()
        {
            _status = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "WebView2: initializing...",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            };

            _webView = new WebView2
            {
                Dock = DockStyle.Fill,
            };

            this.Controls.Add(_webView);
            this.Controls.Add(_status);

            InitializeWebViewAsync();
        }

        private async void InitializeWebViewAsync()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAiAddIn", "WebView2");
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(environment);

                string webRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "web");
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    webRoot,
                    CoreWebView2HostResourceAccessKind.Allow);

                _webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                _webView.Source = new Uri("https://appassets.local/index.html");
                _status.Text = "WebView2: navigated (spike 2 - agent loop)";
            }
            catch (Exception ex)
            {
                _status.Text = "WebView2 init failed: " + ex.Message;
            }
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                using (JsonDocument doc = JsonDocument.Parse(e.WebMessageAsJson))
                {
                    var root = doc.RootElement;
                    string kind = root.GetProperty("kind").GetString();
                    if (kind != "tool-call") return;

                    string requestId = root.GetProperty("requestId").GetString();
                    string toolName = root.GetProperty("toolName").GetString();
                    JsonElement input = root.GetProperty("input");

                    _status.Text = "Executing tool: " + toolName;
                    ToolResult result = WordTools.Execute(toolName, input);
                    _status.Text = "Tool done: " + toolName + (result.IsError ? " (error)" : "");

                    string resultJson = JsonSerializer.Serialize(new
                    {
                        kind = "tool-result",
                        requestId,
                        output = result.Output,
                        isError = result.IsError,
                        mutated = result.Mutated,
                        summary = result.Summary,
                    });
                    _webView.CoreWebView2.PostWebMessageAsJson(resultJson);
                }
            }
            catch (Exception ex)
            {
                _status.Text = "message handling error: " + ex.Message;
            }
        }
    }
}
