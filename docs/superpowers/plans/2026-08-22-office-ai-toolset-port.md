# Office AI Toolset Port Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port genoffice's AI document-manipulation tool set (Word, Excel, PowerPoint) and a redesigned chat UX ("Airchat Office") into real Microsoft Office, building on the VSTO + WebView2 hybrid architecture already validated in `C:\dev\officeoffice\WordAiAddIn` (spikes 1-3).

**Architecture:** Three separate VSTO COM add-ins (Word/Excel/PowerPoint — VSTO add-ins are per-host-application, so they cannot share one binary), each hosting a `CustomTaskPane` with a WebView2 control. All three load the same web bundle pattern: genoffice's `AgentLoop` + `ai-provider` streaming code (vendored once, shared, unmodified) plus a shared, purpose-built chat UI component (design finalized via `shared/chat-ui/mockup.html`, reviewed and iterated with the user), plus an app-specific `entry.ts` that defines that app's tool schemas and bridges tool calls to a `.NET` executor. Tool execution is real COM automation (`Microsoft.Office.Interop.{Word,Excel,PowerPoint}`), reached via the WebMessage JSON bridge proven in spike 3. A shared `.NET` class library (`OfficeAi.Shared`) removes the need to duplicate the WebView2/bridge hosting code three times, and now also owns a small chat-history persistence layer (divider-bounded, see Task 7) and a generic "push message to JS" channel used for live selection tracking (Task 12).

**Tech Stack:** VSTO (.NET Framework 4.8, C# 7.3, old-style `.csproj`), WebView2 (`Microsoft.Web.WebView2.WinForms`), `System.Text.Json`, TypeScript + esbuild (bundled, no framework), `Microsoft.Office.Interop.{Word,Excel,PowerPoint}` COM Interop.

**Spec:** `C:\Users\Itay\.claude\plans\i-was-wondering-if-floating-puffin.md` (feasibility report + spike results) and `C:\dev\officeoffice\shared\chat-ui\mockup.html` (the reviewed, approved chat UI design — the source of truth for Task 5's visual/behavioral implementation, including all UX decisions from the design-review conversation: no attachment support, editing-mode selector instead of a boolean track-changes toggle, selection-aware scope hint, divider-bounded chat history, header chrome that doesn't mirror under RTL, language-changes-on-Save, and no-fade historic messages).

## Global Constraints

- Every VSTO `.csproj` targets `TargetFrameworkVersion=v4.8`, `LangVersion` defaults to C# 7.3 (old-style project) — no `using`-declarations, no target-typed `new()`; use classic `using (...) { }` blocks and explicit types.
- `WebView2Environment` **must** be created with an explicit `userDataFolder` under `%LOCALAPPDATA%` — the default (host `.exe`'s own folder) is a read-only `Program Files` path for every Office host and throws `E_ACCESSDENIED`.
- All chart/shape/COM object-model code that isn't 100%-certain of its exact Interop type name should use `dynamic` typing rather than guessing overloads.
- **No automated unit tests for COM-executor methods** (`*Tools.cs` `Execute()` bodies) — they require a live, licensed Office host with a real open document. Each COM-tool task's "test" step is a **manual verification script** instead. Pure-logic code (JSON protocol, persistence, chat-ui DOM rendering) gets real automated tests as normal.
- The web-src TypeScript vendored from genoffice (`packages/agent-core/src/{loop,types,skill}.ts`, `packages/ai-provider/src/*.ts`) must stay **byte-for-byte unmodified** — module resolution is handled entirely via esbuild `--alias` flags.
- **No file-attachment support anywhere in this port** (explicit product decision) — genoffice's attach/paperclip UI and file-picker flow are not ported. If this changes later, it's a new, separate plan.
- **No multi-session chat browser** (explicit product decision, confirmed against genoffice's actual behavior which also has none) — exactly one continuous, divider-segmented history per document. See Task 7 for the divider design, which is a deliberate *improvement* over genoffice's unbounded-reload behavior, not a port of it.
- **Editing-mode gating is enforced server-side (C#), not just by hiding tools client-side.** The tool list offered to the model is filtered per mode as a first line of defense (smaller prompts, fewer wasted turns), but `*Tools.Execute()` must independently refuse any out-of-mode mutating call even if the model requests it anyway — defense in depth, since nothing stops a misbehaving or malicious model response from calling a tool that wasn't offered.
- Every task's manual verification step for a mutating tool must be run against a **scratch document**, not a document with real content.

---

## Phase 0 — Shared foundation

### Task 1: Initialize the officeoffice repo and top-level structure

**Files:**
- Create: `C:\dev\officeoffice\.gitignore`
- Create: `C:\dev\officeoffice\shared\` (already contains `chat-ui/mockup.html` from design review)

**Interfaces:** none (repo plumbing only).

- [ ] **Step 1: Initialize git and commit the existing spike work + approved mockup as the baseline**

```bash
cd C:/dev/officeoffice
git init
```

Create `.gitignore`:
```
bin/
obj/
node_modules/
*.user
*.suo
web/bundle.js
web/bundle.js.map
```

- [ ] **Step 2: Commit**

```bash
git add -A
git commit -m "chore: initial commit of WordAiAddIn spikes 1-3 and approved chat-ui mockup"
```

- [ ] **Step 3: Create the remaining shared/ subdirectories**

```bash
mkdir -p shared/web-src/agent-core shared/web-src/ai-provider
```

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "chore: scaffold shared/web-src/"
```

---

### Task 2: Move vendored agent-core/ai-provider into shared/, used by all future apps

**Files:**
- Move: `WordAiAddIn/web-src/agent-core/*.ts` → `shared/web-src/agent-core/*.ts`
- Move: `WordAiAddIn/web-src/ai-provider/*.ts` → `shared/web-src/ai-provider/*.ts`
- Modify: `WordAiAddIn/tsconfig.json`

**Interfaces:**
- Produces: `../shared/web-src/agent-core/index.ts` exporting `AgentLoop`, `AgentSkill`, `AgentTransport`, `AgentStreamHandle`, `ToolExecution`, `AgentToolCall`, `AgentToolDef`, `AgentMessage` (unchanged).
- Produces: `../shared/web-src/ai-provider/index.ts` exporting `streamOpenAiCompatible`, `AiProviderConfig` (unchanged).

- [ ] **Step 1: Move the files**

```bash
cd C:/dev/officeoffice
git mv WordAiAddIn/web-src/agent-core/index.ts shared/web-src/agent-core/index.ts
git mv WordAiAddIn/web-src/agent-core/loop.ts shared/web-src/agent-core/loop.ts
git mv WordAiAddIn/web-src/agent-core/skill.ts shared/web-src/agent-core/skill.ts
git mv WordAiAddIn/web-src/agent-core/types.ts shared/web-src/agent-core/types.ts
git mv WordAiAddIn/web-src/ai-provider/index.ts shared/web-src/ai-provider/index.ts
git mv WordAiAddIn/web-src/ai-provider/types.ts shared/web-src/ai-provider/types.ts
git mv WordAiAddIn/web-src/ai-provider/stream.ts shared/web-src/ai-provider/stream.ts
git mv WordAiAddIn/web-src/ai-provider/fetch.ts shared/web-src/ai-provider/fetch.ts
git mv WordAiAddIn/web-src/ai-provider/watchdog.ts shared/web-src/ai-provider/watchdog.ts
git mv WordAiAddIn/web-src/ai-provider/http-error.ts shared/web-src/ai-provider/http-error.ts
git mv WordAiAddIn/web-src/ai-provider/providers.ts shared/web-src/ai-provider/providers.ts
```

- [ ] **Step 2: Update WordAiAddIn's tsconfig.json paths**

Edit `WordAiAddIn/tsconfig.json`, change the `paths` block to:
```json
"paths": {
  "@genoffice/agent-core": ["../shared/web-src/agent-core/index.ts"],
  "@genoffice/ai-provider": ["../shared/web-src/ai-provider/index.ts"]
}
```

- [ ] **Step 3: Canonical build command (documented here, referenced by later tasks)**

```bash
npx esbuild web-src/entry.ts \
  --bundle \
  --outfile=web/bundle.js \
  --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts \
  --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts \
  --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts \
  --target=chrome100 \
  --format=iife \
  --sourcemap
```

- [ ] **Step 4: Verify WordAiAddIn still builds and typechecks**

```bash
cd WordAiAddIn
npx tsc --noEmit
npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --target=chrome100 --format=iife --sourcemap
```
Expected: no errors, `web/bundle.js` written (~39KB).

- [ ] **Step 5: Manual verification — full regression of spike 3's demo**

Rebuild the VSTO project, close/reopen Word, open a scratch document, send "please read the document and add a chart" in the task pane. Expected: identical behavior to spike 3.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor: move vendored agent-core/ai-provider into shared/web-src"
```

---

### Task 3: Shared .NET library (OfficeAi.Shared) — tool protocol, WebView2 bridge host, and a generic push channel

**Files:**
- Create: `OfficeAi.Shared/OfficeAi.Shared.csproj`
- Create: `OfficeAi.Shared/ToolProtocol.cs`
- Create: `OfficeAi.Shared/WebViewBridgeHost.cs`
- Test: `OfficeAi.Shared.Tests/ToolProtocolTests.cs`
- Test: `OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj`

**Interfaces:**
- Produces: `OfficeAi.Shared.ToolResult` — `struct { string Output; bool IsError; bool Mutated; string Summary; }`
- Produces: `OfficeAi.Shared.ToolExecutor` — `delegate ToolResult ToolExecutor(string toolName, System.Text.Json.JsonElement input)`
- Produces: `OfficeAi.Shared.ToolProtocol.ParseToolCall(string json)` → `(string requestId, string toolName, JsonElement input)`, throws `FormatException` if `kind != "tool-call"` or a required field is missing.
- Produces: `OfficeAi.Shared.ToolProtocol.SerializeToolResult(string requestId, ToolResult result)` → `string` (JSON).
- Produces: `OfficeAi.Shared.WebViewBridgeHost` — constructor `WebViewBridgeHost(Control host, ToolExecutor executor, string appDataFolderName, Action<string> setStatus, OtherMessageHandler onOtherMessage = null)` where `OtherMessageHandler` is `delegate void OtherMessageHandler(string kind, JsonElement root)`, invoked for every incoming WebMessage whose `kind` is not `"tool-call"` (used by Task 11 for `set-mode` and Task 12 for selection-related messages — this app-agnostic bridge doesn't know what those mean, it just routes them). Also exposes `void PostMessage(object payload)` — serializes `payload` and posts it to the WebView2 side unprompted (used by Task 12 to push live selection updates and by Task 7 conceptually, though Task 7's persistence responses go through the same channel too).

- [ ] **Step 1: Create the class library project**

`OfficeAi.Shared/OfficeAi.Shared.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <RootNamespace>OfficeAi.Shared</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2903.40" />
    <PackageReference Include="System.Text.Json" Version="8.0.5" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: Write the tool-protocol JSON parse/serialize helpers and their tests first**

`OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\OfficeAi.Shared\OfficeAi.Shared.csproj" />
  </ItemGroup>
</Project>
```

`OfficeAi.Shared.Tests/ToolProtocolTests.cs`:
```csharp
using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class ToolProtocolTests
{
    [Fact]
    public void ParseToolCall_ExtractsFields()
    {
        string json = "{\"kind\":\"tool-call\",\"requestId\":\"abc\",\"toolName\":\"insert_content\",\"input\":{\"text\":\"hi\"}}";
        var (requestId, toolName, input) = ToolProtocol.ParseToolCall(json);
        Assert.Equal("abc", requestId);
        Assert.Equal("insert_content", toolName);
        Assert.Equal("hi", input.GetProperty("text").GetString());
    }

    [Fact]
    public void ParseToolCall_ThrowsOnWrongKind()
    {
        string json = "{\"kind\":\"tool-result\",\"requestId\":\"abc\"}";
        Assert.Throws<System.FormatException>(() => ToolProtocol.ParseToolCall(json));
    }

    [Fact]
    public void SerializeToolResult_RoundTrips()
    {
        var result = new ToolResult { Output = "done", IsError = false, Mutated = true, Summary = "insert_content" };
        string json = ToolProtocol.SerializeToolResult("abc", result);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        Assert.Equal("tool-result", root.GetProperty("kind").GetString());
        Assert.Equal("abc", root.GetProperty("requestId").GetString());
        Assert.Equal("done", root.GetProperty("output").GetString());
        Assert.True(root.GetProperty("mutated").GetBoolean());
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

```bash
cd OfficeAi.Shared.Tests
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" OfficeAi.Shared.Tests.csproj -t:restore
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" OfficeAi.Shared.Tests.csproj -t:Build
"C:/Program Files/Microsoft Visual Studio/2022/Community/Common7/IDE/CommonExtensions/Microsoft/TestWindow/vstest.console.exe" bin/Debug/net48/OfficeAi.Shared.Tests.dll
```
Expected: build fails (`ToolProtocol` doesn't exist yet).

- [ ] **Step 4: Implement ToolProtocol.cs**

```csharp
using System;
using System.Text.Json;

namespace OfficeAi.Shared
{
    public struct ToolResult
    {
        public string Output;
        public bool IsError;
        public bool Mutated;
        public string Summary;
    }

    public delegate ToolResult ToolExecutor(string toolName, JsonElement input);
    public delegate void OtherMessageHandler(string kind, JsonElement root);

    public static class ToolProtocol
    {
        public static (string requestId, string toolName, JsonElement input) ParseToolCall(string json)
        {
            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                JsonElement root = doc.RootElement.Clone();
                string kind = root.GetProperty("kind").GetString();
                if (kind != "tool-call")
                {
                    throw new FormatException("Expected kind=tool-call, got: " + kind);
                }
                string requestId = root.GetProperty("requestId").GetString();
                string toolName = root.GetProperty("toolName").GetString();
                JsonElement input = root.GetProperty("input").Clone();
                return (requestId, toolName, input);
            }
        }

        public static string SerializeToolResult(string requestId, ToolResult result)
        {
            return JsonSerializer.Serialize(new
            {
                kind = "tool-result",
                requestId,
                output = result.Output,
                isError = result.IsError,
                mutated = result.Mutated,
                summary = result.Summary,
            });
        }
    }
}
```

- [ ] **Step 5: Run tests to verify they pass** (repeat step 3's commands — expect 3/3 pass)

- [ ] **Step 6: Implement WebViewBridgeHost.cs**

```csharp
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
                        _setStatus("Tool done: " + toolName + (result.IsError ? " (error)" : ""));
                        PostMessage(new
                        {
                            kind = "tool-result",
                            requestId,
                            output = result.Output,
                            isError = result.IsError,
                            mutated = result.Mutated,
                            summary = result.Summary,
                        });
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
```

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: OfficeAi.Shared class library (tool protocol + WebView2 bridge host)"
```

---

### Task 4: Refactor WordAiAddIn to consume OfficeAi.Shared

**Files:**
- Modify: `WordAiAddIn/WordAiAddIn.csproj`
- Modify: `WordAiAddIn/TaskPaneHost.cs`
- Modify: `WordAiAddIn/WordTools.cs` (use `OfficeAi.Shared.ToolResult`)

**Interfaces:**
- Consumes: `OfficeAi.Shared.WebViewBridgeHost`, `OfficeAi.Shared.ToolExecutor`, `OfficeAi.Shared.ToolResult` (Task 3).

- [ ] **Step 1: Add the project reference** — in `WordAiAddIn/WordAiAddIn.csproj`, before `</Project>`:
```xml
<ItemGroup>
  <ProjectReference Include="..\OfficeAi.Shared\OfficeAi.Shared.csproj" />
</ItemGroup>
```

- [ ] **Step 2: Update WordTools.cs** — remove the local `ToolResult` struct, add `using OfficeAi.Shared;`.

- [ ] **Step 3: Shrink TaskPaneHost.cs**

```csharp
using System.Windows.Forms;
using OfficeAi.Shared;

namespace WordAiAddIn
{
    public partial class TaskPaneHost : UserControl
    {
        private readonly Label _status;
        private readonly WebViewBridgeHost _bridge;

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

            _bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s);
        }
    }
}
```
(This constructor grows in Tasks 7, 11, and 12 to pass a real `onOtherMessage` handler and to wire selection events — left minimal here since those pieces don't exist yet.)

- [ ] **Step 4: Rebuild and manually verify no regression** — same check as Task 2 Step 5.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "refactor: WordAiAddIn consumes OfficeAi.Shared.WebViewBridgeHost"
```

---

## Phase 1 — Chat UX

### Task 5: Build the real chat-ui component from the approved mockup

**Precondition:** `shared/chat-ui/mockup.html` is the approved design — every behavior below is taken directly from it and the design-review conversation, not invented fresh here.

**Files:**
- Create: `shared/chat-ui/chat-ui.css` (extracted from the mockup's `<style>` block, minus the mockup-only `.mockup-controls`/`.host-window` scaffolding)
- Create: `shared/chat-ui/chat-ui.ts`
- Test: `shared/chat-ui/chat-ui.test.ts` (vitest + jsdom)
- Create: `shared/chat-ui/package.json`, `shared/chat-ui/vitest.config.ts`

**Interfaces:**
```ts
export type EditingMode = 'readOnly' | 'commentOnly' | 'trackChanges' | 'fullAutonomy'

export interface ChatUIOptions {
  title: string
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; lang: 'en' | 'he' }) => void
}

export interface ToolStepHandle {
  complete(result: { output: string; isError?: boolean; mutated?: boolean }): void
}

export interface ToolGroupHandle {
  addStep(toolName: string, input: Record<string, unknown>): ToolStepHandle
  end(): void
}

export interface ChatUIHandle {
  addUserMessage(text: string): void
  beginAssistantMessage(): void
  updateAssistantMessage(cumulativeText: string): void
  endAssistantMessage(finalText: string): void
  beginToolGroup(): ToolGroupHandle
  setBusy(busy: boolean): void
  showError(message: string): void
  /** clears the live conversation back to the empty/starter-pills state (called after onNewChat's own persistence work is done) */
  resetToEmpty(): void
  /** renders a divider + the given messages above it as historic (full opacity, no fade - see chat-ui notes), used on mount when Task 7's persistence layer returns prior messages */
  showHistoric(messages: Array<{ role: 'user' | 'assistant'; text: string }>): void
  /** "Whole document" or a live selection preview like `Selection: "..."` - see Task 12 */
  setScopeHint(label: string): void
}

export function mountChatUI(root: HTMLElement, options: ChatUIOptions): ChatUIHandle
```

Note what's deliberately **not** here versus the original spike UI: no attachment methods/UI at all (product decision), no boolean track-changes flag (replaced by `onModeChange` with 4 modes), no settings-are-instant behavior (settings only commit via `onSettingsSave`, called from the dropdown's Save button, matching "language changes on Save").

- [ ] **Step 1: Set up the test project**

`shared/chat-ui/package.json`:
```json
{
  "name": "@officeai/chat-ui",
  "private": true,
  "type": "module",
  "devDependencies": {
    "vitest": "^4.1.0",
    "jsdom": "^25.0.0",
    "typescript": "^5.9.3"
  }
}
```
`shared/chat-ui/vitest.config.ts`:
```ts
import { defineConfig } from 'vitest/config'
export default defineConfig({ test: { environment: 'jsdom' } })
```
```bash
cd shared/chat-ui
npm install
```

- [ ] **Step 2: Write the failing tests**

```ts
import { describe, expect, it, vi } from 'vitest'
import { mountChatUI } from './chat-ui'

function setup() {
  const root = document.createElement('div')
  document.body.appendChild(root)
  const onSend = vi.fn()
  const onModeChange = vi.fn()
  const onSettingsSave = vi.fn()
  const onNewChat = vi.fn()
  const handle = mountChatUI(root, { title: 'Airchat Office', onSend, onModeChange, onSettingsSave, onNewChat })
  return { root, onSend, onModeChange, onSettingsSave, onNewChat, handle }
}

describe('mountChatUI', () => {
  it('renders the title and no attachment button', () => {
    const { root } = setup()
    expect(root.textContent).toContain('Airchat Office')
    expect(root.querySelector('.ai-attach-btn')).toBeNull()
    expect(root.querySelector('input[type="file"]')).toBeNull()
  })

  it('sending calls onSend and clears the textarea', () => {
    const { root, onSend } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    textarea.value = 'do the thing'
    root.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
    expect(onSend).toHaveBeenCalledWith('do the thing')
    expect(textarea.value).toBe('')
  })

  it('clicking a mode menu item calls onModeChange with that mode and marks it selected', () => {
    const { root, onModeChange } = setup()
    root.querySelector<HTMLButtonElement>('.ai-mode-btn')!.click()
    root.querySelector<HTMLElement>('[data-mode="trackChanges"]')!.click()
    expect(onModeChange).toHaveBeenCalledWith('trackChanges')
    expect(root.querySelector('[data-mode="trackChanges"]')!.classList.contains('selected')).toBe(true)
  })

  it('settings only call onSettingsSave when Save is clicked, not on field input', () => {
    const { root, onSettingsSave } = setup()
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    const baseUrlInput = root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!
    baseUrlInput.value = 'http://localhost:9000/v1'
    baseUrlInput.dispatchEvent(new Event('input'))
    expect(onSettingsSave).not.toHaveBeenCalled()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ baseUrl: 'http://localhost:9000/v1' }))
  })

  it('a tool group renders a step per addStep call and reflects completion, collapsed by default', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    expect(root.querySelector('.ai-work-group')!.classList.contains('open')).toBe(false)
    const step = group.addStep('insert_content', { text: 'hi' })
    step.complete({ output: 'Inserted text: hi', mutated: true })
    expect(root.querySelector('.ai-applied-tag')).not.toBeNull()
  })

  it('showHistoric renders messages above a divider with full opacity (no fade class)', () => {
    const { root, handle } = setup()
    handle.showHistoric([{ role: 'user', text: 'earlier question' }, { role: 'assistant', text: 'earlier answer' }])
    expect(root.querySelector('.ai-history-sep')).not.toBeNull()
    expect(root.textContent).toContain('earlier question')
    expect(root.querySelector('.ai-history-faded')).toBeNull()
  })

  it('setScopeHint updates the hint label text', () => {
    const { root, handle } = setup()
    handle.setScopeHint('Selection: "Q3 revenue grew..."')
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selection: "Q3 revenue grew..."')
  })

  it('resetToEmpty calls onNewChat is NOT implied - resetToEmpty just clears the DOM to the empty state', () => {
    const { root, handle } = setup()
    handle.addUserMessage('x')
    handle.resetToEmpty()
    expect(root.querySelector('.ai-msg-user')).toBeNull()
    expect(root.querySelector('.ai-chat-empty')).not.toBeNull()
  })
})
```

- [ ] **Step 3: Run tests, verify they fail**

```bash
npx vitest run
```
Expected: fails (module doesn't exist).

- [ ] **Step 4: Extract the CSS** from the approved `mockup.html`'s `<style>` block into `chat-ui.css` verbatim, dropping only `.mockup-controls`/`.host-window`/`.host-doc-area`.

- [ ] **Step 5: Implement chat-ui.ts**

```ts
import './chat-ui.css'

export type EditingMode = 'readOnly' | 'commentOnly' | 'trackChanges' | 'fullAutonomy'

export interface ChatUIOptions {
  title: string
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; lang: 'en' | 'he' }) => void
}

export interface ToolStepHandle {
  complete(result: { output: string; isError?: boolean; mutated?: boolean }): void
}

export interface ToolGroupHandle {
  addStep(toolName: string, input: Record<string, unknown>): ToolStepHandle
  end(): void
}

export interface ChatUIHandle {
  addUserMessage(text: string): void
  beginAssistantMessage(): void
  updateAssistantMessage(cumulativeText: string): void
  endAssistantMessage(finalText: string): void
  beginToolGroup(): ToolGroupHandle
  setBusy(busy: boolean): void
  showError(message: string): void
  resetToEmpty(): void
  showHistoric(messages: Array<{ role: 'user' | 'assistant'; text: string }>): void
  setScopeHint(label: string): void
}

const MODES: EditingMode[] = ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']

function escapeHtml(s: string): string {
  const div = document.createElement('div')
  div.textContent = s
  return div.innerHTML
}

export function mountChatUI(root: HTMLElement, options: ChatUIOptions): ChatUIHandle {
  root.innerHTML = `
    <div class="ai-panel">
      <div class="ai-panel-header">
        <div class="ai-panel-title"><span class="ai-logo">A</span><span>Airchat Office</span></div>
        <div class="ai-header-actions">
          <button class="ai-header-btn" data-t-title="newChat">+</button>
          <button class="ai-header-btn" data-t-title="settings">&#9881;</button>
          <button class="ai-header-btn" data-t-title="collapse">&#x276E;</button>
        </div>
        <div class="ai-settings-panel" id="settingsPanel">
          <h4>Airchat Office Settings</h4>
          <div class="ai-field"><label>API Base URL</label><input data-field="baseUrl" type="text" /></div>
          <div class="ai-field"><label>API Key</label><input data-field="apiKey" type="password" /></div>
          <div class="ai-field"><label>Model name</label><input data-field="model" type="text" /></div>
          <div class="ai-field">
            <label>Language</label>
            <div class="ai-lang-toggle">
              <button data-lang="en" class="active">English</button>
              <button data-lang="he">עברית</button>
            </div>
          </div>
          <div class="ai-settings-actions"><button class="ai-btn-primary">Save</button></div>
        </div>
      </div>
      <div class="ai-chat"></div>
      <div class="ai-composer">
        <div class="ai-input-box">
          <span class="ai-scope-hint"><span class="dot"></span><span class="label" id="scopeHintLabel">Whole document</span></span>
          <textarea class="ai-textarea" rows="1" placeholder="Ask Airchat Office to edit this document..."></textarea>
          <div class="ai-input-footer">
            <div style="position: relative;">
              <button class="ai-mode-btn"><span class="dot"></span><span id="modeBtnLabel">Full autonomy</span></button>
              <div class="ai-mode-menu" id="modeMenu">
                <div class="ai-mode-menu-item" data-mode="readOnly"><span>Read only</span><span class="desc">AI can only read, never edit</span></div>
                <div class="ai-mode-menu-item" data-mode="commentOnly"><span>Comment only</span><span class="desc">Adds comments, no content edits</span></div>
                <div class="ai-mode-menu-item" data-mode="trackChanges"><span>Track changes</span><span class="desc">Edits as reviewable revisions</span></div>
                <div class="ai-mode-menu-item selected" data-mode="fullAutonomy"><span>Full autonomy</span><span class="desc">Edits applied directly</span></div>
              </div>
            </div>
            <button class="ai-send-btn" data-t-title="send">&#10148;</button>
          </div>
        </div>
      </div>
    </div>
  `

  const chatEl = root.querySelector<HTMLDivElement>('.ai-chat')!
  const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
  const sendBtn = root.querySelector<HTMLButtonElement>('.ai-send-btn')!
  const newChatBtn = root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!
  const settingsBtn = root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!
  const settingsPanel = root.querySelector<HTMLDivElement>('#settingsPanel')!
  const modeBtn = root.querySelector<HTMLButtonElement>('.ai-mode-btn')!
  const modeMenu = root.querySelector<HTMLDivElement>('#modeMenu')!
  const modeBtnLabel = root.querySelector<HTMLSpanElement>('#modeBtnLabel')!
  const scopeHintLabel = root.querySelector<HTMLSpanElement>('#scopeHintLabel')!

  let assistantBubble: HTMLDivElement | null = null
  let pendingLang: 'en' | 'he' = 'en'

  function scrollToBottom(): void {
    chatEl.scrollTop = chatEl.scrollHeight
  }

  function doSend(): void {
    const text = textarea.value.trim()
    if (!text) return
    textarea.value = ''
    options.onSend(text)
  }

  sendBtn.addEventListener('click', doSend)
  textarea.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      doSend()
    }
  })
  newChatBtn.addEventListener('click', () => options.onNewChat())

  settingsBtn.addEventListener('click', () => settingsPanel.classList.toggle('open'))
  root.querySelectorAll<HTMLButtonElement>('.ai-lang-toggle button').forEach((btn) => {
    btn.addEventListener('click', () => {
      pendingLang = btn.dataset.lang as 'en' | 'he'
      root.querySelectorAll('.ai-lang-toggle button').forEach((b) => b.classList.toggle('active', b === btn))
    })
  })
  root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.addEventListener('click', () => {
    options.onSettingsSave({
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
      lang: pendingLang,
    })
    settingsPanel.classList.remove('open')
  })

  modeBtn.addEventListener('click', () => modeMenu.classList.toggle('open'))
  root.querySelectorAll<HTMLElement>('.ai-mode-menu-item').forEach((item) => {
    item.addEventListener('click', () => {
      const mode = item.dataset.mode as EditingMode
      root.querySelectorAll('.ai-mode-menu-item').forEach((el) => el.classList.toggle('selected', el === item))
      modeBtnLabel.textContent = item.querySelector('span')!.textContent
      modeMenu.classList.remove('open')
      modeBtn.classList.toggle('accent', mode === 'trackChanges')
      options.onModeChange(mode)
    })
  })

  function renderMessage(role: 'user' | 'assistant', text: string): HTMLDivElement {
    const div = document.createElement('div')
    div.className = role === 'user' ? 'ai-msg-user' : 'ai-msg-assistant'
    div.textContent = text
    chatEl.appendChild(div)
    return div
  }

  return {
    addUserMessage(text) {
      renderMessage('user', text)
      scrollToBottom()
    },
    beginAssistantMessage() {
      assistantBubble = renderMessage('assistant', '')
      assistantBubble.classList.add('streaming')
      scrollToBottom()
    },
    updateAssistantMessage(cumulativeText) {
      if (assistantBubble) assistantBubble.textContent = cumulativeText
      scrollToBottom()
    },
    endAssistantMessage(finalText) {
      if (assistantBubble) {
        assistantBubble.textContent = finalText
        assistantBubble.classList.remove('streaming')
      }
      assistantBubble = null
      scrollToBottom()
    },
    beginToolGroup() {
      const groupEl = document.createElement('div')
      groupEl.className = 'ai-work-group'
      groupEl.innerHTML = `<div class="ai-work-group-summary"><span class="caret">&#9656;</span><span class="label">Running tools...</span></div><div class="ai-work-group-body"><div class="steps"></div></div>`
      chatEl.appendChild(groupEl)
      groupEl.querySelector('.ai-work-group-summary')!.addEventListener('click', () => groupEl.classList.toggle('open'))
      const summaryEl = groupEl.querySelector<HTMLElement>('.label')!
      const stepsEl = groupEl.querySelector<HTMLElement>('.steps')!
      let count = 0
      scrollToBottom()
      return {
        addStep(toolName, input) {
          count++
          summaryEl.textContent = `Running ${count} tool${count > 1 ? 's' : ''}...`
          const rowEl = document.createElement('div')
          rowEl.className = 'ai-step-row'
          rowEl.innerHTML = `<div class="ai-step-icon">&#8987;</div><div class="ai-step-title">${escapeHtml(toolName)}(${escapeHtml(JSON.stringify(input))})</div>`
          stepsEl.appendChild(rowEl)
          scrollToBottom()
          return {
            complete(result) {
              const iconEl = rowEl.querySelector<HTMLElement>('.ai-step-icon')!
              iconEl.textContent = result.isError ? '\u2717' : '\u2713'
              iconEl.classList.toggle('error', !!result.isError)
              if (result.mutated) {
                const tag = document.createElement('div')
                tag.className = 'ai-applied-tag'
                tag.textContent = '\u2713 Applied'
                stepsEl.appendChild(tag)
              }
              scrollToBottom()
            },
          }
        },
        end() {
          summaryEl.textContent = `Ran ${count} tool${count === 1 ? '' : 's'}`
        },
      }
    },
    setBusy(busy) {
      sendBtn.disabled = busy
      textarea.disabled = busy
    },
    showError(message) {
      const div = document.createElement('div')
      div.className = 'ai-msg-error'
      div.textContent = message
      chatEl.appendChild(div)
      scrollToBottom()
    },
    resetToEmpty() {
      chatEl.innerHTML = `<div class="ai-chat-empty"><div class="ai-chat-empty-title">What can I help with?</div><div class="ai-starters"></div></div>`
    },
    showHistoric(messages) {
      for (const m of messages) renderMessage(m.role, m.text)
      const sep = document.createElement('div')
      sep.className = 'ai-history-sep'
      sep.textContent = 'Earlier conversation'
      chatEl.appendChild(sep)
      scrollToBottom()
    },
    setScopeHint(label) {
      scopeHintLabel.textContent = label
    },
  }
}
```

- [ ] **Step 6: Run tests, verify they pass**

```bash
npx vitest run
```
Expected: all tests pass.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: real chat-ui component from the approved mockup (no attachments, mode selector, scope hint, divider history)"
```

---

### Task 6: Wire chat-ui into WordAiAddIn

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`
- Modify: `WordAiAddIn/web/index.html`

**Interfaces:**
- Consumes: `mountChatUI` (Task 5) via `@officeai/chat-ui`.

- [ ] **Step 1: Shrink index.html**

```html
<!DOCTYPE html>
<html>
<head><meta charset="utf-8" /><title>Airchat Office</title></head>
<body style="margin:0;height:100vh;">
  <div id="root" style="height:100%;"></div>
  <script src="bundle.js"></script>
</body>
</html>
```

- [ ] **Step 2: Rewrite entry.ts's UI wiring**

Replace the old ad hoc transcript code with:
```ts
import { mountChatUI, type EditingMode } from '@officeai/chat-ui'

const root = document.getElementById('root')!
const ui = mountChatUI(root, {
  title: 'Airchat Office',
  onSend: (text) => {
    if (loop.busy) return
    ui.addUserMessage(text)
    ui.beginAssistantMessage()
    ui.setBusy(true)
    loop.run(text)
  },
  onNewChat: () => {
    // Task 7 wires the actual divider persistence call here.
    ui.resetToEmpty()
  },
  onModeChange: (mode: EditingMode) => {
    // Task 11 wires the actual bridge call + tool-list filtering here.
  },
  onSettingsSave: (settings) => {
    // Not yet wired to the transport/provider config - deferred (Phase 5).
  },
})

let currentToolGroup: ReturnType<typeof ui.beginToolGroup> | null = null
const activeSteps = new Map<string, ReturnType<ReturnType<typeof ui.beginToolGroup>['addStep']>>()
```

Update the `AgentLoop` events block:
```ts
events: {
  onText: (text) => ui.updateAssistantMessage(text),
  onToolStart: (call) => {
    if (!currentToolGroup) currentToolGroup = ui.beginToolGroup()
    activeSteps.set(call.id, currentToolGroup.addStep(call.name, call.input))
  },
  onToolExecuted: (event) => {
    activeSteps.get(event.call.id)?.complete({
      output: event.execution.output,
      isError: event.execution.isError,
      mutated: event.execution.mutated,
    })
    activeSteps.delete(event.call.id)
  },
  onTurnEnd: () => {
    currentToolGroup?.end()
    currentToolGroup = null
  },
  onDone: (result) => {
    ui.endAssistantMessage(result.text || '(no text)')
    ui.setBusy(false)
  },
  onError: (error) => {
    ui.showError(error)
    ui.setBusy(false)
  },
},
```

Remove the spike-2 diagnostic per-chunk logging from `makeTransport()`'s `onDelta` (restore it to `onDelta: callbacks.onDelta`).

- [ ] **Step 3: Typecheck, rebuild, verify**

```bash
cd WordAiAddIn
npx tsc --noEmit
npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" WordAiAddIn.csproj -t:Build -p:Configuration=Debug
```

- [ ] **Step 4: Manual verification** — close/reopen Word, compare visually against `mockup.html`, re-run the spike 3 chat instruction, confirm the tool-call timeline renders through the real component, collapsed by default.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: wire real chat-ui into WordAiAddIn"
```

---

### Task 7: Chat persistence with divider-bounded history

**Context:** genoffice persists one unbounded, continuously-growing JSONL log per document with no concept of a reset point — confirmed during design review to be a real UX problem (reopening a long-lived document replays its *entire* history, up to a 200-message cap, with no way to start fresh while keeping the option to resume). This task builds our deliberately different design: "New chat" writes a **divider** record into the same per-document log. On next open, only messages **after the last divider** are loaded, shown, and fed back into the model — i.e., exactly the most recent conversation segment, not everything ever said. Still one continuous log per document (no multi-session browser), just with a real, working reset point.

**Files:**
- Create: `OfficeAi.Shared/ChatStore.cs`
- Test: `OfficeAi.Shared.Tests/ChatStoreTests.cs`
- Modify: `WordAiAddIn/TaskPaneHost.cs` (wire `onOtherMessage` to `ChatStore`)
- Modify: `WordAiAddIn/web-src/entry.ts` (load history on mount, persist per turn, send divider on New Chat)

**Interfaces:**
- Produces: `OfficeAi.Shared.ChatRecord` — `struct { string Role; string Text; long Ts; }` (`Role` is `"user"`, `"assistant"`, or `"divider"`; `Text` is empty for dividers).
- Produces: `OfficeAi.Shared.ChatStore.ChatIdForFile(string filePath) -> string` (first 16 hex chars of SHA-256, mirrors genoffice's `ProjectStore.chatIdForFile`).
- Produces: `OfficeAi.Shared.ChatStore.AppendMessage(string appDataFolderName, string chatId, string role, string text)`.
- Produces: `OfficeAi.Shared.ChatStore.AppendDivider(string appDataFolderName, string chatId)`.
- Produces: `OfficeAi.Shared.ChatStore.LoadSinceLastDivider(string appDataFolderName, string chatId) -> List<ChatRecord>` — returns every record after the last divider line in the file (or every record if there is no divider yet); returns an empty list if the file doesn't exist.
- New WebMessage kinds (JS ↔ .NET, routed through `WebViewBridgeHost`'s `onOtherMessage`): `"load-history"` (JS→.NET request, no payload beyond `kind`) answered with `PostMessage({kind:"history-loaded", messages: ChatRecord[]})`; `"append-message"` (JS→.NET, `{kind, role, text}`, fire-and-forget); `"new-chat-divider"` (JS→.NET, fire-and-forget).
- Scoping decision (stated explicitly): persisted/restored records carry only `role` + `text` — tool-call structure is **not** persisted or replayed. Restored history is fed to `AgentLoop.restore()` as plain `{role:'user'|'assistant', text}` messages, which is valid input for `restore()` and sufficient for conversational continuity; it does not attempt to reconstruct exact `tool_use`/`tool_result` pairing from a prior session.

- [ ] **Step 1: Write ChatStoreTests.cs first**

```csharp
using System;
using System.IO;
using Xunit;
using OfficeAi.Shared;

public class ChatStoreTests : IDisposable
{
    private readonly string _testFolder = "OfficeAiTests_" + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _testFolder);
        if (Directory.Exists(root)) Directory.Delete(root, true);
    }

    [Fact]
    public void ChatIdForFile_IsStableAndSixteenHexChars()
    {
        string id1 = ChatStore.ChatIdForFile(@"C:\docs\report.docx");
        string id2 = ChatStore.ChatIdForFile(@"C:\docs\report.docx");
        Assert.Equal(id1, id2);
        Assert.Equal(16, id1.Length);
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEmptyForMissingFile()
    {
        var result = ChatStore.LoadSinceLastDivider(_testFolder, "nochat");
        Assert.Empty(result);
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEverythingWhenNoDividerYet()
    {
        string chatId = "chat1";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "hello");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "hi there");
        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);
        Assert.Equal(2, result.Count);
        Assert.Equal("hello", result[0].Text);
    }

    [Fact]
    public void LoadSinceLastDivider_OnlyReturnsRecordsAfterTheLastDivider()
    {
        string chatId = "chat2";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "first session question");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "first session answer");
        ChatStore.AppendDivider(_testFolder, chatId);
        ChatStore.AppendMessage(_testFolder, chatId, "user", "second session question");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "second session answer");

        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);

        Assert.Equal(2, result.Count);
        Assert.Equal("second session question", result[0].Text);
        Assert.Equal("second session answer", result[1].Text);
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEmptyImmediatelyAfterADividerWithNothingAfterIt()
    {
        string chatId = "chat3";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "question");
        ChatStore.AppendDivider(_testFolder, chatId);

        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);

        Assert.Empty(result);
    }
}
```

- [ ] **Step 2: Run tests, verify they fail** (same runner commands as Task 3 Step 3, adjusted for the new test file — expect a build failure since `ChatStore` doesn't exist).

- [ ] **Step 3: Implement ChatStore.cs**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAi.Shared
{
    public struct ChatRecord
    {
        public string Role;
        public string Text;
        public long Ts;
    }

    public static class ChatStore
    {
        public static string ChatIdForFile(string filePath)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(filePath));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }

        private static string ChatPath(string appDataFolderName, string chatId)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                appDataFolderName, "ChatHistory");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, chatId + ".jsonl");
        }

        private static void AppendRecord(string appDataFolderName, string chatId, string role, string text)
        {
            string json = JsonSerializer.Serialize(new
            {
                role,
                text,
                ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
            });
            File.AppendAllText(ChatPath(appDataFolderName, chatId), json + "\n");
        }

        public static void AppendMessage(string appDataFolderName, string chatId, string role, string text)
        {
            AppendRecord(appDataFolderName, chatId, role, text);
        }

        public static void AppendDivider(string appDataFolderName, string chatId)
        {
            AppendRecord(appDataFolderName, chatId, "divider", "");
        }

        public static List<ChatRecord> LoadSinceLastDivider(string appDataFolderName, string chatId)
        {
            string path = ChatPath(appDataFolderName, chatId);
            var all = new List<ChatRecord>();
            if (!File.Exists(path)) return all;

            foreach (string line in File.ReadAllLines(path))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try
                {
                    using (JsonDocument doc = JsonDocument.Parse(line))
                    {
                        JsonElement root = doc.RootElement;
                        all.Add(new ChatRecord
                        {
                            Role = root.GetProperty("role").GetString(),
                            Text = root.GetProperty("text").GetString(),
                            Ts = root.GetProperty("ts").GetInt64(),
                        });
                    }
                }
                catch (JsonException)
                {
                    // skip a malformed line rather than losing the whole file
                }
            }

            int lastDivider = -1;
            for (int i = 0; i < all.Count; i++)
            {
                if (all[i].Role == "divider") lastDivider = i;
            }
            return all.Skip(lastDivider + 1).Where(r => r.Role != "divider").ToList();
        }
    }
}
```

- [ ] **Step 4: Run tests, verify they pass** (5/5).

- [ ] **Step 5: Wire into TaskPaneHost.cs**

```csharp
using System.Text.Json;
using System.Windows.Forms;
using OfficeAi.Shared;

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

            string filePath = Globals.ThisAddIn.Application.ActiveDocument.FullName;
            _chatId = ChatStore.ChatIdForFile(filePath);

            _bridge = new WebViewBridgeHost(this, WordTools.Execute, "WordAiAddIn", s => _status.Text = s, OnOtherMessage);
        }

        private void OnOtherMessage(string kind, JsonElement root)
        {
            switch (kind)
            {
                case "load-history":
                    var records = ChatStore.LoadSinceLastDivider("WordAiAddIn", _chatId);
                    _bridge.PostMessage(new
                    {
                        kind = "history-loaded",
                        messages = records.ConvertAll(r => new { role = r.Role, text = r.Text }),
                    });
                    break;
                case "append-message":
                    string role = root.GetProperty("role").GetString();
                    string text = root.GetProperty("text").GetString();
                    ChatStore.AppendMessage("WordAiAddIn", _chatId, role, text);
                    break;
                case "new-chat-divider":
                    ChatStore.AppendDivider("WordAiAddIn", _chatId);
                    break;
            }
        }
    }
}
```
(`Globals.ThisAddIn.Application.ActiveDocument.FullName` requires the document to already be saved to disk — an unsaved document has no stable path to key history off of. For an unsaved document, fall back to a per-session temp id, e.g. `"unsaved-" + Process.GetCurrentProcess().Id` — matches genoffice's own `tempChatId` fallback concept; wire this fallback if testing against an unsaved scratch document surfaces it.)

- [ ] **Step 6: Wire into entry.ts**

```ts
interface OtherMessage { kind: string; [key: string]: unknown }

chrome.webview.addEventListener('message', (ev) => {
  const data = ev.data as OtherMessage & ToolResultMessage
  if (data.kind === 'tool-result') {
    // existing tool-result handling from spike 3
    return
  }
  if (data.kind === 'history-loaded') {
    const messages = data.messages as Array<{ role: 'user' | 'assistant'; text: string }>
    if (messages.length > 0) {
      ui.showHistoric(messages)
      loop.restore(messages.map((m) => ({ role: m.role, text: m.text })))
    }
  }
})

function requestHistory(): void {
  chrome.webview.postMessage({ kind: 'load-history' })
}

function persistMessage(role: 'user' | 'assistant', text: string): void {
  chrome.webview.postMessage({ kind: 'append-message', role, text })
}

requestHistory()
```

Update `onSend` and the loop's `onDone` to persist:
```ts
onSend: (text) => {
  if (loop.busy) return
  ui.addUserMessage(text)
  ui.beginAssistantMessage()
  ui.setBusy(true)
  persistMessage('user', text)
  loop.run(text)
},
```
```ts
onDone: (result) => {
  const finalText = result.text || '(no text)'
  ui.endAssistantMessage(finalText)
  ui.setBusy(false)
  persistMessage('assistant', finalText)
},
```
Update `onNewChat`:
```ts
onNewChat: () => {
  chrome.webview.postMessage({ kind: 'new-chat-divider' })
  loop.reset()
  ui.resetToEmpty()
},
```
(`loop.reset()` — already exists on `AgentLoop`, drops in-memory history; matches genoffice's own `newChat()` doing `loopRef.current?.reset()`.)

- [ ] **Step 7: Manual verification**

1. Rebuild. Close/reopen Word on a **saved** scratch `.docx`. Send one exchange (e.g. "say hello"). Close Word entirely (not just the document — quit the process) and reopen the same file. Expected: the prior exchange renders above a divider (full opacity, matching Task 5's `showHistoric`), and the input is empty and ready for a new message.
2. Without closing, click "+" (New chat), send a second exchange ("say goodbye"), close and reopen Word again. Expected: **only** the "goodbye" exchange shows as historic — the earlier "hello" exchange is gone from view (it's still in the file, just before the divider that "New chat" wrote).
3. Inspect `%LOCALAPPDATA%\WordAiAddIn\ChatHistory\<chatId>.jsonl` directly in a text editor to confirm it contains all three records (hello exchange, a `"role":"divider"` line, goodbye exchange) in one continuous file.

- [ ] **Step 8: Commit**

```bash
git add -A
git commit -m "feat: divider-bounded chat history persistence (ChatStore)"
```

---

## Phase 2 — Word tool set completion

Three of six `apps/docs` tools are already implemented from spikes 1-3 (`get_document_context`, `insert_content`, `edit_chart`). Per the feasibility report's Markdown-app finding, `apps/markdown`'s tools are not ported separately (folded into these same Word tools).

### Task 8: read_blocks tool

**Files:** Modify `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`.

**Interfaces:** `read_blocks({startIndex, endIndex})` → paragraph text for that inclusive 0-based range, one per line.

- [ ] **Step 1: Add the tool schema**
```ts
{
  name: 'read_blocks',
  description: 'Reads paragraphs [startIndex, endIndex] (0-based, inclusive) of the active document, one per line prefixed with its index.',
  inputSchema: {
    type: 'object',
    properties: { startIndex: { type: 'number' }, endIndex: { type: 'number' } },
    required: ['startIndex', 'endIndex'],
  },
},
```

- [ ] **Step 2: Implement**
```csharp
case "read_blocks":
    return ReadBlocks(input);
```
```csharp
private static ToolResult ReadBlocks(JsonElement input)
{
    int startIndex = input.GetProperty("startIndex").GetInt32();
    int endIndex = input.GetProperty("endIndex").GetInt32();
    Word.Document doc = ActiveDoc;
    Word.Paragraphs paragraphs = doc.Paragraphs;
    int count = paragraphs.Count;
    endIndex = Math.Min(endIndex, count - 1);
    if (startIndex < 0 || startIndex > endIndex)
    {
        return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "read_blocks" };
    }
    var sb = new System.Text.StringBuilder();
    for (int i = startIndex; i <= endIndex; i++)
    {
        Word.Paragraph p = paragraphs[i + 1];
        string text = p.Range.Text.TrimEnd('\r', '\a', '\n');
        sb.AppendLine($"[{i}] {text}");
    }
    return new ToolResult { Output = sb.ToString(), Summary = "read_blocks" };
}
```

- [ ] **Step 3: Manual verification** — rebuild, scratch doc with 5+ distinct paragraphs, temporarily set the mock server's demo args for `read_blocks` to `{"startIndex":0,"endIndex":2}`, confirm the tool output lists exactly those 3 paragraphs.

- [ ] **Step 4: Commit** — `git commit -m "feat(word): read_blocks tool"`

---

### Task 9: replace_blocks tool

**Files:** Modify `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`.

**Interfaces:** `replace_blocks({startIndex, endIndex, text})` → replaces that paragraph range with `text` (empty deletes).

- [ ] **Step 1: Add the tool schema**
```ts
{
  name: 'replace_blocks',
  description: 'Replaces paragraphs [startIndex, endIndex] (0-based, inclusive) with new text (empty text deletes the range).',
  inputSchema: {
    type: 'object',
    properties: { startIndex: { type: 'number' }, endIndex: { type: 'number' }, text: { type: 'string' } },
    required: ['startIndex', 'endIndex', 'text'],
  },
},
```

- [ ] **Step 2: Implement**
```csharp
case "replace_blocks":
    return ReplaceBlocks(input);
```
```csharp
private static ToolResult ReplaceBlocks(JsonElement input)
{
    int startIndex = input.GetProperty("startIndex").GetInt32();
    int endIndex = input.GetProperty("endIndex").GetInt32();
    string text = input.GetProperty("text").GetString() ?? "";
    Word.Document doc = ActiveDoc;
    Word.Paragraphs paragraphs = doc.Paragraphs;
    endIndex = Math.Min(endIndex, paragraphs.Count - 1);
    if (startIndex < 0 || startIndex > endIndex)
    {
        return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "replace_blocks" };
    }
    Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
    range.Text = text;
    return new ToolResult { Output = $"Replaced paragraphs {startIndex}-{endIndex} with: {text}", Mutated = true, Summary = "replace_blocks" };
}
```

- [ ] **Step 3: Manual verification** — demo args `{"startIndex":0,"endIndex":0,"text":"REPLACED"}`, confirm paragraph 0 becomes exactly `REPLACED`.

- [ ] **Step 4: Commit** — `git commit -m "feat(word): replace_blocks tool"`

---

### Task 10: apply_commands tool (bold/italic/heading/find-replace)

**Files:** Modify `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`.

- [ ] **Step 1: Add the tool schema**
```ts
{
  name: 'apply_commands',
  description:
    'Applies a batch of formatting/editing commands. Each command has a "kind": ' +
    '"set_bold"/"set_italic" (fields: startIndex, endIndex, value:boolean), ' +
    '"set_heading" (fields: index, level:0-9, 0=Normal style), ' +
    '"find_replace" (fields: find:string, replace:string, matchCase?:boolean).',
  inputSchema: { type: 'object', properties: { commands: { type: 'array', items: { type: 'object' } } }, required: ['commands'] },
},
```

- [ ] **Step 2: Implement**
```csharp
case "apply_commands":
    return ApplyCommands(input);
```
```csharp
private static ToolResult ApplyCommands(JsonElement input)
{
    var lines = new System.Text.StringBuilder();
    bool anyMutated = false;
    bool anyError = false;
    foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
    {
        string kind = cmd.GetProperty("kind").GetString();
        try
        {
            switch (kind)
            {
                case "set_bold":
                    SetRunProperty(cmd, (range, value) => range.Bold = value ? 1 : 0);
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "set_italic":
                    SetRunProperty(cmd, (range, value) => range.Italic = value ? 1 : 0);
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "set_heading":
                    SetHeading(cmd);
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "find_replace":
                    int replacements = FindReplace(cmd);
                    lines.AppendLine($"{kind}: {replacements} replacement(s)");
                    if (replacements > 0) anyMutated = true;
                    break;
                default:
                    lines.AppendLine(kind + ": unknown command kind"); anyError = true; break;
            }
        }
        catch (Exception ex)
        {
            lines.AppendLine(kind + ": ERROR - " + ex.Message); anyError = true;
        }
    }
    return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "apply_commands" };
}

private static void SetRunProperty(JsonElement cmd, Action<Word.Range, bool> apply)
{
    int startIndex = cmd.GetProperty("startIndex").GetInt32();
    int endIndex = cmd.GetProperty("endIndex").GetInt32();
    bool value = cmd.GetProperty("value").GetBoolean();
    Word.Document doc = ActiveDoc;
    Word.Paragraphs paragraphs = doc.Paragraphs;
    endIndex = Math.Min(endIndex, paragraphs.Count - 1);
    Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
    apply(range, value);
}

private static void SetHeading(JsonElement cmd)
{
    int index = cmd.GetProperty("index").GetInt32();
    int level = cmd.GetProperty("level").GetInt32();
    Word.Paragraph p = ActiveDoc.Paragraphs[index + 1];
    p.Range.set_Style(level == 0 ? "Normal" : "Heading " + level);
}

private static int FindReplace(JsonElement cmd)
{
    string find = cmd.GetProperty("find").GetString();
    string replace = cmd.GetProperty("replace").GetString();
    bool matchCase = cmd.TryGetProperty("matchCase", out var mc) && mc.GetBoolean();
    Word.Find findObj = ActiveDoc.Content.Find;
    findObj.ClearFormatting();
    findObj.Text = find;
    findObj.Replacement.ClearFormatting();
    findObj.Replacement.Text = replace;
    findObj.MatchCase = matchCase;
    bool found = findObj.Execute(Replace: Word.WdReplace.wdReplaceAll);
    return found ? 1 : 0;
}
```

- [ ] **Step 3: Manual verification** — test `set_bold`, `set_heading`, and `find_replace` each with a dedicated demo-args run, per the pattern established in Tasks 8-9.

- [ ] **Step 4: Commit** — `git commit -m "feat(word): apply_commands tool (bold/italic/heading/find-replace)"`

**Backlog (follow-up plan):** remaining `apply_commands` kinds (list conversion, image properties), TOC insertion, AI-author attribution in tracked changes.

---

### Task 11: Editing-mode control (Read Only / Comment Only / Track Changes / Full Autonomy) + add_comment tool

**Context:** the chat-ui mode selector (Task 5) needs real teeth. Four modes, enforced in `WordTools.cs` regardless of what the client-side tool list offers the model (defense in depth, per Global Constraints):
- **Read only** — only `get_document_context`/`read_blocks` execute; every mutating tool call is refused.
- **Comment only** — same read tools, plus the new `add_comment` tool; every *other* mutating tool is refused.
- **Track changes** — full tool set; `Document.TrackRevisions = true` is set before any mutating tool runs, so edits land as reviewable revisions.
- **Full autonomy** — full tool set; `Document.TrackRevisions = false`.

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`
- Modify: `WordAiAddIn/TaskPaneHost.cs` (route `"set-mode"` via `onOtherMessage`)
- Modify: `WordAiAddIn/web-src/entry.ts` (send `set-mode` on `onModeChange`, filter the tool list sent to the model per mode)

**Interfaces:**
- Produces: `WordTools.Mode` — a static `EditingMode` enum (`ReadOnly`, `CommentOnly`, `TrackChanges`, `FullAutonomy`), defaulting to `FullAutonomy`.
- Produces: `add_comment` tool (`{anchorText, commentText}`) — real Word comment via `Document.Comments.Add`.
- New WebMessage kind: `"set-mode"` (JS→.NET, `{kind, mode: 'readOnly'|'commentOnly'|'trackChanges'|'fullAutonomy'}`, fire-and-forget).

- [ ] **Step 1: Add the mode enum and gating to WordTools.cs**

```csharp
public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

public static class WordTools
{
    public static EditingMode Mode = EditingMode.FullAutonomy;

    private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
    {
        "get_document_context", "read_blocks",
    };

    public static ToolResult Execute(string name, JsonElement input)
    {
        try
        {
            bool isMutating = !AlwaysAllowedTools.Contains(name) && name != "add_comment";
            if (Mode == EditingMode.ReadOnly && isMutating)
            {
                return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
            }
            if (Mode == EditingMode.CommentOnly && isMutating && name != "add_comment")
            {
                return new ToolResult { Output = "Blocked: editing mode is Comment Only - use add_comment instead of editing content directly.", IsError = true, Summary = name };
            }
            if (isMutating)
            {
                ActiveDoc.TrackRevisions = (Mode == EditingMode.TrackChanges);
            }

            switch (name)
            {
                case "get_document_context": return GetDocumentContext();
                case "insert_content": return InsertContent(input);
                case "edit_chart": return EditChart(input);
                case "read_blocks": return ReadBlocks(input);
                case "replace_blocks": return ReplaceBlocks(input);
                case "apply_commands": return ApplyCommands(input);
                case "add_comment": return AddComment(input);
                default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
            }
        }
        catch (Exception ex)
        {
            return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
        }
    }

    private static ToolResult AddComment(JsonElement input)
    {
        string anchorText = input.GetProperty("anchorText").GetString();
        string commentText = input.GetProperty("commentText").GetString();
        Word.Document doc = ActiveDoc;
        Word.Range range = doc.Content;
        range.Find.ClearFormatting();
        range.Find.Text = anchorText;
        bool found = range.Find.Execute();
        if (!found)
        {
            return new ToolResult { Output = $"Could not find text to anchor comment: '{anchorText}'", IsError = true, Summary = "add_comment" };
        }
        doc.Comments.Add(range, commentText);
        return new ToolResult { Output = "Comment added.", Mutated = true, Summary = "add_comment" };
    }

    // ... existing GetDocumentContext/InsertContent/EditChart/ReadBlocks/ReplaceBlocks/ApplyCommands unchanged ...
}
```
(This replaces the `switch` that lived directly in `Execute` since Tasks 8-10 added cases to it — the gating checks now wrap that same switch rather than sitting beside it.)

- [ ] **Step 2: Add the add_comment tool schema in entry.ts**

```ts
{
  name: 'add_comment',
  description: 'Adds a Word comment anchored to the first occurrence of the given text, without changing document content. Available in every editing mode.',
  inputSchema: {
    type: 'object',
    properties: { anchorText: { type: 'string' }, commentText: { type: 'string' } },
    required: ['anchorText', 'commentText'],
  },
},
```

- [ ] **Step 3: Route `set-mode` in TaskPaneHost.cs's `OnOtherMessage`**

```csharp
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
```

- [ ] **Step 4: Wire onModeChange and client-side tool filtering in entry.ts**

```ts
let editingMode: EditingMode = 'fullAutonomy'

const READ_ONLY_TOOL_NAMES = new Set(['get_document_context', 'read_blocks'])

function toolsForMode(): typeof wordSkill.tools {
  if (editingMode === 'readOnly') {
    return wordSkill.tools.filter((t) => READ_ONLY_TOOL_NAMES.has(t.name))
  }
  if (editingMode === 'commentOnly') {
    return wordSkill.tools.filter((t) => READ_ONLY_TOOL_NAMES.has(t.name) || t.name === 'add_comment')
  }
  return wordSkill.tools
}
```
Update `wordSkill` to compute `tools` dynamically instead of a fixed array — since `AgentSkill.tools` is a plain property (not a getter) per `packages/agent-core/src/skill.ts`, wrap it: build a fresh skill object per turn, or (simpler) make `wordSkill` a `getState()`-style object where `tools` is read live at `AgentLoop.startTurn()` time. `startTurn()` reads `this.options.skill.tools` fresh on every call (see `packages/agent-core/src/loop.ts`, `tools: this.finalizing ? [] : this.options.skill.tools`), so defining `tools` as a JS getter on the skill object works without any agent-core changes:
```ts
const wordSkill: AgentSkill = {
  id: 'word-tools',
  systemPrompt: '...',
  get tools() { return toolsForMode() },
  executeTool: (call) => callDotNetTool(call.name, call.input),
}
```
And in `onModeChange`:
```ts
onModeChange: (mode) => {
  editingMode = mode
  chrome.webview.postMessage({ kind: 'set-mode', mode })
},
```

- [ ] **Step 5: Manual verification**

1. Full autonomy (default): send an edit instruction, confirm it applies directly with no revision marks.
2. Switch to Track Changes, send another edit instruction, confirm Word's Review tab shows it as a tracked insertion/deletion.
3. Switch to Read Only, ask the AI to edit something: confirm the model either doesn't attempt a mutating call (tool list was filtered) or, if it does, the tool result shows the "Blocked: editing mode is Read Only" message and the document is unchanged.
4. Switch to Comment Only, ask for a comment on specific text (`_DEMO_TOOL_ARGS["add_comment"] = {"anchorText": "<some word in the doc>", "commentText": "test comment"}`), confirm a real Word comment appears anchored to that text and no content changed.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(word): editing-mode control (read-only/comment/track-changes/full-autonomy) + add_comment tool"
```

---

### Task 12: Selection-aware scope hint + context injection

**Context:** confirmed genoffice mechanism (`apps/docs/src/renderer/ai/protocol.ts`, `getSelectionScope`/`buildDocContext`): when the user has a real range selected, genoffice injects the *actual selected content* (serialized HTML, capped at 24,000 chars) into the per-turn context sent to the model, under `"Content selected by the user (blocks X-Y):\n<html>"` — not just an index range. This task mirrors that (as plain text, since we don't have genoffice's HTML block model) and makes it visible in the UI via the scope hint, which was previously static.

**Files:**
- Modify: `WordAiAddIn/ThisAddIn.cs` (subscribe to `Application.WindowSelectionChange`)
- Modify: `WordAiAddIn/web-src/entry.ts` (cache latest selection, feed into `skill.buildContext`, update scope hint)

**Interfaces:**
- New WebMessage kind (.NET→JS, pushed unprompted via `WebViewBridgeHost.PostMessage`): `{kind: "selection-changed", hasSelection: bool, preview: string, fullText: string}` — `preview` is the first ~40 chars (for the UI pill), `fullText` is capped at 24,000 chars (for the model context, matching genoffice's `SELECTION_MAX_CHARS`).
- `AgentSkill.buildContext` (already part of the `AgentSkill` interface in `packages/agent-core/src/skill.ts`, called once per `AgentLoop.run()`) returns the cached selection's `fullText` when present, else `''`.

- [ ] **Step 1: Subscribe to selection changes in ThisAddIn.cs**

```csharp
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    _taskPaneControl = new TaskPaneHost();
    _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "Airchat Office");
    _taskPane.Width = 420;
    _taskPane.Visible = true;

    this.Application.WindowSelectionChange += Application_WindowSelectionChange;
}

private void ThisAddIn_Shutdown(object sender, EventArgs e)
{
    this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
}

private void Application_WindowSelectionChange(Word.Selection selection)
{
    _taskPaneControl.OnSelectionChanged(selection);
}
```

- [ ] **Step 2: Add OnSelectionChanged to TaskPaneHost.cs**

```csharp
public void OnSelectionChanged(Word.Selection selection)
{
    bool hasSelection = selection.Start != selection.End;
    string fullText = hasSelection ? selection.Text : "";
    if (fullText.Length > 24000) fullText = fullText.Substring(0, 24000);
    string preview = fullText.Length > 40 ? fullText.Substring(0, 40) : fullText;
    _bridge.PostMessage(new
    {
        kind = "selection-changed",
        hasSelection,
        preview,
        fullText,
    });
}
```
(Needs `using Word = Microsoft.Office.Interop.Word;` added to `TaskPaneHost.cs`'s usings.)

- [ ] **Step 3: Cache the selection and wire it into the skill's context in entry.ts**

```ts
let latestSelection: { hasSelection: boolean; preview: string; fullText: string } = {
  hasSelection: false,
  preview: '',
  fullText: '',
}

chrome.webview.addEventListener('message', (ev) => {
  const data = ev.data as OtherMessage
  if (data.kind === 'selection-changed') {
    latestSelection = data as typeof latestSelection
    ui.setScopeHint(
      latestSelection.hasSelection ? `Selection: "${latestSelection.preview}..."` : 'Whole document',
    )
  }
  // ... existing tool-result / history-loaded branches ...
})
```

Update `wordSkill` to inject the selection into context:
```ts
const wordSkill: AgentSkill = {
  id: 'word-tools',
  systemPrompt: '...',
  get tools() { return toolsForMode() },
  buildContext: () =>
    latestSelection.hasSelection ? `Content selected by the user:\n${latestSelection.fullText}` : '',
  executeTool: (call) => callDotNetTool(call.name, call.input),
}
```
(`AgentLoop.run()` already calls `this.options.skill.buildContext?.()` and appends it to the user's instruction via `formatUserMessage` — no `AgentLoop` changes needed, this is exactly the extension point genoffice's own `buildDocContext` plugs into, just on our skill instead of theirs.)

- [ ] **Step 4: Manual verification**

1. Rebuild. Open a scratch doc, don't select anything — confirm the scope hint reads "Whole document".
2. Select a sentence in the document — confirm the scope hint updates live to `Selection: "<first ~40 chars>..."` without needing to click into the task pane.
3. With that selection still active, ask the AI something referencing "the selected text" (e.g. "make the selected text bold") and confirm — via a temporary debug log of the outgoing request, or by observing correct behavior — that the actual selected text reached the model as context, not just its position.
4. Click elsewhere to collapse the selection to a caret — confirm the hint reverts to "Whole document".

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat(word): selection-aware scope hint + selected-text context injection"
```

---

## Phase 3 — Excel add-in

### Task 13: ExcelAiAddIn project scaffold

Mirrors `WordAiAddIn`'s current (post-Phase-2) structure, substituting the Excel VSTO template and `Microsoft.Office.Interop.Excel`.

**Files:**
- Create: `ExcelAiAddIn/ExcelAiAddIn.csproj`, `ThisAddIn.cs`/`.Designer.cs`/`.Designer.xml`, `Properties/AssemblyInfo.cs`, `App.config`, `TaskPaneHost.cs`, `ExcelTools.cs` (stub), `web-src/entry.ts`, `web/index.html`, `package.json`, `tsconfig.json`.

**Interfaces:**
- Consumes: `OfficeAi.Shared.WebViewBridgeHost`, `OfficeAi.Shared.ChatStore`, `mountChatUI` — identical wiring pattern to `WordAiAddIn`, including chat persistence (Task 7's design) from day one, not bolted on later. Editing-mode gating (Task 11's pattern) also applies from day one, scoped to whatever tools exist at each step (initially just readers, which are always-allowed).

- [ ] **Step 1: Copy the Excel VSTO project template as the starting point**

```bash
cp -r "C:/Program Files/Microsoft Visual Studio/2022/Community/Common7/IDE/ProjectTemplates/CSharp/Office/Addins/1033/VSTOExcel15AddInV4" /tmp/excel-template
```
Inspect `/tmp/excel-template/ExcelAddIn.csproj` for the exact `HostName="Excel"` template tokens (same shape as the Word template, substituting `Word`→`Excel`, `Microsoft.Office.Interop.Word`→`Microsoft.Office.Interop.Excel`).

- [ ] **Step 2: Author ExcelAiAddIn.csproj** from `WordAiAddIn/WordAiAddIn.csproj`'s current (Task 4+) version, substituting Excel types/HostName/HostPackage GUID (copy the real value from the Excel template, don't guess), `DebugInfoExeName` → `#Software\Microsoft\Office\15.0\Excel\InstallRoot\Path#EXCEL.EXE`, keeping the `OfficeAi.Shared` `ProjectReference` and manifest-signing properties (reuse the same dev cert thumbprint from spike 1).

- [ ] **Step 3: Author ThisAddIn.cs/.Designer.cs/.Designer.xml** — copy Word's versions, `Word.Application`→`Excel.Application`, `Microsoft.Office.Tools.Word.ApplicationFactory`→`Microsoft.Office.Tools.Excel.ApplicationFactory`. Panel title "Airchat Office" (same across all three apps).

- [ ] **Step 4: Author TaskPaneHost.cs** — same structure as Word's post-Task-7 version (chat persistence wired), calling `ExcelTools.Execute` and `"ExcelAiAddIn"` as the app-data folder name. Selection tracking (Task 12's pattern) is Excel-specific (different event: `Excel.Application.SheetSelectionChange`) — deferred to a follow-up task once the core tool set exists, not required for this scaffold.

- [ ] **Step 5: Author a stub ExcelTools.cs**
```csharp
using System.Text.Json;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static class ExcelTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
        }
    }
}
```

- [ ] **Step 6: Author web-src/entry.ts, web/index.html, package.json, tsconfig.json** — copy `WordAiAddIn`'s current versions verbatim except `wordSkill`→`excelSkill` with `tools: []` for now (Task 14+), and no editing-mode/selection wiring yet (added once Excel has tools worth gating — follow-up task, same pattern as Task 11/12).

```bash
cd ExcelAiAddIn
npm init -y
npm install --save-dev esbuild typescript @types/node
```

- [ ] **Step 7: Build and manually verify the empty scaffold loads in Excel**

```bash
npx tsc --noEmit
npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" -t:restore ExcelAiAddIn.csproj
"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug
```
Open Excel, confirm via `[System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application').COMAddIns` that `ExcelAiAddIn` shows `Connect=True`, and the task pane renders the chat-ui shell.

- [ ] **Step 8: Commit** — `git commit -m "feat(excel): ExcelAiAddIn project scaffold"`

---

### Task 14: Excel reader tools (get_workbook_context, read_range, read_cells)

**Files:** Modify `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`.

- [ ] **Step 1: Add the tool schemas**
```ts
{
  name: 'get_workbook_context',
  description: "Reads the active sheet's name, used range, and current selection address.",
  inputSchema: { type: 'object', properties: {} },
},
{
  name: 'read_range',
  description: 'Reads cell values in a rectangular range (e.g. "A1:C10"), max 2000 cells. Optional sheet name defaults to the active sheet.',
  inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
},
{
  name: 'read_cells',
  description: 'Reads specific scattered cell addresses (e.g. ["A1","C5"]).',
  inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, addresses: { type: 'array', items: { type: 'string' } } }, required: ['addresses'] },
},
```

- [ ] **Step 2: Implement**
```csharp
using System;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static class ExcelTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                switch (name)
                {
                    case "get_workbook_context": return GetWorkbookContext();
                    case "read_range": return ReadRange(input);
                    case "read_cells": return ReadCells(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static Excel.Worksheet Sheet(JsonElement input)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            if (input.ValueKind == JsonValueKind.Object && input.TryGetProperty("sheet", out var s) && s.ValueKind == JsonValueKind.String)
            {
                return (Excel.Worksheet)app.ActiveWorkbook.Sheets[s.GetString()];
            }
            return (Excel.Worksheet)app.ActiveSheet;
        }

        private static ToolResult GetWorkbookContext()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string usedRange = sheet.UsedRange.Address[false, false];
            string selection = ((Excel.Range)Globals.ThisAddIn.Application.Selection).Address[false, false];
            return new ToolResult { Output = $"Sheet: {sheet.Name}\nUsedRange: {usedRange}\nSelection: {selection}", Summary = "get_workbook_context" };
        }

        private static ToolResult ReadRange(JsonElement input)
        {
            string address = input.GetProperty("address").GetString();
            Excel.Range range = Sheet(input).Range[address];
            if (range.Cells.Count > 2000)
            {
                return new ToolResult { Output = "Range exceeds 2000-cell cap.", IsError = true, Summary = "read_range" };
            }
            object[,] values = range.Value2 as object[,];
            var sb = new System.Text.StringBuilder();
            if (values == null)
            {
                sb.Append(range.Value2 ?? "");
            }
            else
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    var cells = Enumerable.Range(1, values.GetLength(1)).Select(c => values[r, c]?.ToString() ?? "");
                    sb.AppendLine(string.Join("\t", cells));
                }
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_range" };
        }

        private static ToolResult ReadCells(JsonElement input)
        {
            Excel.Worksheet sheet = Sheet(input);
            var sb = new System.Text.StringBuilder();
            foreach (JsonElement addr in input.GetProperty("addresses").EnumerateArray())
            {
                string a = addr.GetString();
                object value = ((Excel.Range)sheet.Range[a]).Value2;
                sb.AppendLine($"{a}: {value}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_cells" };
        }
    }
}
```

- [ ] **Step 3: Manual verification** — scratch workbook with values in `A1:C3`; demo-args each tool in turn and confirm correct output.

- [ ] **Step 4: Commit** — `git commit -m "feat(excel): reader tools (get_workbook_context, read_range, read_cells)"`

---

### Task 15: Excel core mutation tool — propose_operations (set_cell, set_formula, set_range, format_range)

**Files:** Modify `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`.

- [ ] **Step 1: Add the tool schema**
```ts
{
  name: 'propose_operations',
  description:
    'Applies a batch of spreadsheet operations. Each has a "kind": ' +
    '"set_cell" (sheet?, address, value), "set_formula" (sheet?, address, formula), ' +
    '"set_range" (sheet?, address, values: value[][]), ' +
    '"format_range" (sheet?, address, bold?, italic?, numberFormat?, fillColor? - hex like "#FFFF00").',
  inputSchema: { type: 'object', properties: { operations: { type: 'array', items: { type: 'object' } } }, required: ['operations'] },
},
```

- [ ] **Step 2: Implement**
```csharp
case "propose_operations":
    return ProposeOperations(input);
```
```csharp
private static ToolResult ProposeOperations(JsonElement input)
{
    var lines = new System.Text.StringBuilder();
    bool anyMutated = false;
    bool anyError = false;
    foreach (JsonElement op in input.GetProperty("operations").EnumerateArray())
    {
        string kind = op.GetProperty("kind").GetString();
        try
        {
            switch (kind)
            {
                case "set_cell":
                    Sheet(op).Range[op.GetProperty("address").GetString()].Value2 = JsonValueToObject(op.GetProperty("value"));
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "set_formula":
                    Sheet(op).Range[op.GetProperty("address").GetString()].Formula = op.GetProperty("formula").GetString();
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "set_range":
                    SetRangeValues(op);
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                case "format_range":
                    FormatRange(op);
                    lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                default:
                    lines.AppendLine(kind + ": unknown operation kind"); anyError = true; break;
            }
        }
        catch (Exception ex)
        {
            lines.AppendLine(kind + ": ERROR - " + ex.Message); anyError = true;
        }
    }
    return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "propose_operations" };
}

private static object JsonValueToObject(JsonElement v)
{
    switch (v.ValueKind)
    {
        case JsonValueKind.String: return v.GetString();
        case JsonValueKind.Number: return v.GetDouble();
        case JsonValueKind.True: return true;
        case JsonValueKind.False: return false;
        default: return null;
    }
}

private static void SetRangeValues(JsonElement op)
{
    string address = op.GetProperty("address").GetString();
    JsonElement rows = op.GetProperty("values");
    int rowCount = rows.GetArrayLength();
    int colCount = rows[0].GetArrayLength();
    object[,] grid = new object[rowCount, colCount];
    for (int r = 0; r < rowCount; r++)
    {
        JsonElement row = rows[r];
        for (int c = 0; c < colCount; c++) grid[r, c] = JsonValueToObject(row[c]);
    }
    Excel.Range topLeft = Sheet(op).Range[address];
    topLeft.Resize[rowCount, colCount].Value2 = grid;
}

private static void FormatRange(JsonElement op)
{
    Excel.Range range = Sheet(op).Range[op.GetProperty("address").GetString()];
    if (op.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean();
    if (op.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean();
    if (op.TryGetProperty("numberFormat", out var nf)) range.NumberFormat = nf.GetString();
    if (op.TryGetProperty("fillColor", out var fc))
    {
        string hex = fc.GetString().TrimStart('#');
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
    }
}
```

- [ ] **Step 3: Manual verification** — one demo-args run per op kind (`set_cell`, `set_formula`, `set_range`, `format_range`), confirming the exact expected cell state each time.

- [ ] **Step 4: Commit** — `git commit -m "feat(excel): propose_operations - set_cell/set_formula/set_range/format_range"`

---

### Task 16: Excel structural ops (insert/delete rows/cols, add_chart) + editing-mode gating

**Files:** Modify `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/TaskPaneHost.cs`.

Now that Excel has real mutating tools, this task also brings over the Task 11 editing-mode pattern (Read Only / Comment Only / Track Changes / Full Autonomy) — Excel's track-changes equivalent is `Workbook.Highlight­ChangesOnScreen`/shared-workbook change tracking (`Workbook.HighlightChangesOptions`), which is more limited than Word's; scope this task's Track Changes mode to simply gating mutations the same way (allow/block), and note the shared-workbook-based revision UI as a smaller follow-up rather than blocking this task on it.

- [ ] **Step 1: Add the tool schema additions** (extend the `propose_operations` description from Task 15)
```ts
// append: ' "insert_rows"/"delete_rows" (sheet?, startRow:number 1-based, count:number), "insert_cols"/"delete_cols" (sheet?, startCol:number 1-based, count:number), "add_chart" (sheet?, dataRange:string, chartType?:"column"|"line"|"pie", title?:string).'
```

- [ ] **Step 2: Extend ProposeOperations's switch**
```csharp
case "insert_rows": InsertDeleteRows(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
case "delete_rows": InsertDeleteRows(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
case "insert_cols": InsertDeleteCols(op, insert: true); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
case "delete_cols": InsertDeleteCols(op, insert: false); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
case "add_chart": AddChart(op); lines.AppendLine(kind + ": ok"); anyMutated = true; break;
```
```csharp
private static void InsertDeleteRows(JsonElement op, bool insert)
{
    int startRow = op.GetProperty("startRow").GetInt32();
    int count = op.GetProperty("count").GetInt32();
    Excel.Range rows = Sheet(op).Range[$"{startRow}:{startRow + count - 1}"];
    if (insert) rows.EntireRow.Insert(); else rows.EntireRow.Delete();
}

private static void InsertDeleteCols(JsonElement op, bool insert)
{
    int startCol = op.GetProperty("startCol").GetInt32();
    int count = op.GetProperty("count").GetInt32();
    string startLetter = ColumnLetter(startCol);
    string endLetter = ColumnLetter(startCol + count - 1);
    Excel.Range cols = Sheet(op).Range[$"{startLetter}:{endLetter}"];
    if (insert) cols.EntireColumn.Insert(); else cols.EntireColumn.Delete();
}

private static string ColumnLetter(int col)
{
    string result = "";
    while (col > 0)
    {
        int rem = (col - 1) % 26;
        result = (char)('A' + rem) + result;
        col = (col - 1) / 26;
    }
    return result;
}

private static void AddChart(JsonElement op)
{
    Excel.Worksheet sheet = Sheet(op);
    string dataRange = op.GetProperty("dataRange").GetString();
    dynamic chartObjects = sheet.ChartObjects();
    dynamic chartObj = chartObjects.Add(100, 20, 400, 250);
    dynamic chart = chartObj.Chart;
    chart.SetSourceData(sheet.Range[dataRange]);
    int chartTypeCode = 51; // xlColumnClustered
    if (op.TryGetProperty("chartType", out var ct))
    {
        string t = ct.GetString();
        chartTypeCode = t == "line" ? 4 : t == "pie" ? 5 : 51;
    }
    chart.ChartType = chartTypeCode;
    if (op.TryGetProperty("title", out var title))
    {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title.GetString();
    }
}
```

- [ ] **Step 3: Add editing-mode gating** (same shape as Task 11 — `ExcelTools.Mode` static field, `AlwaysAllowedTools = {"get_workbook_context", "read_range", "read_cells"}`, wrap the tool switch in `Execute` with the same 4-way check). No Excel `add_comment`-equivalent tool in this pass — Comment Only mode simply allows no mutating tools yet for Excel (documented gap, follow-up task can add one via `Range.AddComment`/`Range.CommentThreaded`).

- [ ] **Step 4: Manual verification** — `insert_rows`/`delete_cols`/`add_chart` each with a dedicated demo-args run; repeat the Task 11-style mode check (Read Only blocks a mutating call, Full Autonomy allows it).

- [ ] **Step 5: Commit** — `git commit -m "feat(excel): insert/delete rows/cols, add_chart, editing-mode gating"`

**Backlog (follow-up plan):** the remaining ~40 `propose_operations` kinds from genoffice's real DSL (`merge_cells`, `sort_range`, `set_freeze`, `set_filter`, `add_conditional_format`, `set_data_validation`, `add_defined_name`, `protect_sheet`, `add_sheet`, `add_pivot`, `add_table`, `add_sparkline`, etc.) plus the remaining reader tools (`read_formats`, `read_sheet_features`, `find_cells`, `select_range`, `trace_precedents`/`trace_dependents`, `load_guide`) and selection-aware scope tracking for Excel (via `Excel.Application.SheetSelectionChange`, same pattern as Task 12).

---

## Phase 4 — PowerPoint add-in

### Task 17: PowerPointAiAddIn project scaffold

Same process as Task 13, substituting `VSTOPowerPoint15AddInV4` and `Microsoft.Office.Interop.PowerPoint`. `DebugInfoExeName` → `#Software\Microsoft\Office\15.0\PowerPoint\InstallRoot\Path#POWERPNT.EXE`. If the `CustomTaskPane` doesn't render on first launch (PowerPoint's task pane API is pickier about timing than Word/Excel's), fall back to creating it inside a `PowerPoint.Application.WindowActivate` handler instead of directly in `ThisAddIn_Startup` — try the simple path first.

- [ ] **Steps 1-6:** identical structure to Task 13, substituting PowerPoint types/paths.
- [ ] **Step 7:** build + manually verify via `[System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application').COMAddIns`, applying the `WindowActivate` fallback if needed.
- [ ] **Step 8: Commit** — `git commit -m "feat(powerpoint): PowerPointAiAddIn project scaffold"`

---

### Task 18: PowerPoint reader tools (get_deck_context, read_slide)

**Files:** Modify `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`.

- [ ] **Step 1: Add the tool schemas**
```ts
{
  name: 'get_deck_context',
  description: 'Reads a one-line-per-slide outline: slide index and a text preview of its shapes.',
  inputSchema: { type: 'object', properties: {} },
},
{
  name: 'read_slide',
  description: 'Reads full text of every shape on one slide (0-based index).',
  inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
},
```

- [ ] **Step 2: Implement**
```csharp
using System;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static class PowerPointTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static PowerPoint.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;

        private static string ShapeText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return shape.TextFrame.TextRange.Text;
            }
            return "";
        }

        private static ToolResult GetDeckContext()
        {
            var sb = new StringBuilder();
            int i = 0;
            foreach (PowerPoint.Slide slide in ActivePresentation.Slides)
            {
                var texts = new System.Collections.Generic.List<string>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string t = ShapeText(shape).Replace("\r", " ").Trim();
                    if (t.Length > 0) texts.Add(t);
                }
                string preview = string.Join(" | ", texts);
                if (preview.Length > 120) preview = preview.Substring(0, 120) + "...";
                sb.AppendLine($"[{i}] {preview}");
                i++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "get_deck_context" };
        }

        private static ToolResult ReadSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
            {
                return new ToolResult { Output = "Invalid slide index.", IsError = true, Summary = "read_slide" };
            }
            PowerPoint.Slide slide = slides[slideIndex + 1];
            var sb = new StringBuilder();
            int shapeIndex = 0;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                sb.AppendLine($"[{shapeIndex}] {shape.Name}: {ShapeText(shape)}");
                shapeIndex++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_slide" };
        }
    }
}
```

- [ ] **Step 3: Manual verification** — scratch deck with 2-3 text-box slides; verify `get_deck_context` and `read_slide` output.
- [ ] **Step 4: Commit** — `git commit -m "feat(powerpoint): get_deck_context, read_slide"`

---

### Task 19: PowerPoint core mutation tools + editing-mode gating

**Files:** Modify `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`.

- [ ] **Step 1: Add the tool schemas** (`set_element_text`, `set_element_style`, `set_element_transform`, `add_text_box`, `add_shape`, `delete_element` — same shapes as previously specified: `{slideIndex, shapeIndex, ...}` addressing, `left/top/width/height` in points, `color` as hex).

```ts
{ name: 'set_element_text', description: 'Replaces the text content of one shape (0-based slideIndex, 0-based shapeIndex within that slide).', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, text: { type: 'string' } }, required: ['slideIndex', 'shapeIndex', 'text'] } },
{ name: 'set_element_style', description: 'Changes text formatting of one shape without changing its text.', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, bold: { type: 'boolean' }, italic: { type: 'boolean' }, fontSize: { type: 'number' }, color: { type: 'string' } }, required: ['slideIndex', 'shapeIndex'] } },
{ name: 'set_element_transform', description: 'Moves/resizes/rotates one shape (values in points; rotation in degrees).', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, rotation: { type: 'number' } }, required: ['slideIndex', 'shapeIndex'] } },
{ name: 'add_text_box', description: 'Creates a new text box on the given slide.', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, text: { type: 'string' } }, required: ['slideIndex', 'left', 'top', 'width', 'height', 'text'] } },
{ name: 'add_shape', description: 'Creates a shape (rectangle/oval/roundRect) with optional text.', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeType: { type: 'string', enum: ['rectangle', 'oval', 'roundRect'] }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, text: { type: 'string' } }, required: ['slideIndex', 'shapeType', 'left', 'top', 'width', 'height'] } },
{ name: 'delete_element', description: 'Deletes one shape from a slide.', inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } }, required: ['slideIndex', 'shapeIndex'] } },
```

- [ ] **Step 2: Implement**
```csharp
case "set_element_text": return SetElementText(input);
case "set_element_style": return SetElementStyle(input);
case "set_element_transform": return SetElementTransform(input);
case "add_text_box": return AddTextBox(input);
case "add_shape": return AddShape(input);
case "delete_element": return DeleteElement(input);
```
```csharp
private static PowerPoint.Shape ResolveShape(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    int shapeIndex = input.GetProperty("shapeIndex").GetInt32();
    return ActivePresentation.Slides[slideIndex + 1].Shapes[shapeIndex + 1];
}

private static ToolResult SetElementText(JsonElement input)
{
    ResolveShape(input).TextFrame.TextRange.Text = input.GetProperty("text").GetString();
    return new ToolResult { Output = "Text updated.", Mutated = true, Summary = "set_element_text" };
}

private static ToolResult SetElementStyle(JsonElement input)
{
    PowerPoint.TextRange range = ResolveShape(input).TextFrame.TextRange;
    if (input.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    if (input.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    if (input.TryGetProperty("fontSize", out var fontSize)) range.Font.Size = (float)fontSize.GetDouble();
    if (input.TryGetProperty("color", out var color))
    {
        string hex = color.GetString().TrimStart('#');
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        range.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
    }
    return new ToolResult { Output = "Style updated.", Mutated = true, Summary = "set_element_style" };
}

private static ToolResult SetElementTransform(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    if (input.TryGetProperty("left", out var left)) shape.Left = (float)left.GetDouble();
    if (input.TryGetProperty("top", out var top)) shape.Top = (float)top.GetDouble();
    if (input.TryGetProperty("width", out var width)) shape.Width = (float)width.GetDouble();
    if (input.TryGetProperty("height", out var height)) shape.Height = (float)height.GetDouble();
    if (input.TryGetProperty("rotation", out var rotation)) shape.Rotation = (float)rotation.GetDouble();
    return new ToolResult { Output = "Transform updated.", Mutated = true, Summary = "set_element_transform" };
}

private static ToolResult AddTextBox(JsonElement input)
{
    PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
    float left = (float)input.GetProperty("left").GetDouble();
    float top = (float)input.GetProperty("top").GetDouble();
    float width = (float)input.GetProperty("width").GetDouble();
    float height = (float)input.GetProperty("height").GetDouble();
    PowerPoint.Shape shape = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
    shape.TextFrame.TextRange.Text = input.GetProperty("text").GetString();
    return new ToolResult { Output = "Text box added.", Mutated = true, Summary = "add_text_box" };
}

private static ToolResult AddShape(JsonElement input)
{
    PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
    string shapeType = input.GetProperty("shapeType").GetString();
    Microsoft.Office.Core.MsoAutoShapeType autoShapeType =
        shapeType == "oval" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval :
        shapeType == "roundRect" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle :
        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle;
    float left = (float)input.GetProperty("left").GetDouble();
    float top = (float)input.GetProperty("top").GetDouble();
    float width = (float)input.GetProperty("width").GetDouble();
    float height = (float)input.GetProperty("height").GetDouble();
    PowerPoint.Shape shape = slide.Shapes.AddShape(autoShapeType, left, top, width, height);
    if (input.TryGetProperty("text", out var text)) shape.TextFrame.TextRange.Text = text.GetString();
    return new ToolResult { Output = "Shape added.", Mutated = true, Summary = "add_shape" };
}

private static ToolResult DeleteElement(JsonElement input)
{
    ResolveShape(input).Delete();
    return new ToolResult { Output = "Shape deleted.", Mutated = true, Summary = "delete_element" };
}
```

- [ ] **Step 3: Add editing-mode gating** — same pattern as Task 11/16 (`PowerPointTools.Mode`, `AlwaysAllowedTools = {"get_deck_context", "read_slide"}`). No PowerPoint comment-equivalent tool in this pass (`Slide.Comments.Add` exists and is a reasonable follow-up, same shape as Word's `add_comment`).

- [ ] **Step 4: Manual verification** — `add_text_box` → `set_element_style` on it → `set_element_transform` → `add_shape` → `delete_element`, confirming each visible change on a scratch deck; repeat the mode-gating check.

- [ ] **Step 5: Commit** — `git commit -m "feat(powerpoint): core mutation tools + editing-mode gating"`

**Backlog (follow-up plan):** `execute_slide_script`, `set_element_fill`/`set_element_stroke`, image tools (`insert_web_image`/`crop_image`/`set_picture_opacity`/`replace_image`), `delete_slide`/`add_slide`, `add_chart`/`edit_chart` (same shared chart engine + `dynamic` approach as Word/Excel), `add_smartart`, table tools, `set_slide_background`, `ungroup_element`, selection-aware scope tracking (PowerPoint's selection-change event).

---

## Phase 5 — Explicitly deferred

- **LLM connection wiring** — the Settings dropdown (Task 5/6) collects base URL/API key/model but doesn't yet feed `AiProviderConfig`/`BASE_URL` in `entry.ts` (still hardcoded to the local mock server, as spikes 2-3 left it). Wiring `onSettingsSave` to actually reconfigure the transport, plus persisting those settings (likely via the same `ChatStore`-adjacent `.NET` file-storage pattern from Task 7) is a self-contained follow-up.
- **GPO/SCCM deployment packaging** — infrastructure-dependent, needs a representative air-gapped test image.
- **PDF and cross-app Markdown tooling** — explicitly excluded per the feasibility report (PDF isn't Office-plugin-shaped at all; Markdown folds into Word's OOXML/plain-text tools, already reflected in Phase 2).
- **Comment-mode tools for Excel/PowerPoint** (`Range.AddComment`, `Slide.Comments.Add`) — noted as backlog items in Phases 3/4; Word's `add_comment` (Task 11) establishes the pattern.
- **Excel/PowerPoint selection-aware scope tracking** — Word's version (Task 12) establishes the pattern (`WindowSelectionChange` → push → cache → `buildContext`); Excel's and PowerPoint's equivalents use different native events and are backlogged in Phases 3/4 rather than blocking those phases' core tool work.

---

## Self-Review Notes

- **Spec coverage**: every design-review decision from the mockup-iteration conversation has a task: no attachments (Global Constraints + Task 5 explicitly tests their absence), divider-bounded persistence (Task 7, with the exact bug in genoffice's real behavior identified and tests proving the fix), 4-mode editing control wired to real `TrackRevisions`/tool-gating (Task 11, replicated in Tasks 16/19 for Excel/PowerPoint), selection-aware scope hint + real context injection mirroring genoffice's confirmed `buildDocContext` mechanism (Task 12), settings-changes-on-Save and non-mirroring header chrome (both encoded directly in Task 5's `chat-ui.ts`/tests, sourced from the approved mockup).
- **Placeholder scan**: every code block is complete and consistent with the established patterns (COM addressing, JSON protocol, mode gating) — the three backlog sections are explicit scope boundaries, not disguised placeholders.
- **Type consistency**: `OfficeAi.Shared.ToolResult`/`ToolExecutor`/`OtherMessageHandler` (Task 3) flow unchanged through Tasks 4, 7, 11, 12 and into Excel/PowerPoint (Tasks 13, 17). `ChatRecord`/`ChatStore` (Task 7) are consumed unchanged by `TaskPaneHost.OnOtherMessage`. `EditingMode` (TS, Task 5) matches the four mode strings used in `onModeChange`/`set-mode` wiring in Task 11 and replicated in Tasks 16/19. `ChatUIHandle.showHistoric`/`setScopeHint` (Task 5) are called with matching shapes in Tasks 7 and 12 respectively.
