# officeoffice (OpenDocs)

VSTO add-ins that put an LLM chat pane directly into Word, Excel, PowerPoint, and
Outlook — Word/Excel/PowerPoint/Outlook run as real, already-installed desktop
Microsoft Office applications, driven via COM interop
(`Microsoft.Office.Interop.{Word,Excel,PowerPoint,Outlook}`), not a web renderer or a
clone. The chat UI is a WebView2 page hosted in a CustomTaskPane; tool calls cross a
`chrome.webview.postMessage` ⇄ `CoreWebView2.PostWebMessageAsJson` JSON bridge into C#
handlers that call the Office object model directly. The deployment target is
**air-gapped** — there is no web search, image generation, or media analysis anywhere
in the repo, and every tool that touches a file (`add_image`, `replace_image`,
`get_attachment`, …) is local-file-only, never a remote URL.

For the full, current-state catalog of every tool the AI can call in each app — read
tools, editing tools, editing-mode gating, known gaps — see
**[`docs/ai-tool-surface.md`](docs/ai-tool-surface.md)**. If you're an AI coding agent
working in this repo, also read **[`CLAUDE.md`](CLAUDE.md)** first.

## Layout

| Path | What it is |
|---|---|
| `WordAiAddIn/`, `ExcelAiAddIn/`, `PowerPointAiAddIn/`, `OutlookAiAddIn/` | One VSTO add-in project per Office app. Each is a classic `.csproj` (see Building, below) with its own `web-src/entry.ts` (tool definitions, system prompt, starters) and C# `*Tools*.cs` partial-class files implementing the tool handlers over COM. |
| `OfficeAi.Shared/` | Cross-app C# library: shared pure logic (`TextUtil`, `ColorUtil`, `ShapeTypes`, `ChartTypes`, `GeometryUtil`, `JsonUtil`, the `EditingMode` enum, the WebView JSON tool protocol) and a few COM-touching helpers that still need to be split per app (e.g. Outlook's EWS layer). New logic that doesn't need COM types belongs here by default. |
| `OfficeAi.Shared.Tests/` | The one unit-test project in the repo (`dotnet test`), covering `OfficeAi.Shared`'s pure logic. |
| `shared/chat-ui` | The chat panel's UI code (`chat-ui.ts` + CSS), shared by all four add-ins, with its own `vitest` suite. |
| `shared/web-src` | `agent-core`/`ai-provider` (copied from the sibling `genoffice` project — same `AgentLoop`, same multi-provider transport types) and `app-shell` (the shared WebView2 bridge, settings screen, and provider transport wiring every `entry.ts` calls into via `startAddIn`). |
| `deploy/` | `package.ps1` (build all four add-ins in Release, stage a signed, zippable package) and `install.ps1`/`uninstall.ps1` (run on a target machine to register/unregister the add-ins — no internet access required). |
| `docs/` | `ai-tool-surface.md` (the tool reference), its changelog, and `superpowers/` (a dated archive of completed planning/verification work — see `docs/superpowers/README.md`). |
| `tools/` | Small repo-maintenance scripts, e.g. `split-partial.py` (splits an oversized `*Tools.cs` file into partial-class files while preserving the exact member set). |

There is no `.sln` file — each `.csproj` is built individually.

## Building

This is a classic VSTO project set: **Visual Studio 2022** (Community or better), with
the Office/SharePoint development workload, is the expected toolchain. There is no
`.sln`; each `.csproj` is built independently:

```
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" <App>AiAddIn\<App>AiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Each add-in's TypeScript (`web-src/entry.ts`) is bundled with esbuild **before**
MSBuild runs — there is no MSBuild pre-build step that does this for you:

```
cd <App>AiAddIn
node_modules/.bin/esbuild web-src/entry.ts --bundle --outfile=web/bundle.js \
  --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts \
  --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts \
  --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts \
  --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts \
  --target=chrome100 --format=iife --sourcemap
```

To build and stage a distributable package (Release config, signs manifests, exports
the public half of the signing certificate) rather than building for local debugging:

```
deploy\package.ps1                # all four add-ins
deploy\package.ps1 -App Word      # just one
```

`deploy\install.ps1` (run from inside the resulting package, on the target machine)
checks for .NET Framework 4.8, the VSTO 2010 Runtime, and the WebView2 Runtime, trusts
the signing certificate for the current user, and registers each add-in — no admin
rights or internet access needed.

## Testing

C# unit tests (the pure logic in `OfficeAi.Shared`):

```
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj
```

Shared chat-UI tests (`shared/chat-ui`, vitest — there's no `npm test` script defined,
so invoke vitest directly):

```
cd shared/chat-ui
npx vitest run
```

There is **no automated test coverage for the COM-calling code itself** (the
`*Tools*.cs` handlers) — that layer can only really be exercised against a live Office
application. See `docs/ai-tool-surface.md`'s per-app "unverified"/"not confirmed
against live Office" notes, and `docs/superpowers/verification/` for the historical
record of what has and hasn't been manually checked.

## Further reading

- **[`docs/ai-tool-surface.md`](docs/ai-tool-surface.md)** — the detailed, current-state
  reference for every tool the AI can call in each app, including the `EditingMode`
  gating tiers and known gaps vs. the sibling `genoffice` project.
- **[`docs/ai-tool-surface-changelog.md`](docs/ai-tool-surface-changelog.md)** — the
  dated history behind that reference.
- **[`docs/superpowers/`](docs/superpowers/README.md)** — an archive of past planning
  and verification documents for completed work, not live instructions.
