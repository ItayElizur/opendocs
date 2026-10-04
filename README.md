# OpenDocs

**An AI assistant that lives inside Word, Excel, PowerPoint, and Outlook — not next to
them.**

OpenDocs adds a chat pane directly to the real, already-installed Microsoft Office
desktop apps you already use. There's no separate app to learn, no copy-pasting
content into a browser tab, and no document conversion step that mangles your
formatting: the assistant reads and edits your actual Word document, Excel workbook,
PowerPoint deck, or Outlook mailbox through the same COM automation layer Office's own
macros use — so a table it adds is a real native table, a comment it leaves is a real
native comment, and an edit it makes shows up in Track Changes like anyone else's.

**Why it's different:**

- **Drives the real app, not a clone.** No from-scratch document renderer to fall
  behind Office's own feature set — every edit goes through the genuine Word/Excel/
  PowerPoint/Outlook object model, so it looks and behaves exactly like a native
  change, including undo/redo.
- **You decide how much autonomy it gets.** A four-tier editing-mode control — Read
  Only, Comment Only, Track Changes, Full Autonomy (Outlook: Read only, Draft only,
  Automate approvals, Full autonomy) — governs what the assistant is allowed to touch,
  enforced on both ends: the client only offers tools for the current tier, and the
  add-in independently refuses anything beyond it even if asked.
- **Built for air-gapped and on-prem environments.** No web search, no image
  generation, no media analysis, and every file-touching tool is local-file-only —
  nothing leaves your machine except the chat request itself, to whichever AI
  provider you pick.
- **Bring your own model.** Anthropic, OpenAI, Gemini, DeepSeek, Genspark, or any
  OpenAI-compatible endpoint (self-hosted vLLM, etc.) — switch providers from the
  settings panel, with a one-click connection test.
- **Deep, real tool coverage.** Native Word comments and a true auto-paginating TOC,
  Excel pivot tables with calculated fields, PowerPoint Slide Master editing and
  native-clipboard cross-slide shape copy, and full Outlook mailbox/calendar
  automation — including shared calendars, recurring series, and meeting responses.
  See [`docs/tools/`](docs/tools/) for the complete, current-state catalog per app.

For how it's built — the COM/WebView2 bridge, the shared provider transport, and the
editing-mode gating — see [`docs/architecture.md`](docs/architecture.md). If you're an
AI coding agent working in this repo, read [`CLAUDE.md`](CLAUDE.md) first.

## Layout

| Path | What it is |
|---|---|
| `WordAiAddIn/`, `ExcelAiAddIn/`, `PowerPointAiAddIn/`, `OutlookAiAddIn/` | One VSTO add-in project per Office app. Each is a classic `.csproj` (see Building, below) with its own `web-src/entry.ts` (tool definitions, system prompt, starters) and C# `*Tools*.cs` partial-class files implementing the tool handlers over COM. |
| `OfficeAi.Shared/` | Cross-app C# library: shared pure logic (`TextUtil`, `ColorUtil`, `ShapeTypes`, `ChartTypes`, `GeometryUtil`, `JsonUtil`, the `EditingMode` enum, the WebView JSON tool protocol) and a few COM-touching helpers that still need to be split per app (e.g. Outlook's EWS layer). New logic that doesn't need COM types belongs here by default. |
| `OfficeAi.Shared.Tests/` | The one unit-test project in the repo (`dotnet test`), covering `OfficeAi.Shared`'s pure logic. |
| `shared/chat-ui` | The chat panel's UI code (`chat-ui.ts` + CSS), shared by all four add-ins, with its own `vitest` suite. |
| `shared/web-src` | `agent-core`/`ai-provider` (copied from the sibling `genoffice` project — same `AgentLoop`, same multi-provider transport types) and `app-shell` (the shared WebView2 bridge, settings screen, and provider transport wiring every `entry.ts` calls into via `startAddIn`). |
| `deploy/` | `package.ps1` (build all four add-ins in Release, stage a signed, zippable package) and `install.ps1`/`uninstall.ps1` (run on a target machine to register/unregister the add-ins — no internet access required). |
| `docs/` | `architecture.md` (how the system is put together), `tools/` (one tool reference per app), and `comparison-with-genoffice.md` (high-level gaps/advantages vs. the sibling project). |
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
application. See each per-app doc under `docs/tools/` for its "unverified"/"not
confirmed against live Office" notes — Outlook's is the most extensive
(`docs/tools/outlook.md`'s "Unproven at runtime" section).

## Further reading

- **[`docs/architecture.md`](docs/architecture.md)** — how the system is put together:
  the COM/WebView2 bridge, the shared provider/transport layer, and the editing-mode
  gating that applies across all four apps.
- **[`docs/tools/`](docs/tools/)** — the detailed, current-state reference for every
  tool the AI can call, one file per app (`word.md`, `excel.md`, `powerpoint.md`,
  `outlook.md`).
- **[`docs/comparison-with-genoffice.md`](docs/comparison-with-genoffice.md)** — the
  high-level gaps and advantages vs. the sibling `genoffice` project.
