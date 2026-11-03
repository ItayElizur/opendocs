# OpenDocs

**AI that works directly inside Microsoft Office.**

OpenDocs puts an AI chat panel inside **Word, Excel, PowerPoint, and Outlook**. Describe what you want in plain language, and OpenDocs performs the operation directly in the Office application you already have open.

It doesn't export your document, modify a copy, or replace Office with a web editor.

**It operates on the actual Office document through the native desktop object model.**

So when OpenDocs adds a table, you get a real Word table. When it edits a presentation, it edits the actual slides and slide masters. When it makes a change with Track Changes enabled, you get normal tracked changes. And when you press **Ctrl+Z, Office undoes the change normally.**

---

## See what it can do

### Word

> **"Go through section 3 and leave a comment on every sentence that makes a claim without a source."**

OpenDocs can comment without changing the underlying text.

> **"Rewrite the executive summary to half the length, with Track Changes on."**

Review the result exactly as you would review a colleague's edits.

> **"Add a table of contents and fix the heading levels."**

OpenDocs creates a real Word table of contents that behaves like one.

### Excel

> **"Build a pivot table of revenue by region and quarter, plus a margin column."**

Create and manipulate real Excel pivot tables, including calculated fields.

> **"Chart this range as a stacked column and put it next to the summary."**

Work with charts and worksheet layout directly.

> **"Add data validation to column D so only the values in the Status list are accepted."**

Make structural changes to the workbook, not just changes to cell contents.

### PowerPoint

> **"Replace the logo on the slide master and move the footer text to the left."**

Edit the master itself so the change applies across the presentation.

> **"Copy the pricing table from slide 4 to slides 7 and 9, same formatting."**

Manipulate slides and their objects directly.

> **"Add speaker notes to every slide from the bullet points."**

Work with presentation content beyond what is visible on the slide.

### Outlook

> **"Find the emails from last week about the Berlin contract and summarize them."**

Search and work with your actual Outlook mailbox.

> **"Find a 30 minute slot for me, Dana and Sam on Thursday or Friday and draft the invite."**

OpenDocs can prepare the meeting while leaving the final send action under your control.

> **"Draft a reply declining the second meeting and suggesting next week."**

Work with email as part of the same AI workflow.

---

## Why OpenDocs?

### Native Office, not Office in a browser

Most AI Office integrations are built around **Office.js** and a web-based task pane.

OpenDocs takes a different approach: its VSTO add-ins communicate with the installed Office applications through **COM**, the same interface used by Office automation and macros.

That gives OpenDocs access to capabilities that are difficult or impossible to expose through a browser-based Office integration, including things such as:

* Word comments and Track Changes
* Real Word tables and tables of contents
* Excel pivot tables and calculated fields
* PowerPoint slide masters
* Presentation objects and formatting
* Outlook mailbox and calendar operations

The result is important:

**The AI is operating Office itself, rather than creating an approximation of Office functionality.**

---

## One AI interface across Office

OpenDocs provides the same chat experience across:

**Word · Excel · PowerPoint · Outlook**

A shared interface, provider configuration, and permission model let you move between applications without learning a different AI tool for each one.

---

## Built for local and web-gapped environments

OpenDocs is particularly well suited to **web-gapped, restricted, and on-premises environments** where sending documents or conversations to cloud services is not an option.

The application runs entirely on the user's Windows machine, with **no OpenDocs server or internet connection required**. The AI provider can be hosted locally or inside your organization's network, including an OpenAI-compatible endpoint backed by a self-hosted model such as **vLLM**.

OpenDocs deliberately does not depend on built-in web search, image generation, or external media services. Its file tools operate on local files, and the only AI traffic is the request to the provider you configure.

This makes it possible to deploy the complete workflow inside a closed environment.

**Your Office documents stay in your environment, and you choose where the AI model runs.**

This is especially useful for organizations with strict **data residency, security, or network-isolation requirements**.

---

## Built with safety boundaries

AI that can modify documents needs more than a chat interface.

OpenDocs has explicit **editing modes** that control what the AI is allowed to do.

| Mode              | What the AI can do                                   |
| ----------------- | ---------------------------------------------------- |
| **Read Only**     | Inspect information without modifying Office         |
| **Comment Only**  | Add comments without changing document content       |
| **Track Changes** | Make edits through Office's tracked-change mechanism |
| **Full Autonomy** | Perform the operations allowed by the application    |

Outlook has its own corresponding modes:

| Mode                   | What the AI can do                                |
| ---------------------- | ------------------------------------------------- |
| **Read Only**          | Inspect mailbox and calendar information          |
| **Draft Only**         | Create drafts without sending                     |
| **Automate Approvals** | Perform configured approved actions               |
| **Full Autonomy**      | Perform the operations allowed by the application |

These aren't merely UI settings.

The chat client only exposes tools permitted by the current mode, and the add-in independently enforces the same restriction.

**A prompt cannot simply ask the model to bypass the selected permission level.**

---

## Use the model you want

OpenDocs is not tied to a single AI provider.

It supports providers including:

* Anthropic
* OpenAI
* Gemini
* DeepSeek
* OpenAI-compatible endpoints
* Self-hosted models such as **vLLM**

Choose the provider from the settings interface and test the connection.

This makes OpenDocs suitable both for cloud-based models and for organizations that need to run models **inside their own infrastructure**.

---

## No cloud service required

OpenDocs is designed to run directly on the user's machine.

There is:

* **No OpenDocs backend**
* **No Docker container required**
* **No hosted web application**
* **No external search service**
* **No required internet connection**

The only external communication is with the AI provider you configure, which can itself be hosted entirely inside your network.

---

## Simple deployment

OpenDocs is distributed as a signed package with a PowerShell installer.

Installation is:

* **Per-user**
* **No administrator privileges required**
* **No internet connection required**
* **No server to operate**

The installer checks for the required Office/.NET/WebView2 components, registers the add-ins, and configures the signing certificate for the current user.

---

## How it differs from other Office AI add-ins

There are already open-source projects that add AI to Microsoft Office. OpenDocs takes a different approach in several important ways.

### Desktop object model instead of Office.js

Office.js runs inside a web view and exposes only part of what desktop Office can do.

OpenDocs uses the native desktop object model through COM, providing access to functionality such as pivot calculated fields, slide masters, and shared calendars.

### All four major Office applications

Many alternatives focus on Word, or Word + Excel + PowerPoint.

OpenDocs covers:

**Word · Excel · PowerPoint · Outlook**

All four share one chat interface, one provider configuration, and one permission model.

### No hosting infrastructure

Some alternatives require a Docker container, Python backend, CDN-hosted page, or separately deployed service.

OpenDocs runs on the user's machine.

### Permissions are actually enforced

Many tools provide a UI setting that controls how much freedom the AI appears to have.

OpenDocs enforces the selected mode in the add-in itself. The AI cannot simply request a tool that the current mode does not permit.

### No built-in internet features

OpenDocs intentionally leaves out built-in web search and other external services.

That makes it a natural fit for **closed, restricted, and web-gapped environments**.

### General purpose

OpenDocs is not built around a single profession, workflow, or company's backend.

It is a general-purpose AI interface for desktop Microsoft Office.

The trade-off is deliberate: **OpenDocs currently runs on Windows with desktop Office.**

---

## Architecture

At a high level, each Office application has its own VSTO add-in. A shared WebView2-based chat interface communicates with a shared AI/provider layer, while application-specific tools translate AI requests into operations on the native Office object model.

```text
                         ┌─────────────────────┐
                         │     AI Provider     │
                         │                     │
                         │ OpenAI / Anthropic  │
                         │ Gemini / vLLM / ... │
                         └──────────┬──────────┘
                                    │
                           Shared provider layer
                                    │
                         ┌──────────▼──────────┐
                         │     Chat / Agent    │
                         │      WebView2       │
                         └──────────┬──────────┘
                                    │
                         Shared tool protocol
                                    │
             ┌──────────────────────┼──────────────────────┐
             │                      │                      │
       ┌─────▼─────┐          ┌─────▼─────┐          ┌─────▼─────┐
       │   Word    │          │   Excel   │          │PowerPoint │
       │   VSTO    │          │   VSTO    │          │   VSTO    │
       └─────┬─────┘          └─────┬─────┘          └─────┬─────┘
             │                      │                      │
             ▼                      ▼                      ▼
        Word COM                Excel COM             PowerPoint COM

                              ┌─────────────┐
                              │   Outlook   │
                              │    VSTO     │
                              └──────┬──────┘
                                     │
                                     ▼
                                Outlook COM
```

The repository contains:

* `WordAiAddIn/` — Word integration
* `ExcelAiAddIn/` — Excel integration
* `PowerPointAiAddIn/` — PowerPoint integration
* `OutlookAiAddIn/` — Outlook integration
* `OfficeAi.Shared/` — shared C# logic and tool protocol
* `OfficeAi.Shared.Tests/` — unit tests for shared pure logic
* `shared/chat-ui/` — shared chat interface
* `shared/web-src/` — shared agent/provider and WebView2 infrastructure
* `deploy/` — packaging and installation
* `docs/` — architecture and tool documentation
* `tools/` — repository-maintenance scripts

There is no `.sln` file. Each add-in is built individually from its own `.csproj`.

---

## Building

OpenDocs is a classic VSTO project set.

### Requirements

* Windows
* Visual Studio 2022 Community or later
* Office/SharePoint development workload
* Node.js
* Microsoft Office desktop applications
* .NET Framework 4.8
* VSTO 2010 Runtime
* WebView2 Runtime

### Build an add-in

For example:

```powershell
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" `
    WordAiAddIn\WordAiAddIn.csproj `
    -t:Build `
    -p:Configuration=Debug `
    -v:minimal
```

Each add-in's TypeScript entry point is bundled with esbuild **before** MSBuild runs:

```powershell
cd WordAiAddIn

node_modules/.bin/esbuild web-src/entry.ts `
    --bundle `
    --outfile=web/bundle.js `
    --alias:@officeai/agent-core=../shared/web-src/agent-core/index.ts `
    --alias:@officeai/ai-provider=../shared/web-src/ai-provider/index.ts `
    --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts `
    --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts `
    --target=chrome100 `
    --format=iife `
    --sourcemap
```

### Create a distributable package

Build and stage all four add-ins:

```powershell
deploy\package.ps1
```

Or package a single application:

```powershell
deploy\package.ps1 -App Word
```

The packaging script builds the Release configuration, signs the manifests, and prepares the distributable package.

### Install

`deploy\install.ps1`, run from the resulting package on the target machine, checks for:

* .NET Framework 4.8
* VSTO 2010 Runtime
* WebView2 Runtime

It then trusts the signing certificate for the current user and registers the add-ins.

**No administrator privileges or internet access are required.**

---

## Testing

### C# unit tests

The pure logic in `OfficeAi.Shared` is covered by the unit-test project:

```powershell
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj
```

### Chat UI tests

The shared chat UI uses Vitest:

```powershell
cd shared/chat-ui
npx vitest run
```

### COM integration testing

There is **no automated test coverage for the COM-calling code itself** (`*Tools*.cs`).

That layer needs to be exercised against a live Office application. The individual application documentation records the current verification status of its tools, including any functionality that has not yet been confirmed against live Office.

---

## Documentation

* **[`docs/architecture.md`](docs/architecture.md)** — how the COM/WebView2 bridge, shared provider/transport layer, and editing-mode gating work.
* **[`docs/tools/`](docs/tools/)** — detailed reference for the AI tools available in Word, Excel, PowerPoint, and Outlook.
* **[`CLAUDE.md`](CLAUDE.md)** — guidance for AI coding agents working in this repository.

---

## Project status

OpenDocs is an active project. The COM tool layer is continuously being expanded and validated against live Office applications.

If you are interested in **AI that operates the real desktop Office environment rather than approximating it through a web interface**, especially in **local, on-premises, restricted, or web-gapped environments**, OpenDocs is designed for exactly that use case.
