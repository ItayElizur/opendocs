# Architecture — officeoffice

officeoffice drives the **real desktop Office applications** (Word, Excel, PowerPoint,
Outlook) via VSTO + COM interop
(`Microsoft.Office.Interop.{Word,Excel,PowerPoint,Outlook}`) — there is no web renderer
or clone involved. The chat UI runs in a WebView2 page inside a CustomTaskPane; tool
calls cross a `chrome.webview.postMessage` ⇄ `CoreWebView2.PostWebMessageAsJson` JSON
bridge (`OfficeAi.Shared/ToolProtocol.cs`) into C# handlers that call the COM object
model directly — there is no Electron/IPC hop. Word/Excel/PowerPoint attach one pane per
document window (keyed by `Hwnd`); Outlook attaches one pane per `Explorer` window
(keyed by COM identity, since `Explorer` has no `Hwnd`) and nothing to Inspectors — see
[`docs/tools/outlook.md`](tools/outlook.md).

`shared/web-src/{agent-core,ai-provider}` provide the agent loop and multi-provider
transport (`genspark`/`anthropic`/`gemini`/`deepseek`/`openai`/`custom`), `maxTurns`
default of 8. Provider/model/key selection, the settings screen, and the transport live
once in `shared/web-src/app-shell/` (`getSettings` / `makeTransport` / `onSettingsSave`,
persisted in WebView `localStorage`), shared by all four add-ins including Outlook —
`makeTransport` routes via `streamForProvider` to whichever provider the user picks in
the settings dropdown, with a "Test connection" button.

There is **no web-search/image-generation/media-analysis capability** anywhere in the
repo. This is a deliberate scope decision — the deployment target is air-gapped (see
`add_image`/`replace_image`'s explicit rejection of remote URLs in the per-app docs) —
not an oversight. **`get_attachment` is a partial exception:** Outlook's `get_attachment`
saves an email attachment to a local file and extracts text from text-family and
OpenXML (`.docx/.xlsx/.pptx`) types via `OfficeAi.Shared/AttachmentText/` — but never
fetches anything remote, never handles PDF or images, and cannot feed a binary to the
model. See [`docs/tools/outlook.md`](tools/outlook.md).

## Editing-mode gating

Every add-in enforces a shared `OfficeAi.Shared.EditingMode` enum (`ReadOnly |
CommentOnly | TrackChanges | FullAutonomy`), filtered client-side (which tools are
advertised to the model) *and* re-enforced server-side in each `Tools.Execute()`
(mutating tools blocked outright in Read Only / Comment Only mode, regardless of what
the model requests). Outlook repurposes all four tiers with its own meaning (Read only /
Draft only / Automate approvals / Full autonomy) rather than dropping two — see
[`docs/tools/outlook.md`](tools/outlook.md) for the mapping.

These two enforcement points (client-side `entry.ts` tool lists, server-side C# gate
sets) are maintained by hand in two different languages and **will silently drift** if a
tool is added to one side and not the other — see `CLAUDE.md` for the concrete example
and the rule to follow when touching either side.

## Provenance / data-source enforcement

There is no `dataSource`/provenance-enforcement mechanism anywhere in officeoffice (a
chart tool could in principle let the model claim data came from somewhere it didn't). A
product-owner decision opted for a prompt-only mitigation rather than a code-level
guardrail, pending real sign-off.

## Schema-vs-implementation discipline

Every tool's wire-level JSON Schema (what the model actually sees) is required to match
what its C# handler reads and does — not a bare `{type: 'object'}` with the real
contract living only in prose, and not a value that's silently accepted but ignored. A
full audit across all three non-Outlook add-ins' gateway tools (Word's `apply_commands`,
Excel's `propose_operations`, PowerPoint's per-tool schemas) was performed once and
found two recurring bug classes — silent no-op with false success, and undocumented
schema — both since fixed; the per-app tool tables in `docs/tools/` reflect the
fixed, current state directly. If you find a new mismatch, fix it and update the
relevant per-app table in the same change — don't leave a dated note describing a
problem that's since been fixed.
