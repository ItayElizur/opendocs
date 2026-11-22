# Architecture — OpenDocs

OpenDocs drives the **real desktop Office applications** (Word, Excel, PowerPoint,
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
transport (`anthropic`/`gemini`/`deepseek`/`openai`/`custom`), `maxTurns`
default of 8. Provider/model/key selection, the settings screen, and the transport live
once in `shared/web-src/app-shell/` (`getSettings` / `makeTransport` / `onSettingsSave`,
persisted in WebView `localStorage`, except the API keys, which the host keeps DPAPI-encrypted: see `settings.ts.md`), shared by all four add-ins including Outlook —
`makeTransport` routes via `streamForProvider` to whichever provider the user picks in
the settings dropdown, with a "Test connection" button. `streamForProvider`
(`ai-provider/stream.ts`) streams one model turn through the Vercel AI SDK (`ai`,
`@ai-sdk/anthropic`, `@ai-sdk/google`, and `@ai-sdk/openai-compatible` for
openai/deepseek/custom, pinned in `shared/package.json`), with no SDK retries and no SDK
tool execution: `AgentLoop` still runs the tools, the turn cap, compaction and cancel.
`shared/chat-ui/transport-contract.test.ts` pins what the transport reports and sends for
every provider; run it after an SDK upgrade.

There is **no web-search/image-generation/media-analysis capability** anywhere in the
repo. This is a deliberate scope decision — the deployment target is air-gapped (see
the explicit rejection of remote URLs in `add_image` and in `edit_picture`'s `add`/`replace` in the per-app docs) —
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
Draft only / Actions with approval / Full autonomy) rather than dropping two — see
[`docs/tools/outlook.md`](tools/outlook.md) for the mapping.

`save` (Word, Excel, PowerPoint) is defined once in `shared/web-src/app-shell/sharedTools.ts`, which exports it in
`SHARED_TOOLS` / `SHARED_TOOL_DISPLAY`; all three `entry.ts` files spread those in, so another shared tool is one more
entry there and no per-app change. It needs no tier-list entry: it is advertised from Track Changes up
because it is not read-only, and each C# gate blocks it below that. Its Save As path is a modal dialog on the
UI thread that Stop cannot cancel. The C# side is shared too: `OfficeAi.Shared/SaveTool.cs` holds the flow and result
messages, each app supplies only its two COM steps, and `saveAs` must be true or false (anything else is an error).

These two enforcement points (client-side `entry.ts` tool lists, server-side C# gate
sets) are maintained by hand in two different languages and **will silently drift** if a
tool is added to one side and not the other — see `CLAUDE.md` for the concrete example
and the rule to follow when touching either side.

### Modes switched off by the administrator

`policy.json` in the install folder (next to the per-app folders; `deploy/policy.json` is the copy the installer ships)
lists modes to switch off per app: `{ "disabledModes": { "Outlook": ["fullAutonomy"], "Word": [] } }`. The app keys are
`Word`, `Excel`, `PowerPoint` and `Outlook` (an app left out has every mode on), and the modes use the wire names
`readOnly`, `commentOnly`, `trackChanges`, `fullAutonomy` (in Outlook these are Read only / Draft only / Automate
approvals / Full autonomy). It works on whole modes, not on individual tools. `OfficeAi.Shared/ModePolicy.cs` reads it
once per Office process and picks the app's list by the add-in's folder name. A file that cannot be read or parsed, or
that names an unknown app or mode, switches every mode off in every app rather than leaving a mode open by mistake. The
host passes the app's list to the page as `window.__opendocsDisabledModes`.
`bootstrap.ts` shows those modes grayed out and unselectable in the menu and the settings scope control, and starts in
the app's default mode if that one is still on, otherwise the first one left. With no mode left, the mode button is hidden, the model is
offered no tools and the chat says the administrator turned the modes off. The server side is enforced once in
`PaneHostBase.ExecuteToolUnderPolicy`, ahead of each app's own tier gate. When any mode is disabled, it refuses every
tool call while the pane is in a disabled mode or has not reported a mode yet (`ModePolicy.Blocks`). A disabled mode
is never passed on to the app's `SetMode`. Before every call the pane stores its own mode again under its current chat
id. That id changes on a new document's first save, when the app's `ModeFor` would otherwise fall back to Full
autonomy, and Outlook's panes share one mode per mailbox. Changes to the policy take effect when Office is restarted.

## Direct tool calls (Developer tools)

An opt-in setting (`developerTools`, saved with the other panel settings, off by default) turns on a `!` palette in the shared chat UI (`shared/chat-ui/tool-palette.ts` for the rules, `chat-ui.ts` for the popup). `!name({json})` runs one tool without the model: `bootstrap.ts` sends it through `callDotNetTool` (the same bridge the model's calls use), so the server-side editing-mode gate applies unchanged and nothing here bypasses it. The client also locks tools outside the current mode, using the same `availableForMode()` sets that filter what the model is offered. The finished call is added to the loop's history (`AgentLoop.recordToolExchange`), so the next model turn sees what was run. No tool or tier list changes are involved.

## Skills

A skill is a folder in the Agent Skills format that Claude Code uses: a `SKILL.md` (frontmatter `name`,
`description`, optional `argument-hint`, `disable-model-invocation`, `user-invocable`, then markdown instructions),
plus any references, templates or images next to it. Each app has its own skills, in two folders: the machine's
`InstallRoot\skills\<App>\` (installed by IT, next to `policy.json`) and the user's
`%LOCALAPPDATA%\<App>AiAddIn\Skills\`. A machine skill wins over a user skill of the same name.

- **Reading.** `OfficeAi.Shared/SkillCatalog.cs` reads both folders once per Office process and the host hands the
  catalog to the page as `window.__opendocsSkills`, before its scripts run (see `SkillCatalog.cs.md`).
- **The model.** `shared/web-src/app-shell/skills.ts` builds every piece of skill text. The system prompt lists the
  skills the model may load (`catalogPrompt`). Two tools exist only when the app has at least one skill:
  `load_skill {name}`, answered by the page from the catalog, and `read_skill_file {skill, path}`, answered by
  `PaneHostBase.ExecuteToolUnderPolicy` after the policy check. `read_skill_file` reads only files from the skill's
  scanned list, never through a link. Both only read, so `bootstrap.ts` counts them read-only in every mode and
  neither is in any app's C# tier set or `entry.ts` list (like `preview_action`; `TierListsAgreeTests` is unaffected).
- **The user.** A message that starts with `/` opens the skill list (the `!` palette's popup, keyboard and styling,
  without Developer tools). `/name args` is shown and saved as typed; the model gets
  `Base directory for this skill: <dir>`, the instructions and the skill's file list, then the arguments (or the
  arguments in place of `$ARGUMENTS`). If the conversation history still holds the skill's full text, the model gets a
  short re-invocation note with the new arguments instead. That check reads the live history, so a skill that
  compaction or trimming cut is sent in full again. A restored chat expands its saved `/name` lines again, in order.
- **Trust.** A skill's text steers the model but grants nothing: every call it leads to still goes through the
  editing mode, the C# tier gate, `disabledModes` and the Approve card. `allowed-tools` and other frontmatter keys
  are ignored.

## Approve before a call (requireApproval)

The administrator's `policy.json` decides which of the model's calls need an Approve from the user, per app, tool and mode (`requireApproval`; see `OfficeAi.Shared/ModePolicy.cs.md`). `"*"` stands for every tool that is not read-only, and a tool's own entry overrides it. Full autonomy never asks, whatever is listed. The third tier (`trackChanges`) is the tier with approvals in every app but Word: in Excel and PowerPoint it is every tool, called "Edits with approval" (they have no tracked changes), and in Outlook it is "Actions with approval" (`respond_meeting` and permanent delete on top of Draft only; `send_email` and `send_event` stay in Full autonomy, since a draft is the reviewed way to send). The shipped defaults ask before every change in that tier in those three apps, and before `delete_email` in Outlook's Draft only; Word asks for nothing. There is no user setting for it. A tool name the app does not have is ignored. The check is client-side, in `executeConfirmed` around the model's `callDotNetTool` call. It adds no model tool or tier-list entry, and the server-side editing-mode check still runs after Approve. Because only the page enforces it, an installed (Release) pane never enables WebView2's developer tools.

How the card looks comes from `AddInConfig.confirmActions` (Outlook's `web-src/confirm.ts`). The first matching `ConfirmRule` wins. A tool with no rule gets a plain "Allow this action?" card with the tool's display name and description, and its arguments: one row per argument, one per item of a list (such as Excel's operations). The card (`confirmAction` in `chat-ui.ts`) shows the rule's title and labelled rows, never the tool call itself or a message body. A rule with `preview` asks the add-in (`preview_action`, a read-only entry in `OutlookTools.ExecuteAsync` that is in no tier set) for the item's type, subject, sender and date, or for the reason the call would be refused, in which case no card is shown and the model gets that refusal. A rule's `fields` add rows from the call's own arguments (for example the recipient of a new email). Deny, Escape, Stop and New chat all resolve to Deny, and the model gets a result telling it not to retry. A call the user typed with `!` never asks.

## Provenance / data-source enforcement

There is no `dataSource`/provenance-enforcement mechanism anywhere in OpenDocs (a
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
