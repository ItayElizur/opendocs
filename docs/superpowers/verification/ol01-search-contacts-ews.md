# OL-1 — Outlook `search_contacts` on EWS + async tool-execution pilot — Verification

**Scope:** Reimplement `search_contacts` as a server-side EWS `ResolveName` call (Contacts
then GAL) authenticated with Windows Integrated Auth, run off the UI thread; make the
shared tool-execution path `Task`-returning so it can. Removes the recursive multi-store
contact-folder crawl that froze then crashed Outlook.

Plan: `C:\Users\Itay\.claude-personal\plans\when-in-outlook-i-dazzling-emerson.md`.

## What changed

**Async plumbing**
- `OfficeAi.Shared/ToolProtocol.cs` — `ToolExecutor` delegate → `Task<ToolResult>`.
- `OfficeAi.Shared/PaneHostBase.cs` — abstract `ExecuteTool` → `Task<ToolResult>`.
- `OfficeAi.Shared/WebViewBridgeHost.cs` — `OnWebMessageReceived` is now `async void`;
  reads `e.WebMessageAsJson` into a local before awaiting; `await _executor(...)`; no
  `ConfigureAwait(false)`; the outer catch now also posts an `IsError` tool-result for the
  faulted `requestId` so the JS promise can't hang.
- Word / Excel / PowerPoint `TaskPaneHost.ExecuteTool` — `return Task.FromResult(XxxTools.Execute(...))`
  (one line each, no behaviour change).
- `OutlookAiAddIn/TaskPaneHost.cs` → `OutlookTools.ExecuteAsync`; `OutlookTools.Execute` →
  `async Task<ToolResult> ExecuteAsync`, only `search_contacts` is actually async.

**EWS contact search**
- `OutlookAiAddIn/OutlookAiAddIn.csproj` — `PackageReference Microsoft.Exchange.WebServices` 2.2.0
  + `<Compile Include="OutlookEws.cs">`.
- `OutlookAiAddIn/OutlookEws.cs` (new) — `ExchangeService` with `UseDefaultCredentials`,
  `Timeout` 15 s; `ResolveNamesAsync` / `DiscoverUrlAsync` via `Task.Run`; `NameResolution`
  → `(name, email)` mapping (GAL `EX`/X500 addresses fall back to the resolved contact's
  own `EmailAddress1..3`; no-`@` entries dropped, mirroring mcp-outlook); `IsTimeout` walks
  the inner-exception chain. No `Microsoft.Office.Interop.Outlook` using (type collisions).
- `OfficeAi.Shared/EwsAutodiscoverXml.cs` (new, pure) — `ParseEwsUrl`: local-name walk of
  `Protocol` elements, `<EwsUrl>` over `<ASUrl>`, `EXCH` over `EXPR` over any; null on
  anything unusable.
- `OfficeAi.Shared/ContactSearchFormat.cs` (new, pure) — dedupe (lowercased address, name
  fallback), truncate to `limit`, byte-for-byte identical output to the old implementation.
- `OutlookAiAddIn/OutlookTools.Contacts.cs` — rewritten: `SearchContactsAsync` +
  `FindExchangeAccountInfo` (COM, UI thread — picks the Exchange account matching the
  signed-in user, reads `Account.AutoDiscoverXml`). Deleted `SearchContacts`,
  `CollectContactFolders`, `ResolveContactFolder`, `WellKnownContacts`, the inline DASL.
- `OutlookAiAddIn/web-src/entry.ts` — `search_contacts` `folder` arg removed, description
  rewritten. Bundle rebuilt.
- `docs/ai-tool-surface.md` — EWS carve-out, `search_contacts` row, native-query bullet,
  "Unproven at runtime" entry.
- `native-query-apis` memory — "On record" note.

## Automated checks — passed

- [x] `dotnet test OfficeAi.Shared.Tests` — **155/155** (was 136; +9 `EwsAutodiscoverXml`,
  +8 `ContactSearchFormat`, +2 theory rows). `ToolProtocolTests` unaffected by the delegate
  change (it never constructs a `ToolExecutor`).
- [x] `MSBuild.exe` **Debug and Release**, all four add-ins + `OfficeAi.Shared` —
  **0 warnings, 0 errors** each. No `App.config` binding redirect needed.
- [x] `Microsoft.Exchange.WebServices.dll` + `.Auth.dll` land in
  `OutlookAiAddIn/bin/{Debug,Release}/` (and `bin/*/web/bundle.js` carries the new
  description).
- [x] `npx tsc --noEmit` in `OutlookAiAddIn/` — clean.
- [x] `npx esbuild web-src/entry.ts` — `web/bundle.js` 146.1 kb, `bundle.css` 23.6 kb, no
  errors.
- [x] grep — zero references to `CollectContactFolders` / `ResolveContactFolder` /
  `WellKnownContacts` / `SearchContacts(` / a `folder` key in the `search_contacts` schema.

## Manual matrix — NOT run (no live on-prem-Exchange Outlook reachable from this environment)

Watch `%TEMP%\AirchatOfficeDebug.log` alongside each.

- [ ] **a.** Chat "find everyone named <common surname>" → multiple GAL hits, formatted
  `- Name <smtp>`, sub-second on a warm cache (first call slower — autodiscover).
- [ ] **b. Headline property:** while (a) runs, switch folders / open a message / scroll —
  Outlook stays fully responsive. To widen the observation window, temporarily add
  `System.Threading.Thread.Sleep(5000)` inside `OutlookEws.ResolveNames`, confirm the UI
  never blocks, then remove. (The old build froze then crashed here.)
- [ ] **c.** A contact that exists only in personal Contacts, not the GAL — still found.
- [ ] **d.** Disconnect VPN (or point the cached `Uri` at a dead host) → clear `IsError`
  within ~15 s, no freeze, no crash, Outlook usable during and after.
- [ ] **e.** Exchange account whose autodiscover fails and no reachable EWS → clear
  `IsError` ("could not reach Exchange autodiscover…"), no hang.
- [ ] **f.** Gmail/IMAP-only profile → immediate clear `IsError` ("needs an on-prem
  Exchange mailbox…").
- [ ] **g.** `O'Brien`-style query with an apostrophe → no error, results returned.
- [ ] **h.** Regression: `list_emails`, `search_emails`, `get_email`, `list_folders`,
  `list_events`, `list_tasks`, one draft tool — behave exactly as before.
- [ ] **i.** Delegate refactor: one tool each in Word / Excel / PowerPoint (e.g. Word
  `insert_content`, Excel `read_range`, PowerPoint `read_slide`) — round-trips, renders, no
  latency/behaviour change.
- [ ] **j.** Two agent turns each calling `search_contacts` — the second reuses the cached
  `Uri` (log shows no second autodiscover).

## Still genuinely unverified (flagged, not discovered after the fact)

Reflection/compile confirms signatures, not runtime semantics or environment:

1. `Account.AutoDiscoverXml` real content shape vs `EwsAutodiscoverXml.ParseEwsUrl` — a
   null return just falls through to `ExchangeService.AutodiscoverUrl` (slower first call),
   so a parse miss degrades rather than fails.
2. `ExchangeService.UseDefaultCredentials` actually completing Kerberos/NTLM to the on-prem
   CAS as the signed-in user.
3. `ResolveName` returning an empty `NameResolutionCollection` vs throwing
   `ServiceResponseException` on no matches — both handled (empty → "No contacts matched").
4. The off-thread `Task.Run` genuinely keeping the Outlook UI unblocked during the HTTP
   call (test b).
5. EWS Managed API timeout surfacing — `IsTimeout` walks bare `TimeoutException`,
   `WebException(Timeout)`, and either wrapped in `ServiceRequestException`.

## Residual / plan 2

- Every other Outlook COM tool still runs on the UI thread — the `Task<ToolResult>`
  delegate added here is the enabler for moving them.
- In-flight tool-call cancellation still unwired (`loop.ts` passes an `AbortSignal`,
  `bootstrap.ts` drops it) — `search_contacts` can only time out at 15 s.
- Exchange Online account in the profile → `UseDefaultCredentials` can't do modern auth;
  surfaces as the clear `IsError`, no crash.
- Internal Exchange TLS cert trust — assumed machine-trusted on a domain box; wiring the
  existing `set-tls-bypass` toggle to `ServicePointManager.ServerCertificateValidationCallback`
  is flagged, not built.

## Fixes after first on-machine test (2026-09-01)

First run against a real LLM on another machine: `search_contacts` logged
`ServiceRequestException: "The underlying connection was closed: An unexpected error
occurred on a send."` and then **Outlook crashed**. Two independent bugs:

1. **TLS.** That exact error is a TLS-protocol mismatch — the process default didn't offer
   TLS 1.2 to the on-prem Exchange endpoint. Fixed: `OutlookEws` static constructor does
   `ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12`.
2. **Crash instead of clean error.** The `await _executor(...)` continuation in
   `WebViewBridgeHost.OnWebMessageReceived` resumed on a **threadpool thread** (an Office
   host thread carries no WinForms `SynchronizationContext`, so a tool that awaited
   `Task.Run` — the EWS call — does not resume on the UI thread). The old code then touched
   the status `Label` and `_webView.CoreWebView2` off-thread; those throw cross-thread, and
   the throw inside the `catch` escaped the `async void` handler → process termination.
   Fixed: the result post-back now goes through `PostToolResult`, which marshals via
   `Control.BeginInvoke` (works with no `SynchronizationContext`) and swallows its own
   failures; the post-await `catch` no longer calls `_setStatus`. This is a shared-bridge
   fix — it hardens every add-in's tool path, not just Outlook's.

Both need re-testing on the on-prem-Exchange machine (matrix items a, b, d).

## Deviations from plan

- `ContactSearchFormat.Format` / `OutlookEws` use `IReadOnlyList<KeyValuePair<string,string>>`
  rather than value tuples, matching the original `SearchContacts` accumulator and the
  surrounding codebase style.
- `ServiceRequestException` is referenced fully-qualified in `OutlookTools.Contacts.cs`
  (no `using`) to keep that file free of the EWS/COM type collisions.
