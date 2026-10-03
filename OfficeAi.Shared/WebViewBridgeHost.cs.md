## WebViewBridgeHost._skipTlsVerify

```
// Off by default. Only settable via the explicit "set-tls-bypass"
// WebMessage the Settings panel's checkbox sends - never enabled
// silently. Intended for testing against an internal/air-gapped LLM
// gateway using a self-signed or otherwise untrusted certificate;
// the correct fix for a real deployment is trusting the real
// certificate in Windows' certificate store instead.
```

## WebViewBridgeHost._environments

```
// One CoreWebView2Environment per app-data-folder name, shared across
// every pane created in this process (keyed defensively; in practice
// one VSTO add-in's AppDomain only ever passes its own fixed name).
// Before PP-1 there was only ever one TaskPaneHost/WebViewBridgeHost
// per process, so CreateAsync was only ever called once - never
// exercised. PP-1 creates one pane per open document window, so
// opening (or Office internally re-activating) more than one window
// can call this constructor again while the first pane's environment
// is still initializing. Calling CoreWebView2Environment.CreateAsync
// a second time for the SAME user-data-folder before the first call
// has finished is exactly what WebView2 rejects with "the group or
// resource is not in the correct state" (HRESULT 0x8007139F,
// confirmed repro) - sharing one environment (the officially
// documented pattern for multiple WebView2 controls against one
// profile) eliminates the race by construction rather than trying to
// serialize around it. Caching the in-flight Task (not just the
// eventual result) matters: a second pane's constructor runs
// synchronously up to this dictionary check/insert, before any
// `await`, so it sees and awaits the SAME in-flight task instead of
// starting a second CreateAsync. Only ever touched from the single
// STA UI thread all Office COM callbacks run on, so no lock is needed.
```

## WebViewBridgeHost.InitializeAsync - cache-busting filter

```
// Post-hoc fix (2026-08-24, user-reported a CSS fix not
// taking effect after rebuild + close/reopen Word): the
// CoreWebView2Environment here uses a PERSISTENT
// userDataFolder (see GetOrCreateEnvironment below), so its
// HTTP disk cache survives across Word restarts even though
// the C# DLLs themselves reload correctly - closing and
// reopening Word recreates the WebView2 CONTROL, but not its
// cache. bundle.js/bundle.css/index.html are referenced with
// no cache-busting query string, so a rebuilt bundle can be
// silently served stale from cache indefinitely. Force every
// request to this virtual host to bypass cache and
// revalidate, matching what a browser's hard-refresh does.
```

## WebViewBridgeHost.InitializeAsync - http scheme for appassets.local

```
// http, not https: appassets.local serves local files only (no
// real network transport), so the scheme is a free choice - and
// an insecure page fetching an insecure resource is not mixed
// content, unlike a secure page doing the same. Without this, a
// user-configured remote LLM endpoint reachable only over plain
// HTTP gets silently blocked by Chromium as mixed content
// before the request ever leaves the process (confirmed via
// WebView2 DevTools: "Mixed Content: ... was loaded over
// HTTPS, but requested an insecure resource"). Fetching an
// HTTPS endpoint from this now-insecure origin is unaffected -
// mixed content is one-directional. The trade-off is
// crypto.randomUUID() (secure-context-gated) - see
// agent-core/id.ts's randomId() fallback, used everywhere this
// codebase previously called crypto.randomUUID() directly.
```

## WebViewBridgeHost.OnWebMessageReceived

```
// async void: this is an event handler bound to CoreWebView2.WebMessageReceived,
// raised on the UI thread. The tool-call branch awaits the (now Task-returning)
// executor so a slow tool - e.g. Outlook's search_contacts, which does its EWS
// call off the UI thread - no longer blocks the message pump. The outer try/catch
// is mandatory: an exception escaping an async void crashes the process. The
// continuation after `await _executor(...)` is NOT guaranteed to resume on the UI
// thread (an Office host thread may carry no WinForms SynchronizationContext, so a
// tool that awaited Task.Run resumes on a threadpool thread) - so the result is
// posted back through PostToolResult, which marshals onto the control's thread.
```
