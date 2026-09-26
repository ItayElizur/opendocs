# Outlook `apply_search` Tool — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** After the model has already run one or more `search_emails`/`list_emails` calls and knows which messages are relevant, let it push that same search into Outlook's own Explorer window — the native Instant Search UI, with the matching messages listed in the reading pane — instead of only describing them in the chat. This is the "show, don't just tell" counterpart to `search_emails`: the model already computed a `query`/date range/sender filter for its own use; `apply_search` re-applies the identical filter to what the user is actually looking at.

**Architecture:** Outlook's `Explorer` object exposes `Explorer.Search(string Query, OlSearchScope SearchScope)` — the programmatic form of typing into the Instant Search box and pressing Enter (scoped via `OlSearchScope`: `olSearchScopeCurrentFolder`, `olSearchScopeSubfolders`, `olSearchScopeAllFolders`, `olSearchScopeAllOutlookItems`). `search_emails` already builds the equivalent filter as a DASL string via `OutlookTools.BuildSearchDasl` (`OutlookTools.cs:168-171`, backed by the unit-tested `OutlookDasl.BuildSearchFilter` in `OfficeAi.Shared`) and applies it with `Items.Restrict("@SQL=" + dasl)` (`OutlookTools.Search.cs:38`). `apply_search`'s job is narrower: get the active `Explorer`, point it at the right folder, and call `.Search(...)` with that same query so the *user's* Explorer window shows what the model already found — no new filter logic, reuse `BuildSearchDasl` verbatim.

Getting the `Explorer` to act on: this codebase already has a documented fallback pattern for exactly this ("Outlook's Explorer has no Hwnd... back to `Application.ActiveExplorer()`", `OutlookAiAddIn/ThisAddIn.cs:24-27`). `apply_search`'s handler uses `App.ActiveExplorer()` (via `OutlookTools.App`, `OutlookTools.cs:95`) the same way — it acts on whichever Explorer window currently has focus, matching how `draft_email`/`reply_email`/etc. already operate on the application's active state rather than routing through a specific pane.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (`OutlookAiAddIn/OutlookTools.Search.cs`, `OutlookTools.cs`), TypeScript (`OutlookAiAddIn/web-src/entry.ts`).

## Global Constraints

- **This is a view-state change, not a data mutation.** It changes what the user's Explorer window is showing (folder + search box contents), exactly like the user pressing Ctrl+E themselves — it does not read, move, delete, or modify any item. Per the existing mode-gate philosophy stated in `OutlookTools.cs:28-29` ("Only Full autonomy permits mutation"), this belongs in `AlwaysAllowedTools` and `readOnlyTools`, same tier as `search_emails` itself.
- Reuse `BuildSearchDasl`/`ResolveFolder` exactly as `SearchEmails` does — do not reimplement filter-building.
- If no `Explorer` is active (`App.ActiveExplorer()` returns null — e.g. the add-in is somehow invoked with no Explorer window, or between window transitions), return a clear `IsError` result rather than throwing.

---

### Task 0: Verify `Explorer.Search`'s query syntax against the real PIA (do this first — it decides Task 1's shape)

**Problem:** `Items.Restrict` definitely accepts `"@SQL=" + dasl` (confirmed, already shipping in `search_emails`). It is **not yet confirmed** that `Explorer.Search`'s `Query` parameter accepts the same `@SQL=`-prefixed DASL syntax, as opposed to only plain AQS (Advanced Query Syntax) free-text search terms — these are two different query languages Outlook understands in different contexts, and guessing here would repeat the exact mistake this project's own history warns against (`docs/superpowers/plans/STATUS.md`'s chart-bug saga: three rounds of plausible-sounding guesses before a signature was actually checked via reflection).

- [ ] **Step 1:** Inspect `Microsoft.Office.Interop.Outlook.Explorer.Search`'s documented parameter contract (XML doc / object browser on the referenced PIA) and, if available, test interactively (Outlook's own Instant Search box accepts `@SQL=` search strings when Advanced Find generates one — confirm this is true for the `Explorer.Search` method call specifically, not just the UI box).
- [ ] **Step 2:** Record the finding at the top of Task 1's handler as a comment, with either:
  - **DASL confirmed working:** pass `"@SQL=" + dasl` straight through, identical to `search_emails`.
  - **DASL not accepted / unreliable:** fall back to passing the plain-text `query` argument only (AQS free-text search), and have the tool's description say explicitly that `apply_search`'s UI result may be less precise than `search_emails`'s own filtering (date/sender narrowing may not carry over to the visible UI search, even though the tool still navigates to the right folder).

**Verification:** a one-line documented decision, made from an actual check, not an assumption — this gates which of Task 1's two branches gets implemented.

---

### Task 1: C# — `apply_search` handler

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Search.cs` (new handler, next to `SearchEmails`)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (dispatcher + `AlwaysAllowedTools`)

- [ ] **Step 1: Add the handler** (shape assumes Task 0 confirmed DASL works; adjust per Task 0's actual finding):

```csharp
private static ToolResult ApplySearch(JsonElement input)
{
    string query = Str(input, "query", "");
    string folderName = Str(input, "folder", "inbox");
    DateTime? start = DateArg(input, "start_date");
    DateTime? end = DateArg(input, "end_date");
    string sender = Str(input, "sender", null);

    Outlook.Folder folder = ResolveFolder(folderName);
    Outlook.Explorer explorer = App.ActiveExplorer();
    if (explorer == null)
        return new ToolResult { Output = "No active Outlook window to apply the search to.", IsError = true, Summary = "apply_search" };

    explorer.CurrentFolder = folder;

    string dasl = BuildSearchDasl(query, start, end, sender);
    if (dasl.Length == 0)
        return new ToolResult { Output = "Showed " + folder.Name + " (no filter criteria given).", Summary = "apply_search" };

    explorer.Search("@SQL=" + dasl, Outlook.OlSearchScope.olSearchScopeCurrentFolder);
    return new ToolResult { Output = "Applied the search to " + folder.Name + " in Outlook - the user can see the results now.", Summary = "apply_search" };
}
```

- [ ] **Step 2: Wire the dispatcher.** Add to `OutlookTools.cs`'s `ExecuteAsync` switch, in the read-only block (`:54-63`):
```csharp
case "apply_search": return ApplySearch(input);
```

- [ ] **Step 3: Add `"apply_search"` to `AlwaysAllowedTools`** (`OutlookTools.cs:30-34`).

**Verification:** `dotnet build` / MSBuild on `OutlookAiAddIn` succeeds. Manual: ask for a search with a date range and sender, confirm the model calls `search_emails` (or already knows the filter from context) then `apply_search` with the same filter — Outlook's own window navigates to the folder and shows the Instant Search box populated/results filtered accordingly.

---

### Task 2: TypeScript — tool schema, display strings, read-only registration

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Add to `ALL_OUTLOOK_TOOLS`**, directly after `search_emails` (`entry.ts:29-47`) — same parameter shape as `search_emails` minus `limit`/`unread_only`/`recipient` (those only make sense for the model's own in-chat listing, not for what gets typed into Outlook's search box):

```ts
{
  name: 'apply_search',
  description:
    "Applies a search to the user's actual Outlook window - navigates to the folder and runs the search there, so the user sees the same results you found. Use after search_emails/list_emails once you know what's relevant; reuses the same query/date/sender filters.",
  inputSchema: {
    type: 'object',
    properties: {
      query: { type: 'string', description: 'Text matched against subject and body (contains).' },
      folder: FOLDER,
      start_date: { type: 'string', description: 'Only messages on/after this date (YYYY-MM-DD).' },
      end_date: { type: 'string', description: 'Only messages on/before this date (YYYY-MM-DD).' },
      sender: { type: 'string', description: 'Sender email or display-name fragment.' },
    },
    required: [],
  },
},
```

- [ ] **Step 2: Add to `OUTLOOK_TOOL_DISPLAY`**, after `search_emails`:
```ts
apply_search: d('Show search in Outlook', 'הצגת חיפוש ב-Outlook', 'Applies the search to the Outlook window itself.', 'מיישם את החיפוש בחלון Outlook עצמו.'),
```

- [ ] **Step 3: Add `'apply_search'` to `readOnlyTools`** (`entry.ts:337-348`), after `'search_emails'`.

- [ ] **Step 4: System prompt.** Add one sentence near the existing "Prefer list_emails / search_emails..." line (`entry.ts:330`) telling the model: once it has found the relevant messages, `apply_search` can show the same results in the user's own Outlook window instead of only listing them in chat.

**Verification:** `tsc --noEmit` clean; bundle rebuilds; tool appears with display strings in the running add-in.

---

### Task 3: Docs

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] Add `apply_search` to the Outlook read-only tools table, next to `search_emails`, noting it's a view-only Explorer-state change (`Mutated: false`) and recording Task 0's DASL-vs-AQS finding so a future reader doesn't have to re-derive it.

**Verification:** doc reads correctly against the implemented behavior.
