# Outlook tools (`OutlookAiAddIn/OutlookTools.*.cs`)

Current-state reference for every tool the AI can call in Outlook. See
[`docs/architecture.md`](../architecture.md) for the shared editing-mode/transport
architecture. The tool set mirrors a self-hostable EWS
MCP reference server as closely as the Outlook COM object model allows. This add-in
drives the **already-running, already-authenticated desktop Outlook client** via
`Microsoft.Office.Interop.Outlook`, not EWS. There is no `Namespace`/credential setup;
it acts as the signed-in user. **Three exceptions use EWS** (see "EWS carve-out"
below): `search_contacts`, `find_meeting_slots`' work-week lookup, and `list_events`'
shared-calendar path — each because the COM object model either can't do the job
(a multi-result directory search) or does it in a way that freezes Outlook.

## Shape of the integration (differs from Word/Excel/PowerPoint)

- **Explorer-only.** The chat pane docks in the main Outlook window (`Explorer`).
  There is **no task pane in Inspector windows** (pop-out read/compose). The ribbon
  button is on the Explorer's Mail tab (`idMso="TabMail"`, via `RibbonBase.HomeTabIdMso`
  / `ProvidesRibbonFor`) and suppressed on every Inspector ribbon surface.
- **Per-mailbox chat.** `GetChatId()` returns `mbx-` + `SHA256(primary SMTP)[:16]` —
  one rolling conversation per mailbox. No file path, so none of Word/Excel/
  PowerPoint's provisional-id / `ChatStore.Migrate` / `DocSettingsStore` save
  lifecycle applies. Primary SMTP is resolved lazily (`ExchangeUser.PrimarySmtpAddress`
  → `PR_SMTP_ADDRESS` proptag → `AddressEntry.Address` → first account → display name).
- **Pane lifecycle.** `Explorer` has no `Hwnd`, so panes are keyed by the COM
  identity pointer (`Marshal.GetIUnknownForObject`). Creation is **lazy** — on an
  explorer's first `Activate` or the ribbon button, never in `ThisAddIn_Startup` —
  because Outlook auto-disables add-ins whose median startup exceeds ~1 s, and
  `install.ps1` also writes `Resiliency\DoNotDisableAddinList`.
- **Editing modes — four tiers.** Outlook repurposes all four `EditingMode` slots with
  its own meaning. Each tier is a strict superset of the one before it:

  | `EditingMode` | Outlook label | Adds |
  |---|---|---|
  | `ReadOnly` | **Read only** | The 12 read tools (`readOnlyTools` in `entry.ts`, `AlwaysAllowedTools` in `OutlookTools.cs`). |
  | `CommentOnly` | **Draft only** (the default for a fresh session) | 17 tools that mutate the mailbox or open a draft but never leave anything unreviewed: triage (`mark_email_read`/`mark_email_unread`, `flag_email_important`, `move_email`, `delete_email` — non-permanent only, see below), tasks/reminders (`create_task`, `update_task`, `set_reminder`, `set_email_reminder`), event tagging/status (`set_event_categories`, `set_category_color`, `set_event_availability`), `apply_search`, every draft tool (`draft_email`, `draft_event`), and `undo_last_action`/`redo_last_action` (`commentOnlyExtraTools` / `DraftTierTools`). |
  | `TrackChanges` | **Automate approvals** | `respond_meeting` — it calls `resp.Send()` to notify the organizer, so it gets its own explicit tier rather than hiding in Draft only or Full autonomy (`trackChangesExtraTools` / `ApprovalTierTools`). |
  | `FullAutonomy` | **Full autonomy** | The two auto-send tools — `send_email`, `send_event` (`SendTierTools`; on the `entry.ts` side, simply absent from every lower tier's list) — plus `delete_email` with `permanent: true`. |

  `entry.ts` passes `availableModes: ['readOnly','commentOnly','trackChanges',
  'fullAutonomy']` and a `modeOverrides` map so the mode menu shows Outlook's own
  labels/descriptions instead of Word/Excel/PowerPoint's generic "Comment only"/"Track
  changes" text — those apps' own use of `CommentOnly`/`TrackChanges` for real editing
  is untouched (Outlook's relabeling is purely its own `entry.ts` config, not a
  shared-string change). Gate logic (`OutlookTools.cs`'s `ExecuteAsync`) is an ordinal
  check on the enum's own declared order (`ReadOnly < CommentOnly < TrackChanges <
  FullAutonomy`), name-based except for one narrow input-aware special case:
  `delete_email` with `permanent: true` is blocked below Full autonomy (the tool name
  itself stays in `DraftTierTools`, since its default non-permanent path is a
  reversible move). A fresh session defaults to **Draft only** (`defaultMode:
  'commentOnly'`, resolved via `chat-ui.ts`'s `defaultModeFor()` — the same helper
  Word/Excel/PowerPoint use to default to their own `trackChanges`).
- **Selection context.** The Explorer's `SelectionChange` pushes the selected mail
  item(s) / conversation into per-turn context as a `mail` `SelectionContext` variant
  (`shared/web-src/app-shell/bootstrap.ts`) carrying each item's `EntryID`, so the
  model prefers the selection over searching. The scope pill shows the subject (one
  message) or a count (`scopeUnit: 'mailbox'`).
- **`message_id` / `event_id` / `task_id` = Outlook `EntryID`**, re-resolved via
  `Namespace.GetItemFromID` on every call. EntryID is stable within a folder but
  **changes on `Move` and across stores** — `move_email` / `delete_email` return the
  new id. Recurring calendar instances all share the master appointment's EntryID —
  see "Recurring events" below for how the calendar-editing tools target a single
  occurrence anyway. An `event_id` from someone else's shared calendar also needs its
  `store_id` — see "Shared calendars" below.
- **Native query APIs are mandatory.** `list_*` use `Folder.GetTable` (an in-memory
  rowset, no per-item COM object); `search_emails` uses `Items.Sort` →
  `Items.Restrict("@SQL=" + DASL)`; a capped linear scan is a fallback only when
  `Restrict` rejects the filter. The DASL builder is pure and unit-tested
  (`OfficeAi.Shared/OutlookDasl.cs`). This is a hard requirement, not a style
  preference — see `CLAUDE.md`'s "Outlook specifically" section for the real
  performance incident (a positional-indexing bug in Word's `find_text`/`get_headings`
  that froze Word on a large document) this rule exists to avoid repeating here, since
  Outlook's mailbox/calendar collections have the exact same trap. `search_contacts`
  and `list_events`' shared-calendar path are the deliberate, documented exceptions:
  both go to EWS instead of a COM scan — see the EWS carve-out below.
- **EWS carve-out: `search_contacts`, `find_meeting_slots`' work-week default, and
  `list_events`' shared-calendar path are the only things that touch EWS.** Every
  other Outlook tool is pure `Microsoft.Office.Interop.Outlook` COM against the running
  client. All EWS-dependent code lives apart from the pure-COM tool files, in two
  layers: `OutlookEws.cs` (the raw EWS Managed API wire calls — `ResolveNamesAsync`,
  `GetWorkingHoursAsync`, `GetSharedCalendarEventsAsync`) and `OutlookTools.Ews.cs`
  (the tool-facing orchestration on top: `SearchContactsAsync`, `ResolveWorkWeekAsync`,
  and the shared endpoint/account resolution every caller uses,
  `ResolveEwsUrlAsync`/`FindExchangeAccountInfo`). All EWS calls use EWS Managed API
  2.2 (`Microsoft.Exchange.WebServices` 2.2.0) with
  `ExchangeService.UseDefaultCredentials` (Windows Integrated Auth as the signed-in
  user; no stored credentials). Endpoint is parsed from the cached
  `Outlook.Account.AutoDiscoverXml` (`<EwsUrl>`/`<ASUrl>`, `EXCH` preferred over
  `EXPR`; pure parser `OfficeAi.Shared/EwsAutodiscoverXml.cs`), falling back to
  `ExchangeService.AutodiscoverUrl`, then cached in a process-static `Uri` shared by
  every EWS-dependent tool. Calls run off the UI thread (`await Task.Run`,
  `svc.Timeout` 15 s) so Outlook stays responsive. **On-prem Exchange only.** EWS
  unreachable / SSPI failure / endpoint not found / timeout → a clear `IsError` result
  for `search_contacts` and for `list_events` with `mailbox` (EWS isn't optional
  there); `find_meeting_slots` instead falls back to a hardcoded Sun-Thu/9-18 default,
  since EWS is an enhancement over an already-working COM-only path there, not the
  only way to do the job. These three are the only async tools — the shared
  `ToolExecutor` delegate is `Task<ToolResult>`-returning
  (`WebViewBridgeHost.OnWebMessageReceived` is `async`); every other tool in all four
  add-ins is still synchronous, wrapped in `Task.FromResult`.

## Read tools (12 — always allowed, never gated)

| Tool | Notes |
|---|---|
| `list_emails` | `Folder.GetTable`; columns EntryID/Subject/ReceivedTime/SenderName/UnRead + `PR_HASATTACH` proptag; `[UnRead] = true` restriction when `unread_only`; sorted newest-first; non-mail rows filtered by `MessageClass` not starting `IPM.Note`. Args: `folder`, `limit` (20), `unread_only`. |
| `search_emails` | `Items.Sort("[ReceivedTime]")` then `Restrict("@SQL=" + BuildSearchFilter(...))`. `LIKE '%q%'` on subject + body; UTC-ISO date range; sender by `fromemail =` OR `fromname LIKE` (Exchange senders carry `legacyExchangeDN`, not SMTP — display-name fuzzy match). `ci_phrasematch`/`ci_startswith` are **not** usable via `Restrict` (they throw). `recipient` is a client-side filter, capped 500. Fallback: capped linear scan on a malformed filter. |
| `get_email` | Full `Body` (≤ 40k), To/CC via `Recipient.Type`, `ConversationID`/`ConversationTopic`, importance, unread, and an `attachments` array — `{index (1-based), name, type (byValue/embeddedItem/ole/reference), size}` — feed the index to `get_attachment`. |
| `open_email` | `MailItem.Display(false)` on an existing item resolved via `ItemById` — opens the message in its own Outlook reading window, unmodified. `Mutated: false`. |
| `get_attachment` | `Attachment.SaveAsFile` into `%LOCALAPPDATA%\OutlookAiAddIn\Attachments\`; returns the path. `extracted_text` (≤ 40k) is populated **only** for text-family extensions (`.txt .csv .tsv .md .json .xml .log`, `.html` tag-stripped) and OpenXML (`.docx .xlsx .pptx`), via the swappable `OfficeAi.Shared/AttachmentText/` module (`DocumentFormat.OpenXml` 2.20.0). **No PDF, no images, no vision** — those return the path + type only. `olOLE` throws (rejected); `olByReference` has no data (rejected); `olEmbeddedItem` saves as `.msg`. |
| `list_folders` | Recursive walk of every store's `Folders`, mail folders only (`DefaultItemType == olMailItem`), with item + unread counts; capped ~800 / depth 8. |
| `search_contacts` | **EWS `ResolveName` called twice — `ContactsOnly`, then `DirectoryOnly` (GAL) — both with `returnContactDetails: true`, results merged** (server-side ANR), run off the UI thread. Querying both locations unconditionally matters: Exchange's ANR on the Contacts folder only matches `DisplayName`, while the Directory phase also matches given name/surname, so a single `ContactsThenDirectory` call (which stops at the first Contacts hit) could miss the person entirely. Each `NameResolution` mapped to a match carrying both the person's **full name** (GivenName + Surname) and the directory's **display name** — the output leads with the full name (what a query like "John Doe" actually matches against) and shows the display name as a parenthetical only when it adds information, since an org-formatted directory label (e.g. `"Dept/Unit/Title"`) can look nothing like the person's name. GAL X500/`EX` addresses fall back to the resolved contact's own `EmailAddress1..3`; entries with no `@` address are dropped. Deduped by lowercased address (name fallback), capped at `limit` (default 10). **No `folder`/scope arg.** On-prem Exchange only; unreachable / auth failure / no endpoint / 15 s timeout → a specific `IsError` message, no COM fallback (a recursive multi-store contact-folder crawl froze then crashed Outlook, which is why there is none). |
| `list_events` | **Own calendar** (no `mailbox`): `Items.Sort("[Start]")` → `Items.IncludeRecurrences = true` → `Items.Restrict("[Start] <= end AND [End] >= start")` — **this order is load-bearing** and rules out `GetTable`. Recurring instances share the master `event_id`; each row carries its own `start` to disambiguate (and to pass as `occurrence_date`). **Shared calendar** (`mailbox` given — an email address): resolves via `Ns.CreateRecipient(mailbox).Resolve()`, then queries EWS `FindAppointments` + `CalendarView` (server-side date filtering and recurrence expansion, off the UI thread) instead of enumerating COM `Items` — see "Shared calendars" below. Args: `start_date` (today), `end_date` (+7d), `mailbox` (optional), `limit` (50). |
| `get_event` | `Body` (≤ 40k), `RequiredAttendees`/`OptionalAttendees`, organizer, response status, recurring flag. Optional `store_id` resolves an `event_id` from a shared calendar (see "Shared calendars" below); omit it for the caller's own events. |
| `find_meeting_slots` | `Recipient.FreeBusy(anchor, 30, true)` — a per-30-min status string — for `Namespace.CurrentUser` + each resolved attendee; then `OfficeAi.Shared.MeetingSlots.Rank` (pure, unit-tested) slides a `duration_minutes` window in 30-min steps across each work day's `[start_hour, end_hour)` and scores each candidate by how many people are free (so a best partial match still comes back). **Work week/hours are read from the mailbox's own EWS `GetUserAvailability` → `AttendeeAvailability.WorkingHours` (`OutlookEws.GetWorkingHoursAsync`), not hardcoded** — falls back to Sun–Thu 09:00–18:00 only if that call fails (non-Exchange profile, EWS unreachable, etc.), cached per process like the EWS URL. Default range is today through the end of the current contiguous work-day run, max 28 days. Times past the returned free/busy window are assumed free. Async only because of the EWS work-week lookup; the FreeBusy/ranking work itself is still synchronous COM. `Recipient.FreeBusy` is all local-time, so no cross-timezone math is needed. Args: `attendees` (req), `duration_minutes` (req), `start_date`, `end_date`, `start_hour`, `end_hour` (both default to the resolved work hours, or 9/18 as a last resort), `limit` (5). |
| `list_tasks` | `Folder.GetTable` over the default Tasks folder; open tasks only unless `include_completed`. Columns EntryID/Subject/Due/Start/Status/PercentComplete/Complete/ReminderTime. |
| `list_color_categories` | `Namespace.Categories` — the profile's master color-tag ("Category") list shared by mail/calendar/tasks, same list Outlook's Categorize picker shows. Each entry: `{name, color}`; color is one of the 26 `OlCategoryColor` values (None/Red/Orange/…/Dark Maroon), mapped to a friendly display name in `OutlookTools.Categories.cs` (not in `OfficeAi.Shared` — that project doesn't reference the Outlook PIA, same split as `ColorUtil`). |

## Mailbox, task, and event actions (13 — Draft only or higher)

All of these act directly (no review window) but stay local to the mailbox — nothing is
sent. All except `apply_search` set `Mutated = true` and record an undo entry (see
"Undo/redo tools" below).

| Tool | Notes |
|---|---|
| `mark_email_read` / `mark_email_unread` | `MailItem.UnRead` + `.Save()`. |
| `flag_email_important` | `Importance = olImportanceHigh/Normal` + `.Save()`. `important` defaults true. |
| `move_email` | `MailItem.Move(ResolveFolder(destination))`; returns `{message_id: <new EntryID>, old_message_id}`. |
| `delete_email` | Non-permanent (default) → `Move` to Deleted Items (returns new id) — Draft only, same risk class as `move_email`. `permanent: true` → then `.Delete()` from there (no single-call hard delete in the OM — documented as "may still be server-recoverable") — **requires Full autonomy**, via the input-aware check in `ExecuteAsync` described under "Editing modes" above; the tool description says so, so the model doesn't attempt it at a lower tier. |
| `apply_search` | View-only (`Mutated: false`, no undo entry) — pushes the same `BuildSearchDasl` filter `search_emails` computes into `Explorer.CurrentFolder` + `Explorer.Search("@SQL=" + dasl, olSearchScopeCurrentFolder)`, so the user's own Outlook window shows the results. Draft tier rather than always-allowed even though it never changes data: it visibly takes over the user's real Outlook window, which Read only promises never to do. `_Explorer.Search(string Query, OlSearchScope SearchScope)`'s signature was confirmed via .NET reflection against the referenced PIA; the `@SQL=` DASL string itself has **not yet been exercised against a live Outlook session**. |
| `set_event_categories` | `AppointmentItem.Categories` (comma-separated tag names, the color shown on the event in the calendar grid) + `.Save()`; empty/omitted `categories` clears all tags. A name outside the master list is auto-added by Outlook on `Save` with an arbitrary color — call `set_category_color` first to control it. Optional `store_id` (shared calendars). |
| `set_category_color` | `Namespace.Categories[name]` — updates `.Color` if the tag exists, else `Categories.Add(name, color)` creates it. Same master list `list_color_categories` reads. |
| `set_event_availability` | `AppointmentItem.BusyStatus` (the calendar's "Show As" dropdown — one of the 5 `OlBusyStatus` values: Free/Tentative/Busy/Out of Office/Working Elsewhere) + `.Save()`. Case/spacing-insensitive friendly-name lookup; an unrecognized value returns `IsError` listing the valid names instead of throwing. Purely local, never `.Send()` — **verified against a live Outlook client**, including on an organized meeting with attendees (no prompt, no notification). Optional `store_id` (shared calendars). |
| `create_task` | `Application.CreateItem(olTaskItem)` + `.Save()` — no window (a task doesn't send anything, so it follows the mutate-directly pattern, not draft-and-display). Args: `subject` (req), `body`, `due_date`, `start_date`, `reminder_time`, `importance`. |
| `update_task` | `(TaskItem)GetItemFromID`; only passed fields change; `mark_complete: true` → `Complete = true` + `PercentComplete = 100`. |
| `set_reminder` | `ReminderSet` / `ReminderTime` on an appointment **or** task, addressed by its `item_id` (EntryID); `clear: true` turns it off. |
| `set_email_reminder` | `MailItem.MarkAsTask(mapped interval)` + `TaskStartDate`/`TaskDueDate` + `ReminderSet`/`ReminderTime` + `.Save()` — the confirmed COM path for "flag an email for follow-up with a reminder". `MailItem` does expose `ReminderSet`/`ReminderTime`. |

## Undo/redo tools (2 — Draft only or higher; `Mutated` on success)

A **custom** undo/redo stack of the assistant's own actions (`OutlookTools.Undo.cs`,
stack logic in the pure, unit-tested `OfficeAi.Shared/ActionHistory.cs`), not a wrapper
around Outlook's native Undo. Native Undo (`Explorer.CommandBars.ExecuteMso("Undo")`)
was tried by hand and can't be used: it's a single slot tied to the window, it doesn't
see object-model changes, and after one use both it and the ribbon button failed with
"The operation cannot be performed because the message has changed." Outlook has no API
to put an object-model change onto that slot.

| Tool | Notes |
|---|---|
| `undo_last_action` | Reverses the assistant's most recent recorded action; call it again to step further back. Refuses if the item was changed since the assistant's action (by the user or anything else) rather than overwrite it, and never touches the user's own manual changes. A failed or refused entry is dropped, never retried. Stops at a barrier (below) instead of reaching past it. |
| `redo_last_action` | Re-applies the most recently undone action, with the same "changed since" check. The redo list is cleared by any new recorded action or barrier. |

How it works: an in-memory stack per mailbox chat (keyed like `ModeByMailbox`, capped at
50 entries, lost on restart). Each mutating handler records one entry after its change
succeeds:

- **Property snapshots** (before/after values, restored and `Save()`d):
  `mark_email_read/unread` (`UnRead`), `flag_email_important` (full `Importance`, so Low
  is preserved), `set_event_categories`, `set_event_availability` (`BusyStatus`),
  `set_reminder`, `update_task` (task fields as a group, or the flagged-mail fields),
  and `send_event` (edit) whenever the call doesn't send or change a whole recurring series'
  time — a snapshot of only the fields that call actually touched (`Start`/`End`,
  `Subject`, `Body`, `Location`). This includes an occurrence-level edit via
  `occurrence_date`: `RecurrencePattern.GetOccurrence()` returns a real item with its
  own `EntryID` once saved, so the same mechanism covers it.
- **`set_email_reminder`**: if the message wasn't flagged before, undo calls
  `ClearTaskFlag()`, and redo calls `MarkAsTask` again.
- **Moves**: `move_email`, non-permanent `delete_email`, and whole-event `send_event` (cancel)
  on a plain appointment or an already-canceled event (nobody to notify, so canceling
  it is just a move to Deleted Items). The folder is resolved by its own
  EntryID/StoreID via `Namespace.GetFolderFromID`, and the new EntryID after each move
  is rewritten into every entry for that item.
- **Created items**: `create_task`, and `send_event` (new) without attendees (including a
  recurring series — deleting the master deletes the whole series). Undo moves the item
  to Deleted Items (recoverable); redo moves it back.
- **`set_category_color`**: undo restores the old color, or removes a tag the assistant
  created.
- **Barriers** (recorded, but not reversible — undo reports it can't go past them):
  every `send_email`; `send_event` (new) with attendees; `respond_meeting`;
  `delete_email permanent:true`; `send_event` (edit) and `send_event` (cancel) whenever they `.Send()`
  a notice (a still-active meeting the user organizes, including `send_event` (edit)
  converting a plain event into a meeting by adding attendees for the first time);
  **any occurrence-level cancellation** via `send_event` (cancel)'s `occurrence_date`, plain
  appointment or meeting alike (`RecurrencePattern.Exceptions`/`Exception` are
  read-only via COM, confirmed by reflection — there is no API to restore a deleted
  occurrence); and `send_event` (edit) on a **whole recurring series whenever it includes a
  time change** (Outlook rejects `Start`/`End` on a recurring master, so that path
  writes `RecurrencePattern` fields, which the snapshot mechanism can't read/write — if
  the same call also changes other fields, the whole call stays a barrier; there is no
  partial undo of a mixed change).
- **Known gap — `set_event_categories`**: the snapshot only covers the appointment's own
  `Categories` string. If the assigned name wasn't already in the mailbox's master
  category list, Outlook auto-adds it on `Save()` with an arbitrary color; undo restores
  the appointment but does not remove that auto-created master category entry.

Draft tools never record anything — nothing is saved or sent until the user acts on the
opened window. **Not verified against a live Outlook client**, except
`set_event_availability`'s undo/redo round trip and occurrence-level reschedule/
cancellation on a plain (non-meeting) recurring series — see "Unproven at runtime".

## Draft-and-display tools (2 — Draft only or higher; open a native Outlook window for the user to review and send; `Mutated = false`)

| Tool | Notes |
|---|---|
| `draft_email` | One handler, `ComposeEmail(mbxKey, input, send: false)`, shared with `send_email`. `action` (default `new`): **`new`** — `CreateItem(olMailItem)`, set `To`/`Subject`/`Body`, `Display(false)`, `"— Created with OpenDocs"` appended to a non-empty body. **`reply` / `reply_all` / `forward`** — act on exactly one of `message_id` (an email; `orig.Reply()`/`ReplyAll()`/`Forward()`, `body` HTML-encoded and prepended above the quoted original, optional `to` on a forward) or `event_id` (a calendar meeting: see below). Never sends; the user sends from the opened window. |
| `draft_event` · `action: new` (default) | `CreateItem(olAppointmentItem)`; with `required_attendees`/`optional_attendees` → `MeetingStatus = olMeeting`, `Recipients.Add(...).Type`, each `Recipient` resolved individually via `r.Resolve()` right after it's added (the collection-level `Recipients.ResolveAll()` doesn't reliably resolve these — confirmed live); `Display(false)`. Same `"— Created with OpenDocs"` signature appended to a non-empty body, plus a trailing note naming any attendee address that failed to resolve (so the caller can flag a possible typo). Optional `recurrence` object — same shape and handling as `send_event` (new)'s (see "Recurring events" below). |
| `draft_event` · `action: edit` | Same inputs, refusals, and resolution as `send_event` edit (below) — `event_id`, optional `occurrence_date`/`store_id`, at least one of `start`+`end` (together), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` — but applies every change **unsaved** and ends in `appt.Display(false)` for the user to review. Never touches `ForceUpdateToAllAttendees`, never calls `.Send()`/`.Save()`, never records undo/redo. Runs the same `CheckOccurrenceReorderCollision` pre-check before an occurrence move. |
| `draft_event` · `action: cancel` | Resolves `event_id` (optional `occurrence_date`/`store_id`) and `Display(false)`s it **unchanged** — never touches `MeetingStatus` or anything else, because Outlook persists an unsaved `MeetingStatus` change when the window closes, even without the user clicking Send Cancellation (confirmed live). The user cancels it themselves from the opened window (Cancel Meeting/Send Cancellation for a meeting they organize, Delete for a plain appointment). Refuses (`IsError`) on an already-canceled event (points at `send_event` (cancel), the only way to dismiss one) or a still-active meeting the user only attends (points at `draft_email` with `event_id` to write to the organizer, or `respond_meeting` with `response: decline` one tier up). |

**These never call `.Send()` (mail) or save a calendar event.** The user sends from the
opened Outlook window.

**Reply, reply-all and forward on a calendar event (`draft_email` / `send_email` with `event_id`).**
`AppointmentItem` has no `Reply` or `Forward()` (confirmed against the live COM object — it only has
`ForwardAsVcal()`); only the invitation `MeetingItem` in the Inbox does. `FindInviteMessage` finds it:
first `appt.GetConversation().GetRootItems()` (the `IPM.Schedule.Meeting.Request` item), then by
matching `PidLidGlobalObjectId` against the first 3000 Inbox items. Reply/reply-all open that
message's native `Reply()`/`ReplyAll()` (a `MailItem`); forward opens its native meeting forward
(`((_MeetingItem)invite).Forward()` — the cast avoids the method/event name clash), with `to` added
through `Recipients` (each resolved; unresolved addresses are reported, and block a send). `body` is
prepended to the forward's plain `Body`. When no invitation message exists (deleted/archived, or the
user organizes the event — an organizer has none) the tool returns `IsError`; there is deliberately
no lookalike fallback and no `.ics` attachment. This does not respond to the invitation — that is
`respond_meeting`. Verified live against a received invite: the lookup finds the message and
`Reply`/`ReplyAll`/`Forward` create the expected items. Opening the window from the add-in, sending,
and whether prepending `body` keeps the forwarded invitation intact were not exercised.

## Meeting-response tools (1 — Automate approvals or higher; `Mutated = true`)

| Tool | Notes |
|---|---|
| `respond_meeting` | `response` (req): `accept`, `tentative` or `decline`, mapped to the `OlMeetingResponse` value by the one `RespondMeeting` helper. Resolves to `AppointmentItem` via `ResolveMeetingAppointment` (`MeetingItem.GetAssociatedAppointment(false)` when the id is a meeting request; optional `store_id` for shared calendars). Refuses (`IsError`) up front on an already-canceled meeting (`IsCanceledMeeting`) or an item that isn't a meeting the user was actually invited to (an organizer's own `olMeeting` copy, or a plain `olNonMeeting` appointment). Otherwise captures the subject, then `appt.Respond(olMeetingAccepted/Declined/Tentative, true, false)`, then `.Send()` on the response if non-null. Optional `message` sets `resp.Body` before `.Send()` — a short comment for the organizer (**unverified live** whether the organizer actually sees it). If `.Send()` fails, the result says so explicitly ("...but the response could not be sent to the organizer") — the local `Respond()` still went through, but the organizer was never notified. Recorded as undo barriers. Act on the whole series for a recurring meeting (no `occurrence_date`). |

There is no draft counterpart: `Respond()` commits a real calendar change at call time, so a draft
version that pre-selects a response would break the draft contract. To write to the organizer
instead, use `draft_email` with `event_id`.

## Auto-send tools (2 — Full autonomy only; send/create immediately, no review window; `Mutated = true`)

`draft_event` and `send_event` take the same arguments; `action` (`new` default, `edit`, `cancel`) picks the
operation. Each action has a draft tail (changes left unsaved, `Display(false)`, nothing persisted, no undo
record) and a send tail (`Save`/`Send` plus the undo records described elsewhere in this file). What the two
modes share is real code: `ResolveEventTarget` (event/occurrence lookup), `ParseEventEdit` (the refusals and
validation for an edit), `ApplyEventFields` (time/subject/body/location) and `BuildNewEvent` (new-event
construction). `EventAction` dispatches on `action`.

Every draft/compose tool above stays draft-and-display-only in every mode; these are
separate tools, reachable only at the top tier.

| Tool | Notes |
|---|---|
| `send_email` | The same `ComposeEmail` handler as `draft_email` with `send: true`: `m.Send()` instead of `Display(false)`, and an undo barrier is recorded. Same `action` (`new`/`reply`/`reply_all`/`forward`) and same `message_id`/`event_id` targets. `to` is required for `new` and `forward` (there is no window for a human to fill it in); a meeting forward with an unresolvable recipient is discarded and returns `IsError` instead of sending. |
| `send_event` · `action: new` | Same construction as `draft_event` (the shared `BuildNewEvent`), but `start`/`end` are required. No attendees → `a.Save()` (a plain calendar entry, nobody to notify). Attendees present → `MeetingStatus = olMeeting` then `a.Send()`, dispatching the invite. Both `AppointmentItem.Send()`/`.Save()` confirmed present via .NET reflection against the referenced PIA. Optional `recurrence` object builds a repeating series (see "Recurring events" below). |
| `send_event` · `action: edit` | General-purpose event edit. `event_id` (req), optional `occurrence_date` and `store_id`; at least one of `start`+`end` (required together), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` — else a clean `IsError` naming all seven. **Organizer-only:** refuses (`IsError`) on a canceled event (`olMeetingCanceled`/`olMeetingReceivedAndCanceled` — nothing to edit) or a meeting the user only attends (`olMeetingReceived` — points at Outlook's own "Propose New Time", see "Excluded / deferred"); these checks run against the recurring **master**, since an occurrence item from `GetOccurrence()` doesn't reliably report `MeetingStatus`. Time changes: a whole recurring series is shifted via `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` (Outlook throws `0xAF620009` on `Start`/`End` set directly on a recurring master — confirmed live); an occurrence or non-recurring event gets `appt.Start`/`.End` directly, after the `CheckOccurrenceReorderCollision` pre-check (below). `subject`/`body`/`location` are set directly (occurrence-level allowed). Attendee fields are whole-series/non-recurring only (`occurrence_date` + an attendee field is an `IsError`) and **replace the list wholesale per category** via `ReplaceAttendees` — `required_attendees` clears and re-adds only the required category, `optional_attendees` only the optional one, the organizer is never touched — re-adding via the same per-recipient `r.Resolve()` call `draft_event`/`send_event` (new) use (not `Recipients.ResolveAll()`); read the current list with `get_event` first to keep anyone. The result text names any attendee address that failed to resolve. Adding attendees to a plain event flips it to `olMeeting`. `.Send()`s if the result is or becomes a meeting — setting `ForceUpdateToAllAttendees = false` explicitly first (Outlook's own default: notify only added/removed attendees) — else `.Save()`s. A `COMException` from an occurrence `.Save()` is re-verified via `ResolveOccurrenceTarget` before being reported as failure, since Outlook can throw "Cannot save this item." after the change has already persisted. Undo: barrier when it sends or shifts a whole series' time, else a snapshot of the touched fields. |
| `send_event` · `action: cancel` | `event_id` (req), optional `occurrence_date` and `store_id`. Still-active organized meeting → `MeetingStatus = olMeetingCanceled` then `.Send()` (the cancellation notice), then moves the item to Deleted Items (best-effort — a failed move is reported, with the notice still sent; see "Unproven at runtime"). Plain appointment, or an **already-canceled event** (either status — the only way to dismiss one) → moves straight to Deleted Items (recoverable), nobody to notify. Refuses (`IsError`) only on a still-active meeting the user only attends (points at `respond_meeting` with `response: decline`, or `draft_email` with `event_id`). Same master-based organizer checks and false-negative `.Delete()` re-verification as `send_event` (edit). Undo: a move entry for a whole plain/already-canceled event; a barrier for a still-active meeting (the notice already went out) and for **every** occurrence-level cancellation. |

Gated by `OutlookTools.cs`'s `SendTierTools` set, requiring `EditingMode.FullAutonomy`
exactly (the ordinal check's top tier) — not reachable from Draft only or Automate
approvals. The model is told to prefer the draft tools unless the user clearly asked for
something to go out immediately.

## Recurring events

- **Occurrence targeting.** `draft_event`/`send_event` (edit)/`draft_event`/`send_event` (cancel) take an optional `occurrence_date` (a date from `list_events`'
  per-row `start`). A shared `ResolveOccurrenceTarget` helper resolves it via
  `RecurrencePattern.GetOccurrence(DateTime)` (confirmed present by reflection) into
  that one occurrence; omitting it acts on the whole series (or a non-recurring event).
  `get_event` and the meeting-response tools have no `occurrence_date` and act on the
  master.
- **Occurrence reorder pre-check.** Outlook rejects moving an occurrence past a later
  occurrence of the same series, or onto a day another occurrence already occupies —
  both confirmed live, both surfacing from automation only as a generic
  `COMException "Cannot save this item."` with no inner detail.
  `CheckOccurrenceReorderCollision` queries the calendar for other occurrences between
  the original and target dates before `.Save()` and returns an exact error naming the
  valid range instead.
- **Recurrence creation.** `draft_event`/`send_event` (new) take an optional `recurrence`
  object: `type` (required — `daily`/`weekly`/`monthly`/`monthlyNth`/`yearly`/
  `yearlyNth`, all 6 `OlRecurrenceType` values), `interval` (default 1),
  `days_of_week`, `day_of_month`, `instance` (1-4, or 5 for "last"), `month_of_year`,
  and at most one of `count`/`until` (neither = no end date). Type-specific fields
  omitted by the caller default from `start`'s own date. Validated before any COM call
  by the pure, unit-tested `OfficeAi.Shared.RecurrenceValidator` — every failure names
  the specific field and the fix — then applied via `AppointmentItem.GetRecurrencePattern()`
  after attendee/`MeetingStatus` setup, before `.Save()`/`.Send()`.
- **Undo asymmetry.** Occurrence-level edits are undo-able; occurrence-level
  cancellations and whole-series time changes are always barriers (see "Undo/redo
  tools").

## Shared calendars

- **Reading.** `list_events`' `mailbox` param lists someone else's calendar. The
  recipient is resolved via `Ns.CreateRecipient(mailbox).Resolve()` (unresolvable →
  `IsError`), then `Ns.GetSharedDefaultFolder(recipient, olFolderCalendar)` is called
  **once, only to read `.Store.StoreID`** — never its `.Items` — and the events
  themselves come from EWS `FindAppointments` + `CalendarView`, off the UI thread.
  Enumerating the shared folder's COM `Items` froze Outlook (confirmed live, even for a
  one-day range): that folder normally isn't cached offline, so every property read
  could be a blocking round trip to Exchange on Outlook's STA UI thread. Folder-open
  failure → an `IsError` suggesting the calendar isn't shared with the user or needs
  adding via Outlook's own "Open Calendar" first; a `StoreID` read failure alone just
  omits `store_id`.
- **Output.** Same per-event text shape as the own-calendar path (via
  `OfficeAi.Shared.SharedCalendarEventFormat`, deliberately separate code from the COM
  path's inline formatting), plus `calendar_owner: <name>` — the **resolved
  recipient's display name**, not the raw `mailbox` string — and `store_id: <StoreID>`.
  The zero-results message names the owner too. One accepted difference:
  `response`/`meeting_status` show EWS's label names (`Accept`, `Meeting`,
  `Cancelled`) rather than the COM path's `Ol*` enum names — display text only,
  nothing parses it.
- **Acting on a shared-calendar event.** `get_event`, `draft_event`/`send_event` (edit),
  `draft_event`/`send_event` (cancel), `respond_meeting`, `draft_email` (with `event_id`), `set_event_categories`, and
  `set_event_availability` all take an optional `store_id`, passed straight through to
  `ItemById`/`GetItemFromID` — `ItemById` only searches the caller's own default store
  without it. This is a lookup aid only.
- **No authorization or redaction layer of this add-in's own** (an explicit
  project-owner decision): whatever the caller's real Exchange sharing permissions let
  them do through Outlook's own UI — view (including a private item, if their access
  exposes it), edit, cancel, respond — these tools allow too; Outlook/Exchange itself is
  the only thing that refuses. Subject/location are passed through as Outlook resolves
  them, private items included. A permission failure surfaces as whatever exception
  `.Save()`/`.Send()` naturally throws, caught by each tool's existing error handling.

## Excluded / deferred

- **`update_event` / `delete_event`** — present in the reference EWS server's API but
  not ported under those names. `send_event` (edit) and `send_event` (cancel) (plus their `draft_`
  versions) cover the same ground.
- **Proposing a new time on a meeting the user only attends** (`MeetingStatus ==
  olMeetingReceived`) — real Outlook's "Propose New Time" feature, deliberately
  *not* implemented. `draft_event`/`send_event` (edit) refuse outright on a received
  meeting rather than attempt anything, because there's no clean way to honor the
  request: a plain `.Send()` here wouldn't be an authoritative reschedule Outlook
  actually honors (the user isn't the organizer), and .NET reflection against the
  actually-referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA found no
  "Propose"/"Counter"/"NewTime"-named member on `AppointmentItem`/
  `_AppointmentItem` or `MeetingItem`/`_MeetingItem` to call instead —
  `_AppointmentItem.Respond` only accepts `OlMeetingResponse`
  (Accept/Decline/Tentative). "Propose New Time" appears to be a ribbon/UI-level
  feature (MAPI counter-proposal properties) not cleanly exposed through classic
  COM automation. A future implementation would likely need raw `PropertyAccessor`
  MAPI-property manipulation rather than a typed object-model call — untried here,
  and risky without a live mailbox to validate against.

## Object Model Guard

In-process VSTO add-ins that use the **VSTO-supplied `Application`** object are trusted
by default — reading `Body` / `Recipients` / `SenderEmailAddress` /
`AddressEntry.PrimarySmtpAddress`, calling `PropertyAccessor.GetProperty` or
`Attachment.SaveAsFile` do **not** raise the "a program is trying to access…" prompt on
default settings. Prompts appear only under Trust Center → Programmatic Access set to
"Always warn", or an Exchange public-folder security form. `MailItem.Send` — the
highest-risk call — is used only by the Full-autonomy auto-send tools (`send_email`); whether it triggers a guard prompt
under a stricter Trust Center setting has not been checked live.

## Structural fragility

Everything is addressed by `EntryID`. It is stable while an item stays put but changes
on `Move` and is store-specific; the mutating tools that move items return the new id,
and every other tool re-resolves via `GetItemFromID` each call (with `store_id` for a
shared-calendar item). All recurring occurrences of a calendar series **share one
`EntryID`**, so `get_event` / `respond_meeting`
/ `draft_email` (with `event_id`) cannot target a single occurrence — they act on the master
series. The calendar-editing tools can, via `occurrence_date` (see "Recurring events");
without it they too act on the whole series. `list_events` carries each occurrence's
`start` as the disambiguator.

**Known limitation, confirmed live (Gmail-connected calendar):** on a mailbox connected
via Google's Gmail/Google Workspace sync, moving a calendar item to Deleted Items does
not appear to be durable the way it is on Exchange — a `send_event` (cancel) call that
reported success moving an item to Deleted Items was followed immediately (no other
action in between) by that item relocating itself to a `Drafts` folder, with no code in
this add-in touching it a second time. Most likely Google Calendar's own sync
reconciling the move shortly after, outside this add-in's control.
`undo_last_action`'s own "did the item change since?" conflict check caught the
mismatch and refused rather than guessing or overwriting — the designed safety net
worked, and no data was lost — but this means `RecordMove`'s core assumption (an item
stays wherever the last recorded move put it, until this add-in moves it again) does
not reliably hold for Gmail-connected calendars specifically. Treated as a known
account-type-specific risk, not a code bug in `send_event` (cancel), `send_event` (edit), or the undo
stack.

## Unproven at runtime

The project builds clean (MSBuild Debug + Release, 0 warnings). But **most COM paths
have not been exercised in a real Outlook** — the Explorer pane lifecycle /
IUnknown-identity keying, WebView2 rendering inside an Explorer task pane,
`Folder.GetTable` column names, the `list_events` recurrence ordering,
`Recipient.FreeBusy`'s string format / month coverage (`find_meeting_slots`), and
several enum-name / method-signature assumptions (`olEmbeddeditem` casing,
`AppointmentItem.Respond` argument types, `MailItem.MarkAsTask`) compiled against the
interop assembly but are not yet confirmed live. Early manual testing via the mock
server exercised `list_emails`, `search_emails`, `list_folders`, `list_events`, and
`list_tasks` against a real mailbox successfully. **Verified live:** the color-tag
tools (category enumeration, tag creation, event `Categories` round-trip),
`set_event_availability` (all 5 values, no attendee notification on an organized
meeting) and its undo/redo, and the recurring-series items listed below.

**EWS paths.** `search_contacts`'s EWS path is **compiled but not confirmed against a
live on-prem Exchange**: the `ResolveName` calls and their `NameResolution` mapping,
`ExchangeService.UseDefaultCredentials` (Windows Integrated Auth to on-prem CAS),
`Account.AutoDiscoverXml` shape / `EwsAutodiscoverXml.ParseEwsUrl`, the
`AutodiscoverUrl` fallback, and the off-thread `Task.Run` actually keeping Outlook
responsive during the call. The async delegate refactor (`ToolExecutor` →
`Task<ToolResult>`, `async void OnWebMessageReceived`) builds clean for all four
add-ins; a runtime smoke of one tool per app confirming nothing regressed is still
pending. For `list_events`' shared-calendar EWS path, one assumption still needs live
confirmation: that the single remaining `GetSharedDefaultFolder` call (used only for
`StoreID`) was never itself part of the freeze — if Outlook still freezes, that call is
the next suspect.

**Shared calendars — permission tiers untested.** Behavior has not been exercised
against a real second mailbox with different Exchange sharing levels (Full Access /
Editor / Reviewer / free-busy only / titles+locations only / no access): neither the
exact read-path error messages, nor whether partial-read tiers degrade gracefully, nor
what a genuinely unpermitted cross-mailbox write (e.g. `send_event` (edit) with a `store_id`
from a Reviewer-only calendar) actually returns — assumed to be a `COMException` from
`.Save()`/`.Send()`, untested.

**`draft_event`/`send_event` (edit)** — compiled, not exercised live for their meeting
paths. Ranked by consequence of being wrong:

1. **`ReplaceAttendees`'s recipient-clearing** — highest risk, since a wrong
   organizer-exclusion could drop the user's own organizer entry or fail to actually
   clear old attendees, feeding directly into a `.Send()` with real people.
2. **Converting a plain event into a meeting via `send_event` (edit)** — adding attendees to an
   *existing*, previously-saved item (unlike `send_event` (new), which only ever adds
   attendees to a brand-new one) is an untested combination.
3. **`subject`/`body`/`location` edits on a recurring master** — likely fine (none of
   these are part of `RecurrencePattern`), but unverified.

Also unverified: whether Outlook accepts a plain `.Send()` on a modified organizer-owned
meeting as a real update notice in every case (e.g. an all-day flag change). **Open
question:** whether `draft_event` (edit)'s unsaved changes (mutate an existing item,
`Display(false)`, never `.Save()`/`.Send()`) have the same "persists on window-close
without an explicit action" risk that `MeetingStatus` turned out to have for
`draft_event` (cancel). If so, closing a `draft_event` (edit) window on a meeting the user
organizes without clicking Send could silently move it in the organizer's own calendar
while attendees still see the old time. Not confirmed either way.

**`send_event` (cancel)** — its organized-meeting branch calls `.Move()` immediately after
`.Send()` on the same item, a sequence no other tool here performs. Whether Outlook still
permits relocating a just-canceled-and-sent appointment is unconfirmed; the code catches
a failure there and reports the (genuinely-sent) cancellation succeeded regardless,
telling the user to delete the stray calendar entry manually if the move didn't take.

**Recurring series — partly live-verified.** Confirmed live, against a plain
(non-meeting) recurring series: occurrence resolution via
`ResolveOccurrenceTarget`/`RecurrencePattern.GetOccurrence`, occurrence-level reschedule
(including its snapshot-based undo), occurrence-level cancellation (as a barrier), the
false-negative `.Save()`/`.Delete()` re-verification, the two occurrence-reordering
rejections `CheckOccurrenceReorderCollision` guards against, and that `Start`/`End`
cannot be set directly on a recurring master. Still unconfirmed:

- **The master-vs-occurrence organizer-authority check has never been exercised against
  a genuine recurring *meeting* occurrence with attendees** — every live test used a
  plain series. Checking `master.MeetingStatus`/`IsCanceledMeeting(master)`/
  `IsReceivedMeeting(master)` instead of the occurrence's own value is a reasoned fix
  based on `GetOccurrence()`'s known-unreliable behavior for this property, not one
  confirmed by observing the bug on a real meeting occurrence.
- **The whole-series `RecurrencePattern` time change**: whether
  `PatternStartDate`/`StartTime`/`EndTime` interact correctly with the series' existing
  `DayOfWeekMask`/other pattern fields when only the time (not the day) changes, whether
  `.Send()`/`.Save()` on the master after mutating the pattern behaves identically to
  mutating it directly, and what the barrier's undo-refusal messaging looks like in
  practice.
- **`CheckOccurrenceReorderCollision` may not cover every reason Outlook can reject an
  occurrence reorder** — any rejection reason it doesn't anticipate still reaches the
  model as the same generic "Cannot save this item."
- **`recurrence` on `draft_event`/`send_event` (new)**: `RecurrenceValidator`'s unit tests
  cover only the pure validation logic. The COM application path —
  `GetRecurrencePattern()` after attendee/`MeetingStatus` setup, before
  `.Save()`/`.Send()` — is unverified live, as is whether that ordering matters.

**Meeting responses.** Whether the organizer actually sees `respond_meeting`'s `message`
text on the delivered response (vs. it being dropped or overwritten by Outlook's own
response template) has not been confirmed against a real received invite.
