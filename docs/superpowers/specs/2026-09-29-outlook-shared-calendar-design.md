# Design: view another user's calendar via `list_events`

**Status:** approved design, not yet implemented. Branch: `feature/outlook-shared-calendar` (fresh branch from `main`, post-PR#27).

## Problem

`list_events` only ever reads the current user's own default calendar (`Ns.GetDefaultFolder(olFolderCalendar)`). There's no way to view a colleague's calendar even when the user has been granted delegate/reviewer access to it in Exchange — a common real-world need ("what does my manager have on Thursday").

## Ground truth (verified this session)

- `NameSpace.GetSharedDefaultFolder(Recipient, OlDefaultFolders)` exists on the referenced PIA (confirmed via .NET reflection against `_NameSpace`) — this is the correct, standard mechanism for opening a folder someone else has shared/delegated, matching what Outlook's own "Open Calendar → Open Shared Calendar" UI does internally.
- **Not verified live**: exact behavior across Exchange calendar-sharing permission tiers (no access / "can view when I'm busy" (free-busy only) / "can view titles and locations" / "can view all details" / full editor). No second mailbox with configured sharing was available to test against during design. The user has confirmed they can test this against a real shared calendar once built, elsewhere.
- Reasoned-but-unconfirmed expectation, based on how Exchange permission-scoped calendar sharing is generally known to behave: opening the folder itself should succeed regardless of permission tier (Exchange doesn't hide the folder's existence), but individual item properties beyond what the granted tier allows are expected to come back redacted/generic (e.g., `Subject`/`Location` shown as blank or a placeholder like "Busy") rather than the call throwing — this needs to be confirmed against a real shared calendar, not assumed correct.
- If the current user has **no** access at all to the target calendar, `GetSharedDefaultFolder` is expected to throw a `COMException` — also unconfirmed live, must be caught defensively regardless of the exact HRESULT.

## Design

### 1. Schema

`list_events` gains one new optional parameter:

```
mailbox: string  -- email address of the calendar owner to view. Omit to view your own calendar (unchanged default behavior).
```

No other tool's schema changes. `get_event`, `edit_event`, `cancel_event`, etc. are explicitly **out of scope** for this change — acting on an event_id from someone else's calendar (rescheduling/canceling something you don't own) is a separate, much bigger authority question not addressed here. `list_events` against a shared calendar is read-only, same as it already is for your own.

### 2. Resolution and folder access

```
if mailbox is given:
    recipient = Ns.CreateRecipient(mailbox)
    if !recipient.Resolve():
        return IsError: "Could not resolve \"" + mailbox + "\" - check the email address."
    try:
        cal = (Outlook.Folder) Ns.GetSharedDefaultFolder(recipient, olFolderCalendar)
    catch (COMException ex):
        DebugLog.WriteException("ListEvents GetSharedDefaultFolder", ex)
        return IsError: "Could not open " + mailbox + "'s calendar - you may not have been granted access to view it, or need to add it via Outlook's own \"Open Calendar\" first. (" + ex.Message + ")"
else:
    cal = (Outlook.Folder) Ns.GetDefaultFolder(olFolderCalendar)   // existing, unchanged
```

Everything after folder resolution — `Sort("[Start]")` → `IncludeRecurrences = true` → `Restrict(filter)` → the per-event output loop — is **completely unchanged**, reusing the exact same load-bearing query order `ListEvents` already uses. This is a single resolution-point change, not a rewrite.

### 3. Redacted/limited-visibility events

No special-case code for different permission tiers. Whatever Outlook itself returns for `Subject`/`Location`/etc. (full text, blank, or a generic placeholder) is passed through as-is in the existing per-event output format — Outlook is the one enforcing the permission tier at the COM layer, not this code. If live testing shows something Outlook returns that's actively confusing (e.g., a blank line with no indication *why* it's blank), a follow-up can add a one-line note, but no such handling is included speculatively before that's confirmed necessary.

Add one addition to the per-event output when `mailbox` was given: a `calendar_owner: <mailbox>` line, so a caller listing several people's calendars in the same conversation (or reading back a mixed transcript) can tell which calendar each event came from. `event_id`/`get_event` continues to work exactly as it does today for these events (EntryIDs are globally unique across stores; no new store-hint plumbing is introduced, since none is needed for `Ns.GetItemFromID`'s existing no-store-hint call pattern already used everywhere else in this codebase).

### 4. Error cases

1. `mailbox` doesn't resolve to a valid recipient → `IsError`, as above.
2. `mailbox` resolves but the current user has no sharing access → `IsError` (catches the `COMException`), as above.
3. `mailbox` resolves and some access exists, but the range has no events → existing "No events between X and Y." message, with the addition of naming whose calendar: `"No events on " + mailbox + "'s calendar between " + ... + "."` when `mailbox` was given.

### 5. Risk / verification

**Explicitly unverified, flagged for the user's own live test once built** (per their offer to test elsewhere): the exact behavior at each Exchange sharing tier — whether an event property comes back blank, a placeholder, or the field is simply omitted; whether `GetSharedDefaultFolder` throws immediately on zero access or only fails lazily when `.Items` is enumerated; whether a *self*-shared calendar (viewing your own address via `mailbox`) behaves identically to the existing no-`mailbox` path (expected to, since it'd resolve to the same underlying folder, but not explicitly tested).

## Out of scope

- Writing to another user's calendar (creating/editing/canceling events on their behalf) — a distinct, much larger authority question.
- `find_meeting_slots`-style free/busy-only queries for a named person are already covered by the existing `find_meeting_slots` tool; this design is specifically about full event listings, not availability.
- Any UI/chat-ui changes — this is a pure `list_events` schema + `OutlookTools.Calendar.cs` change, same shape as prior additions this project.
