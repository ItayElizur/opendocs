# Design: rewrite `list_events`' shared-calendar path onto EWS

**Status:** approved design, not yet implemented. Branch: `feature/outlook-shared-calendar-ews` (fresh branch from `main`, post the `NullReferenceException`/`Store` fix landed 2026-09-30).

## Problem

`list_events` with a `mailbox` argument (the shared-calendar feature from `2026-09-29-outlook-shared-calendar-design.md`) freezes Outlook — confirmed live by the project owner even for a single-day range, so this isn't proportional to event count in any simple way; it's the mechanism itself that's heavy.

Root cause: the shared-calendar path is 100% synchronous `Microsoft.Office.Interop.Outlook` COM, running on Outlook's own UI thread (Outlook Interop objects are STA-bound - they can't be moved to `Task.Run` the way this add-in's EWS calls already are, since EWS is a separate, thread-agnostic HTTP API). The same COM shape is used for the caller's own calendar too, but it's fast there because Cached Exchange Mode already has the primary mailbox's data offline; a folder opened via `GetSharedDefaultFolder` is normally **not** cached offline, so every property read during enumeration is potentially a live, blocking round trip to Exchange. `IncludeRecurrences = true` (load-bearing for recurrence expansion, see `ListEvents`' own comment) makes this worse: it's a well-documented Outlook COM pitfall that `IncludeRecurrences` forces recurrence expansion *before* `Restrict` narrows anything down, so a mailbox with several recurring series can mean a large expansion cost even for a one-day query window.

This add-in already has a working precedent for exactly this class of problem: `search_contacts` was rewritten off Outlook COM onto EWS specifically because "the object model can't do a multi-result directory search, and running EWS off the UI thread... is what keeps Outlook responsive" (`OutlookTools.Ews.cs`'s own file-header comment) - the previous COM-walk implementation there "froze/crashed Outlook," per the same comment. This is the same fix, applied to the same class of bug, on a different tool.

## Ground truth (established this session, from the existing EWS code in this project)

- `OutlookEws.cs` already has a working `ExchangeService` setup (`NewService(Uri)`), TLS 1.2 fix, `UseDefaultCredentials`, 15s timeout, and a `Task.Run`-based async wrapper convention (`ResolveNamesAsync`, `GetWorkingHoursAsync`). The new call follows the exact same shape.
- `OutlookTools.Ews.cs` already has `ResolveEwsUrlAsync()` (cached EWS endpoint discovery, shared by every EWS-dependent tool) and the exact exception-handling shape for turning EWS failures into `IsError` results (timeout / `ServiceResponseException` / `WebException` / generic) — reused verbatim, not reinvented.
- `SmtpOf(Outlook.AddressEntry)` (`OutlookTools.cs`) already resolves a COM `AddressEntry` to a clean SMTP address (handles the `"EX"`/legacyExchangeDN case via `GetExchangeUser().PrimarySmtpAddress`) — reused to get the shared mailbox's real SMTP address from the already-resolved `Recipient`, rather than trusting the caller's raw (possibly a bare name, not an address) `mailbox` argument.

## Ground truth (reasoned from the EWS Managed API 2.2 surface this project already references — NOT verified live in this session; no live Exchange/Outlook available in this environment)

- `new Ews.FolderId(Ews.WellKnownFolderName.Calendar, new Ews.Mailbox(smtp))` is the standard EWS mechanism for referencing another mailbox's default Calendar folder without opening a COM store — the EWS equivalent of `GetSharedDefaultFolder`, but pure HTTP.
- `ExchangeService.FindAppointments(FolderId, CalendarView)` with a `new Ews.CalendarView(start, end, maxItemsReturned)` expands recurring appointments into individual occurrences **server-side**, within the given date window — the EWS equivalent of `IncludeRecurrences = true`, without the "expand everything, then filter" COM cost.
- `ExchangeService.ConvertId(new Ews.AlternateId(Ews.IdFormat.EwsId, item.Id.UniqueId, smtp), Ews.IdFormat.EntryId)` converts an EWS item ID to the classic Outlook/MAPI EntryID string format — needed so `event_id` values from this path stay compatible with the existing `get_event`/`edit_event`/`cancel_event`/etc. tools' `Ns.GetItemFromID(entryId, storeId)` resolution, completely unchanged.
- These three are all real, documented members of the EWS Managed API 2.2 surface (`Microsoft.Exchange.WebServices.Data`) already referenced by this project (see `OutlookEws.cs`'s existing `Ews.ResolveNameSearchLocation`, `Ews.NameResolutionCollection`, etc. usage) — a clean build is the first real check on member names/signatures being right (the compiler will name exactly what's wrong against the actual referenced PIA/DLL version, same as this project's established "confirmed via .NET reflection against the referenced PIA" verification standard elsewhere).
- `Appointment.MyResponseType` (`Ews.MeetingResponseType`), `Appointment.IsMeeting`, `Appointment.IsCancelled`, `Appointment.AppointmentType` (`Ews.AppointmentType`: `Single`/`Occurrence`/`Exception`/`RecurringMaster`) are the properties used to synthesize `response`/`meeting_status`/`recurring` — these don't map 1:1 to the COM enums' `.ToString()` text (e.g. EWS's `"Accept"` vs. COM's `"olResponseAccepted"`). **Accepted, deliberate difference**: these are human/LLM-readable display strings only — confirmed (via `grep`) that nothing else in this codebase parses `response:`/`meeting_status:` text back out of `list_events`' output, so a differently-worded but equally clear label is not a functional regression.

## Open question this design does NOT resolve, and the decision made anyway

**Can `store_id` be obtained without any COM call at all?** `get_event`'s `store_id` parameter exists because `Ns.GetItemFromID(entryId, storeId)` cannot resolve a cross-store item from a bare EntryID alone (see `GetEvent`'s own comment, added in PR #28 review). EWS's ID-conversion surface has no direct equivalent of a raw MAPI "StoreID" hex string (that's a COM/MAPI-specific concept), so there's no pure-EWS way found to obtain it.

**Decision:** keep exactly one COM call per `list_events` invocation — the existing `Ns.GetSharedDefaultFolder(recipient, olFolderCalendar)` — but use it *only* to read `.Store.StoreID`, never touch `.Items` on it (no `Sort`/`Restrict`/enumeration — that's the part EWS replaces). Opening a folder object without enumerating its items is a single, bounded round trip, not the N-items-times-M-properties cost the freeze report points at.

**This is the one assumption in this plan that most needs live confirmation** (Task 3 has an explicit manual-verification step for it). If `GetSharedDefaultFolder` *itself* turns out to be the dominant freeze cost (not the enumeration after it), this design would not fully fix the freeze, and a follow-up step would be needed to find or accept a different way to get `store_id` — explicitly out of scope for this plan; report back with findings if this assumption doesn't hold.

## Scope

**In scope:** `list_events`' shared-calendar (`mailbox` argument given) path only, in `OutlookTools.Calendar.cs`. Rewritten to resolve events via EWS instead of COM `Items.Sort/IncludeRecurrences/Restrict`. Output format (the `- event_id: ...` block, the zero-results message, `calendar_owner`/`store_id` lines) stays byte-for-byte identical to today.

**Out of scope (explicitly deferred, not this plan):**
- `list_events`' own-calendar (no `mailbox`) path — stays COM, unchanged (it's fast; no reason to touch it).
- `get_event`, `edit_event`, `cancel_event`, `accept_meeting`/`decline_meeting`/`tentative_meeting`, and their draft counterparts — all keep resolving via `Ns.GetItemFromID(entryId, storeId)` exactly as today. Nothing about how a returned `event_id`/`store_id` pair gets *used* changes, only how `list_events` *produces* one.
- Any change to `entry.ts`'s tool schema — `mailbox` already exists as a parameter; no new parameters are added.
- Any UI/chat-ui change.

## Design

### 1. New pure formatter (`OfficeAi.Shared/SharedCalendarEventFormat.cs`)

A pure, unit-testable function reproducing the per-event text block `QueryCalendarItems` already writes inline, used by the *new* EWS path — same architectural role as `ContactSearchFormat`/`RecurrenceValidator`/`MeetingSlots` already play for `search_contacts`/`create_event`/`find_meeting_slots`.

**Deliberately not wired into `QueryCalendarItems` (the existing COM / own-calendar path) too**, despite the duplication that leaves: that method's `try { ...Iso(appt.Start)... } catch { }` around the start/end line is load-bearing (a malformed item's property-read exception silently omits just that one line, not the whole event) and has no equivalent need on the EWS side, where `Start`/`End` are already safely materialized `DateTime` values by the time they reach `SharedCalendarEventRow`. Forcing both call sites through one shared function would mean either re-deriving that per-line-omission behavior inside the shared formatter (which doesn't know how to skip a sub-line, only whole rows) or dropping it — not worth the risk to an already-correct, already-shipped code path for a DRY win. `QueryCalendarItems` stays exactly as it is today; only the new EWS path uses `SharedCalendarEventFormat`.

```csharp
public struct SharedCalendarEventRow
{
    public string EntryId;
    public string Subject;
    public DateTime Start;
    public DateTime End;
    public string Location;
    public string Organizer;
    public bool AllDay;
    public bool Recurring;
    public string ResponseStatus;
    public string MeetingStatus;
}

public static class SharedCalendarEventFormat
{
    public static string Format(IReadOnlyList<SharedCalendarEventRow> rows, int limit, string calendarOwner, string storeId);
}
```

`Format` reproduces `QueryCalendarItems`' existing per-event text exactly (same field order, same `calendar_owner`/`store_id` lines only when `calendarOwner != null`), so the calling code path (COM or EWS) is invisible in the tool's output.

### 2. New EWS wire call (`OutlookAiAddIn/OutlookEws.cs`)

```csharp
public static Task<IReadOnlyList<SharedCalendarEventRow>> GetSharedCalendarEventsAsync(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit)
```

Same `Task.Run` + `NewService(url)` shape as `ResolveNamesAsync`/`GetWorkingHoursAsync`. Builds a cross-mailbox `FolderId`, a date-scoped `CalendarView` (server-side recurrence expansion), requests exactly the properties the formatter needs, and converts each item's ID to a classic EntryID via `ConvertId`. Returns plain `SharedCalendarEventRow` values — no `Ews.*` type crosses out of this file, same boundary `ResolveNamesAsync` already keeps (returns `KeyValuePair<string,string>`, not an EWS type).

### 3. Orchestration + `ListEvents` rewrite (`OutlookAiAddIn/OutlookTools.Calendar.cs`, `OutlookTools.Ews.cs`)

`ListEvents` becomes `ListEventsAsync` (async, per the existing `case "list_events": return await ListEventsAsync(input);` convention already used for `search_contacts`/`find_meeting_slots`). The own-calendar (no `mailbox`) branch is untouched. The shared-calendar branch becomes:

1. Resolve `recipient` (unchanged: `Ns.CreateRecipient(mailbox).Resolve()`).
2. Open the shared calendar folder *only* to read `.Store.StoreID` (see the Open Question above) — same `GetSharedDefaultFolder` call as today, same error message on failure, but nothing past `.Store` is touched.
3. Resolve the shared mailbox's real SMTP via `SmtpOf(recipient.AddressEntry)`.
4. `await ResolveEwsUrlAsync()` (shared, cached after the first call by any EWS tool).
5. `await OutlookEws.GetSharedCalendarEventsAsync(url, sharedSmtp, start, end, limit)`, wrapped in the same timeout/`ServiceRequestException`/`WebException`/generic catch shape `SearchContactsAsync` already uses, worded for `list_events`.
6. `SharedCalendarEventFormat.Format(rows, limit, displayName, storeId)` → existing `BuildListEventsResult`, unchanged.

### 4. Docs (`docs/ai-tool-surface.md`)

New dated update block (matching the existing `> **Update 2026-09-29 (...)**` convention) describing the EWS rewrite, why (the freeze), and re-flagging the `store_id`-via-`GetSharedDefaultFolder` assumption as the thing to watch if a freeze is *still* reported after this change.

## Risk / verification

**Everything under "Ground truth (reasoned... NOT verified live)" above needs a real shared mailbox to confirm.** Task 3 of the implementation plan has an explicit live-verification checklist: confirm the freeze is actually gone, confirm returned events match what the COM path would have shown, confirm a `get_event` call against a returned `event_id`/`store_id` pair still resolves correctly (proving the `ConvertId` + kept-`GetSharedDefaultFolder`-for-`store_id` combination actually round-trips).
