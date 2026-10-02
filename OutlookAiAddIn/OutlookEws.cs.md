# OutlookEws.cs

## Class-level comment - why this file exists and its scope

The non-COM calls in the Outlook add-in - the raw EWS wire layer.
Tool-facing orchestration built on top of this (which account/endpoint
to use, turning a result into a ToolResult) lives in
OutlookTools.Ews.cs, not here; this file only wraps the EWS Managed API
itself. Two operations: ResolveNamesAsync (search_contacts - a
server-side Ambiguous Name Resolution over Contacts then the GAL,
exactly what the native Address Book dialog does and what the
mcp-outlook reference's account.protocol.resolve_names does) and
GetWorkingHoursAsync (find_meeting_slots' work-week default - see its
own comment). Both are deliberately NOT Microsoft.Office.Interop.Outlook:
the COM object model can do neither a multi-result directory search nor
expose a mailbox's configured work week, and EWS is plain HTTP with no
STA affinity so both run off the UI thread (Task.Run) and never freeze
Outlook.

Auth is ExchangeService.UseDefaultCredentials (Windows Integrated Auth as
the signed-in user) - the .NET equivalent of mcp-outlook's auth_type=sspi.
No stored credentials, no impersonation. On-prem Exchange only.

This file has no `using Microsoft.Office.Interop.Outlook`: the two
namespaces collide on Contact, EmailAddress, Folder, Task, Item, ...

## `GetWorkingHoursAsync`

GetUserAvailability's WorkingHours is EWS's documented, server-side
source for a mailbox's configured work days/hours - the same data
Outlook itself uses to shade "outside working hours" in the
scheduling assistant. Unlike Outlook's local Calendar Options
dialog, this has no COM equivalent at all: confirmed via .NET
reflection against the referenced Microsoft.Office.Interop.Outlook
PIA that no Application.CalendarOptions property (or any
WorkDay*/FirstDayOfWeek member) exists anywhere in that assembly,
so this EWS call is the only real, non-hardcoded source for it.

## `GetSharedCalendarEventsAsync`

list_events' shared-calendar path (mailbox parameter): EWS CalendarView
expands recurring appointments server-side within [start, end) - the EWS
equivalent of the COM path's IncludeRecurrences=true, without that path's
"expand everything, then filter" cost, and off the UI thread like every
other EWS call in this file. See docs/superpowers/specs/
2026-09-30-outlook-shared-calendar-ews-design.md for why this exists (the
COM path froze Outlook, confirmed live even for a one-day range).

## `GetSharedCalendarEvents` - convertIdBroken

A ConvertId timeout on one item means the connection is broken for
the rest of this call too - once that happens, stop calling ConvertId
entirely rather than letting every remaining item pay its own full
timeout (up to ~12 minutes for 50 items on a mid-loop connection
drop, with list_events just hanging the whole time). A non-timeout
failure on a single item is left alone (still tried for subsequent
items) since that's plausibly just a one-off glitch for that item.

## `TryGetBool` / `TryGetResponseType` / `TryGetAppointmentType` / `TryGetOrganizerLabel`

Some PropertySet-requested properties aren't reliably populated by every
Exchange server for every item type returned from CalendarView - confirmed
live 2026-10-01: Appointment.IsCancelled threw ServiceObjectPropertyException
("This property was requested, but it wasn't returned by the server") for a
plain, non-meeting appointment, despite IsCancelled being in the PropertySet
above. TryGetProperty is EWS's documented safe-read for exactly this case -
these helpers wrap it so a missing property degrades to a sensible default
instead of crashing the whole list_events call. Applied to every
meeting-specific property read here (IsAllDayEvent/AppointmentType/
MyResponseType/IsMeeting/IsCancelled/Organizer) since they're all plausibly
absent for some item type, unlike Subject/Start/End/Location which are core
enough to every calendar item that direct access is kept as-is.

## `ConvertToEntryId`

Converts this item's EWS id to the classic Outlook/MAPI EntryID format so
it stays resolvable via the existing Ns.GetItemFromID(entryId, storeId)
every read/write tool already uses (get_event's own store_id parameter) -
list_events' shared-calendar output must not change shape just because
this path now fetches data via EWS instead of COM. Failure here (should
be rare - reasoned from the EWS Managed API surface, not yet verified
live) degrades to an empty event_id rather than dropping the whole event:
the caller still sees the event exists, just can't act on it via
get_event/edit_event until this is investigated.

A timeout is deliberately NOT caught here - it propagates to
GetSharedCalendarEvents' loop, which needs to distinguish "this was a
timeout" (stop calling ConvertId for the rest of the batch) from any
other failure (fine to keep trying subsequent items).

## `ResolveNames`

Queries Contacts and the Directory (GAL) separately and merges both,
rather than the single-call ResolveNameSearchLocation.ContactsThenDirectory
this used to use. ContactsThenDirectory short-circuits: Exchange's ANR
match against a personal Contacts folder only checks DisplayName (not
GivenName/Surname/company/etc.), and if that phase returns anything at
all, the fuller GAL ANR (which does cover first/last name) never runs -
confirmed as the cause of search_contacts matching display name only.
ContactSearchFormat.Format (the caller) already dedupes by email/name
and applies the limit, so merging both lists here is safe.

## `DisplayNameFrom`

The directory's own display name - may be org-formatted (e.g. a GAL
entry whose AD displayName attribute reads "Dept/Unit/Title") and look
nothing like the person's actual name. Kept separate from
FullNameFrom (below) specifically so a caller can show both when they
disagree - confirmed live 2026-10-01: an agent saw only this field,
it didn't resemble the searched name, and it wrongly concluded the
search had failed even though the match was correct.
