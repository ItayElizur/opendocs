# OutlookTools.Ews.cs

## Class-level comment - EWS tool orchestration posture

Every tool-level thing in this add-in that needs Exchange/EWS to be
reachable, in one place, separate from the pure-COM tools in the other
OutlookTools.*.cs files (which only need the already-running, already-
signed-in Outlook client and work the same regardless of mailbox type
or network state). Low-level EWS wire calls themselves live one layer
further down, in OutlookEws.cs; this file is the tool-facing
orchestration on top of it - resolving which account/endpoint to use,
and turning a raw EWS result into a ToolResult or a plain value another
tool can consume.

On-prem Exchange only, same posture everywhere in this file: no
Exchange account in the profile, or EWS unreachable, is either a clear
IsError result (search_contacts, where EWS is the only way to do the
job) or a graceful fallback to a COM-only default (find_meeting_slots'
work-week lookup, where a COM-derived value already exists to fall
back to). Never a silent partial success.

## `SearchContactsAsync`

search_contacts resolves a name/email fragment through EWS ResolveName
(server-side ANR over Contacts then the GAL) - the same thing the
native Address Book does, and what the mcp-outlook reference does. It
is one of two tools that aren't Outlook COM: the object model can't do
a multi-result directory search, and running EWS off the UI thread
(OutlookEws.ResolveNamesAsync -> Task.Run) is what keeps Outlook
responsive during the call. The previous implementation walked every
contact folder in every store on the UI thread and froze/crashed
Outlook. See docs/tools/outlook.md and OutlookEws.cs.

## `_cachedWorkWeek` / `_workWeekResolved`

Cached once per process, same lifetime/posture as OutlookEws.CachedUrl
- the mailbox's configured work days/hours don't change mid-session,
and each lookup is a network round trip via EWS. Mirrors CachedUrl's
posture exactly: _workWeekResolved only latches true on a SUCCESSFUL
lookup, never on failure - a transient EWS/network blip on the first
find_meeting_slots call must not permanently disable the real
work-week lookup (falling back to Sun-Thu/9-18) for the rest of the
Outlook session, which can run for days.

## `ResolveWorkWeekAsync`

Best-effort: the real work week/hours come from EWS's
GetUserAvailability (see OutlookEws.GetWorkingHoursAsync's own
comment for why - no COM equivalent exists). On-prem Exchange only,
same as search_contacts; any failure (no Exchange account, EWS
unreachable, etc.) falls through to `LocalWorkWeek` below, and only
if that also finds nothing does it return null, so the callers (the
prompt and find_meeting_slots) fall back to "ask the user" or a
COM-only default rather than failing - unlike search_contacts, where
EWS isn't optional, here it's an enhancement over an already-working
default.

## `LocalWorkWeek`

Outlook's File > Options > Calendar > Work time settings, read from
`HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Calendar`:
`WorkDay` (a day bitmask, decoded by WorkWeekPrompt.DaysFromOutlookMask),
`CalDefStart` and `CalDefEnd` (minutes after midnight). A profile with
only Gmail/IMAP/POP accounts has no Exchange mailbox, so this is its
only source; on a live test machine a Sun-Thu work week set in Outlook
showed up there as `WorkDay = 248`. Missing values (never changed
from Outlook's defaults) or an impossible range return null. It is
not cached like the EWS result: a registry read is cheap, and the user
may change the settings while Outlook runs.

## `ResolveEwsUrlAsync`

Shared by every EWS-dependent tool: resolve the endpoint once
(OutlookEws.CachedUrl short-circuits every call after the first,
process-wide, regardless of which tool triggered the discovery).
Always throws InvalidOperationException on failure, with a message
specific enough for search_contacts to surface directly -
ResolveWorkWeekAsync above just catches and discards it (a graceful
null is all it wants).
