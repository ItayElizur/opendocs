# OutlookTools.Compose.cs

## `ApplyRecurrence`

Applies a validated RecurrenceSpec (OfficeAi.Shared - pure, unit
tested in RecurrenceValidatorTests) to an AppointmentItem, converting
to the actual Outlook enums/flags OfficeAi.Shared can't reference
directly (it has no Outlook PIA reference, same split as
ColorUtil/BusyStatus in OutlookTools.Categories.cs). Called after
attendees are set but before .Save()/.Send() in both create_event
and draft_event - unverified whether GetRecurrencePattern() before
vs. after setting MeetingStatus/attendees matters; this ordering
(identity first, schedule second) matches natural Outlook UI flow
and is the one to verify live in Task 8.

## `ReadRecurrence` - defaulting from the event's start date

Fill in defaults from the event's own start date when the
caller omits a field the chosen type needs, rather than
forcing every call to spell out values already implied by
"recurring starting from <start>". Only fills what's
genuinely missing - an explicitly-provided value is never
overridden. RecurrenceValidator.Parse stays the strict, pure
validator; by the time it runs here the spec already looks
complete for whichever type was requested. An empty
days_of_week array is treated the same as omitted (above) -
it can't satisfy any type's requirement as-is, and a caller
sending [] almost certainly meant "use the default", not
"explicitly zero days".

## `AddAttendees` - per-recipient Resolve()

Recipients.ResolveAll() (called by every caller of this
method, previously) does NOT reliably resolve these -
confirmed live via COM: Resolved stayed False and Address
stayed empty for every address tested (including ones
already known-real), with no exception thrown, and a real
Send() using an unresolved recipient this way never
actually arrived at a real external mailbox. Resolving
each Recipient individually right after adding it does
work reliably (confirmed live: Resolved=True, Address
populated, and delivery confirmed to a real external
account) - do it here so every caller gets a genuinely
resolved recipient without needing its own resolve step.

## `ComposeEmail` (send mode) - `to` required

to is required, unlike draft_email's - draft_email opens a
compose window where a human can add a missing recipient
before anything goes out; send_email has no such window, so
Send()-ing with no recipient set would either throw or, worse,
block on a native Outlook resolution prompt this add-in isn't
expecting a response to.

## `CreateEvent` - `start`/`end` required

Unlike draft_event (where an omitted start/end just leaves
Outlook's own new-appointment default - "now", 30 min - sitting
in a review window the user sees before sending/saving),
create_event has no review step at all: a caller that forgets
either would otherwise create a real, immediately-live calendar
entry at an unintended time with nothing surfacing that it was
defaulted. Required here specifically.
