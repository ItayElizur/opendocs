# OutlookTools.Calendar.cs

## `ListEventsAsync` - shared calendar folder is only used for StoreID

2026-09-30: this used to also enumerate the folder's Items (Sort ->
IncludeRecurrences -> Restrict -> foreach), which froze Outlook - that
COM enumeration against a shared, normally-not-cached-offline mailbox
could mean a live round trip to Exchange per property per event. Now
this call is ONLY used to read .Store.StoreID (get_event's own
store_id parameter needs it - see its comment); the actual event data
comes from EWS below, off the UI thread. See
docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md's
"Open question" section: if Outlook still freezes after this change,
this GetSharedDefaultFolder call itself - not the enumeration it used
to do - is the next thing to investigate.

## `ListEventsAsync` - StoreID read split from the GetSharedDefaultFolder try/catch

Split from the GetSharedDefaultFolder try/catch above on purpose
(restores the shape from commit ec53129, predating this branch): a
StoreID read failure here means the folder itself opened fine, so it
should degrade to a missing store_id (costing only get_event
usability for these specific events) rather than being misreported
as "you may not have been granted access" - see ec53129 for the
original bug this shape fixes.

## `QueryCalendarItems` - private items pass through unrestricted

Whatever Outlook itself resolves for Subject/Location is passed
through as-is, including for private items on a shared calendar -
this add-in doesn't add its own visibility restriction on top of the
caller's real Exchange permissions. If the caller's access level
would let them see this in Outlook's own UI (or if the object model
exposes more than the UI would, which is a known Outlook behavior for
the "Private" flag), that's between the caller and the calendar
owner's actual sharing settings, not something this code
second-guesses.

## Shared `store_id` parameter pattern

Used by GetEvent, RespondMeeting, DraftRespondMeeting, EditEvent,
DraftEditEvent, DraftCancelEvent, and CancelEvent:

Only needed for an event_id from someone else's shared calendar
(returned by list_events' mailbox parameter) - ItemById/GetItemFromID
can't find an item outside the caller's own default store without it.
Omit for your own events, same as before this parameter existed.
Whether the caller actually has permission to act on the resulting
item is entirely up to Outlook/Exchange - this add-in doesn't add its
own authorization check on top of that.

## `DefaultWorkRange`

Generalization of mcp-outlook's scheduling.default_range (originally
"today through Thursday of this work week, rolling to next Sun-Thu
if today is Fri/Sat") for an arbitrary work-days set: walk forward
from today to the next work day (today itself if it already is
one), then extend through the following contiguous run of work
days (capped at a week) - correct for any shape, including a
work week that wraps past a calendar-week boundary.

## `DraftRespondMeeting`

Draft-tier counterpart to RespondMeeting - redesigned after a code
review confirmed (via Respond()'s documented behavior, and this
project's own prior history with draft_cancel_event hitting the
identical shape of bug) that calling appt.Respond() commits a real
calendar change at call time - a new EntryID on accept/tentative,
a move to Deleted Items on decline - independent of whether the
resulting response is ever sent or the window ever closed with an
action taken. That directly broke this codebase's "draft tools
persist nothing until the user acts" guarantee, the same way an
earlier version of draft_cancel_event did with an unsaved
MeetingStatus change. Fixed the same way that was: never mutate
the item at all here. Opens the original item completely
unchanged; the user picks Accept/Tentative/Decline themselves
from Outlook's own native ribbon buttons. Since this tool no
longer calls Respond() with a specific response type, one unified
tool replaces what used to be three separate ones
(draft_accept_meeting/draft_decline_meeting/draft_tentative_meeting).
message can no longer be pre-filled into a response body (that
would require calling Respond() to get the MeetingItem, the exact
call this redesign avoids) - it's returned in the output text
instead, for the user to paste in themselves if they use
Outlook's own "Edit response before sending" option.

## `IsReceivedMeeting`

Shared by draft_edit_event/edit_event: an olMeetingReceived
(or olMeetingReceivedAndCanceled) appointment is one the user only
attends, not organizes - Outlook gives attendees no authority to
unilaterally move someone else's meeting. Confirmed via .NET
reflection against the referenced Microsoft.Office.Interop.Outlook
15.0.0.0 PIA that neither AppointmentItem/_AppointmentItem nor
MeetingItem/_MeetingItem expose any "Propose New Time" member -
_AppointmentItem.Respond only takes an OlMeetingResponse
(Accept/Decline/Tentative), no counter-proposal overload. Real
Outlook's "Propose New Time" is a ribbon/UI feature (MAPI
counter-proposal properties), not one the classic COM object model
exposes cleanly, so there is no authoritative or even semi-authoritative
way to honor a reschedule request on a received meeting here - both
tools refuse up front instead of silently attempting a Save()/Send()
Outlook wouldn't actually honor as a real reschedule.

OlMeetingStatus has 5 values: olNonMeeting (0), olMeeting (1),
olMeetingReceived (3), olMeetingCanceled (5),
olMeetingReceivedAndCanceled (7) - checking only the exact
olMeetingReceived value here would let an attendee's copy of a
meeting the organizer has since canceled (olMeetingReceivedAndCanceled)
through silently, so both "received" statuses count.

## `ResolveOccurrenceTarget`

Shared by all four occurrence-aware tools (reschedule/cancel, draft
and immediate). occurrence_date lets a caller target one instance
of a recurring series instead of the whole master.
RecurrencePattern.GetOccurrence(DateTime) confirmed present via
.NET reflection against the referenced PIA (not assumed) - this
had been undocumented capability until this addition, even though
every occurrence list_events returns shares the master's EntryID.

## `CheckOccurrenceReorderCollision`

Checks whether moving an occurrence from `original` to `target` would
cross or land on the same day as another occurrence of the same series -
Outlook rejects both (confirmed live 2026-09-28 by reproducing the
specific error manually in Outlook's UI: "Cannot reschedule an occurrence
of the recurring appointment ... if it skips over a later occurrence of
the same appointment" - both cases surface only as a generic
COMException "Cannot save this item." from .Save(), with no way to
distinguish the cause from the exception itself). Checking this
ourselves first gives an exact, deterministic answer instead of relying
on Outlook's collapsed generic exception message.

Queries the Calendar folder the same way ListEvents does (Sort("[Start]")
-> IncludeRecurrences = true -> Restrict, in that order - load-bearing,
see ListEvents' own comment) over the range between `original` and
`target` (inclusive of both endpoint days), then keeps only occurrences
of THIS series (matching master's EntryID) other than the one being
moved (excluded by day - it's still sitting at `original` since nothing
has been saved yet). If any remain, the closest one to `original` is the
binding obstruction; returns null if the move is clear.

## `CheckOccurrenceReorderCollision` - master's own Parent folder

master's own Parent folder, not the caller's default calendar -
for an event resolved via store_id from someone else's shared
calendar, those are different folders entirely. Using the
caller's own default calendar here meant this check silently
never found a collision for a shared-calendar event (found via
PR review, 2026-09-29).

## `ReplaceAttendees`

Shared by edit_event/draft_edit_event: replaces the required and/or
optional attendee list wholesale (not a diff/merge - the caller
supplies the full new list each time, reading the current one
first via get_event if they need to preserve someone). The two
categories are independently optional: a null argument leaves that
attendee category completely untouched; a non-null argument
(including "") fully replaces it - clearing the existing entries
in that category and re-adding via the same AddAttendees helper
create_event/draft_event already use. The organizer recipient is
never touched. Recipients indices are 1-based (confirmed via .NET
reflection against the referenced PIA, matching every other
Outlook collection in this codebase); iterating downward from
Count avoids skipping an element after Remove shifts the rest down.

## `CountAttendeeRecipients`

Counts real attendees only, excluding the organizer - used by
EditEvent/DraftEditEvent to decide whether clearing attendees
should revert MeetingStatus back to olNonMeeting. Confirmed live
via COM that the organizer does NOT appear in Recipients even
after a real .Send() (Recipients.Count matched exactly the number
of real invited attendees, organizer never counted) - this
exclusion is defensive insurance for any account/Exchange
configuration where that might differ, not a fix for an observed
bug.

## `EditEvent` - deferred revert to olNonMeeting

If clearing brought the real attendee count to zero, the
event should revert to olNonMeeting - otherwise it stays
permanently "a meeting" (and thus a permanent undo barrier)
even after every attendee is removed. The actual flip is
deferred to just after .Send() below (not done here) - a
code review of this method noted that flipping MeetingStatus
to olNonMeeting BEFORE .Send() risks Outlook not treating
the send as a cancellation notice to the just-removed
attendees. Deferring is safe: reaching recipientsAfter == 0
from a meeting always means wasMeetingBefore was true, which
always forces isMeetingNow/mustBarrier true below, so the
.Send() branch always runs when this flag is set - there is
no code path where revertToNonMeeting is set but .Send() is
skipped.

## `EditEvent` - Send() then revert then Save()

Send while still flagged as a meeting so Outlook treats
this as a real meeting update/cancellation to whoever
was just removed, THEN apply the deferred revert to
olNonMeeting (if attendees were cleared to zero), THEN
Save. .Send() alone can leave the item's own Saved flag
stuck False even after a successful send - confirmed
live via COM (create a recurring meeting, clear its
attendees, revert MeetingStatus, .Send(): Saved reads
False even on a fresh re-fetch by EntryID, causing
Outlook to prompt "save changes?" if the user later just
opens and closes the item with nothing to change). An
explicit .Save() right after .Send() clears it -
confirmed the same sequence with the extra Save() reads
Saved=True on a fresh re-fetch. The Save() is wrapped
since it runs after the irreversible Send() has already
succeeded - a failure here must not be reported as a
failed edit_event call (the update already went out).

## `EditEvent` - RecordSnapshot's props is always non-empty

RecordSnapshot reads appt.EntryID via ItemEntryIdOf(appt) - for an
occurrence, that's the real, resolvable EntryID GetOccurrence's
returned item gets once saved, so undo/redo works via the exact
same SnapshotEntry mechanism as the retired reschedule tool - no new
undo-entry type needed. props is always non-empty here: reaching
this branch requires mustBarrier == false, which means attendees
were never touched (that forces isMeetingNow, hence a barrier)
and, if a time change was requested, it's occurrence/non-recurring
(whole-series-recurring also forces a barrier) - so at least one
of Start/End/Subject/Body/Location is always in props.

## `DraftCancelEvent`

Originally set MeetingStatus = olMeetingCanceled unsaved before
Display(false), mirroring the retired draft reschedule tool's unsaved
Start/End - removed 2026-09-28 after live testing showed Outlook
persists that change when the Inspector closes even without the
user clicking "Send Cancellation" (unlike Start/End, an unsaved
MeetingStatus change apparently isn't purely cosmetic here). That
silently canceled the meeting locally with attendees never
notified - worse than doing nothing, since it also means this tool
can no longer be un-done or reattempted (cancel_event/edit_event
both refuse on an already-canceled item). Never mutates the item at
all now.

## `CancelEvent`

Full-autonomy-only counterpart to draft_cancel_event above: for a
meeting the user organizes, sends the cancellation notice
immediately (no review step, mirroring edit_event's/
create_event's attendee-present branches) then removes it from the
user's own calendar; for a plain appointment, just removes it -
nobody to notify. Either way the item is moved to Deleted Items
(recoverable there), same as delete_email's non-permanent path, not
permanently deleted.

An already-canceled event (IsCanceledMeeting - either the
organizer's own olMeetingCanceled copy, or an attendee's stale
olMeetingReceivedAndCanceled one) is treated as cleanup, not
refused: there's nothing new to notify anyone of, so it just moves
straight to Deleted Items like a plain appointment. This is the
only way to dismiss a canceled event at all - draft_cancel_event
still refuses on one (nothing to preview/review for a cleanup).
