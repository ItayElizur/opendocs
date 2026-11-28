# OutlookTools.Draft.cs

The item window's own tools: `get_current_item`, `edit_current_body`,
`set_current_fields`, `read_main_chat`, and the attendee default for
`find_meeting_slots`. `ItemPaneHost` calls them with the one item its pane
holds, after its own gate (see `ItemPaneHost.cs.md`). None takes an item id,
so the model in one window can never reach another window's draft.

## Why the body is edited through Word, not `HTMLBody`

A compose or meeting window's body is a Word document (`Inspector.EditorType`
is `olEditorWord`, and `Inspector.WordEditor` returned a Word `Document`,
measured on a live Outlook on 2026-11-28). Rewriting `HTMLBody` would lose the
signature's formatting, the quoted thread's layout, the cursor and the user's
Ctrl+Z. One `Range.Text` write per call keeps all of them, and is one step on
the window's own undo list.

## Finding the passage

The whole body is read once per call as `Range.Text`, with field codes and
hidden text included, so that every character has its document position.
`TextAnchors.Clean(raw, keepBreaks: true)` drops what the reader does not see
(field codes, reference marks) and keeps paragraph marks as `\n`, so a passage
can span paragraphs. `TextAnchors.DocumentPositions` maps each kept character
back to a position; the one known exception to one character per position is
an end-of-cell mark (`\r\a`, one position), and HTML replies carry tables.

Two checks keep an edit from landing anywhere else:

- The mapped end of the text must equal `Content.End`. If it does not, the
  call refuses, rather than edit through a map it cannot trust.
- After the match is mapped, the target range is read again and cleaned, and
  must equal the passage. If it does not, the call refuses.

The body is read fresh at every call, so text the user typed elsewhere since
the model last read it does not move the edit. Text the user changed inside the
passage is simply not found, and the model is told to read the item again.

`get_current_item` returns the same cleaned text, so a passage the model copies
from it matches.

## Recipients

`RecipientEdits.Plan` (OfficeAi.Shared, unit-tested) works out the changes from
the recipient list as read: kind changes, then removals (highest index first),
then additions, which keeps every index valid while applying. Adding someone
already on the item under another kind moves them; the organizer is never
touched. Each new name goes through `Recipient.Resolve()`, which leaves an
unknown or ambiguous name unresolved; such a recipient is deleted again and
reported, so the model asks `search_contacts` or the user instead of guessing.
`search_contacts` itself is not called here: it is EWS-only and slow, and the
model can call it when it needs candidates.

## Appointments

Times, subject and location go through `draft_event`'s `ApplyEventFields`, so a
series moves through its `RecurrencePattern` (Outlook refuses `Start`/`End` on a
series). A start without an end keeps the duration. Recurrence reuses
`ReadRecurrence` and `ApplyRecurrence`, and attendee changes end in
`SyncMeetingStatus`, both shared with `draft_event`. `DescribeRecurrence` calls
`GetRecurrencePattern` only on a series, because on a single appointment that
call creates a pattern. A meeting the user only attends, or a canceled one, is
refused before anything changes.

Nothing is saved or sent. A sent meeting changed here keeps Outlook's own Send
Update step.
