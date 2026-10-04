# OutlookTools.Undo.cs

## Class-level comment - why not Outlook's native Undo

Why not Outlook's native Undo (CommandBars.ExecuteMso("Undo")): tried
and tested by hand on 2026-09-28. Outlook's Undo is a single slot tied
to the Explorer window that toggles undo/redo, it doesn't see
object-model changes (a move_email followed by the ribbon Undo did
nothing), and after one ExecuteMso("Undo") both our tool and the
ribbon button failed with "The operation cannot be performed because
the message has changed." There is no API to put an object-model
change onto that slot.
