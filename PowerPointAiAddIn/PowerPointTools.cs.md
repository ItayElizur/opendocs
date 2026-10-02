# PowerPointTools.cs

## `ModeByDoc` - editing-mode gating

Editing-mode gating (mirrors the Word/Excel Task 11/16 pattern this plan establishes
elsewhere): the tool list offered to the model is filtered client-side per mode
(web-src/entry.ts, first line of defense - smaller prompts, fewer wasted turns), but
Execute() independently re-checks mode here as defense-in-depth, since nothing stops a
misbehaving or malicious model response from calling a tool that wasn't offered.

Per-document since PP-1 - see WordTools.cs's identical pattern for the rationale.

## `Execute` - StartNewUndoEntry (one tool call = one undo step)

PowerPoint's native undo manager can coalesce several
back-to-back COM-driven mutations into a single undo entry
when there's no UI tick between them (confirmed live,
2026-10-02: add_table immediately followed by
edit_table_structure's delete-row got merged - one
undo_last_action removed the whole table, not just the
row). StartNewUndoEntry() (confirmed via reflection against
the referenced PIA - Microsoft.Office.Interop.PowerPoint
has no Document-level undo, so this is the only available
boundary) forces a fresh entry before each mutating tool
call, so one tool call always maps to exactly one undo
step. Not called for always-allowed (read-only) tools -
nothing to barrier there.

Deliberately NOT excluded here (unlike Word's analogous
fix in WordTools.cs, which also excludes undo_last_action/
redo_last_action): live-tested, 2026-10-02 - calling
StartNewUndoEntry() immediately before PowerPoint's own
ExecuteMso("Undo"/"Redo") (PowerPointTools.History.cs) with
nothing undo-worthy pending did not disturb the redo stack.
This is a genuine behavioral difference from Word, not
copy-paste drift - Word's UndoRecord.StartCustomRecord
wrapping Document.Undo()/Redo() would be meaningless by
construction, whereas PowerPoint's barrier is just inert
when empty.

## `IsMutationAllowed`

Read Only and Comment Only modes block all mutating tools (PowerPoint has no
comment-equivalent tool in this pass - see plan backlog - so Comment Only currently
behaves identically to Read Only: no mutating tools available). Track Changes is scoped
to simple allow/block gating for now (same as Excel's Task 16 scoping note) rather than a
native PowerPoint revision-tracking UI. Full Autonomy allows everything.
