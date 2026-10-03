# PowerPointTools.History.cs

## `UndoLastAction` / `RedoLastAction` / `ExecuteRibbonUndoRedo`

undo_last_action / redo_last_action.

PowerPoint's Application object has NO Undo()/Redo() method
anywhere in the interop surface - confirmed absent (see
docs/ai-tool-surface.md's PowerPoint section). That earlier finding
still stands, but it missed a distinct mechanism this file uses:
Application.CommandBars.ExecuteMso(string)/GetEnabledMso(string) -
the generic Office-2007+ ribbon-command dispatch API
(Microsoft.Office.Core.CommandBars, from the "Office" PIA
reference already in this project). Confirmed real and callable
via .NET reflection against the exact 15.0.0.0 PIA this project
references: PowerPointAiAddIn's Application.CommandBars property
returns Microsoft.Office.Core.CommandBars, and that type's
_CommandBars interface declares both `Void ExecuteMso(String)` and
`Boolean GetEnabledMso(String)`. This dispatches by ribbon-command
ID string - the same mechanism the real Undo/Redo buttons and
Ctrl+Z/Ctrl+Y use internally - so it works despite there being no
direct Undo/Redo method to call. GetEnabledMso is checked FIRST so
the result honestly reports "nothing to undo/redo" rather than
always claiming success (ExecuteMso itself returns nothing to
signal whether anything happened).

See PowerPointTools.Master.cs's RemoveMasterElement (2026-09-22,
live-tested incident) for why this repo takes "can we actually
undo this" seriously for PowerPoint: a slide-master placeholder
deletion had no code path at all to undo at the time (Application.
Undo() doesn't exist, and nothing else was tried), leaving two
theme placeholders gone with no way back - RemoveMasterElement now
refuses to delete a placeholder outright rather than risk a repeat.
This tool is what generally closes that "no way back" gap for
whatever the model does in a session (via the ExecuteMso mechanism
above), but does not change RemoveMasterElement's own placeholder
refusal, which stands for its own separate, still-valid reason
(a Slide Master placeholder deletion may not even affect slides
that carry their own independent layout copy - see that method's
comment).
