# WordTools.History.cs

## `UndoLastAction` / `RedoLastAction`

undo_last_action / redo_last_action (added for the undo/redo tooling
pass across Word/Excel/Outlook - PowerPoint has no such tool, see
PowerPointTools.History.cs for why it does, and OutlookTools.Undo.cs
for Outlook's unrelated small inverse-op mechanism).

IMPORTANT correction vs. the initial plan for this feature: Word's
Application object has NO Undo()/Redo() method at all - confirmed
via .NET reflection against the exact 15.0.0.0 PIA this project
references (WordAiAddIn.csproj), not assumed. Application only
exposes an UndoRecord property (for grouping a batch of automation
edits into one user-visible undo entry - unrelated to invoking
undo/redo). The real, callable methods live one level down, on
Document: `Document.Undo(ref object Times)` / `Document.Redo(ref
object Times)`, both returning bool. The Times ref parameter is
COM-optional (reflection confirms ParameterInfo.IsOptional == true
on it), so C# lets it be omitted entirely - Document.Undo()/Redo()
with no arguments is valid and is exactly what real-world Word
automation code calls. The bool return is used below for an honest
"Nothing to undo/redo" instead of always claiming success.
