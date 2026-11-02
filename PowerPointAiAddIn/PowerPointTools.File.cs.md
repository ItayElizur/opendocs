# PowerPointTools.File.cs

## `save`

`save` is advertised from Track Changes up. It is in neither `readOnlyTools` nor
`commentOnlyExtraTools` in `entry.ts`, and the C# gate treats every tool outside
`AlwaysAllowedTools` as mutating, so Read Only and Comment Only block it server-side
too. No tier-list edits are needed for it.

`saveAs` defaults to false (save in place, overwriting the file). It must be true
whenever the user wants a copy, a new name, another folder or another format - the
schema text says so explicitly. A never-saved file (empty `Path`) goes to the Save As
UI even when `saveAs` is false, because there is no path to save to.

PowerPoint has no Dialogs collection, so the Save As path uses Application.FileDialog[msoFileDialogSaveAs]
(present on _Application in the referenced PIA). FileDialog.Show() only picks the target (-1 confirmed, 0 cancelled);
the save itself happens in FileDialog.Execute(). `Execute` in PowerPointTools.cs skips
StartNewUndoEntry() for `save` (nothing to undo).

The Save As UI is modal and runs synchronously on PowerPoint's UI thread, inside the
WebView2 `WebMessageReceived` handler, so the chat pane is frozen until the user saves
or cancels. There is no tool-call timeout (`callDotNetTool` is a bare promise) and the
AbortSignal is never forwarded to C#, so Stop cannot dismiss the dialog. This is
accepted. A cancel returns a non-error "cancelled" result with `Mutated = false`, so
the model does not retry in a loop.

NOT VERIFIED AGAINST LIVE POWERPOINT: compiles clean in Debug and Release, and the dialog
members were confirmed to exist in the referenced PIA by reflection, but neither the
in-place save nor the dialog (including cancel detection) was exercised against a
running instance.

## Shared flow and strict `saveAs`

The flow (save in place, or Save As when `saveAs` is true or the file has never been saved) and the result messages
live in `OfficeAi.Shared/SaveTool.cs`; this file only supplies PowerPoint's three COM steps (`Presentation.Save` and the Save As file dialog). `saveAs` is parsed by `ToolArgs.OptionalFlag`:
missing = false, `true`/`false` (or the numbers 1/0) accepted, anything else - a string like "true", another number,
null - throws "saveAs must be true or false", so a mistyped flag is never guessed into an in-place overwrite.
A read-only file is not special-cased: `Save()` behaves as the app itself does, and the user can ask for Save As.
