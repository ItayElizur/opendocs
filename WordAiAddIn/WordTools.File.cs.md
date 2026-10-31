# WordTools.File.cs

## `save`

`save` is advertised from Track Changes up. It is in neither `readOnlyTools` nor
`commentOnlyExtraTools` in `entry.ts`, and the C# gate treats every tool outside
`AlwaysAllowedTools` as mutating, so Read Only and Comment Only block it server-side
too. No tier-list edits are needed for it.

`saveAs` defaults to false (save in place, overwriting the file). It must be true
whenever the user wants a copy, a new name, another folder or another format - the
schema text says so explicitly. A never-saved file (empty `Path`) goes to the Save As
UI even when `saveAs` is false, because there is no path to save to.

Word's Application.Dialogs[wdDialogFileSaveAs].Show() returns -1 when the user confirmed (Word has already
saved by then) and 0 on cancel. `Execute` in WordTools.cs excludes `save` from the forced
TrackRevisions toggle (it is not a content edit) and from the UndoRecord wrapper (there is nothing
to undo, and a custom record around a modal dialog is pointless).

The Save As UI is modal and runs synchronously on Word's UI thread, inside the
WebView2 `WebMessageReceived` handler, so the chat pane is frozen until the user saves
or cancels. There is no tool-call timeout (`callDotNetTool` is a bare promise) and the
AbortSignal is never forwarded to C#, so Stop cannot dismiss the dialog. This is
accepted. A cancel returns a non-error "cancelled" result with `Mutated = false`, so
the model does not retry in a loop.

NOT VERIFIED AGAINST LIVE WORD: compiles clean in Debug and Release, and the dialog
members were confirmed to exist in the referenced PIA by reflection, but neither the
in-place save nor the dialog (including cancel detection) was exercised against a
running instance.
