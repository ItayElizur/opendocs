# ExcelTools.File.cs

## `save`

`save` is advertised from Track Changes up. It is in neither `readOnlyTools` nor
`commentOnlyExtraTools` in `entry.ts`, and the C# gate treats every tool outside
`AlwaysAllowedTools` as mutating, so Read Only and Comment Only block it server-side
too. No tier-list edits are needed for it.

`saveAs` defaults to false (save in place, overwriting the file). It must be true
whenever the user wants a copy, a new name, another folder or another format - the
schema text says so explicitly. A never-saved file (empty `Path`) goes to the Save As
UI even when `saveAs` is false, because there is no path to save to.

The Save As step uses Application.GetSaveAsFilename (picks a name only; returns the path as a string, or the
boolean false on cancel) followed by Workbook.SaveAs with a format taken from the extension
(xlsx, xlsm, xlsb, xls, csv; anything else falls back to xlsx). It deliberately does NOT use
Application.Dialogs[xlDialogSaveAs].Show(): with that legacy built-in dialog the mouse pointer vanished
when the command was sent while the pointer was over the chat pane (hidden over the whole Excel window, and
back only after the pointer re-entered the pane) - observed live, Excel only, and not fixed by running the
dialog after WebView2's message callback returned. The GetSaveAsFilename change is NOT VERIFIED beyond a
clean build until the pointer symptom is retested. Excel's Track Changes tier behaves like Full
Autonomy (see ExcelTools.cs.md), so `save` is available from Track Changes up exactly like
`propose_operations`.

The Save As UI is modal and runs synchronously on Excel's UI thread, inside the
WebView2 `WebMessageReceived` handler, so the chat pane is frozen until the user saves
or cancels. There is no tool-call timeout (`callDotNetTool` is a bare promise) and the
AbortSignal is never forwarded to C#, so Stop cannot dismiss the dialog. This is
accepted. A cancel returns a non-error "cancelled" result with `Mutated = false`, so
the model does not retry in a loop.

NOT VERIFIED AGAINST LIVE EXCEL: compiles clean in Debug and Release, and the dialog
members were confirmed to exist in the referenced PIA by reflection, but neither the
in-place save nor the dialog (including cancel detection) was exercised against a
running instance.

## Hidden mouse pointer

Symptom (Excel only): send `save` with `saveAs` while the mouse is over the chat pane and the pointer disappears
over every Excel window when Save As opens; it comes back only after the pointer re-enters the pane (Alt+Tab also
resets it). Word and PowerPoint do not do this. Not caused by the dialog API or by where the dialog is opened from:
swapping Dialogs[xlDialogSaveAs] for GetSaveAsFilename + SaveAs, and running the dialog after WebView2's message
callback returned, each changed nothing.

Measured with GetCursorInfo/ShowCursor just before the dialog (three runs, pointer over the pane twice):
over the pane the pointer was already invisible on Excel's UI thread before the dialog existed - `showing=False`,
`hCursor=0` (a null cursor), and once also a ShowCursor display counter of -1. With the pointer over the sheet
the state was normal. `ResetHiddenPointer` winds a negative counter back to 0 and replaces a null cursor with the
arrow just before the dialog; after that the pointer stayed visible in all three runs. A line is written to
%TEMP%\OpenDocsDebug.log only when it actually repaired something.

Why the cursor is null is not established. Likely the chat pane's browser clearing the cursor (for example
hide-while-typing) on the shared UI thread; not confirmed. Why Word/PowerPoint are unaffected is also unknown.
NOT VERIFIED beyond those three manual runs.

## Shared flow and strict `saveAs`

The flow (save in place, or Save As when `saveAs` is true or the file has never been saved) and the result messages
live in `OfficeAi.Shared/SaveTool.cs`; this file only supplies Excel's two COM steps (`Workbook.Save` and the Save As picker). `saveAs` is parsed by `ToolArgs.OptionalFlag`:
missing = false, `true`/`false` (or the numbers 1/0) accepted, anything else - a string like "true", another number,
null - throws "saveAs must be true or false", so a mistyped flag is never guessed into an in-place overwrite.
A read-only file is not special-cased: `Save()` behaves as the app itself does, and the user can ask for Save As.
