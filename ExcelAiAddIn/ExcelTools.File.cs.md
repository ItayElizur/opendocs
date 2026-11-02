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

The Save As step is Application.Dialogs[xlDialogSaveAs].Show(), which returns true when the user confirmed (Excel
has already saved by then) and false on cancel; Excel itself handles the file name, every file format and the
overwrite prompt. Just before it opens, `HiddenPointer.Reset()` runs (see "Hidden mouse pointer" below). Excel's
Track Changes tier behaves like Full Autonomy (see ExcelTools.cs.md), so `save` is available from Track Changes up
exactly like `propose_operations`.

The Save As UI is modal and runs synchronously on Excel's UI thread, inside the
WebView2 `WebMessageReceived` handler, so the chat pane is frozen until the user saves
or cancels. There is no tool-call timeout (`callDotNetTool` is a bare promise) and the
AbortSignal is never forwarded to C#, so Stop cannot dismiss the dialog. This is
accepted. A cancel returns a non-error "cancelled" result with `Mutated = false`, so
the model does not retry in a loop.

NOT VERIFIED AGAINST LIVE EXCEL for the current code: it compiles clean in Debug and Release and the dialog member
was confirmed to exist in the referenced PIA by reflection. An earlier version that used GetSaveAsFilename + SaveAs
(not kept) was exercised by hand, including cancel; the xlDialogSaveAs + HiddenPointer.Reset combination was not
(see below).

## Hidden mouse pointer

Sending `save` with `saveAs` while the mouse is over the chat pane left the pointer invisible once Save As opened
(Excel only). The cause was the pane leaving the thread's cursor in a bad state before the dialog existed, not the
dialog API, so the fix is `HiddenPointer.Reset()` just before the dialog. The measurements and what is still unknown
are in `OfficeAi.Shared/HiddenPointer.cs.md`. It was verified by hand with a GetSaveAsFilename-based dialog; with
`Dialogs[xlDialogSaveAs]` it is expected to behave the same but NOT VERIFIED - retest by sending the command with the
mouse over the pane.

## Shared flow and strict `saveAs`

The flow (save in place, or Save As when `saveAs` is true or the file has never been saved) and the result messages
live in `OfficeAi.Shared/SaveTool.cs`; this file only supplies Excel's three COM steps (`Workbook.Save`, the Save As dialog, and reading the full name). `saveAs` is parsed by `ToolArgs.OptionalFlag`:
missing = false, `true`/`false` (or the numbers 1/0) accepted, anything else - a string like "true", another number,
null - throws "saveAs must be true or false", so a mistyped flag is never guessed into an in-place overwrite.
A read-only file is not special-cased: `Save()` behaves as the app itself does, and the user can ask for Save As.
