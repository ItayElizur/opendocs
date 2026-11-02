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
