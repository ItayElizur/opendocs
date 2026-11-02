# HiddenPointer.cs

## Why this exists

Symptom (found with Excel's Save As): send a tool that opens a dialog while the mouse is over the chat pane and the
pointer disappears over every Excel window once the dialog opens. It comes back only after the pointer re-enters the
pane (Alt+Tab also resets it). Word and PowerPoint did not show it with their own dialogs.

It is not caused by the dialog API or by where the dialog is opened from: running the dialog after WebView2's message
callback returned, and swapping `Dialogs[xlDialogSaveAs]` for `GetSaveAsFilename` + `SaveAs`, each changed nothing.

Measured with `GetCursorInfo`/`ShowCursor` just before the dialog (three runs, pointer over the pane twice): over the
pane the pointer was already invisible on Excel's UI thread before the dialog existed - `showing=False`, `hCursor=0`
(a null cursor), and once also a ShowCursor display counter of -1. With the pointer over the sheet the state was
normal. `Reset` winds a negative counter back to 0 and replaces a null cursor with the arrow; after that the pointer
stayed visible in all three runs (those runs used `GetSaveAsFilename`). A line goes to %TEMP%\OpenDocsDebug.log only
when something was actually repaired.

Why the cursor is null is not established. Likely the chat pane's browser clearing the cursor (for example
hide-while-typing) on the thread it shares with the host app; not confirmed. Why Word/PowerPoint are unaffected is
also unknown. Only Excel's Save As calls this today.
