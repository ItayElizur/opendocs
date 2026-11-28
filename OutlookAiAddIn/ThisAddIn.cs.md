# ThisAddIn.cs

## `GetOfficeUiLanguageId`

The one real COM call for Office's UI display language in this
app - Ribbon.cs and TaskPaneHost.cs each need their own copy of
this (their base classes' GetOfficeUiLanguageId hooks are
abstract, since neither shared assembly can see this app's own
Globals class), but delegate here rather than re-issuing the COM
call themselves, so there is exactly one place per app that can
fail and exactly one place that guards against it. A theme-
detection bug must never break pane creation (OfficeTheme.cs's own
stated posture) - same reasoning applies here: if
LanguageSettings throws (an unusual COM/host state), degrade to
the code that already means "not Hebrew" rather than letting the
ribbon render a blank label or the "load-language" bridge message
die silently with no reply ever sent (that one-shot message has no
retry - see PaneHostBase's "load-language" case).

## Item window panes (`ItemPaneEntry`, `ToggleItemPane`, `RemoveItemPane`)

A mail compose window and the appointment or meeting editor (both
Inspectors) get their own chat pane, created on the first press of the
ribbon button in that window. It is never created when the window opens:
Outlook opens an Inspector for every double-clicked email, and each
pane starts a WebView2 page. The ribbon passes the window it was pressed
in (`IRibbonControl.Context`), which is how `TogglePane` tells an item
window from the main one.

Item panes are keyed by the Inspector's COM identity, like the main
window's panes. Outlook shows an item in at most one window, so keying
by window also means an item's chat never has two live panes, and the
two never write to one chat file at once. The main window does not have
this guarantee: two main windows share the mailbox chat.

`RemoveItemPane` runs from `Inspector.Close`, while the item is still
reachable, and calls `FlushChatIdMigration` first. A new draft saved by
the close itself (the user chose Save) gets its EntryID only then, and
that last chat id check moves its provisional chat onto the saved id.
