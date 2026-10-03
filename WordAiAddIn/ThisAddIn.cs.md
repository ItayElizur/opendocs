# ThisAddIn.cs

## `_paneCreationInProgress`

Guards against reentrancy into EnsurePaneFor for the same hwnd -
confirmed repro (Excel/PowerPoint, single document open):
CustomTaskPanes.Add can pump the Windows message queue internally,
which lets a nested WindowActivate for the window already being set
up reenter this method before the outer call has returned and
written _panes[hwnd]. Without this guard that constructs a SECOND
TaskPaneHost/WebViewBridgeHost for the one window, and both race to
create a CoreWebView2Environment against the identical user-data
folder - which WebView2 rejects with "the group or resource is not
in the correct state" (HRESULT 0x8007139F).

## `ThisAddIn_Startup`

The startup window, as today - every subsequently-opened window
gets its own pane via Application_WindowActivate below. Guarded
(unlike every other EnsurePaneFor call site, all of which are
already wrapped) because Word can start on its own "Start
Screen" template chooser rather than a real document - a state
ActiveWindow may not represent as a normal, fully-formed
Word.Window (confirmed repro of this exact shape in Excel/
PowerPoint: this call, unguarded, left the add-in showing a
blank/gray pane). If that happens here, no pane is created for
the Start Screen at all - the first real document (Ctrl+N,
File > Open, etc.) still gets a working pane via
Application_WindowActivate regardless.

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

## `Application_WindowSelectionChange` catch

Selection-change notifications are best-effort; never let one
crash out of a COM event sink and kill the add-in connection.
Diagnostic addition (2026-08-24): this used to be a truly
silent catch-all - if a shape-selection event threw HERE
(e.g. accessing selection.Document.ActiveWindow.Hwnd during
some transient shape-selection state), it would vanish with
zero trace, which could fully explain "only one
OnSelectionChanged logged despite several clicks" - the
other clicks may never have reached OnSelectionChanged at
all. Now logged instead of silently dropped.
