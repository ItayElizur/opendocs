# ThisAddIn.cs

## `_paneCreationInProgress`

Guards against reentrancy into EnsurePaneFor for the same hwnd -
confirmed repro (single document open): CustomTaskPanes.Add can
pump the Windows message queue internally, which lets a nested
WindowActivate for the window already being set up reenter this
method before the outer call has returned and written
_panes[hwnd]. Without this guard that constructs a SECOND
TaskPaneHost/WebViewBridgeHost for the one window, and both race to
create a CoreWebView2Environment against the identical user-data
folder - which WebView2 rejects with "the group or resource is not
in the correct state" (HRESULT 0x8007139F).

## `ThisAddIn_Startup` - startup window guard

The startup window, as today - every subsequently-opened window
gets its own pane via Application_WindowActivate below. Guarded
(unlike every other EnsurePaneFor call site, all of which are
already wrapped) because Excel can start on its own "Start
Screen" template chooser rather than a real workbook - a state
ActiveWindow may not represent as a normal, fully-formed
Excel.Window (confirmed repro: this call, unguarded, left the
add-in showing a blank/gray pane). If that happens here, no
pane is created for the Start Screen at all - the first real
workbook (Ctrl+N, File > Open, etc.) still gets a working pane
via Application_WindowActivate regardless.

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
