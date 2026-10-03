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
