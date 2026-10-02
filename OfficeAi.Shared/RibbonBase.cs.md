## RibbonBase.GetGroupLabel / GetButtonLabel

```
// Office's UI display language decides whether the group/button show
// the transliterated Hebrew brand name or the Latin one. getLabel
// callbacks (rather than a static label attribute) are a standard,
// documented Ribbon XML mechanism for exactly this.
//
// GetOfficeUiLanguageId() is abstract - and expected to delegate to
// Globals.ThisAddIn's own guarded copy, not re-issue the COM call
// here - for the same reason PaneHostBase's identically-named/
// -shaped hook is: this shared assembly has no access to any app's
// own VSTO-generated Globals class. Consolidated onto ThisAddIn
// (2026-09-27) after this LCID lookup turned out to be duplicated
// 12 times across Ribbon.cs/TaskPaneHost.cs/ThisAddIn.cs (3 per app x
// 4 apps) with no shared exception guard - see ThisAddIn.
// GetOfficeUiLanguageId's own comment for the failure mode that fixed.
```

## RibbonBase.GetLogoImage

```
// Reads the same web/logo.png the WebView2-hosted header uses (copied
// there at build time from shared/chat-ui/logo.png) - one physical
// image file drives both surfaces, so editing it updates both.
// AppDomain.CurrentDomain.BaseDirectory is the host process's own
// directory (e.g. WordAiAddIn/bin/Debug), not this shared assembly's -
// that still resolves correctly here since it's a property of the
// running AppDomain, not of whichever assembly happens to read it.
```
