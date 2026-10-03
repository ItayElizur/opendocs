## OfficeLanguage

```
/// Resolves an Office LCID (Application.LanguageSettings.LanguageID[
/// MsoAppLanguageID.msoLanguageIDUI], confirmed via reflection against
/// the referenced Office core PIA - it's an indexed property, not a
/// method) to the panel's own supported UI language. Only "he"/"en" are
/// supported; any LCID that isn't Hebrew (1037 / 0x040D) resolves to
/// "en", mirroring OfficeTheme's "unknown degrades to the safe default"
/// posture rather than throwing.
```

## OfficeLanguage.ResolveBrandName

```
// The one place the brand name's per-language rendering is decided -
// previously duplicated as the same ternary in RibbonBase.BrandLabel()
// and each app's own ThisAddIn.PaneTitle() (5 copies total, none of
// them tested). "אופן דוקס" is a phonetic transliteration, not a
// translation - matches chat-ui.ts's own panelTitle Hebrew string
// (harmonized there per user request after PR #11's rebrand testing).
```
