## EditingMode

```
// Previously declared identically as three separate app-namespaced enums
// (WordAiAddIn.EditingMode, ExcelAiAddIn.EditingMode, PowerPointAiAddIn.EditingMode)
// with the same four members. PaneHostBase's SetEditingMode hook needs one
// shared type to dispatch through, so this replaces all three - each app's
// *Tools.cs now references this instead of declaring its own.
```

## PaneHostBase

```
// Shared across Word/Excel/PowerPoint's TaskPaneHost.cs - the three copies
// differed only in the COM type used to resolve the owning document (the
// abstract hooks below) and app-data folder name (now a constructor
// parameter). Everything else - the status label, the WebView2 bridge,
// and the OnOtherMessage branches that don't need app-specific
// resolution - lives once, here.
```

## PaneHostBase._selectionTimer

```
// FT-2 Task 1: shared debounced selection dispatch. A WinForms Timer
// (not System.Timers.Timer) ticks on the UI thread, where this
// UserControl and its WebView2 already live - no cross-thread
// marshaling needed. 200ms was chosen empirically: holding an arrow
// key down in Excel settles the pill once, after the burst ends,
// rather than flickering on every keystroke.
```

## PaneHostBase constructor - lazy COM resolution

```
// This base constructor never touches the owning document/
// workbook/presentation - it only takes the app-data folder name.
// Each subclass's own constructor stores its COM document
// reference WITHOUT dereferencing it (no .Path/.FullName access
// there); see GetChatId()'s doc comment on each subclass for the
// confirmed repro of why an eager read silently kills the whole
// VSTO connection (no exception, no error - just Connect=False
// forever) when done at construction time instead of lazily.
```

## PaneHostBase.PostSelection

```
// FT-2 Task 1: coalesces bursts of selection-change events and drops
// exact repeats. `signature` is a cheap string identifying the
// selection (e.g. "Sheet1!B2:D40", "slides:2,3") - when it matches the
// last one actually posted, the event is dropped outright rather than
// even restarting the timer, so an event storm that never changes the
// selection (e.g. re-entrant COM notifications) costs nothing.
```

## PaneHostBase.GetChatId

```
// The per-document chat-history/mode key. Lazily computed and cached
// by each subclass on first actual use (never in the constructor -
// see the constructor comment above). A subclass's override re-checks
// a still-provisional ("unsaved-...") id on every call and migrates
// ChatStore/DocSettingsStore onto the real id the moment the document
// is saved (FT-1 Task 7b) - callers never need to know this happens.
```

## PaneHostBase.FlushChatIdMigration

```
// FT-1 Task 7b Step 2: called from each app's document-close handler
// (ThisAddIn.cs, alongside pane disposal) to force one last GetChatId()
// check before the pane goes away - covers "save, then immediately
// close" without needing a save-then-close-specific hook. ThisAddIn
// cannot call the protected GetChatId() directly (different class,
// not a subclass), hence this public wrapper.
```

## PaneHostBase.GetOfficeUiLanguageId

```
// Office's UI display language (Application.LanguageSettings.
// LanguageID[MsoAppLanguageID.msoLanguageIDUI]) - needs each
// subclass's own Globals.ThisAddIn.Application, same reason
// GetChatId()/SetEditingMode are abstract here rather than
// implemented once (this shared assembly has no access to any
// app's own VSTO-generated Globals class). See OfficeLanguage.cs
// for how the returned LCID maps to a supported UI language.
```
