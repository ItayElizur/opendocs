# bridge.ts — relocated comment history / rationale

## `RawSelectionPayload`

```
FT-2: the raw 'selection-changed' WebMessage, before any app-specific
interpretation. Every field below is app-specific and optional - `app`
distinguishes Excel/PowerPoint's payloads from Word's (which carries no
`app` field at all, only `hasSelection`/`preview`/`fullText`). This bridge
module stays app-agnostic (per its own file-header rule) - only
bootstrap.ts's per-app describeSelection/classification logic interprets
these fields into a SelectionContext.
```

## `RawSelectionPayload.layoutName`

```
PowerPoint (user-requested, 2026-09-22): the selected slide's current
layout name (custom-theme layouts, e.g. "Title Slide") - lets the model
address add_master_element/read_master_elements/etc.'s layoutName
directly from context instead of a separate read_slide call. Sent for
all three selKind variants (a shapes/shapeText selection is always
within exactly one slide, so this is unambiguous there too).
```

## `BridgeHandlers.onOfficeLanguageLoaded`

```
Office's own UI display language, read once via Application.
LanguageSettings.LanguageID(msoLanguageIDUI) when the pane boots and
sent once in response to requestOfficeLanguage() - never re-sent later
(by design, see OfficeAi.Shared/OfficeLanguage.cs), so this fires
exactly once per pane lifetime. Only "he"/"en" are supported UI
languages; any other Office UI language resolves to "en" server-side.
```
