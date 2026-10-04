# Word tools (`WordAiAddIn/WordTools*.cs`)

Current-state reference for every tool the AI can call in Word. See
[`docs/architecture.md`](../architecture.md) for the shared editing-mode/transport
architecture. "genoffice" below refers to a sibling from-scratch web-based Office clone
project this tool design was originally ported from — comparisons against it are kept
where they explain a real scope or capability difference; see
[`docs/comparison-with-genoffice.md`](../comparison-with-genoffice.md) for the
high-level summary.

## Top-level tools (19)

| Tool | Kind | Notes |
|---|---|---|
| `get_document_context` | Read | Paragraph/word count + a flat 300-char text preview. No block-indexed list (index\|type\|preview) like genoffice's version. |
| `find_text` | Read | Read-only substring or `.NET` regex search across paragraph text. Returns `[index] text` using the exact same 0-based paragraph index `read_blocks`/`replace_blocks`/`apply_commands`' `Target.blockIndexes` use. Uses Word's native `Range.Find` engine for the plain-substring path — the first version used positional `Paragraphs[i]` indexing, which isn't a real array access in Word's COM object model, so each read re-walked the document from the start (effectively O(n²)); a user hit this as Word visibly freezing on a large document. Fixed to use `Range.Find`/`Range.GoTo` and a shared forward-chaining paragraph resolver instead — see `CLAUDE.md`'s "Outlook specifically" section, which documents this incident as the reason Outlook's own read tools use native query APIs from the start. |
| `get_headings` | Read | Lists every heading-styled paragraph in order, Navigation-Pane style: `[index] H<level>: text`. Uses `Range.GoTo(wdGoToHeading, wdGoToNext)`, never a per-paragraph scan. |
| `read_blocks` | Read | Paragraph-indexed range read, capped at 1000 paragraphs. `format:"html"` emits a restricted HTML subset (headings/bold/italic/underline/list membership — capped at 100 paragraphs, since per-paragraph formatting reads are slower); plain text is `[i] text` lines. |
| `read_chart` | Read | Reads an existing chart's title, type, categories, and per-series names/values — call before an incremental `edit_chart` edit, since `edit_chart` replaces the whole dataset when given. |
| `read_table` | Read | Reads a table's cell contents, one row per line. `tableIndex` (0-based, document order) defaults to the first table. |
| `read_smartart` | Read | Reads SmartArt node text; omitting `smartArtIndex` reads every diagram in the document in one call. |
| `insert_content` | Write | Supply `text` (plain) or `html` (restricted subset: `<p> <h1>-<h3> <ul>/<ol>/<li> <b>/<strong> <i>/<em> <u> <br/>`, no attributes, no nested lists). `afterBlockIndex` anchors the insertion (0-based paragraph index; `-1` = start of document; omit = end). |
| `replace_blocks` | Write | Replaces paragraphs `[startIndex, endIndex]` with `text` or `html` (same restricted subset). Empty text deletes the range. `preserveFormatting` (default true) reapplies the first replaced paragraph's style to the result. |
| `add_image` | Write | Inserts an image from a **local file path only** (no URLs — air-gapped deployment). Inline by default (addressable afterward via `apply_commands`' `updateImageProperties`), or `floating:true`. |
| `add_table` | Write | Adds a native Word table, optionally pre-filled with cell text (row-major array of arrays). |
| `edit_table` | Write | `kind`: `set_cell`, `insert_row`/`delete_row`/`insert_col`/`delete_col`, `set_style` (styleName/headerRow/bandedRows/borders/borderColor), `set_shading` (cell/row/col/table scope, hex color). Structural edits shift later indices — re-read before a second structural edit in the same run. Row/col index 0 is just the first physical row/column, including a header row — there's no separate header concept in the index space. |
| `add_smartart` | Write | `layout`: list/process/cycle/hierarchy/pyramid/matrix/venn. Flat item list, one per top-level node. |
| `edit_smartart` | Write | `kind`: `set_text`, `add_node`, `delete_node`, `set_style` (colorName/quickStyleName, substring-matched against the install's real gallery), `set_layout`. `delete_node` shifts later node indices. |
| `edit_chart` | Write — gateway | Creates or edits a native Word chart. Supports categories + one or more named series (multi-series, chart-type selection: column/columnStacked/bar/barStacked/line/area/pie/doughnut). `chartIndex` addresses an existing chart (0-based, inline then floating shapes, document order); `create:true` always adds a new one. `afterBlockIndex` anchors a new chart inline instead of floating at document origin. |
| `apply_commands` | Write — gateway, see below | |
| `add_comment` | Write (every mode) | Anchors a real Word comment to the first match of given text. **Not in genoffice's docs surface at all.** Available in every editing mode, including Comment Only. |
| `undo_last_action` / `redo_last_action` | Write | Word's own Ctrl+Z/Ctrl+Y over the document's real undo history, via `Document.Undo()`/`Document.Redo()` (`WordTools.History.cs` — Word's `Application` has no Undo/Redo method); each returns a bool, so an empty stack reports "Nothing to undo/redo" with `Mutated: false`. Every other mutating tool call is wrapped in `Application.UndoRecord.StartCustomRecord`/`EndCustomRecord`, so **one tool call = one undo step** (a pre-filled `add_table` or a whole `apply_commands` batch backs out in a single undo). Read tools and undo/redo themselves are not wrapped. Track Changes mode and up. |

## `apply_commands` command kinds (15)

8 kinds with a genoffice equivalent, 2 officeoffice-only additions with no genoffice
counterpart (`copyBlocks`, `copyFormat`), and 5 officeoffice-only shorthand aliases
(`set_bold`, `set_italic`, `set_heading`, `find_replace`, `set_bullet`) — all cross-checked
exhaustively against `WORD_COMMAND_SCHEMAS` in `entry.ts` (a structural per-kind schema;
no bare `{type:'object'}` items left).

| Command | Notes |
|---|---|
| `updateTextStyle` | bold/italic/underline/strike/sizeHalfPoints/font/color/baselineOffset/link/highlight — full field parity with genoffice. `highlight` is a fixed 16-entry `WdColorIndex` palette (not arbitrary RGB like `color`); an unrecognized `fields` entry errors by name instead of silently no-opping. |
| `updateParagraphStyle` | align/lineSpacing/indentLeft/indentRight/indentFirstLine/spaceBefore/spaceAfter/pageBreakBefore/shadingFill/borders — full parity. |
| `deleteBlocks` | Same `Target` matcher (nodeType/headingLevel/containsText/blockIndexes/scope). Deleting every paragraph clears content instead, leaving one empty paragraph. Errors if the target matches nothing, instead of reporting `ok` having changed nothing. |
| `moveBlocks` | Captures moved paragraphs as OOXML snapshots before deleting, reinserts via `InsertXML` — preserves formatting through the move. Rejects an empty `blockIndexes` array. |
| `copyBlocks` | **No genoffice equivalent.** Duplicates paragraphs matched by the same `Target` matcher `deleteBlocks` uses, leaving originals in place. `afterBlockIndex` may reference one of the copied paragraphs' own indices (allowed, unlike `moveBlocks`). |
| `copyFormat` | **No genoffice equivalent.** Format-painter semantics: copies Font (bold/italic/underline/strike/size/name/color/superscript/subscript/highlight) and ParagraphFormat/Shading/4-side Borders from one source paragraph onto one or more targets, atomically. Hyperlinks are never copied (matches real Word Format Painter). Whole-paragraph granularity only. |
| `createParagraphBullets` / `deleteParagraphBullets` | 7 named `bulletPreset` values (`BULLET_DISC_CIRCLE_SQUARE`, `BULLET_DIAMOND_X`, `BULLET_CHECKBOX`, `NUMBERED_DECIMAL`, `NUMBERED_DECIMAL_ALPHA_ROMAN`, `NUMBERED_UPPERALPHA`, `NUMBERED_UPPERROMAN`). An unrecognized preset errors listing valid names. Heading paragraphs matched-but-skipped, non-list paragraphs matched-but-skipped on delete — both report the skipped count instead of a bare `ok`. `NUMBERED_DECIMAL_ALPHA_ROMAN` only cycles decimal→alpha→roman at the standard 3 outline levels Word itself defines for that scheme — a 4th+ nesting level falls back to decimal rather than continuing the cycle. `NUMBERED_DECIMAL_ALPHA_ROMAN` renders level 1 as plain decimal only — this flat per-paragraph model has no multi-level nesting to show alpha/roman sub-levels. Two Wingdings-glyph presets (`BULLET_DIAMOND_X`/`BULLET_CHECKBOX`) are unverified against a real render. |
| `updateImageProperties` | width/height (proportional scale from current size)/align, targets `doc.InlineShapes` by index. |
| `insertToc` | Uses Word's native `TablesOfContents.Add(UseHeadingStyles: true)` — real, auto-paginating, a more direct native equivalent than a hand-built TOC field-XML workaround (real Word already paginates). |
| `set_bold` / `set_italic` | Alias for `updateTextStyle`'s bold/italic fields, addressed by paragraph-index range rather than a `Target`. |
| `set_heading` | Alias, ≈ genoffice's `setHeadingLevel`. |
| `find_replace` | Alias, ≈ genoffice's `replaceAllText`. Loops and counts replacements accurately. |
| `set_bullet` | Alias turning bullets on/off in one command (`value:true/false`, default true), matching the `set_bold`/`set_italic`/`set_heading` shape — shorthand for `createParagraphBullets`/`deleteParagraphBullets`. |

A malformed command fails only itself (per-command try/catch) and the batch result is a
numbered, per-command report with a summary header (`"Applied 5 of 7 command(s) (2
failed)."`) — not an all-or-nothing abort.

No `dataSource`/provenance-enforcement mechanism exists (see
[`docs/architecture.md`](../architecture.md)). Every genoffice `apply_commands` kind has
a real implementation here.
