# AI Tool Surface Reference — officeoffice

This document catalogs every mechanism by which the LLM (the AI assistant embedded in
the Word/Excel/PowerPoint/Outlook VSTO add-ins) can read or mutate a real Microsoft
Office document — or, for Outlook, the mailbox and calendar — in this repo. It is a
**current-state reference**: every table below reflects the tool surface as it exists
in `entry.ts`/`*Tools.cs` today, not a point-in-time snapshot. For the history of how it
got here (dated bug fixes, phased rollouts, a schema audit and its fixes), see
**`docs/ai-tool-surface-changelog.md`** — nothing in that file should be treated as
current without checking the tables here first.

The Word/Excel/PowerPoint sections compare against the equivalent surface documented for
`genoffice` (`C:\dev\genoffice\docs\ai-tool-surface.md`, a from-scratch web-based
Office clone suite that this project ports its tool design from); the **Outlook**
section has no genoffice counterpart and mirrors `C:\dev\mcp-outlook` instead.

## Architecture

officeoffice drives the **real desktop Office applications** via VSTO + COM interop
(`Microsoft.Office.Interop.{Word,Excel,PowerPoint,Outlook}`), unlike genoffice's
from-scratch web renderers. The chat UI runs in a WebView2 page inside a CustomTaskPane;
tool calls cross a `chrome.webview.postMessage` ⇄ `CoreWebView2.PostWebMessageAsJson`
JSON bridge (`OfficeAi.Shared/ToolProtocol.cs`) into C# handlers that call the COM
object model directly — there is no Electron/IPC hop. Word/Excel/PowerPoint attach one
pane per document window (keyed by `Hwnd`); Outlook attaches one pane per `Explorer`
window (keyed by COM identity, since `Explorer` has no `Hwnd`) and nothing to
Inspectors — see the Outlook section.

`packages/agent-core` and `packages/ai-provider` from genoffice are copied verbatim
into `shared/web-src/{agent-core,ai-provider}` (same `AgentLoop`, same `AgentSkill`
contract, same `maxTurns` default of 8, same multi-provider types including
`genspark`/`anthropic`/`gemini`/`deepseek`/`openai`/`custom`). Provider/model/key
selection, the settings screen, and the transport live once in
`shared/web-src/app-shell/` (`getSettings` / `makeTransport` / `onSettingsSave`,
persisted in WebView `localStorage`), shared by all four add-ins including Outlook —
this is live, not a stub; `makeTransport` routes via `streamForProvider` to whichever
provider the user picks in the settings dropdown, with a "Test connection" button.

There is **no `packages/ai-search` equivalent** in officeoffice: no `web_search`,
`image_search`, `generate_image`, or `analyze_media` anywhere in the repo. This is a
deliberate scope decision (the deployment target is air-gapped — see
`add_image`/`replace_image`'s explicit rejection of remote URLs below), not an
oversight. **`get_attachment` is a partial exception:** Outlook's `get_attachment`
saves an email attachment to a local file and extracts text from text-family and
OpenXML (`.docx/.xlsx/.pptx`) types via `OfficeAi.Shared/AttachmentText/` — but never
fetches anything remote, never handles PDF or images, and cannot feed a binary to the
model. See the Outlook section.

Each add-in has a governance layer genoffice's docs surface doesn't have in this
form: a shared `OfficeAi.Shared.EditingMode` enum (`ReadOnly | CommentOnly |
TrackChanges | FullAutonomy`), filtered client-side (which tools are advertised to the
model) *and* re-enforced server-side in each `Tools.Execute()` (mutating tools blocked
outright in Read Only / Comment Only mode, regardless of what the model requests).
Outlook repurposes all four tiers with its own meaning (Read only / Draft only /
Automate approvals / Full autonomy) rather than dropping two — see the Outlook section
for the mapping.

There is no `dataSource`/provenance-enforcement mechanism anywhere in officeoffice
(genoffice's slides app gates chart data-source claims at the tool layer). A
product-owner decision (see `docs/superpowers/plans/STATUS.md`, PP-7) opted for a
prompt-only mitigation rather than a code-level guardrail, pending real sign-off.

---

## Word (`WordAiAddIn/WordTools*.cs`)

### Top-level tools (19)

| Tool | Kind | Notes |
|---|---|---|
| `get_document_context` | Read | Paragraph/word count + a flat 300-char text preview. No block-indexed list (index\|type\|preview) like genoffice's version. |
| `find_text` | Read | Read-only substring or `.NET` regex search across paragraph text. Returns `[index] text` using the exact same 0-based paragraph index `read_blocks`/`replace_blocks`/`apply_commands`' `Target.blockIndexes` use. Uses Word's native `Range.Find` engine for the plain-substring path (not positional indexing — see the changelog's 2026-08-27 performance-fix entries). |
| `get_headings` | Read | Lists every heading-styled paragraph in order, Navigation-Pane style: `[index] H<level>: text`. Uses `Range.GoTo(wdGoToHeading, wdGoToNext)`, never a per-paragraph scan. |
| `read_blocks` | Read | Paragraph-indexed range read, capped at 1000 paragraphs. `format:"html"` emits a restricted HTML subset (headings/bold/italic/underline/list membership — capped at 100 paragraphs, since per-paragraph formatting reads are slower); plain text is `[i] text` lines. |
| `read_chart` | Read | Reads an existing chart's title, type, categories, and per-series names/values — call before an incremental `edit_chart` edit, since `edit_chart` replaces the whole dataset when given. |
| `read_table` | Read | Reads a table's cell contents, one row per line. `tableIndex` (0-based, document order) defaults to the first table. |
| `read_smartart` | Read | Reads SmartArt node text; omitting `smartArtIndex` reads every diagram in the document in one call. |
| `insert_content` | Write | Supply `text` (plain) or `html` (restricted subset: `<p> <h1>-<h3> <ul>/<ol>/<li> <b>/<strong> <i>/<em> <u> <br/>`, no attributes, no nested lists). `afterBlockIndex` anchors the insertion (0-based paragraph index; `-1` = start of document; omit = end). No longer append-only or plain-text-only (both fixed by PP-10). |
| `replace_blocks` | Write | Replaces paragraphs `[startIndex, endIndex]` with `text` or `html` (same restricted subset). Empty text deletes the range. `preserveFormatting` (default true) reapplies the first replaced paragraph's style to the result. |
| `add_image` | Write | Inserts an image from a **local file path only** (no URLs — air-gapped deployment). Inline by default (addressable afterward via `apply_commands`' `updateImageProperties`), or `floating:true`. |
| `add_table` | Write | Adds a native Word table, optionally pre-filled with cell text (row-major array of arrays). |
| `edit_table` | Write | `kind`: `set_cell`, `insert_row`/`delete_row`/`insert_col`/`delete_col`, `set_style` (styleName/headerRow/bandedRows/borders/borderColor), `set_shading` (cell/row/col/table scope, hex color). Structural edits shift later indices — re-read before a second structural edit in the same run. Row/col index 0 is just the first physical row/column, including a header row - there's no separate header concept in the index space (PR review, 2026-10-02). |
| `add_smartart` | Write | `layout`: list/process/cycle/hierarchy/pyramid/matrix/venn. Flat item list, one per top-level node. |
| `edit_smartart` | Write | `kind`: `set_text`, `add_node`, `delete_node`, `set_style` (colorName/quickStyleName, substring-matched against the install's real gallery), `set_layout`. `delete_node` shifts later node indices. |
| `edit_chart` | Write — gateway | Creates or edits a native Word chart. Supports **categories + one or more named series** (multi-series, chart-type selection: column/columnStacked/bar/barStacked/line/area/pie/doughnut) — no longer single-series/title-only (fixed by PP-9). `chartIndex` addresses an existing chart (0-based, inline then floating shapes, document order); `create:true` always adds a new one. `afterBlockIndex` anchors a new chart inline instead of floating at document origin. |
| `apply_commands` | Write — gateway, see below | |
| `add_comment` | Write (every mode) | Anchors a real Word comment to the first match of given text. **Not in genoffice's docs surface at all.** Available in every editing mode, including Comment Only. |
| `undo_last_action` / `redo_last_action` | Write | Word's own Ctrl+Z/Ctrl+Y over the document's real undo history, via `Document.Undo()`/`Document.Redo()` (`WordTools.History.cs` — Word's `Application` has no Undo/Redo method); each returns a bool, so an empty stack reports "Nothing to undo/redo" with `Mutated: false`. Every other mutating tool call is wrapped in `Application.UndoRecord.StartCustomRecord`/`EndCustomRecord`, so **one tool call = one undo step** (a pre-filled `add_table` or a whole `apply_commands` batch backs out in a single undo). Read tools and undo/redo themselves are not wrapped. Track Changes mode and up. |

### `apply_commands` command kinds (15)

8 kinds with a genoffice equivalent, 2 officeoffice-only additions with no genoffice
counterpart (`copyBlocks`, `copyFormat`), and 5 officeoffice-only shorthand aliases
(`set_bold`, `set_italic`, `set_heading`, `find_replace`, `set_bullet`) — all genuinely
implemented, all cross-checked exhaustively against `WORD_COMMAND_SCHEMAS` in
`entry.ts` (PP-5's structural per-kind schema; no bare `{type:'object'}` items left).

| Command | Notes |
|---|---|
| `updateTextStyle` | bold/italic/underline/strike/sizeHalfPoints/font/color/baselineOffset/link/**highlight** — full 10/10 field parity with genoffice. `highlight` is a fixed 16-entry `WdColorIndex` palette (not arbitrary RGB like `color`) and was the one missing field until PP-12; an unrecognized `fields` entry now errors by name instead of silently no-opping. |
| `updateParagraphStyle` | align/lineSpacing/indentLeft/indentRight/indentFirstLine/spaceBefore/spaceAfter/pageBreakBefore/shadingFill/borders — full parity. |
| `deleteBlocks` | Same `Target` matcher (nodeType/headingLevel/containsText/blockIndexes/scope). Deleting every paragraph clears content instead, leaving one empty paragraph (mirrors genoffice's own guard). Errors if the target matches nothing, instead of reporting `ok` having changed nothing. |
| `moveBlocks` | Captures moved paragraphs as OOXML snapshots before deleting, reinserts via `InsertXML` — preserves formatting through the move. Rejects an empty `blockIndexes` array. |
| `copyBlocks` | **No genoffice equivalent.** Duplicates paragraphs matched by the same `Target` matcher `deleteBlocks` uses, leaving originals in place. `afterBlockIndex` may reference one of the copied paragraphs' own indices (allowed, unlike `moveBlocks`). |
| `copyFormat` | **No genoffice equivalent.** Format-painter semantics: copies Font (bold/italic/underline/strike/size/name/color/superscript/subscript/highlight) and ParagraphFormat/Shading/4-side Borders from one source paragraph onto one or more targets, atomically. Hyperlinks are never copied (matches real Word Format Painter). Whole-paragraph granularity only. |
| `createParagraphBullets` / `deleteParagraphBullets` | 7 named `bulletPreset` values (`BULLET_DISC_CIRCLE_SQUARE`, `BULLET_DIAMOND_X`, `BULLET_CHECKBOX`, `NUMBERED_DECIMAL`, `NUMBERED_DECIMAL_ALPHA_ROMAN`, `NUMBERED_UPPERALPHA`, `NUMBERED_UPPERROMAN`) — no longer collapsing to a generic bullet for anything but `NUMBERED_*` (fixed by PP-12); an unrecognized preset now errors listing valid names. Heading paragraphs matched-but-skipped, non-list paragraphs matched-but-skipped on delete — both report the skipped count instead of a bare `ok`. **Two Wingdings-glyph presets (`BULLET_DIAMOND_X`/`BULLET_CHECKBOX`) are unverified against a real render** — see `docs/superpowers/verification/pp12.md`. |
| `updateImageProperties` | width/height (proportional scale from current size)/align, targets `doc.InlineShapes` by index. |
| `insertToc` | Uses Word's **native** `TablesOfContents.Add(UseHeadingStyles: true)` — real, auto-paginating, more direct than genoffice's hand-built TOC field-XML. |
| `set_bold` / `set_italic` | Alias for `updateTextStyle`'s bold/italic fields, addressed by paragraph-index range rather than a `Target`. |
| `set_heading` | Alias, ≈ genoffice's `setHeadingLevel`. |
| `find_replace` | Alias, ≈ genoffice's `replaceAllText`. Loops and counts replacements accurately (previously always reported 0 or 1 regardless of actual count). |
| `set_bullet` | Alias turning bullets on/off in one command (`value:true/false`, default true), matching the `set_bold`/`set_italic`/`set_heading` shape — shorthand for `createParagraphBullets`/`deleteParagraphBullets`. |

A malformed command fails only itself (per-command try/catch, PP-5) and the batch
result is a numbered, per-command report with a summary header (`"Applied 5 of 7
command(s) (2 failed)."`, PP-12) — not an all-or-nothing abort.

No `dataSource`/provenance-enforcement mechanism exists (see Architecture). Every
genoffice `apply_commands` kind has a real implementation here.

---

## Excel (`ExcelAiAddIn/ExcelTools*.cs`)

### Top-level tools (10)

9 read/query tools plus `propose_operations` — full 1:1 parity with genoffice's naming
and shape for the read side (`get_workbook_context`, `read_range`, `read_cells`,
`select_range`, `read_formats`, `read_sheet_features`, `find_cells`, `trace_precedents`,
`trace_dependents`). `load_guide` has no equivalent (deliberately out of scope —
genoffice's is an internal prompt-budget mechanism for managing its larger op count in
context, not needed at officeoffice's current scale).

`find_cells` and `propose_operations`' `find_replace` op both default to **the active
sheet only** (matching Ctrl+F/Ctrl+H's default "Within: Sheet"), with an `allSheets`
boolean to search the whole workbook ("Within: Workbook") or `sheetId` to name one
specific sheet. Both use Excel's native `Range.Find`/`FindNext` for the plain-substring
path (one native pass per `LookIn` mode, not a per-cell `foreach` scan) — `regex:true`
still needs a per-cell scan, since Excel's `Find` has no regex mode. `find_replace`
only ever touches literal cell **values**, never formulas — use `find_cells` to locate
formulas, then `set_formula` to edit them.

No `undo_last_action`/`redo_last_action` for Excel: Excel clears its undo stack on any
object-model write, so there is no native history for such a tool to walk (see the
changelog's 2026-09-27 entry for the tools' brief existence and removal).

Notable native-COM advantage: `find_cells`'s `errors_only` mode uses
`Range.SpecialCells(xlCellTypeFormulas, xlErrors)` — a genuinely native error-cell
scan the code comments call out as the categorical VSTO/COM advantage over
Office.js's wildcard-only `Range.find`.

### `propose_operations` operation kinds (54) — all named kinds from genoffice's list are implemented

| Group | Kinds |
|---|---|
| Writing | `set_cell`, `set_formula`, `set_range`, `clear_cell`, `clear_range`, `find_replace` |
| Formatting | `format_range` |
| Layout | `sort_range`, `merge_cells`, `unmerge_cells`, `copy_range`, `move_range`, `set_row_height`, `set_col_width`, `set_rows_hidden`, `set_cols_hidden`, `set_freeze`, `set_page_setup` |
| Structure | `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`, `add_sheet`, `delete_sheet`, `duplicate_sheet`, `set_sheet_hidden`, `move_sheet`, `protect_sheet`, `rename_sheet` |
| Charts/visuals | `add_chart`, `edit_chart`, `delete_visual`, `add_sparkline`, `add_shape`, `edit_shape`, `add_image` |
| Tables | `add_table`, `add_table_row`, `add_table_column`, `delete_table_row`, `delete_table_column`, `delete_table` |
| Pivot | `add_pivot`, `refresh_pivot` |
| Data | `set_hyperlink`, `set_note`, `add_defined_name`, `delete_defined_name`, `set_filter`, `clear_filter`, `set_filter_criteria`, `add_conditional_format`, `clear_conditional_formats`, `set_data_validation` |

`copy_range`/`move_range` are a superset addition beyond genoffice's own list —
genoffice has no equivalent named kind for duplicating or relocating an arbitrary
rectangular range (capped at 2000 source cells; floating objects inside the range are
not moved; `move_range` clears the source as part of the move, native Excel Cut
behavior, and other formulas referencing the moved cells auto-update).

Every kind's real per-field JSON Schema lives in `EXCEL_OPS` (`ExcelAiAddIn/web-src/
entry.ts`) — the single source of truth for both the wire schema and the human-readable
description (PP-5), cross-checked exhaustively against `ExcelTools.cs`'s
`ProposeOperations` switch.

Where officeoffice's version is still **narrower** than genoffice's:

| Op | Gap |
|---|---|
| `add_image` | **Local file paths only** — remote URLs throw `NotSupportedException` ("air-gapped deployment"). genoffice downloads from `image_search`/`generate_image` results. Deliberate scope boundary, not a bug. |
| `add_shape` | 26 named preset types + textbox (case-insensitive match, unknown name errors listing valid ones) — narrower than genoffice's "full OOXML preset-geometry set" but substantial. |
| `set_data_validation` | `checkbox` kind explicitly rejected — Excel's Data Validation COM API (verified via reflection against the referenced PIA) has no boolean-checkbox validation type; only 5 kinds exist total (`list`/`listRef`/`numberBetween`/`dateBetween`/`formula`), none map to it. |

`format_range` (previously missing ~7 of ~11 style properties genoffice supports) and
`add_chart` (previously column/line/pie only, with the richer vocabulary only on
`edit_chart`) have both since reached parity: `format_range` now covers font
name/size/color, strikethrough, underline (enum-or-boolean), horizontal/vertical
alignment, wrap, rotation, indent, and full border control (preset+edges+style+color);
`add_chart` now shares the same widened `ExcelChartTypeMap` as `edit_chart`
(column/columnStacked/bar/barStacked/line/area/pie/doughnut) and supports rebinding to
a new range (`dataRange`+`dataSheet`+`plotBy`).

No `dataSource`/provenance-enforcement mechanism exists anywhere (genoffice's slides
app gates chart data-source claims; nothing analogous exists in officeoffice's Excel
or PowerPoint chart tools — see Architecture).

---

## PowerPoint (`PowerPointAiAddIn/PowerPointTools*.cs`)

### Tools (51 total, all genuinely implemented)

The largest and fastest-growing tool surface of the three apps. Organized here by area
rather than as one flat table.

**Reading (8, always allowed, never gated — `READER_TOOLS` in `entry.ts`)**

| Tool | Notes |
|---|---|
| `get_deck_context` | One-line-per-slide outline: slide index + text preview of its shapes. No per-element type/id inventory like genoffice's version. |
| `read_slide` | Full text of every shape on one slide, plus its layout, transition, animation count, and speaker notes. Shapes listed back-to-front (z-order). A group is shown as one line with a child count — see `read_group`. |
| `read_group` | Lists a group's children recursively (nested groups expanded in place), each prefixed with a dotted path (e.g. `"3.1.0"`) usable as `shapeIndex` for `set_element_text`/`set_element_style`/`set_element_fill`/`set_element_stroke`. Positional/structural edits still require `ungroup_element` first. |
| `read_animations` | A slide's animations in play order (shape, effect, entrance/exit, trigger, timing) — `animationIndex` in `edit_animation` addresses this same order. |
| `find_text` | Read-only substring/regex search across every slide's shape text (text boxes, placeholders, table cells, SmartArt node text) and speaker notes. Returns `[slide i, shape j] text` or `[slide i, notes] text`. |
| `read_smartart` | Node text of SmartArt diagrams on a slide; omitting `smartArtIndex` reads every diagram on the slide. |
| `list_layouts` | Every layout name in the presentation (across every design/theme), plus how many slides use each — call before passing `layoutName` elsewhere instead of guessing. |
| `read_master_elements` | Shapes placed directly on the Slide Master, or one specific layout's own shapes (`layoutName`, substring-matched). |

**Slide management**

| Tool | Notes |
|---|---|
| `add_slide` | Clones a source slide's layout as a new slide inserted after it, optional text-clear. |
| `delete_slide` | Deletes one slide. Cannot delete the last remaining slide; later slides shift down — re-read before deleting another in the same run. **Now implemented and in-scope** (PP-19) — the retired `tool-surface-todo.md` had bundled this with the (still out-of-scope) scripting DSL, which was a real inconsistency; this is resolved. |
| `move_slide` | Moves a slide to a new 0-based position. |
| `duplicate_slide` | Inserts a copy of a slide (content included) directly after it — use `add_slide` instead for layout-only, content-free duplication. |
| `set_slide_layout` | `kind:"classic"` (16 standard `PpSlideLayout` values) or `kind:"custom"` (free-text `layoutName`, matched against the deck's own theme layouts). |
| `set_slide_transition` | 31 curated entry-effect values, duration, click/timed advance (independent toggles). |
| `set_slide_background` | Solid color for one slide or `slideIndex:-1` for every slide. |
| `set_headers_footers` | Slide number/footer/date toggles, deck-wide (`slideIndex:-1`) or per-slide — mirrors PowerPoint's native "Insert Header and Footer" dialog. `skipTitleSlide`/`startNumber` are always deck-wide regardless of `slideIndex` (PowerPoint has no per-slide version of either). |

**Shape editing**

| Tool | Notes |
|---|---|
| `set_element_text` | Replaces one shape's text. `bulleted:true/false` controls real PowerPoint bullets — the model is instructed never to type a literal "•"/"-"/"*". |
| `set_element_style` | bold/italic/underline/shadow/fontSize/fontName/color/alignment/baselineOffset. **Strikethrough deliberately not implemented** — this PIA has no `TextFrame2` member on `Shape` at all (confirmed via a `CS0234` compile failure), unlike Excel's newer text model. |
| `set_element_transform` | left/top/width/height/rotation. |
| `set_element_order` | z-order: `bringToFront`/`sendToBack`/`bringForward`/`sendBackward`. Shifts other shapes' `shapeIndex` on the slide — re-read before addressing another shape by index in the same run. |
| `set_element_fill` / `set_element_stroke` | Solid fill or none; outline color/width or remove. |
| `set_slide_notes` | Replaces a slide's speaker notes. |
| `delete_element` | Deletes one shape. |
| `crop_image` | Fractional crop (0..1) against current on-slide size. Documented imprecision under repeated crops. |
| `replace_image` | **Local file path only** (same air-gapped constraint as Excel/Word) — swaps a picture's content in place, keeping position/size/rotation/approximate z-order. |
| `set_picture_opacity` | Via `Fill.Transparency`. |

**Shape creation, duplication, and cross-slide copy/move**

| Tool | Notes |
|---|---|
| `add_text_box` | New text box; `bulleted:true` for a real bulleted list. |
| `add_shape` | 26 presets + textbox, shares Excel's ported `ShapeTypeMap`; unrecognized name errors listing valid ones (no longer a silent rectangle fallback). No fill/line params — chain `set_element_fill`/`set_element_stroke`. |
| `duplicate_element` | Copies a shape on the **same** slide — every shape kind, including groups/pictures/tables/charts/SmartArt. Offset or exact position; auto-deduped name. |
| `copy_element` | Copies a shape to a **different** slide using PowerPoint's own native copy/paste (the real Windows clipboard — an explicit exception to this codebase's usual rule against it, since it reproduces the exact underlying OOXML the same way Ctrl+C/Ctrl+V does). Every shape kind supported with full native fidelity, including groups nesting SmartArt. Refuses `targetSlideIndex == slideIndex` (use `duplicate_element` for that). |
| `move_element` | Same as `copy_element`, but also removes the original — atomic (only removed once the copy fully succeeds). |
| `copy_element_style` | Format painter: native `Shape.PickUp`/`Apply` (not the clipboard) for broad fidelity (fill/outline/shadow/3-D/bevel), plus targeted per-run text formatting, text outline/strikethrough/glow/reflection, text-box anchor/margins, gradient/patterned fill, dash/arrowhead outlines, rotation, flip, and AutoShape adjustment handles. Not position or size. Targets can span any slide. |
| `group_element` / `ungroup_element` | Groups two or more top-level shapes, or promotes a group's children back to top-level. Both shift other shapes' indices on the slide — re-read before another index-addressed edit in the same run. |

**Tables**

| Tool | Notes |
|---|---|
| `add_table` | Native `Shapes.AddTable`, optional pre-filled cell text. |
| `edit_table_cell` | Replaces one cell's text. Row/col index 0 is just the first physical row/column, including a header row - there's no separate header concept in the index space (PR review, 2026-10-02). |
| `edit_table_structure` | Insert/delete row or column; bounds-checked (an out-of-range index errors specifically rather than a raw COM error). Shifts later indices. Row/col index 0 is just the first physical row/column, including a header row - there's no separate header concept in the index space (PR review, 2026-10-02). |
| `edit_table_style` | firstRow/bandRow, shading, `borderPreset`: `all` (every cell edge) / `outline` (outer perimeter only) / `none`. |

**Charts**

| Tool | Notes |
|---|---|
| `add_chart` | column/columnStacked/bar/barStacked/line/area/pie/doughnut, writes real data into the chart's embedded Excel workbook. `kind` is enum'd from the corrected chart-type map (see below); a series/categories length mismatch errors instead of drawing a silently wrong chart. Returns the new shape's index for a follow-up `edit_chart`. |
| `edit_chart` | `chartType`, `title`, `legendPos` (`none`/`right`/`top`/`left`/`bottom` + short aliases `r`/`t`/`l`/`b`), `dataLabels` (`none`/`value`/`percent`), `gridlines` (errors with a clear message on chart types with no value axis, e.g. pie/doughnut). Result names exactly which properties were applied. **The `"bar"` → `xlColumnClustered` mapping bug (silent wrong chart type on a successful call) and `legendPos`'s terminal-else-to-bottom silent fallback are both fixed** — unrecognized values now throw, listing valid ones from the map itself. |

**SmartArt**

| Tool | Notes |
|---|---|
| `add_smartart` | 7 layout keys (list/process/cycle/hierarchy/pyramid/matrix/venn) mapped to native SmartArt layouts by display name; flat item list only. Unrecognized key errors listing the 7 valid ones; a valid key with no matching gallery layout gives a distinct "may be a non-English Office install" error instead of silently falling back to Basic Block List. |
| `edit_smartart` | `set_text`, `add_node`, `delete_node`, `set_style` (colorName/quickStyleName, substring-matched against the gallery), `set_layout`. `smartArtIndex` is 0-based **within that slide**. `delete_node` shifts later indices. |

**Animations & transitions** — see also `set_slide_transition` above

| Tool | Notes |
|---|---|
| `add_animation` | Entrance (default) or exit, 20 effect names, `trigger` (`onClick`/`withPrevious`/`afterPrevious`), duration/delay. No directional variants (e.g. "wipe from the left") — base effect only. Returns the new `animationIndex`. |
| `edit_animation` | `delete`, `set_timing` (duration/delay/trigger), `reorder` (`toIndex`). `animationIndex` addresses current play order (call `read_animations` first); delete/reorder shift later indices. |

**Slide Master**

| Tool | Notes |
|---|---|
| `add_master_element` | Pins a persistent icon or text label to the Slide Master (shows on every slide) or one specific layout (`layoutName`). Corner placement (inset by `marginPt`) or exact `left`+`top`. Known limits: only the default (first) Slide Master is reachable without `layoutName`; a layout with "Hide Background Graphics" won't show a master-level element. |
| `set_master_element_transform` | Moves/resizes/rotates an existing master element by `masterShapeIndex` (from `read_master_elements`) — works on theme placeholders too. |
| `remove_master_element` | Deletes a master shape by index. Refuses to delete a theme placeholder (marked `(placeholder)` in `read_master_elements`) — use `set_headers_footers`'s `*Visible:false` to turn one off instead. |

**Search/replace across the deck**

| Tool | Notes |
|---|---|
| `replace_text` | Every text-frame shape (text boxes, title/body placeholders) on every slide, plus speaker notes unless `includeNotes:false`. **Not** table cells or SmartArt node text — use `edit_table_cell`/`edit_smartart`. `regex:true` supports `$1`-style backreferences. Reports the number of occurrences actually replaced. |

**Undo/redo (2)**

| Tool | Notes |
|---|---|
| `undo_last_action` / `redo_last_action` | PowerPoint's own Ctrl+Z/Ctrl+Y. PowerPoint's object model has no Undo/Redo method, so these go through the generic ribbon-command dispatch, `Application.CommandBars.ExecuteMso("Undo"/"Redo")`, with a `GetEnabledMso` pre-check so an empty stack reports "Nothing to undo/redo" instead of a silent no-op (`PowerPointTools.History.cs`). `Application.StartNewUndoEntry()` runs before **every** non-read tool call — undo/redo included, an intentional difference from Word, whose `UndoRecord` wrapping excludes them — so one tool call = one undo step. Track Changes mode and up. |

### Missing entirely (confirmed absent from both the C# switch and the advertised tool list)

- `execute_slide_script` — no scripting DSL; every multi-property/multi-element edit
  must go tool-by-tool, rather than genoffice's atomic AST-interpreted batch script.
  A 2026-08-24 product-owner-absent decision (`docs/superpowers/plans/STATUS.md`,
  PP-19 Task 2) opted for a middle-ground "batch-ops + layout-check" approach as the
  eventual replacement, but it is **not yet built** — this remains a real gap today.
- The entire deck-generation pipeline: `ask_clarification`, `plan_deck`,
  `generate_deck`, `regenerate_slide`, `save_style_template`, `list_style_templates`.
- No automatic post-edit audit/QC pass — genoffice's `auditSlideLayout` (geometric
  overflow/overlap/bounds check after every script run) and `slide-qc.ts` (vision-based
  QC sub-loop after generated pages) have no counterpart; nothing here checks the
  result of an edit automatically.
- No `add_comment`-equivalent for PowerPoint (unlike Word), so Comment Only mode
  currently behaves identically to Read Only.

### Structural fragility note

Shapes are addressed by **positional index** (`slideIndex`, `shapeIndex` into
`slide.Shapes`, or a dotted path like `"3.1.0"` for a shape inside a group) rather than
a stable id — indices shift whenever shapes are added/removed/reordered/grouped, unlike
genoffice's `sourceId`-based addressing. Several tools' descriptions explicitly tell the
model to re-read the slide after a structural change, but there's no protection against
acting on a stale index otherwise.

---

## Outlook (`OutlookAiAddIn/OutlookTools.*.cs`)

genoffice has no mail or calendar app, so there is no "vs. genoffice" column here. The
tool set mirrors `C:\dev\mcp-outlook` (a self-hostable EWS MCP server) as closely as the
Outlook COM object model allows — but this add-in drives the **already-running,
already-authenticated desktop Outlook client** via
`Microsoft.Office.Interop.Outlook`, not EWS. There is no `Namespace`/credential setup;
it acts as the signed-in user. **Three exceptions use EWS** (see "EWS carve-out"
below): `search_contacts`, `find_meeting_slots`' work-week lookup, and `list_events`'
shared-calendar path — each because the COM object model either can't do the job
(a multi-result directory search) or does it in a way that freezes Outlook.

### Shape of the integration (differs from Word/Excel/PowerPoint)

- **Explorer-only.** The chat pane docks in the main Outlook window (`Explorer`).
  There is **no task pane in Inspector windows** (pop-out read/compose). The ribbon
  button is on the Explorer's Mail tab (`idMso="TabMail"`, via `RibbonBase.HomeTabIdMso`
  / `ProvidesRibbonFor`) and suppressed on every Inspector ribbon surface.
- **Per-mailbox chat.** `GetChatId()` returns `mbx-` + `SHA256(primary SMTP)[:16]` —
  one rolling conversation per mailbox. No file path, so none of Word/Excel/
  PowerPoint's provisional-id / `ChatStore.Migrate` / `DocSettingsStore` save
  lifecycle applies. Primary SMTP is resolved lazily (`ExchangeUser.PrimarySmtpAddress`
  → `PR_SMTP_ADDRESS` proptag → `AddressEntry.Address` → first account → display name).
- **Pane lifecycle.** `Explorer` has no `Hwnd`, so panes are keyed by the COM
  identity pointer (`Marshal.GetIUnknownForObject`). Creation is **lazy** — on an
  explorer's first `Activate` or the ribbon button, never in `ThisAddIn_Startup` —
  because Outlook auto-disables add-ins whose median startup exceeds ~1 s, and
  `install.ps1` also writes `Resiliency\DoNotDisableAddinList`.
- **Editing modes — four tiers.** Outlook repurposes all four `EditingMode` slots with
  its own meaning. Each tier is a strict superset of the one before it:

  | `EditingMode` | Outlook label | Adds |
  |---|---|---|
  | `ReadOnly` | **Read only** | The 12 read tools (`readOnlyTools` in `entry.ts`, `AlwaysAllowedTools` in `OutlookTools.cs`). |
  | `CommentOnly` | **Draft only** (the default for a fresh session) | 23 tools that mutate the mailbox or open a draft but never leave anything unreviewed: triage (`mark_email_read`/`mark_email_unread`, `flag_email_important`, `move_email`, `delete_email` — non-permanent only, see below), tasks/reminders (`create_task`, `update_task`, `set_reminder`, `set_email_reminder`), event tagging/status (`set_event_categories`, `set_category_color`, `set_event_availability`), `apply_search`, every draft tool (`draft_email`, `reply_email`, `reply_all_email`, `forward_email`, `draft_event`, `draft_edit_event`, `draft_cancel_event`, `draft_respond_meeting`), and `undo_last_action`/`redo_last_action` (`commentOnlyExtraTools` / `DraftTierTools`). |
  | `TrackChanges` | **Automate approvals** | `accept_meeting`/`decline_meeting`/`tentative_meeting` — these call `resp.Send()` to notify the organizer, so they get their own explicit tier rather than hiding in Draft only or Full autonomy (`trackChangesExtraTools` / `ApprovalTierTools`). |
  | `FullAutonomy` | **Full autonomy** | The seven auto-send tools — `send_email`, `send_reply`, `send_reply_all`, `send_forward`, `create_event`, `edit_event`, `cancel_event` (`SendTierTools`; on the `entry.ts` side, simply absent from every lower tier's list) — plus `delete_email` with `permanent: true`. |

  `entry.ts` passes `availableModes: ['readOnly','commentOnly','trackChanges',
  'fullAutonomy']` and a `modeOverrides` map so the mode menu shows Outlook's own
  labels/descriptions instead of Word/Excel/PowerPoint's generic "Comment only"/"Track
  changes" text — those apps' own use of `CommentOnly`/`TrackChanges` for real editing
  is untouched (Outlook's relabeling is purely its own `entry.ts` config, not a
  shared-string change). Gate logic (`OutlookTools.cs`'s `ExecuteAsync`) is an ordinal
  check on the enum's own declared order (`ReadOnly < CommentOnly < TrackChanges <
  FullAutonomy`), name-based except for one narrow input-aware special case:
  `delete_email` with `permanent: true` is blocked below Full autonomy (the tool name
  itself stays in `DraftTierTools`, since its default non-permanent path is a
  reversible move). A fresh session defaults to **Draft only** (`defaultMode:
  'commentOnly'`, resolved via `chat-ui.ts`'s `defaultModeFor()` — the same helper
  Word/Excel/PowerPoint use to default to their own `trackChanges`).
- **Selection context.** The Explorer's `SelectionChange` pushes the selected mail
  item(s) / conversation into per-turn context as a `mail` `SelectionContext` variant
  (`shared/web-src/app-shell/bootstrap.ts`) carrying each item's `EntryID`, so the
  model prefers the selection over searching. The scope pill shows the subject (one
  message) or a count (`scopeUnit: 'mailbox'`).
- **`message_id` / `event_id` / `task_id` = Outlook `EntryID`**, re-resolved via
  `Namespace.GetItemFromID` on every call. EntryID is stable within a folder but
  **changes on `Move` and across stores** — `move_email` / `delete_email` return the
  new id. Recurring calendar instances all share the master appointment's EntryID —
  see "Recurring events" below for how the calendar-editing tools target a single
  occurrence anyway. An `event_id` from someone else's shared calendar also needs its
  `store_id` — see "Shared calendars" below.
- **Native query APIs are mandatory** (see the slow Word-search incident in the
  changelog — the same class of bug, avoided here from the start): `list_*` use
  `Folder.GetTable` (an in-memory rowset, no per-item COM object); `search_emails` uses
  `Items.Sort` → `Items.Restrict("@SQL=" + DASL)`; a capped linear scan is a fallback
  only when `Restrict` rejects the filter. The DASL builder is pure and unit-tested
  (`OfficeAi.Shared/OutlookDasl.cs`). `search_contacts` and `list_events`' shared-calendar
  path are the deliberate, documented exceptions: both go to EWS instead of a COM scan —
  see the EWS carve-out below.
- **EWS carve-out: `search_contacts`, `find_meeting_slots`' work-week default, and
  `list_events`' shared-calendar path are the only things that touch EWS.** Every
  other Outlook tool is pure `Microsoft.Office.Interop.Outlook` COM against the running
  client. All EWS-dependent code lives apart from the pure-COM tool files, in two
  layers: `OutlookEws.cs` (the raw EWS Managed API wire calls — `ResolveNamesAsync`,
  `GetWorkingHoursAsync`, `GetSharedCalendarEventsAsync`) and `OutlookTools.Ews.cs`
  (the tool-facing orchestration on top: `SearchContactsAsync`, `ResolveWorkWeekAsync`,
  and the shared endpoint/account resolution every caller uses,
  `ResolveEwsUrlAsync`/`FindExchangeAccountInfo`). All EWS calls use EWS Managed API
  2.2 (`Microsoft.Exchange.WebServices` 2.2.0) with
  `ExchangeService.UseDefaultCredentials` (Windows Integrated Auth as the signed-in
  user — the .NET equivalent of `mcp-outlook`'s `auth_type=sspi`; no stored
  credentials). Endpoint is parsed from the cached `Outlook.Account.AutoDiscoverXml`
  (`<EwsUrl>`/`<ASUrl>`, `EXCH` preferred over `EXPR`; pure parser
  `OfficeAi.Shared/EwsAutodiscoverXml.cs`), falling back to
  `ExchangeService.AutodiscoverUrl`, then cached in a process-static `Uri` shared by
  every EWS-dependent tool. Calls run off the UI thread (`await Task.Run`,
  `svc.Timeout` 15 s) so Outlook stays responsive. **On-prem Exchange only.** EWS
  unreachable / SSPI failure / endpoint not found / timeout → a clear `IsError` result
  for `search_contacts` and for `list_events` with `mailbox` (EWS isn't optional
  there); `find_meeting_slots` instead falls back to a hardcoded Sun-Thu/9-18 default,
  since EWS is an enhancement over an already-working COM-only path there, not the
  only way to do the job. These three are the only async tools — the shared
  `ToolExecutor` delegate is `Task<ToolResult>`-returning
  (`WebViewBridgeHost.OnWebMessageReceived` is `async`); every other tool in all four
  add-ins is still synchronous, wrapped in `Task.FromResult`.

### Read tools (12 — always allowed, never gated)

| Tool | Notes |
|---|---|
| `list_emails` | `Folder.GetTable`; columns EntryID/Subject/ReceivedTime/SenderName/UnRead + `PR_HASATTACH` proptag; `[UnRead] = true` restriction when `unread_only`; sorted newest-first; non-mail rows filtered by `MessageClass` not starting `IPM.Note`. Args: `folder`, `limit` (20), `unread_only`. |
| `search_emails` | `Items.Sort("[ReceivedTime]")` then `Restrict("@SQL=" + BuildSearchFilter(...))`. `LIKE '%q%'` on subject + body; UTC-ISO date range; sender by `fromemail =` OR `fromname LIKE` (Exchange senders carry `legacyExchangeDN`, not SMTP — display-name fuzzy match). `ci_phrasematch`/`ci_startswith` are **not** usable via `Restrict` (they throw). `recipient` is a client-side filter, capped 500. Fallback: capped linear scan on a malformed filter. |
| `get_email` | Full `Body` (≤ 40k), To/CC via `Recipient.Type`, `ConversationID`/`ConversationTopic`, importance, unread, and an `attachments` array — `{index (1-based), name, type (byValue/embeddedItem/ole/reference), size}` — feed the index to `get_attachment`. |
| `open_email` | `MailItem.Display(false)` on an existing item resolved via `ItemById` — opens the message in its own Outlook reading window, unmodified. `Mutated: false`. |
| `get_attachment` | `Attachment.SaveAsFile` into `%LOCALAPPDATA%\OutlookAiAddIn\Attachments\`; returns the path. `extracted_text` (≤ 40k) is populated **only** for text-family extensions (`.txt .csv .tsv .md .json .xml .log`, `.html` tag-stripped) and OpenXML (`.docx .xlsx .pptx`), via the swappable `OfficeAi.Shared/AttachmentText/` module (`DocumentFormat.OpenXml` 2.20.0). **No PDF, no images, no vision** — those return the path + type only. `olOLE` throws (rejected); `olByReference` has no data (rejected); `olEmbeddedItem` saves as `.msg`. |
| `list_folders` | Recursive walk of every store's `Folders`, mail folders only (`DefaultItemType == olMailItem`), with item + unread counts; capped ~800 / depth 8. |
| `search_contacts` | **EWS `ResolveName` called twice — `ContactsOnly`, then `DirectoryOnly` (GAL) — both with `returnContactDetails: true`, results merged** (server-side ANR), run off the UI thread. Querying both locations unconditionally matters: Exchange's ANR on the Contacts folder only matches `DisplayName`, while the Directory phase also matches given name/surname, so a single `ContactsThenDirectory` call (which stops at the first Contacts hit) could miss the person entirely. Each `NameResolution` mapped to a match carrying both the person's **full name** (GivenName + Surname) and the directory's **display name** — the output leads with the full name (what a query like "John Doe" actually matches against) and shows the display name as a parenthetical only when it adds information, since an org-formatted directory label (e.g. `"Dept/Unit/Title"`) can look nothing like the person's name. GAL X500/`EX` addresses fall back to the resolved contact's own `EmailAddress1..3`; entries with no `@` address are dropped (mirrors `mcp-outlook`). Deduped by lowercased address (name fallback), capped at `limit` (default 10). Pure helpers `EwsAutodiscoverXml.ParseEwsUrl` + `ContactSearchFormat.Format` are unit-tested. **No `folder`/scope arg.** On-prem Exchange only; unreachable / auth failure / no endpoint / 15 s timeout → a specific `IsError` message, no COM fallback (a recursive multi-store contact-folder crawl froze then crashed Outlook, which is why there is none). |
| `list_events` | **Own calendar** (no `mailbox`): `Items.Sort("[Start]")` → `Items.IncludeRecurrences = true` → `Items.Restrict("[Start] <= end AND [End] >= start")` — **this order is load-bearing** and rules out `GetTable`. Recurring instances share the master `event_id`; each row carries its own `start` to disambiguate (and to pass as `occurrence_date`). **Shared calendar** (`mailbox` given — an email address): resolves via `Ns.CreateRecipient(mailbox).Resolve()`, then queries EWS `FindAppointments` + `CalendarView` (server-side date filtering and recurrence expansion, off the UI thread) instead of enumerating COM `Items` — see "Shared calendars" below. Args: `start_date` (today), `end_date` (+7d), `mailbox` (optional), `limit` (50). |
| `get_event` | `Body` (≤ 40k), `RequiredAttendees`/`OptionalAttendees`, organizer, response status, recurring flag. Optional `store_id` resolves an `event_id` from a shared calendar (see "Shared calendars" below); omit it for the caller's own events. |
| `find_meeting_slots` | `Recipient.FreeBusy(anchor, 30, true)` — a per-30-min status string — for `Namespace.CurrentUser` + each resolved attendee; then `OfficeAi.Shared.MeetingSlots.Rank` (pure, unit-tested) slides a `duration_minutes` window in 30-min steps across each work day's `[start_hour, end_hour)` and scores each candidate by how many people are free (so a best partial match still comes back). **Work week/hours are read from the mailbox's own EWS `GetUserAvailability` → `AttendeeAvailability.WorkingHours` (`OutlookEws.GetWorkingHoursAsync`), not hardcoded** — falls back to Sun–Thu 09:00–18:00 only if that call fails (non-Exchange profile, EWS unreachable, etc.), cached per process like the EWS URL. Default range is today through the end of the current contiguous work-day run, max 28 days. Times past the returned free/busy window are assumed free. Async only because of the EWS work-week lookup; the FreeBusy/ranking work itself is still synchronous COM. `Recipient.FreeBusy` is all local-time, so no cross-timezone math is needed. Args: `attendees` (req), `duration_minutes` (req), `start_date`, `end_date`, `start_hour`, `end_hour` (both default to the resolved work hours, or 9/18 as a last resort), `limit` (5). |
| `list_tasks` | `Folder.GetTable` over the default Tasks folder; open tasks only unless `include_completed`. Columns EntryID/Subject/Due/Start/Status/PercentComplete/Complete/ReminderTime. |
| `list_color_categories` | `Namespace.Categories` — the profile's master color-tag ("Category") list shared by mail/calendar/tasks, same list Outlook's Categorize picker shows. Each entry: `{name, color}`; color is one of the 26 `OlCategoryColor` values (None/Red/Orange/…/Dark Maroon), mapped to a friendly display name in `OutlookTools.Categories.cs` (not in `OfficeAi.Shared` — that project doesn't reference the Outlook PIA, same split as `ColorUtil`). |

### Mailbox, task, and event actions (13 — Draft only or higher)

All of these act directly (no review window) but stay local to the mailbox — nothing is
sent. All except `apply_search` set `Mutated = true` and record an undo entry (see
"Undo/redo tools" below).

| Tool | Notes |
|---|---|
| `mark_email_read` / `mark_email_unread` | `MailItem.UnRead` + `.Save()`. |
| `flag_email_important` | `Importance = olImportanceHigh/Normal` + `.Save()`. `important` defaults true. |
| `move_email` | `MailItem.Move(ResolveFolder(destination))`; returns `{message_id: <new EntryID>, old_message_id}`. |
| `delete_email` | Non-permanent (default) → `Move` to Deleted Items (returns new id) — Draft only, same risk class as `move_email`. `permanent: true` → then `.Delete()` from there (no single-call hard delete in the OM — documented as "may still be server-recoverable") — **requires Full autonomy**, via the input-aware check in `ExecuteAsync` described under "Editing modes" above; the tool description says so, so the model doesn't attempt it at a lower tier. |
| `apply_search` | View-only (`Mutated: false`, no undo entry) — pushes the same `BuildSearchDasl` filter `search_emails` computes into `Explorer.CurrentFolder` + `Explorer.Search("@SQL=" + dasl, olSearchScopeCurrentFolder)`, so the user's own Outlook window shows the results. Draft tier rather than always-allowed even though it never changes data: it visibly takes over the user's real Outlook window, which Read only promises never to do. `_Explorer.Search(string Query, OlSearchScope SearchScope)`'s signature was confirmed via .NET reflection against the referenced PIA; the `@SQL=` DASL string itself has **not yet been exercised against a live Outlook session**. |
| `set_event_categories` | `AppointmentItem.Categories` (comma-separated tag names, the color shown on the event in the calendar grid) + `.Save()`; empty/omitted `categories` clears all tags. A name outside the master list is auto-added by Outlook on `Save` with an arbitrary color — call `set_category_color` first to control it. Optional `store_id` (shared calendars). |
| `set_category_color` | `Namespace.Categories[name]` — updates `.Color` if the tag exists, else `Categories.Add(name, color)` creates it. Same master list `list_color_categories` reads. |
| `set_event_availability` | `AppointmentItem.BusyStatus` (the calendar's "Show As" dropdown — one of the 5 `OlBusyStatus` values: Free/Tentative/Busy/Out of Office/Working Elsewhere) + `.Save()`. Case/spacing-insensitive friendly-name lookup; an unrecognized value returns `IsError` listing the valid names instead of throwing. Purely local, never `.Send()` — **verified against a live Outlook client**, including on an organized meeting with attendees (no prompt, no notification). Optional `store_id` (shared calendars). |
| `create_task` | `Application.CreateItem(olTaskItem)` + `.Save()` — no window (a task doesn't send anything, so it follows the mutate-directly pattern, not draft-and-display). Args: `subject` (req), `body`, `due_date`, `start_date`, `reminder_time`, `importance`. |
| `update_task` | `(TaskItem)GetItemFromID`; only passed fields change; `mark_complete: true` → `Complete = true` + `PercentComplete = 100`. |
| `set_reminder` | `ReminderSet` / `ReminderTime` on an appointment **or** task, addressed by its `item_id` (EntryID); `clear: true` turns it off. |
| `set_email_reminder` | `MailItem.MarkAsTask(mapped interval)` + `TaskStartDate`/`TaskDueDate` + `ReminderSet`/`ReminderTime` + `.Save()` — the confirmed COM path for "flag an email for follow-up with a reminder". `MailItem` does expose `ReminderSet`/`ReminderTime`. |

### Undo/redo tools (2 — Draft only or higher; `Mutated` on success)

A **custom** undo/redo stack of the assistant's own actions (`OutlookTools.Undo.cs`,
stack logic in the pure, unit-tested `OfficeAi.Shared/ActionHistory.cs`), not a wrapper
around Outlook's native Undo. Native Undo (`Explorer.CommandBars.ExecuteMso("Undo")`)
was tried by hand and can't be used: it's a single slot tied to the window, it doesn't
see object-model changes, and after one use both it and the ribbon button failed with
"The operation cannot be performed because the message has changed." Outlook has no API
to put an object-model change onto that slot.

| Tool | Notes |
|---|---|
| `undo_last_action` | Reverses the assistant's most recent recorded action; call it again to step further back. Refuses if the item was changed since the assistant's action (by the user or anything else) rather than overwrite it, and never touches the user's own manual changes. A failed or refused entry is dropped, never retried. Stops at a barrier (below) instead of reaching past it. |
| `redo_last_action` | Re-applies the most recently undone action, with the same "changed since" check. The redo list is cleared by any new recorded action or barrier. |

How it works: an in-memory stack per mailbox chat (keyed like `ModeByMailbox`, capped at
50 entries, lost on restart). Each mutating handler records one entry after its change
succeeds:

- **Property snapshots** (before/after values, restored and `Save()`d):
  `mark_email_read/unread` (`UnRead`), `flag_email_important` (full `Importance`, so Low
  is preserved), `set_event_categories`, `set_event_availability` (`BusyStatus`),
  `set_reminder`, `update_task` (task fields as a group, or the flagged-mail fields),
  and `edit_event` whenever the call doesn't send or change a whole recurring series'
  time — a snapshot of only the fields that call actually touched (`Start`/`End`,
  `Subject`, `Body`, `Location`). This includes an occurrence-level edit via
  `occurrence_date`: `RecurrencePattern.GetOccurrence()` returns a real item with its
  own `EntryID` once saved, so the same mechanism covers it.
- **`set_email_reminder`**: if the message wasn't flagged before, undo calls
  `ClearTaskFlag()`, and redo calls `MarkAsTask` again.
- **Moves**: `move_email`, non-permanent `delete_email`, and whole-event `cancel_event`
  on a plain appointment or an already-canceled event (nobody to notify, so canceling
  it is just a move to Deleted Items). The folder is resolved by its own
  EntryID/StoreID via `Namespace.GetFolderFromID`, and the new EntryID after each move
  is rewritten into every entry for that item.
- **Created items**: `create_task`, and `create_event` without attendees (including a
  recurring series — deleting the master deletes the whole series). Undo moves the item
  to Deleted Items (recoverable); redo moves it back.
- **`set_category_color`**: undo restores the old color, or removes a tag the assistant
  created.
- **Barriers** (recorded, but not reversible — undo reports it can't go past them):
  every `send_*`; `create_event` with attendees; `accept/decline/tentative_meeting`;
  `delete_email permanent:true`; `edit_event` and `cancel_event` whenever they `.Send()`
  a notice (a still-active meeting the user organizes, including `edit_event`
  converting a plain event into a meeting by adding attendees for the first time);
  **any occurrence-level cancellation** via `cancel_event`'s `occurrence_date`, plain
  appointment or meeting alike (`RecurrencePattern.Exceptions`/`Exception` are
  read-only via COM, confirmed by reflection — there is no API to restore a deleted
  occurrence); and `edit_event` on a **whole recurring series whenever it includes a
  time change** (Outlook rejects `Start`/`End` on a recurring master, so that path
  writes `RecurrencePattern` fields, which the snapshot mechanism can't read/write — if
  the same call also changes other fields, the whole call stays a barrier; there is no
  partial undo of a mixed change).
- **Known gap — `set_event_categories`**: the snapshot only covers the appointment's own
  `Categories` string. If the assigned name wasn't already in the mailbox's master
  category list, Outlook auto-adds it on `Save()` with an arbitrary color; undo restores
  the appointment but does not remove that auto-created master category entry.

Draft tools never record anything — nothing is saved or sent until the user acts on the
opened window. **Not verified against a live Outlook client**, except
`set_event_availability`'s undo/redo round trip and occurrence-level reschedule/
cancellation on a plain (non-meeting) recurring series — see "Unproven at runtime".

### Draft-and-display tools (8 — Draft only or higher; open a native Outlook window for the user to review and send; `Mutated = false`)

| Tool | Notes |
|---|---|
| `draft_email` | `CreateItem(olMailItem)` → set `To`/`Subject`/`Body` → `Display(false)`. `"— Created with OpenDocs"` appended to a non-empty body (mirrors mcp-outlook's signature seed). |
| `reply_email` | `orig.Reply()` (sender only); `body` HTML-encoded and prepended above the quoted original; `Display(false)`. |
| `reply_all_email` | `orig.ReplyAll()`. A **distinct tool**, not a `reply_all` boolean on `reply_email` — clearer for the model, and the user sees the full recipient list before sending. |
| `forward_email` | `orig.Forward()`, optional `To`, `body` prepended; `Display(false)`. |
| `draft_event` | `CreateItem(olAppointmentItem)`; with `required_attendees`/`optional_attendees` → `MeetingStatus = olMeeting`, `Recipients.Add(...).Type`, `Recipients.ResolveAll()`; `Display(false)`. Same `"— Created with OpenDocs"` signature appended to a non-empty body. Optional `recurrence` object — same shape and handling as `create_event`'s (see "Recurring events" below). |
| `draft_edit_event` | Same inputs, refusals, and resolution as `edit_event` (below) — `event_id`, optional `occurrence_date`/`store_id`, at least one of `start`+`end` (together), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` — but applies every change **unsaved** and ends in `appt.Display(false)` for the user to review. Never touches `ForceUpdateToAllAttendees`, never calls `.Send()`/`.Save()`, never records undo/redo. Runs the same `CheckOccurrenceReorderCollision` pre-check before an occurrence move. |
| `draft_cancel_event` | Resolves `event_id` (optional `occurrence_date`/`store_id`) and `Display(false)`s it **unchanged** — never touches `MeetingStatus` or anything else, because Outlook persists an unsaved `MeetingStatus` change when the window closes, even without the user clicking Send Cancellation (confirmed live). The user cancels it themselves from the opened window (Cancel Meeting/Send Cancellation for a meeting they organize, Delete for a plain appointment). Refuses (`IsError`) on an already-canceled event (points at `cancel_event`, the only way to dismiss one) or a still-active meeting the user only attends (points at `draft_respond_meeting`, the Draft-tier-reachable alternative, or `decline_meeting` one tier up). |
| `draft_respond_meeting` | **Never calls `Respond()`.** Resolves `event_id` (optional `store_id`) via the shared `ResolveMeetingAppointment` helper, refuses (`IsError`) on an already-canceled meeting or one the user wasn't actually invited to (same `AlreadyCanceledRespondError`/`NotInvitedError` helpers `RespondMeeting` uses), then `appt.Display(false)`s the original item **completely unchanged** — the user picks Accept/Tentative/Decline from Outlook's own ribbon buttons. One tool rather than one per response type because it never picks an `OlMeetingResponse` up front: `AppointmentItem.Respond()` is known to commit a real calendar change at call time (a new EntryID on accept/tentative, a move to Deleted Items on decline) whether or not the response is ever sent, which would break the draft contract. Optional `message` is returned as suggested text in the tool's own output (for the user to paste via Outlook's "Edit response before sending") rather than pre-filled, since pre-filling would require calling `Respond()`. |

**These never call `.Send()` (mail) or save a calendar event.** The user sends from the
opened Outlook window.

### Meeting-response tools (3 — Automate approvals or higher; `Mutated = true`)

| Tool | Notes |
|---|---|
| `accept_meeting` / `decline_meeting` / `tentative_meeting` | One shared `RespondMeeting` helper, taking the actual `OlMeetingResponse` value. Resolves to `AppointmentItem` via `ResolveMeetingAppointment` (`MeetingItem.GetAssociatedAppointment(false)` when the id is a meeting request; optional `store_id` for shared calendars). Refuses (`IsError`) up front on an already-canceled meeting (`IsCanceledMeeting`) or an item that isn't a meeting the user was actually invited to (an organizer's own `olMeeting` copy, or a plain `olNonMeeting` appointment). Otherwise captures the subject, then `appt.Respond(olMeetingAccepted/Declined/Tentative, true, false)`, then `.Send()` on the response if non-null. Optional `message` sets `resp.Body` before `.Send()` — a short comment for the organizer (**unverified live** whether the organizer actually sees it). If `.Send()` fails, the result says so explicitly ("...but the response could not be sent to the organizer") — the local `Respond()` still went through, but the organizer was never notified. Recorded as undo barriers. Act on the whole series for a recurring meeting (no `occurrence_date`). |

### Auto-send tools (7 — Full autonomy only; send/create immediately, no review window; `Mutated = true`)

Every draft/compose tool above stays draft-and-display-only in every mode; these are
separate tools, reachable only at the top tier.

| Tool | Notes |
|---|---|
| `send_email` | Same construction as `draft_email`, but `m.Send()` instead of `m.Display(false)`. |
| `send_reply` | Same as `reply_email`, but `.Send()`. |
| `send_reply_all` | Same as `reply_all_email`, but `.Send()`. |
| `send_forward` | Same as `forward_email`, but `.Send()`; `to` is required (unlike `forward_email`, where it's optional). |
| `create_event` | Same construction as `draft_event`. No attendees → `a.Save()` (a plain calendar entry, nobody to notify). Attendees present → `MeetingStatus = olMeeting` then `a.Send()`, dispatching the invite. Both `AppointmentItem.Send()`/`.Save()` confirmed present via .NET reflection against the referenced PIA. Optional `recurrence` object builds a repeating series (see "Recurring events" below). |
| `edit_event` | General-purpose event edit. `event_id` (req), optional `occurrence_date` and `store_id`; at least one of `start`+`end` (required together), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` — else a clean `IsError` naming all seven. **Organizer-only:** refuses (`IsError`) on a canceled event (`olMeetingCanceled`/`olMeetingReceivedAndCanceled` — nothing to edit) or a meeting the user only attends (`olMeetingReceived` — points at Outlook's own "Propose New Time", see "Excluded / deferred"); these checks run against the recurring **master**, since an occurrence item from `GetOccurrence()` doesn't reliably report `MeetingStatus`. Time changes: a whole recurring series is shifted via `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` (Outlook throws `0xAF620009` on `Start`/`End` set directly on a recurring master — confirmed live); an occurrence or non-recurring event gets `appt.Start`/`.End` directly, after the `CheckOccurrenceReorderCollision` pre-check (below). `subject`/`body`/`location` are set directly (occurrence-level allowed). Attendee fields are whole-series/non-recurring only (`occurrence_date` + an attendee field is an `IsError`) and **replace the list wholesale per category** via `ReplaceAttendees` — `required_attendees` clears and re-adds only the required category, `optional_attendees` only the optional one, the organizer is never touched — then `Recipients.ResolveAll()`; read the current list with `get_event` first to keep anyone. Adding attendees to a plain event flips it to `olMeeting`. `.Send()`s if the result is or becomes a meeting — setting `ForceUpdateToAllAttendees = false` explicitly first (Outlook's own default: notify only added/removed attendees) — else `.Save()`s. A `COMException` from an occurrence `.Save()` is re-verified via `ResolveOccurrenceTarget` before being reported as failure, since Outlook can throw "Cannot save this item." after the change has already persisted. Undo: barrier when it sends or shifts a whole series' time, else a snapshot of the touched fields. |
| `cancel_event` | `event_id` (req), optional `occurrence_date` and `store_id`. Still-active organized meeting → `MeetingStatus = olMeetingCanceled` then `.Send()` (the cancellation notice), then moves the item to Deleted Items (best-effort — a failed move is reported, with the notice still sent; see "Unproven at runtime"). Plain appointment, or an **already-canceled event** (either status — the only way to dismiss one) → moves straight to Deleted Items (recoverable), nobody to notify. Refuses (`IsError`) only on a still-active meeting the user only attends (points at `draft_respond_meeting`/`decline_meeting`). Same master-based organizer checks and false-negative `.Delete()` re-verification as `edit_event`. Undo: a move entry for a whole plain/already-canceled event; a barrier for a still-active meeting (the notice already went out) and for **every** occurrence-level cancellation. |

Gated by `OutlookTools.cs`'s `SendTierTools` set, requiring `EditingMode.FullAutonomy`
exactly (the ordinal check's top tier) — not reachable from Draft only or Automate
approvals. The model is told to prefer the draft tools unless the user clearly asked for
something to go out immediately.

### Recurring events

- **Occurrence targeting.** `edit_event`/`draft_edit_event`/`cancel_event`/
  `draft_cancel_event` take an optional `occurrence_date` (a date from `list_events`'
  per-row `start`). A shared `ResolveOccurrenceTarget` helper resolves it via
  `RecurrencePattern.GetOccurrence(DateTime)` (confirmed present by reflection) into
  that one occurrence; omitting it acts on the whole series (or a non-recurring event).
  `get_event` and the meeting-response tools have no `occurrence_date` and act on the
  master.
- **Occurrence reorder pre-check.** Outlook rejects moving an occurrence past a later
  occurrence of the same series, or onto a day another occurrence already occupies —
  both confirmed live, both surfacing from automation only as a generic
  `COMException "Cannot save this item."` with no inner detail.
  `CheckOccurrenceReorderCollision` queries the calendar for other occurrences between
  the original and target dates before `.Save()` and returns an exact error naming the
  valid range instead.
- **Recurrence creation.** `create_event`/`draft_event` take an optional `recurrence`
  object: `type` (required — `daily`/`weekly`/`monthly`/`monthlyNth`/`yearly`/
  `yearlyNth`, all 6 `OlRecurrenceType` values), `interval` (default 1),
  `days_of_week`, `day_of_month`, `instance` (1-4, or 5 for "last"), `month_of_year`,
  and at most one of `count`/`until` (neither = no end date). Type-specific fields
  omitted by the caller default from `start`'s own date. Validated before any COM call
  by the pure, unit-tested `OfficeAi.Shared.RecurrenceValidator` — every failure names
  the specific field and the fix — then applied via `AppointmentItem.GetRecurrencePattern()`
  after attendee/`MeetingStatus` setup, before `.Save()`/`.Send()`.
- **Undo asymmetry.** Occurrence-level edits are undo-able; occurrence-level
  cancellations and whole-series time changes are always barriers (see "Undo/redo
  tools").

### Shared calendars

- **Reading.** `list_events`' `mailbox` param lists someone else's calendar. The
  recipient is resolved via `Ns.CreateRecipient(mailbox).Resolve()` (unresolvable →
  `IsError`), then `Ns.GetSharedDefaultFolder(recipient, olFolderCalendar)` is called
  **once, only to read `.Store.StoreID`** — never its `.Items` — and the events
  themselves come from EWS `FindAppointments` + `CalendarView`, off the UI thread.
  Enumerating the shared folder's COM `Items` froze Outlook (confirmed live, even for a
  one-day range): that folder normally isn't cached offline, so every property read
  could be a blocking round trip to Exchange on Outlook's STA UI thread. Folder-open
  failure → an `IsError` suggesting the calendar isn't shared with the user or needs
  adding via Outlook's own "Open Calendar" first; a `StoreID` read failure alone just
  omits `store_id`.
- **Output.** Same per-event text shape as the own-calendar path (via
  `OfficeAi.Shared.SharedCalendarEventFormat`, deliberately separate code from the COM
  path's inline formatting), plus `calendar_owner: <name>` — the **resolved
  recipient's display name**, not the raw `mailbox` string — and `store_id: <StoreID>`.
  The zero-results message names the owner too. One accepted difference:
  `response`/`meeting_status` show EWS's label names (`Accept`, `Meeting`,
  `Cancelled`) rather than the COM path's `Ol*` enum names — display text only,
  nothing parses it.
- **Acting on a shared-calendar event.** `get_event`, `edit_event`/`draft_edit_event`,
  `cancel_event`/`draft_cancel_event`, `accept_meeting`/`decline_meeting`/
  `tentative_meeting`/`draft_respond_meeting`, `set_event_categories`, and
  `set_event_availability` all take an optional `store_id`, passed straight through to
  `ItemById`/`GetItemFromID` — `ItemById` only searches the caller's own default store
  without it. This is a lookup aid only.
- **No authorization or redaction layer of this add-in's own** (an explicit
  project-owner decision): whatever the caller's real Exchange sharing permissions let
  them do through Outlook's own UI — view (including a private item, if their access
  exposes it), edit, cancel, respond — these tools allow too; Outlook/Exchange itself is
  the only thing that refuses. Subject/location are passed through as Outlook resolves
  them, private items included. A permission failure surfaces as whatever exception
  `.Save()`/`.Send()` naturally throws, caught by each tool's existing error handling.

### Excluded / deferred

- **`update_event` / `delete_event`** — present in mcp-outlook's `server.py` but not
  its README; not ported under those names. `edit_event` and `cancel_event` (plus
  their `draft_` versions) cover the same ground.
- **Proposing a new time on a meeting the user only attends** (`MeetingStatus ==
  olMeetingReceived`) — real Outlook's "Propose New Time" feature, deliberately
  *not* implemented. `edit_event`/`draft_edit_event` refuse outright on a received
  meeting rather than attempt anything, because there's no clean way to honor the
  request: a plain `.Send()` here wouldn't be an authoritative reschedule Outlook
  actually honors (the user isn't the organizer), and .NET reflection against the
  actually-referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA found no
  "Propose"/"Counter"/"NewTime"-named member on `AppointmentItem`/
  `_AppointmentItem` or `MeetingItem`/`_MeetingItem` to call instead —
  `_AppointmentItem.Respond` only accepts `OlMeetingResponse`
  (Accept/Decline/Tentative). "Propose New Time" appears to be a ribbon/UI-level
  feature (MAPI counter-proposal properties) not cleanly exposed through classic
  COM automation. A future implementation would likely need raw `PropertyAccessor`
  MAPI-property manipulation rather than a typed object-model call — untried here,
  and risky without a live mailbox to validate against.

### Object Model Guard

In-process VSTO add-ins that use the **VSTO-supplied `Application`** object are trusted
by default — reading `Body` / `Recipients` / `SenderEmailAddress` /
`AddressEntry.PrimarySmtpAddress`, calling `PropertyAccessor.GetProperty` or
`Attachment.SaveAsFile` do **not** raise the "a program is trying to access…" prompt on
default settings. Prompts appear only under Trust Center → Programmatic Access set to
"Always warn", or an Exchange public-folder security form. `MailItem.Send` — the
highest-risk call — is used only by the Full-autonomy auto-send tools (`send_email`/
`send_reply`/`send_reply_all`/`send_forward`); whether it triggers a guard prompt
under a stricter Trust Center setting has not been checked live.

### Structural fragility

Everything is addressed by `EntryID`. It is stable while an item stays put but changes
on `Move` and is store-specific; the mutating tools that move items return the new id,
and every other tool re-resolves via `GetItemFromID` each call (with `store_id` for a
shared-calendar item). All recurring occurrences of a calendar series **share one
`EntryID`**, so `get_event` / `accept_meeting` / `decline_meeting` / `tentative_meeting`
/ `draft_respond_meeting` cannot target a single occurrence — they act on the master
series. The calendar-editing tools can, via `occurrence_date` (see "Recurring events");
without it they too act on the whole series. `list_events` carries each occurrence's
`start` as the disambiguator.

**Known limitation, confirmed live (Gmail-connected calendar):** on a mailbox connected
via Google's Gmail/Google Workspace sync, moving a calendar item to Deleted Items does
not appear to be durable the way it is on Exchange — a `cancel_event` call that
reported success moving an item to Deleted Items was followed immediately (no other
action in between) by that item relocating itself to a `Drafts` folder, with no code in
this add-in touching it a second time. Most likely Google Calendar's own sync
reconciling the move shortly after, outside this add-in's control.
`undo_last_action`'s own "did the item change since?" conflict check caught the
mismatch and refused rather than guessing or overwriting — the designed safety net
worked, and no data was lost — but this means `RecordMove`'s core assumption (an item
stays wherever the last recorded move put it, until this add-in moves it again) does
not reliably hold for Gmail-connected calendars specifically. Treated as a known
account-type-specific risk, not a code bug in `cancel_event`, `edit_event`, or the undo
stack.

### Unproven at runtime

The project builds clean (MSBuild Debug + Release, 0 warnings). But **most COM paths
have not been exercised in a real Outlook** — the Explorer pane lifecycle /
IUnknown-identity keying, WebView2 rendering inside an Explorer task pane,
`Folder.GetTable` column names, the `list_events` recurrence ordering,
`Recipient.FreeBusy`'s string format / month coverage (`find_meeting_slots`), and
several enum-name / method-signature assumptions (`olEmbeddeditem` casing,
`AppointmentItem.Respond` argument types, `MailItem.MarkAsTask`) compiled against the
interop assembly but are not yet confirmed live. Early manual testing via the mock
server exercised `list_emails`, `search_emails`, `list_folders`, `list_events`, and
`list_tasks` against a real mailbox successfully. **Verified live:** the color-tag
tools (category enumeration, tag creation, event `Categories` round-trip),
`set_event_availability` (all 5 values, no attendee notification on an organized
meeting) and its undo/redo, and the recurring-series items listed below.

**EWS paths.** `search_contacts`'s EWS path is **compiled but not confirmed against a
live on-prem Exchange**: the `ResolveName` calls and their `NameResolution` mapping,
`ExchangeService.UseDefaultCredentials` (Windows Integrated Auth to on-prem CAS),
`Account.AutoDiscoverXml` shape / `EwsAutodiscoverXml.ParseEwsUrl`, the
`AutodiscoverUrl` fallback, and the off-thread `Task.Run` actually keeping Outlook
responsive during the call. The async delegate refactor (`ToolExecutor` →
`Task<ToolResult>`, `async void OnWebMessageReceived`) builds clean for all four
add-ins; a runtime smoke of one tool per app confirming nothing regressed is still
pending. For `list_events`' shared-calendar EWS path, one assumption still needs live
confirmation: that the single remaining `GetSharedDefaultFolder` call (used only for
`StoreID`) was never itself part of the freeze — if Outlook still freezes, that call is
the next suspect (see `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md`).

**Shared calendars — permission tiers untested.** Behavior has not been exercised
against a real second mailbox with different Exchange sharing levels (Full Access /
Editor / Reviewer / free-busy only / titles+locations only / no access): neither the
exact read-path error messages, nor whether partial-read tiers degrade gracefully, nor
what a genuinely unpermitted cross-mailbox write (e.g. `edit_event` with a `store_id`
from a Reviewer-only calendar) actually returns — assumed to be a `COMException` from
`.Save()`/`.Send()`, untested.

**`edit_event`/`draft_edit_event`** — compiled, not exercised live for their meeting
paths. Ranked by consequence of being wrong (per
`docs/superpowers/specs/2026-09-28-outlook-edit-event-design.md`):

1. **`ReplaceAttendees`'s recipient-clearing** — highest risk, since a wrong
   organizer-exclusion could drop the user's own organizer entry or fail to actually
   clear old attendees, feeding directly into a `.Send()` with real people.
2. **Converting a plain event into a meeting via `edit_event`** — adding attendees to an
   *existing*, previously-saved item (unlike `create_event`, which only ever adds
   attendees to a brand-new one) is an untested combination.
3. **`subject`/`body`/`location` edits on a recurring master** — likely fine (none of
   these are part of `RecurrencePattern`), but unverified.

Also unverified: whether Outlook accepts a plain `.Send()` on a modified organizer-owned
meeting as a real update notice in every case (e.g. an all-day flag change). **Open
question:** whether `draft_edit_event`'s unsaved changes (mutate an existing item,
`Display(false)`, never `.Save()`/`.Send()`) have the same "persists on window-close
without an explicit action" risk that `MeetingStatus` turned out to have for
`draft_cancel_event`. If so, closing a `draft_edit_event` window on a meeting the user
organizes without clicking Send could silently move it in the organizer's own calendar
while attendees still see the old time. Not confirmed either way.

**`cancel_event`** — its organized-meeting branch calls `.Move()` immediately after
`.Send()` on the same item, a sequence no other tool here performs. Whether Outlook still
permits relocating a just-canceled-and-sent appointment is unconfirmed; the code catches
a failure there and reports the (genuinely-sent) cancellation succeeded regardless,
telling the user to delete the stray calendar entry manually if the move didn't take.

**Recurring series — partly live-verified.** Confirmed live, against a plain
(non-meeting) recurring series: occurrence resolution via
`ResolveOccurrenceTarget`/`RecurrencePattern.GetOccurrence`, occurrence-level reschedule
(including its snapshot-based undo), occurrence-level cancellation (as a barrier), the
false-negative `.Save()`/`.Delete()` re-verification, the two occurrence-reordering
rejections `CheckOccurrenceReorderCollision` guards against, and that `Start`/`End`
cannot be set directly on a recurring master. Still unconfirmed:

- **The master-vs-occurrence organizer-authority check has never been exercised against
  a genuine recurring *meeting* occurrence with attendees** — every live test used a
  plain series. Checking `master.MeetingStatus`/`IsCanceledMeeting(master)`/
  `IsReceivedMeeting(master)` instead of the occurrence's own value is a reasoned fix
  based on `GetOccurrence()`'s known-unreliable behavior for this property, not one
  confirmed by observing the bug on a real meeting occurrence.
- **The whole-series `RecurrencePattern` time change**: whether
  `PatternStartDate`/`StartTime`/`EndTime` interact correctly with the series' existing
  `DayOfWeekMask`/other pattern fields when only the time (not the day) changes, whether
  `.Send()`/`.Save()` on the master after mutating the pattern behaves identically to
  mutating it directly, and what the barrier's undo-refusal messaging looks like in
  practice.
- **`CheckOccurrenceReorderCollision` may not cover every reason Outlook can reject an
  occurrence reorder** — any rejection reason it doesn't anticipate still reaches the
  model as the same generic "Cannot save this item."
- **`recurrence` on `create_event`/`draft_event`**: `RecurrenceValidator`'s unit tests
  cover only the pure validation logic. The COM application path —
  `GetRecurrencePattern()` after attendee/`MeetingStatus` setup, before
  `.Save()`/`.Send()` — is unverified live, as is whether that ordering matters.

**Meeting responses.** Whether the organizer actually sees `accept_meeting`/
`decline_meeting`/`tentative_meeting`'s `message` text on the delivered response (vs. it
being dropped or overwritten by Outlook's own response template) has not been confirmed
against a real received invite. `draft_respond_meeting`'s `message` has no such gap —
it's plain suggested text in the tool output. Whether `AppointmentItem.Respond()` itself
commits local calendar changes at call time remains formally unconfirmed, but is moot
for the draft path, which never calls it.

---

## Explicitly out of scope everywhere (per project scope, not gaps)

This scope boundary originated in the project's original feasibility report and the
toolset-port plan's Global Constraints (previously stated only in the now-retired
`docs/tool-surface-todo.md`'s header — see the changelog).

- `web_search`, `image_search`, `generate_image`, `analyze_media` — no `ai-search`
  equivalent; air-gapped deployment target. (`get_attachment` is a partial exception
  for Outlook — it reads local + OpenXML attachment text, nothing remote, no
  PDF/images; see the Outlook section.)
- The PDF app and the Markdown app have no officeoffice counterpart (Markdown's
  scope is folded into Word).
- PowerPoint's `execute_slide_script` DSL and entire deck-generation pipeline (see
  the PowerPoint section's "Missing entirely" list).

`delete_slide` was originally miscategorized as out of scope (bundled with the DSL and
generation pipeline in the retired `tool-surface-todo.md`) but is a small, in-scope fix
with no dependency on the DSL (`add_slide` already existed; deleting a slide needs no
scripting language). It shipped independently (PP-19 Task 1) and is documented in the
PowerPoint section above as a normal, implemented tool — this inconsistency is now
resolved, not merely flagged.

---

## Summary: what genoffice has that officeoffice doesn't

1. **Web-sourced content** — no search, no AI image generation, no media analysis,
   no chat-attachment reading, anywhere. Image tools are local-file-only by design.
2. **PowerPoint's scripting DSL and generation pipeline** — no `execute_slide_script`,
   no `generate_deck`/`regenerate_slide`, no automatic QC/audit pass, no `dataSource`
   provenance enforcement on charts (Excel or PowerPoint).
3. **Block-indexed document context** — genoffice's `get_document_context`/
   `get_deck_context` return structured per-block/per-element inventories; Word's and
   PowerPoint's officeoffice equivalents return flat text previews only.
4. **`add_shape`'s preset breadth** in Excel/PowerPoint (26 named presets each) is
   still narrower than genoffice's "full OOXML preset-geometry set", and
   `set_data_validation`'s `checkbox` kind is unsupported (no COM equivalent exists).

Previously listed here and **since fixed** (kept for context, not as an open list):
live multi-provider selection (PP-6/FT-1 — the app-shell settings screen wires
provider/model/key end to end now), Word's `insert_content`/`read_blocks`/
`replace_blocks` HTML support and positional insert (PP-10), Word's missing image tool
(PP-11, `add_image`), Word's `edit_chart` single-series/no-categories limitation
(PP-9), Word's `updateTextStyle` missing `highlight` (PP-12), Excel's narrow
`format_range` and `add_chart` (PP-13/PP-15), and PowerPoint's narrow
`set_element_style`/`add_shape` (PP-20).

## Summary: what officeoffice has that genoffice doesn't

1. **Real Word comments** (`add_comment`) — anchored, native, available in every
   editing mode including Comment Only. genoffice's docs surface has no comment tool.
2. **A real native `TablesOfContents` TOC** on `insertToc`, vs. genoffice's hand-built
   TOC field-XML workaround (needed because genoffice's own renderer doesn't
   paginate).
3. **A native error-cell scan** (`find_cells` with `errors_only`, via
   `SpecialCells(xlErrors)`) — a categorical COM-vs-Office.js/web advantage called out
   explicitly in the source.
4. **Server-enforced editing modes** (Read Only / Comment Only / Track Changes / Full
   Autonomy, or Outlook's four relabeled equivalents) as a first-class, uniformly-applied
   gate across all four apps' tool dispatch — genoffice's docs app has Track-Changes-aware
   writes but no equivalently formal, uniform mode-gating system across its apps.
5. **Live selection-push into context** (Word/Excel/PowerPoint/Outlook) — a real Office
   selection event pushes the user's current selection into `buildContext()`
   automatically on every turn, rather than app-side selection-range plumbing.
6. **`add_pivot` with calculated fields** — Excel's pivot op supports
   `PivotTable.CalculatedFields().Add(name, formula)` for formula-derived pivot
   values, in addition to row/column/page/data fields.
7. **PowerPoint's Slide Master tooling** (`add_master_element`/`read_master_elements`/
   `set_master_element_transform`/`remove_master_element`/`set_headers_footers`/
   `list_layouts`) and its native-clipboard cross-slide shape copy/move
   (`copy_element`/`move_element`) — both areas with no genoffice equivalent at all.

---

## Schema-vs-implementation audit

A full schema-vs-handler audit was performed across all three add-ins' gateway tools
(Word's `apply_commands`, Excel's `propose_operations`, PowerPoint's per-tool schemas),
checking whether the JSON Schema advertised to the LLM actually matched what the C#
handler read and did. It found two recurring bug classes — **silent no-op with false
success** (an out-of-range value does nothing useful but the tool still reports
success) and **undocumented schema** (the wire-level schema is a bare `{type:
'object'}`/`{type: 'string'}` with the real contract living only in prose).

**Every finding from that audit has since been fixed by its own follow-up item** — PP-5
(structural per-kind schemas for both gateway tools) and PP-9/PP-12 for Word, PP-13
through PP-18 for Excel, PP-21/PP-22 for PowerPoint. The current tool tables above
already reflect the fixed state (e.g. Word's `highlight` and `bulletPreset`, Excel's
conditional-formatting operators and shape-type validation, PowerPoint's chart-type and
legend-position enums). The full original audit tables, findings, and fix-by-fix detail
are preserved in **`docs/ai-tool-surface-changelog.md`** for reference — if a new
schema/handler mismatch is found, record it there as a new dated entry (or open a table
here if a *current*, unfixed one is found) rather than treating this section as an
open list.

---

## Changelog

See **`docs/ai-tool-surface-changelog.md`** for the full dated history behind this
document: every bug fix, phased rollout, and the original schema-vs-implementation
audit findings.