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

### Top-level tools (17)

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
| `edit_table` | Write | `kind`: `set_cell`, `insert_row`/`delete_row`/`insert_col`/`delete_col`, `set_style` (styleName/headerRow/bandedRows/borders/borderColor), `set_shading` (cell/row/col/table scope, hex color). Structural edits shift later indices — re-read before a second structural edit in the same run. |
| `add_smartart` | Write | `layout`: list/process/cycle/hierarchy/pyramid/matrix/venn. Flat item list, one per top-level node. |
| `edit_smartart` | Write | `kind`: `set_text`, `add_node`, `delete_node`, `set_style` (colorName/quickStyleName, substring-matched against the install's real gallery), `set_layout`. `delete_node` shifts later node indices. |
| `edit_chart` | Write — gateway | Creates or edits a native Word chart. Supports **categories + one or more named series** (multi-series, chart-type selection: column/columnStacked/bar/barStacked/line/area/pie/doughnut) — no longer single-series/title-only (fixed by PP-9). `chartIndex` addresses an existing chart (0-based, inline then floating shapes, document order); `create:true` always adds a new one. `afterBlockIndex` anchors a new chart inline instead of floating at document origin. |
| `apply_commands` | Write — gateway, see below | |
| `add_comment` | Write (every mode) | Anchors a real Word comment to the first match of given text. **Not in genoffice's docs surface at all.** Available in every editing mode, including Comment Only. |

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

### Tools (49 total, all genuinely implemented)

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
| `edit_table_cell` | Replaces one cell's text. |
| `edit_table_structure` | Insert/delete row or column; bounds-checked (an out-of-range index errors specifically rather than a raw COM error). Shifts later indices. |
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
it acts as the signed-in user. **One exception:** `search_contacts` makes EWS calls
(see below) — the COM object model cannot do a multi-result directory search.

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
  its own meaning: `ReadOnly` → **Read only** (unchanged); `CommentOnly` → **Draft
  only** (every mutating/drafting tool that never leaves the mailbox unreviewed —
  `mark_email_read`/`unread`, `flag_email_important`, `move_email`, `delete_email`,
  `create_task`, `update_task`, `set_reminder`, `set_email_reminder`, `draft_email`,
  `reply_email`, `reply_all_email`, `forward_email`, `draft_event`); `TrackChanges` →
  **Automate approvals** (adds `accept_meeting`/`decline_meeting` — these already call
  `resp.Send()` to notify the organizer, so they get their own explicit tier rather
  than hiding in Draft only or Full autonomy); `FullAutonomy` → unchanged name, adds
  the five auto-send tools (see below). Each tier is a strict superset of the one
  before it. `entry.ts` passes `availableModes: ['readOnly','commentOnly',
  'trackChanges','fullAutonomy']` and a `modeOverrides` map so the mode menu shows
  Outlook's own labels/descriptions instead of Word/Excel/PowerPoint's generic
  "Comment only"/"Track changes" text — those two apps' own use of `CommentOnly`/
  `TrackChanges` for real editing is untouched (Outlook's relabeling is purely its own
  `entry.ts` config, not a shared-string change). Gate logic (`OutlookTools.cs`'s
  `ExecuteAsync`) is an ordinal check on the enum's own declared order (`ReadOnly <
  CommentOnly < TrackChanges < FullAutonomy`). A fresh session defaults to **Draft
  only**, not Full autonomy (`defaultMode: 'commentOnly'`, resolved via `chat-ui.ts`'s
  `defaultModeFor()` — the same helper Word/Excel/PowerPoint use to default to their
  own `trackChanges`).
- **Selection context.** The Explorer's `SelectionChange` pushes the selected mail
  item(s) / conversation into per-turn context as a `mail` `SelectionContext` variant
  (`shared/web-src/app-shell/bootstrap.ts`) carrying each item's `EntryID`, so the
  model prefers the selection over searching. The scope pill shows the subject (one
  message) or a count (`scopeUnit: 'mailbox'`).
- **`message_id` / `event_id` / `task_id` = Outlook `EntryID`**, re-resolved via
  `Namespace.GetItemFromID` on every call. EntryID is stable within a folder but
  **changes on `Move` and across stores** — `move_email` / `delete_email` return the
  new id. Recurring calendar instances all share the master appointment's EntryID.
- **Native query APIs are mandatory** (see the slow Word-search incident in the
  changelog — the same class of bug, avoided here from the start): `list_*` use
  `Folder.GetTable` (an in-memory rowset, no per-item COM object); `search_emails` uses
  `Items.Sort` → `Items.Restrict("@SQL=" + DASL)`; a capped linear scan is a fallback
  only when `Restrict` rejects the filter. The DASL builder is pure and unit-tested
  (`OfficeAi.Shared/OutlookDasl.cs`). `search_contacts` is the deliberate, documented
  exception: server-side EWS ANR (`ResolveName`), not a COM scan — see the EWS
  carve-out below.
- **`search_contacts` and `find_meeting_slots`' work-week default are the two things
  that touch EWS.** Every other Outlook tool is pure `Microsoft.Office.Interop.Outlook`
  COM against the running client. All EWS-dependent code lives apart from the pure-COM
  tool files, in two layers: `OutlookEws.cs` (the raw EWS Managed API wire calls —
  `ResolveNamesAsync`, `GetWorkingHoursAsync`) and `OutlookTools.Ews.cs` (the
  tool-facing orchestration on top: `SearchContactsAsync`, `ResolveWorkWeekAsync`, and
  the shared endpoint/account resolution both call, `ResolveEwsUrlAsync`/
  `FindExchangeAccountInfo`). Contact resolution calls **EWS `ResolveName(query,
  ContactsThenDirectory, returnContactDetails: true)`** (EWS Managed API 2.2,
  `Microsoft.Exchange.WebServices` 2.2.0) with `ExchangeService.UseDefaultCredentials`
  (Windows Integrated Auth as the signed-in user — the .NET equivalent of
  `mcp-outlook`'s `auth_type=sspi`; no stored credentials). Endpoint is parsed from the
  cached `Outlook.Account.AutoDiscoverXml` (`<EwsUrl>`/`<ASUrl>`, `EXCH` preferred over
  `EXPR`; pure parser `OfficeAi.Shared/EwsAutodiscoverXml.cs`), falling back to
  `ExchangeService.AutodiscoverUrl`, then cached in a process-static `Uri` shared by
  both EWS-dependent tools. The call runs off the UI thread (`await Task.Run`,
  `svc.Timeout` 15 s) so Outlook stays responsive. **On-prem Exchange only.** EWS
  unreachable / SSPI failure / endpoint not found / timeout → a clear `IsError` result
  for `search_contacts` (EWS isn't optional there); `find_meeting_slots` instead falls
  back to its old hardcoded Sun-Thu/9-18 default, since EWS is an enhancement over an
  already-working COM-only path there, not the only way to do the job. This is also the
  pilot for async tool execution — the shared `ToolExecutor` delegate is
  `Task<ToolResult>`-returning (`WebViewBridgeHost.OnWebMessageReceived` is `async`).
  `find_meeting_slots` is the second Outlook tool to go async, for its own EWS
  work-week lookup — see its row below; every other tool in all four add-ins is still
  synchronous, wrapped in `Task.FromResult`.

### Read tools (13 — always allowed, never gated)

| Tool | Notes |
|---|---|
| `list_emails` | `Folder.GetTable`; columns EntryID/Subject/ReceivedTime/SenderName/UnRead + `PR_HASATTACH` proptag; `[UnRead] = true` restriction when `unread_only`; sorted newest-first; non-mail rows filtered by `MessageClass` not starting `IPM.Note`. Args: `folder`, `limit` (20), `unread_only`. |
| `search_emails` | `Items.Sort("[ReceivedTime]")` then `Restrict("@SQL=" + BuildSearchFilter(...))`. `LIKE '%q%'` on subject + body; UTC-ISO date range; sender by `fromemail =` OR `fromname LIKE` (Exchange senders carry `legacyExchangeDN`, not SMTP — display-name fuzzy match). `ci_phrasematch`/`ci_startswith` are **not** usable via `Restrict` (they throw). `recipient` is a client-side filter, capped 500. Fallback: capped linear scan on a malformed filter. |
| `apply_search` | View-only (`Mutated: false`) — pushes the same `BuildSearchDasl` filter `search_emails` computes into `Explorer.CurrentFolder` + `Explorer.Search("@SQL=" + dasl, olSearchScopeCurrentFolder)`, so the user's own Outlook window shows the results. `_Explorer.Search(string Query, OlSearchScope SearchScope)`'s signature was confirmed via .NET reflection against the referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA; the `@SQL=` DASL string itself has **not yet been exercised against a live Outlook session** — first real use should confirm it's accepted as-is by `Explorer.Search`, or fall back to plain-text `query` only. |
| `get_email` | Full `Body` (≤ 40k), To/CC via `Recipient.Type`, `ConversationID`/`ConversationTopic`, importance, unread, and an `attachments` array — `{index (1-based), name, type (byValue/embeddedItem/ole/reference), size}` — feed the index to `get_attachment`. |
| `open_email` | `MailItem.Display(false)` on an existing item resolved via `ItemById` — opens the message in its own Outlook reading window, unmodified. `Mutated: false`. |
| `get_attachment` | `Attachment.SaveAsFile` into `%LOCALAPPDATA%\OutlookAiAddIn\Attachments\`; returns the path. `extracted_text` (≤ 40k) is populated **only** for text-family extensions (`.txt .csv .tsv .md .json .xml .log`, `.html` tag-stripped) and OpenXML (`.docx .xlsx .pptx`), via the swappable `OfficeAi.Shared/AttachmentText/` module (`DocumentFormat.OpenXml` 2.20.0). **No PDF, no images, no vision** — those return the path + type only. `olOLE` throws (rejected); `olByReference` has no data (rejected); `olEmbeddedItem` saves as `.msg`. |
| `list_folders` | Recursive walk of every store's `Folders`, mail folders only (`DefaultItemType == olMailItem`), with item + unread counts; capped ~800 / depth 8. |
| `search_contacts` | **EWS `ResolveName` over Contacts then GAL** (server-side ANR), run off the UI thread via EWS Managed API 2.2 with `UseDefaultCredentials`. Endpoint discovery: `Account.AutoDiscoverXml` → `AutodiscoverUrl` → process-static `Uri` cache. Each `NameResolution` mapped to `(name, email)` — GAL X500/`EX` addresses fall back to the resolved contact's own `EmailAddress1..3`; entries with no `@` address are dropped (mirrors `mcp-outlook`). Deduped by lowercased address (name fallback), capped at `limit` (default 10). Pure helpers `EwsAutodiscoverXml.ParseEwsUrl` + `ContactSearchFormat.Format` are unit-tested. **No `folder`/scope arg.** On-prem Exchange only; unreachable / auth failure / no endpoint / 15 s timeout → a specific `IsError` message, no COM fallback. (The pre-existing recursive multi-store contact-folder crawl froze then crashed Outlook — removed in favor of EWS.) |
| `list_events` | `Items.Sort("[Start]")` → `Items.IncludeRecurrences = true` → `Items.Restrict("[Start] <= end AND [End] >= start")` — **this order is load-bearing** and rules out `GetTable`. Recurring instances share the master `event_id`; each row carries its own `start` to disambiguate. Args: `start_date` (today), `end_date` (+7d), `limit` (50). |
| `get_event` | `Body` (≤ 40k), `RequiredAttendees`/`OptionalAttendees`, organizer, response status, recurring flag. |
| `find_meeting_slots` | `Recipient.FreeBusy(anchor, 30, true)` — a per-30-min status string — for `Namespace.CurrentUser` + each resolved attendee; then `OfficeAi.Shared.MeetingSlots.Rank` (pure, unit-tested) slides a `duration_minutes` window in 30-min steps across each work day's `[start_hour, end_hour)` and scores each candidate by how many people are free (so a best partial match still comes back). **Work week/hours are read from the mailbox's own EWS `GetUserAvailability` → `AttendeeAvailability.WorkingHours`, not hardcoded** — falls back to Sun–Thu 09:00–18:00 only if that call fails (non-Exchange profile, EWS unreachable, etc.), cached per process like `OutlookEws.CachedUrl`. Default range is today through the end of the current contiguous work-day run (generalizes the old "today→Thursday, or next week if Fri/Sat" to any work-days shape), max 28 days. Times past the returned free/busy window are assumed free. Async (like `search_contacts`) only because of the EWS work-week lookup; the FreeBusy/ranking work itself is still synchronous COM. Args: `attendees` (req), `duration_minutes` (req), `start_date`, `end_date`, `start_hour`, `end_hour` (both default to the resolved work hours, or 9/18 as a last resort), `limit` (5). |
| `list_tasks` | `Folder.GetTable` over the default Tasks folder; open tasks only unless `include_completed`. Columns EntryID/Subject/Due/Start/Status/PercentComplete/Complete/ReminderTime. |
| `list_color_categories` | `Namespace.Categories` — the profile's master color-tag ("Category") list shared by mail/calendar/tasks, same list Outlook's Categorize picker shows. Each entry: `{name, color}`; color is one of the 26 `OlCategoryColor` values (None/Red/Orange/…/Dark Maroon), mapped to a friendly display name in `OutlookTools.Categories.cs` (not in `OfficeAi.Shared` — that project doesn't reference the Outlook PIA, same split as `ColorUtil`). |

### Mutating tools (13 — Full autonomy only; `Mutated = true`)

| Tool | Notes |
|---|---|
| `mark_email_read` / `mark_email_unread` | `MailItem.UnRead` + `.Save()`. |
| `flag_email_important` | `Importance = olImportanceHigh/Normal` + `.Save()`. `important` defaults true. |
| `move_email` | `MailItem.Move(ResolveFolder(destination))`; returns `{message_id: <new EntryID>, old_message_id}`. |
| `delete_email` | Non-permanent → `Move` to Deleted Items (returns new id); `permanent: true` → then `.Delete()` from there (no single-call hard delete in the OM — documented as "may still be server-recoverable"). |
| `accept_meeting` / `decline_meeting` | Resolves to `AppointmentItem` (via `MeetingItem.GetAssociatedAppointment(false)` when the id is a meeting request), `appt.Respond(olMeetingAccepted/Declined, true, false)`, then `.Send()` on the response if non-null. |
| `set_event_categories` | `AppointmentItem.Categories` (comma-separated tag names, the color shown on the event in the calendar grid) + `.Save()`; empty/omitted `categories` clears all tags. A name outside the master list is auto-added by Outlook on `Save` with an arbitrary color — call `set_category_color` first to control it. |
| `set_category_color` | `Namespace.Categories[name]` — updates `.Color` if the tag exists, else `Categories.Add(name, color)` creates it. Same master list `list_color_categories` reads. |
| `create_task` | `Application.CreateItem(olTaskItem)` + `.Save()` — no window (a task doesn't send anything, so it follows the mutate-directly pattern, not draft-and-display). Args: `subject` (req), `body`, `due_date`, `start_date`, `reminder_time`, `importance`. |
| `update_task` | `(TaskItem)GetItemFromID`; only passed fields change; `mark_complete: true` → `Complete = true` + `PercentComplete = 100`. |
| `set_reminder` | `ReminderSet` / `ReminderTime` on an appointment **or** task, addressed by its `item_id` (EntryID); `clear: true` turns it off. |
| `set_email_reminder` | `MailItem.MarkAsTask(mapped interval)` + `TaskStartDate`/`TaskDueDate` + `ReminderSet`/`ReminderTime` + `.Save()` — the confirmed COM path for "flag an email for follow-up with a reminder". `MailItem` does expose `ReminderSet`/`ReminderTime`. |

### Draft-and-display tools (5 — Draft only or higher; open a native Outlook window for the user to review and send; `Mutated = false`)

| Tool | Notes |
|---|---|
| `draft_email` | `CreateItem(olMailItem)` → set `To`/`Subject`/`Body` → `Display(false)`. `"— Created with OpenDocs"` appended to a non-empty body (mirrors mcp-outlook's signature seed). |
| `reply_email` | `orig.Reply()` (sender only); `body` HTML-encoded and prepended above the quoted original; `Display(false)`. |
| `reply_all_email` | `orig.ReplyAll()`. A **distinct tool**, not a `reply_all` boolean on `reply_email` — clearer for the model, and the user sees the full recipient list before sending. |
| `forward_email` | `orig.Forward()`, optional `To`, `body` prepended; `Display(false)`. |
| `draft_event` | `CreateItem(olAppointmentItem)`; with `required_attendees`/`optional_attendees` → `MeetingStatus = olMeeting`, `Recipients.Add(...).Type`, `Recipients.ResolveAll()`; `Display(false)`. Same `"— Created with OpenDocs"` signature appended to a non-empty body. |

**These never call `.Send()` (mail) or save a calendar event.** The user sends from the
opened Outlook window. Available from Draft only mode upward (tier 2 — see "Editing
modes" above), not gated behind Full autonomy.

### Auto-send tools (5 — Full autonomy only; send/create immediately, no review window; `Mutated = true`)

Reverses part of an earlier "no auto-send tool" decision — narrowed to Full autonomy
specifically, not removed generally. Every draft/compose tool above is unaffected and
stays draft-and-display-only in every mode.

| Tool | Notes |
|---|---|
| `send_email` | Same construction as `draft_email`, but `m.Send()` instead of `m.Display(false)`. |
| `send_reply` | Same as `reply_email`, but `.Send()`. |
| `send_reply_all` | Same as `reply_all_email`, but `.Send()`. |
| `send_forward` | Same as `forward_email`, but `.Send()`; `to` is required (unlike `forward_email`, where it's optional). |
| `create_event` | Same construction as `draft_event`. No attendees → `a.Save()` (a plain calendar entry, nobody to notify). Attendees present → `MeetingStatus = olMeeting` then `a.Send()`, dispatching the invite. Both `AppointmentItem.Send()`/`.Save()` confirmed present via .NET reflection against the referenced PIA before writing this — not assumed from `draft_event`'s non-sending shape. |

Gated by `OutlookTools.cs`'s `SendTierTools` set, requiring `EditingMode.FullAutonomy`
exactly (the ordinal check's top tier) — not reachable from Draft only or Automate
approvals. `accept_meeting`/`decline_meeting` are **not** in this table; they live one
tier down, in Automate approvals (see "Editing modes" above), since they already
existed before this addition and already call `resp.Send()`.

### Excluded / deferred

- **`update_event` / `delete_event`** — present in mcp-outlook's `server.py` but not
  its README; not ported. Trivial parity adds if wanted.

`Recipient.FreeBusy` is all local-time, so no cross-timezone math is needed in
`find_meeting_slots`; the ranking is a pure, unit-tested helper.

### Object Model Guard

In-process VSTO add-ins that use the **VSTO-supplied `Application`** object are trusted
by default — reading `Body` / `Recipients` / `SenderEmailAddress` /
`AddressEntry.PrimarySmtpAddress`, calling `PropertyAccessor.GetProperty` or
`Attachment.SaveAsFile` do **not** raise the "a program is trying to access…" prompt on
default settings. Prompts appear only under Trust Center → Programmatic Access set to
"Always warn", or an Exchange public-folder security form. `MailItem.Send` — the
highest-risk call — is never used.

### Structural fragility

Everything is addressed by `EntryID`. It is stable while an item stays put but changes
on `Move` and is store-specific; the mutating tools that move items return the new id,
and every other tool re-resolves via `GetItemFromID` each call. All recurring
occurrences of a calendar series **share one `EntryID`**, so `get_event` /
`accept_meeting` / `decline_meeting` cannot target a single occurrence unambiguously —
`list_events` carries each occurrence's `start` as the disambiguator.

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
`list_tasks` against a real mailbox successfully.

`search_contacts`'s EWS path is likewise **compiled but not confirmed against a live
on-prem Exchange**: the `ResolveName` call and its `NameResolution` mapping,
`ExchangeService.UseDefaultCredentials` (Windows Integrated Auth to on-prem CAS),
`Account.AutoDiscoverXml` shape / `EwsAutodiscoverXml.ParseEwsUrl`, the
`AutodiscoverUrl` fallback, and the off-thread `Task.Run` actually keeping Outlook
responsive during the call. The async delegate refactor (`ToolExecutor` →
`Task<ToolResult>`, `async void OnWebMessageReceived`) builds clean for all four
add-ins; a runtime smoke of one tool per app confirming nothing regressed is still
pending.

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
