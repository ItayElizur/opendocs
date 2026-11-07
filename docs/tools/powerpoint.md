# PowerPoint tools (`PowerPointAiAddIn/PowerPointTools*.cs`)

Current-state reference for every tool the AI can call in PowerPoint. See
[`docs/architecture.md`](../architecture.md) for the shared editing-mode/transport
architecture.

## Tools (25 total, all genuinely implemented)

The largest tool surface of the four apps. Organized here by area rather than as one
flat table.

**Reading (4, always allowed, never gated — `READER_TOOLS` in `tools.ts`)**

| Tool | Notes |
|---|---|
| `get_deck_context` | One-line-per-slide outline: slide index + text preview of its shapes. No per-element type/id inventory. |
| `read_slide` | `view:"slide"` (default) — Full text of every shape on one slide, plus its layout, transition, animation count, and speaker notes. Shapes listed back-to-front (z-order). A group is shown as one line with a child count — see `read_slide (view=group)`.; `view:"group"` — Lists a group's children recursively (nested groups expanded in place), each prefixed with a dotted path (e.g. `"3.1.0"`) usable as `shapeIndex` for `set_element_text`/`format_element`. Positional/structural edits still require `group_element` with `action:"ungroup"` first.; `view:"animations"` — A slide's animations in play order (shape, effect, entrance/exit, trigger, timing) — `animationIndex` in `edit_animation` addresses this same order.; `view:"smartart"` — Node text of SmartArt diagrams on a slide; omitting `smartArtIndex` reads every diagram on the slide. |
| `find_text` | Read-only substring/regex search across every slide's shape text (text boxes, placeholders, table cells, SmartArt node text) and speaker notes. Returns `[slide i, shape j] text` or `[slide i, notes] text`. |
| `read_master` | `view:"elements"` (default) — Shapes placed directly on the Slide Master, or one specific layout's own shapes (`layoutName`, substring-matched).; `view:"layouts"` — Every layout name in the presentation (across every design/theme), plus how many slides use each — call before passing `layoutName` elsewhere instead of guessing. |

**Slide management**

| Tool | Notes |
|---|---|
| `manage_slide` | `action` is `add` (clone the slide's layout right after it with every text frame cleared), `duplicate` (full copy right after it), `move` (to 0-based `toIndex`), or `delete` (one slide per call; cannot delete the last remaining slide; later slides shift down — re-read before another index-based edit in the same run). |
| `set_slide` | One change per call (giving fields from more than one group errors): `layout` (16 classic `PpSlideLayout` values) or `layoutName` (free text matched against the deck's own theme layouts; give one, not both); `background` (solid color; `slideIndex:-1` applies to every slide, the only change that allows it); `notes` (replaces the speaker notes); or the entry transition (`effect` — 31 curated values, `durationSeconds`, `advanceOnClick`, `advanceAfterSeconds`; the transition fields count as one change, the two advance options are independent toggles and can change without restating `effect`). One edit per call keeps each call one logical change and avoids half-applied slides. |
| `set_headers_footers` | Slide number/footer/date toggles, deck-wide (`slideIndex:-1`) or per-slide — mirrors PowerPoint's native "Insert Header and Footer" dialog. `skipTitleSlide`/`startNumber` are always deck-wide regardless of `slideIndex` (PowerPoint has no per-slide version of either). |

**Shape editing**

| Tool | Notes |
|---|---|
| `set_element_text` | Replaces one shape's text. `bulleted:true/false` controls real PowerPoint bullets — the model is instructed never to type a literal "•"/"-"/"*". |
| `format_element` | One kind of formatting per call (mixing groups errors): text style (bold/italic/underline/shadow/fontSize/fontName/`color`/alignment/baselineOffset), `fill` (hex or `"none"`), stroke (`strokeColor`/`strokeWidthPt` or `removeStroke`), or `opacity` (via `Fill.Transparency`). **Strikethrough deliberately not implemented** — this PIA has no `TextFrame2` member on `Shape` at all (confirmed via a `CS0234` compile failure), unlike Excel's newer text model. |
| `set_element_transform` | One of: left/top/width/height/rotation, or `order` (z-order: `bringToFront`/`sendToBack`/`bringForward`/`sendBackward`). Changing the order shifts other shapes' `shapeIndex` on the slide — re-read before addressing another shape by index in the same run. Top-level shapes only. |
| `delete_element` | Deletes one shape. |
| `edit_picture` | `kind:"add"` — inserts a picture on a slide from a **full local drive path only** (`PathGuard`, same rule as the other image paths); natural size shrunk to fit the slide, or proportional from one of `width`/`height`, or stretched to both; omitted `left`/`top` centers that axis; returns the new `shapeIndex`; the shape is deleted again if placing it fails. Both `crop` and `replace` work only on a picture (plain or linked); anything else errors with "is not a picture" and changes nothing. `crop` also accepts a placeholder that holds a picture; `replace` refuses any placeholder, because replacing deletes the old shape and PowerPoint recreates an empty placeholder (seen live on a title: its text was lost). `kind:"crop"` — Fractional crop (0..1) against current on-slide size. Documented imprecision under repeated crops.; `kind:"replace"` — **Full local drive path only** (`PathGuard`, same rule as Excel/Word; also `edit_master_element` icons) — swaps a picture's content in place, keeping position/size/rotation/approximate z-order. |

**Shape creation, copy/move, and grouping**

| Tool | Notes |
|---|---|
| `add_shape` | 26 presets + textbox, shares Excel's ported `ShapeTypeMap`; unrecognized name errors listing valid ones (no longer a silent rectangle fallback). **No `shapeType` makes a text box** (`text` required; `bulleted:true` for a real bulleted list). No fill/line params — chain `format_element`. |
| `copy_element` | No `targetSlideIndex` (or the same slide): copies a shape on the **same** slide with the native `Shape.Duplicate` — every shape kind, including groups/pictures/tables/charts/SmartArt; offset or exact position; auto-deduped name. A `targetSlideIndex` on a **different** slide: copies it there using PowerPoint's own native copy/paste (the real Windows clipboard — an explicit exception to this codebase's usual rule against it, since it reproduces the exact underlying OOXML the same way Ctrl+C/Ctrl+V does), every shape kind with full native fidelity. `move:true` (different slide only) also removes the original — atomic (only removed once the copy fully succeeds). |
| `copy_element_style` | Format painter: native `Shape.PickUp`/`Apply` (not the clipboard) for broad fidelity (fill/outline/shadow/3-D/bevel), plus targeted per-run text formatting, text outline/strikethrough/glow/reflection, text-box anchor/margins, gradient/patterned fill, dash/arrowhead outlines, rotation, flip, and AutoShape adjustment handles. Not position or size. Targets can span any slide. |
| `group_element` | `action:"group"` groups two or more top-level shapes (`shapeIndexes`); `action:"ungroup"` promotes a group's children back to top-level (`shapeIndex`). Both shift other shapes' indices on the slide — re-read before another index-addressed edit in the same run. |

**Tables**

| Tool | Notes |
|---|---|
| `edit_table` | `kind` selects the operation: `create` (native `Shapes.AddTable`, optional pre-filled `cells`, `left`/`top`/`width`/`height` default 100/100/400/200), `set_cell` (replaces one cell's text), `insert-row`/`delete-row`/`insert-col`/`delete-col` (bounds-checked — an out-of-range `index` errors specifically rather than a raw COM error; shifts later indices), or `style` (firstRow/bandRow, shading, `borderPreset`: `all` (every cell edge) / `outline` (outer perimeter only) / `none`). Row/col index 0 is just the first physical row/column, including a header row — there's no separate header concept in the index space. |

**Charts**

| Tool | Notes |
|---|---|
| `edit_chart` | `kind:"create"`: column/columnStacked/bar/barStacked/line/area/pie/doughnut (`chartType`, enum'd from the corrected chart-type map), writes real data into the chart's embedded Excel workbook; a series/categories length mismatch errors instead of drawing a silently wrong chart; position/size default 100/100/400/300. `kind:"update"`: `chartType`, `title`, `legendPos` (`none`/`right`/`top`/`left`/`bottom` + short aliases `r`/`t`/`l`/`b`), `dataLabels` (`none`/`value`/`percent`), `gridlines` (errors with a clear message on chart types with no value axis, e.g. pie/doughnut); several properties may be combined in one update. |

**SmartArt**

| Tool | Notes |
|---|---|
| `edit_smartart` | `kind`: `create` (7 layout keys list/process/cycle/hierarchy/pyramid/matrix/venn mapped to native SmartArt layouts by display name; flat `items` list only; unrecognized key errors listing the 7 valid ones; a valid key with no matching gallery layout gives a distinct error), `set_text`, `add_node`, `delete_node`, `set_style` (colorName/quickStyleName, substring-matched against the gallery), `set_layout`. `smartArtIndex` is 0-based **within that slide** and is not used by `create`. `delete_node` shifts later indices. |

**Animations & transitions** — see also `set_slide` above

| Tool | Notes |
|---|---|
| `edit_animation` | `kind`: `add` (entrance by default or `exit:true`, 20 effect names, `trigger` (`onClick`/`withPrevious`/`afterPrevious`), duration/delay; no directional variants, e.g. "wipe from the left" — base effect only; returns the new `animationIndex`), `delete`, `set_timing` (duration/delay/trigger), `reorder` (`toIndex`). `animationIndex` addresses current play order (call `read_slide (view=animations)` first); delete/reorder shift later indices. |

**Slide Master**

| Tool | Notes |
|---|---|
| `edit_master_element` | `kind:"add"` (`contentType` `icon` or `text`) — Pins a persistent icon or text label to the Slide Master (shows on every slide) or one specific layout (`layoutName`). Corner placement (inset by `marginPt`) or exact `left`+`top`. Known limits: only the default (first) Slide Master is reachable without `layoutName`; a layout with "Hide Background Graphics" won't show a master-level element.; `kind:"move"` — Moves/resizes/rotates an existing master element by `masterShapeIndex` (from `read_master`) — works on theme placeholders too.; `kind:"remove"` — Deletes a master shape by index. Refuses to delete a theme placeholder (marked `(placeholder)` in `read_master`) — use `set_headers_footers`'s `*Visible:false` to turn one off instead. |

**Search/replace across the deck**

| Tool | Notes |
|---|---|
| `replace_text` | Every text-frame shape (text boxes, title/body placeholders) on every slide, plus speaker notes unless `includeNotes:false`. **Not** table cells or SmartArt node text — use `edit_table`/`edit_smartart`. `regex:true` supports `$1`-style backreferences. Reports the number of occurrences actually replaced. |

**Undo/redo (2)**

| Tool | Notes |
|---|---|
| `undo_last_action` / `redo_last_action` | PowerPoint's own Ctrl+Z/Ctrl+Y. PowerPoint's object model has no Undo/Redo method, so these go through the generic ribbon-command dispatch, `Application.CommandBars.ExecuteMso("Undo"/"Redo")`, with a `GetEnabledMso` pre-check so an empty stack reports "Nothing to undo/redo" instead of a silent no-op (`PowerPointTools.History.cs`). `Application.StartNewUndoEntry()` runs before **every** non-read tool call — undo/redo included, an intentional difference from Word, whose `UndoRecord` wrapping excludes them — so one tool call = one undo step. Track Changes mode and up. |

**File (1)**

| Tool | Notes |
|---|---|
| `save` | Saves the open file in place, or opens the Save As file dialog (`Application.FileDialog[msoFileDialogSaveAs]` — PowerPoint has no `Dialogs` collection; `Show()` picks the target, `Execute()` saves) when `saveAs` is true (default false) or the file has never been saved. `saveAs` must be true whenever the user wants a copy, a new name, another location or another format. A cancelled dialog returns a non-error "cancelled" result with `Mutated: false`; the dialog is modal on PowerPoint's UI thread and Stop cannot dismiss it (`PowerPointTools.File.cs`). **NOT VERIFIED AGAINST LIVE POWERPOINT** — compiles clean in Debug and Release, dialog members confirmed by reflection against the PIA, never exercised against a running instance. Track Changes mode and up. |

## Missing entirely (confirmed absent from both the C# switch and the advertised tool list)

- `execute_slide_script` — no scripting DSL; every multi-property/multi-element edit
  must go tool-by-tool, rather than as one atomic batch script.
  This remains a real gap today.
- The entire deck-generation pipeline: `ask_clarification`, `plan_deck`,
  `generate_deck`, `regenerate_slide`, `save_style_template`, `list_style_templates`.
- No automatic post-edit audit/QC pass (geometric overflow/overlap/bounds check, or a
  vision-based review of generated slides); nothing here checks the result of an edit
  automatically.
- No `add_comment`-equivalent for PowerPoint (unlike Word), so Comment Only mode
  currently behaves identically to Read Only.

## Structural fragility note

Shapes are addressed by **positional index** (`slideIndex`, `shapeIndex` into
`slide.Shapes`, or a dotted path like `"3.1.0"` for a shape inside a group) rather than
a stable id — indices shift whenever shapes are added/removed/reordered/grouped, Several tools' descriptions explicitly tell the
model to re-read the slide after a structural change, but there's no protection against
acting on a stale index otherwise.
