# PowerPoint tools (`PowerPointAiAddIn/PowerPointTools*.cs`)

Current-state reference for every tool the AI can call in PowerPoint. See
[`docs/architecture.md`](../architecture.md) for the shared editing-mode/transport
architecture.

## Tools (51 total, all genuinely implemented)

The largest tool surface of the four apps. Organized here by area rather than as one
flat table.

**Reading (8, always allowed, never gated — `READER_TOOLS` in `entry.ts`)**

| Tool | Notes |
|---|---|
| `get_deck_context` | One-line-per-slide outline: slide index + text preview of its shapes. No per-element type/id inventory. |
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
| `delete_slide` | Deletes one slide. Cannot delete the last remaining slide; later slides shift down — re-read before deleting another in the same run. |
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
| `edit_table_cell` | Replaces one cell's text. Row/col index 0 is just the first physical row/column, including a header row — there's no separate header concept in the index space. |
| `edit_table_structure` | Insert/delete row or column; bounds-checked (an out-of-range index errors specifically rather than a raw COM error). Shifts later indices. Row/col index 0 is just the first physical row/column, same note as above. |
| `edit_table_style` | firstRow/bandRow, shading, `borderPreset`: `all` (every cell edge) / `outline` (outer perimeter only) / `none`. |

**Charts**

| Tool | Notes |
|---|---|
| `add_chart` | column/columnStacked/bar/barStacked/line/area/pie/doughnut, writes real data into the chart's embedded Excel workbook. `kind` is enum'd from the corrected chart-type map; a series/categories length mismatch errors instead of drawing a silently wrong chart. Returns the new shape's index for a follow-up `edit_chart`. |
| `edit_chart` | `chartType`, `title`, `legendPos` (`none`/`right`/`top`/`left`/`bottom` + short aliases `r`/`t`/`l`/`b`), `dataLabels` (`none`/`value`/`percent`), `gridlines` (errors with a clear message on chart types with no value axis, e.g. pie/doughnut). Result names exactly which properties were applied. Unrecognized values throw, listing valid ones from the map itself (fixed a previous `"bar"` → `xlColumnClustered` silent-wrong-chart-type bug, and `legendPos`'s terminal-else-to-bottom silent fallback). |

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
