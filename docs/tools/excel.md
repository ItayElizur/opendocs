# Excel tools (`ExcelAiAddIn/ExcelTools*.cs`)

Current-state reference for every tool the AI can call in Excel. See
[`docs/architecture.md`](../architecture.md) for the shared editing-mode/transport
architecture. "genoffice" below refers to a sibling from-scratch web-based Office clone
project this tool design was originally ported from — see
[`docs/comparison-with-genoffice.md`](../comparison-with-genoffice.md) for the
high-level summary.

## Top-level tools (10)

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

No `undo_last_action`/`redo_last_action` for Excel. An earlier version dispatched
`Application.CommandBars.ExecuteMso("Undo"/"Redo")`, but that can't work for the
assistant's own edits: every write this add-in makes goes through the object model
(`propose_operations`, find/replace, etc.), and Excel clears its whole undo stack on
any object-model write — so right after an AI edit there is nothing to undo, and the
tool would later undo the *user's* next manual edit instead. `Application.OnUndo` only
takes a VBA macro name, so it isn't usable from a C# add-in. A snapshot-based custom
undo (the approach Outlook's undo/redo uses) was considered and not pursued; the tools
were removed rather than ship something misleading.

Notable native-COM advantage: `find_cells`'s `errors_only` mode uses
`Range.SpecialCells(xlCellTypeFormulas, xlErrors)` — a genuinely native error-cell
scan, the categorical VSTO/COM advantage over Office.js's wildcard-only `Range.find`.

## `propose_operations` operation kinds (54) — all named kinds from genoffice's list are implemented

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
description, cross-checked exhaustively against `ExcelTools.cs`'s `ProposeOperations`
switch.

Where officeoffice's version is still **narrower** than genoffice's:

| Op | Gap |
|---|---|
| `add_image` | **Local file paths only** — remote URLs throw `NotSupportedException` ("air-gapped deployment"). genoffice downloads from `image_search`/`generate_image` results. Deliberate scope boundary, not a bug. |
| `add_shape` | 26 named preset types + textbox (case-insensitive match, unknown name errors listing valid ones) — narrower than genoffice's "full OOXML preset-geometry set" but substantial. |
| `set_data_validation` | `checkbox` kind explicitly rejected — Excel's Data Validation COM API (verified via reflection against the referenced PIA) has no boolean-checkbox validation type; only 5 kinds exist total (`list`/`listRef`/`numberBetween`/`dateBetween`/`formula`), none map to it. |

`format_range` and `add_chart` have both reached parity with genoffice: `format_range`
covers font name/size/color, strikethrough, underline (enum-or-boolean),
horizontal/vertical alignment, wrap, rotation, indent, and full border control
(preset+edges+style+color); `add_chart` shares the same chart-type vocabulary as
`edit_chart` (column/columnStacked/bar/barStacked/line/area/pie/doughnut) and supports
rebinding to a new range (`dataRange`+`dataSheet`+`plotBy`).

No `dataSource`/provenance-enforcement mechanism exists anywhere (see
[`docs/architecture.md`](../architecture.md) — nothing analogous exists in
officeoffice's Excel or PowerPoint chart tools).
