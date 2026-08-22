# Tool surface checklist

Tracks every tool/operation kind from `C:\Dev\genoffice\docs\ai-tool-surface.md` that is
in scope for `officeoffice` (Word/Excel/PowerPoint document manipulation only), against
what's actually implemented in `WordTools.cs`/`ExcelTools.cs`/`PowerPointTools.cs`.

**Explicitly out of scope everywhere** (per the original feasibility report and
project Global Constraints — not tracked below): `web_search`, `image_search`,
`generate_image`, `analyze_media`, `read_attachment`, the PDF app, the Markdown app
(folded into Word), and PowerPoint's `execute_slide_script` DSL + entire deck-generation
pipeline (`ask_clarification`, `plan_deck`, `generate_deck`, `regenerate_slide`,
`delete_slide`, `save_style_template`, `list_style_templates`).

## Word (`WordAiAddIn/WordTools.cs`)

### Top-level tools

- [x] `get_document_context`
- [x] `read_blocks`
- [x] `insert_content`
- [x] `replace_blocks`
- [x] `apply_commands` (gateway tool — see command kinds below)
- [x] `edit_chart` (combines genoffice's separate `insert_chart` + `edit_chart` into one create-or-edit tool)
- [x] `add_comment` (not in genoffice's docs surface — added here specifically for Comment Only mode)

### `apply_commands` command kinds

- [x] `set_bold` (officeoffice-specific name for genoffice's `updateTextStyle` bold field)
- [x] `set_italic` (officeoffice-specific name for genoffice's `updateTextStyle` italic field)
- [x] `set_heading` (≈ genoffice's `setHeadingLevel`)
- [x] `find_replace` (≈ genoffice's `replaceAllText`)
- [ ] `updateTextStyle` (full: font/color/underline/strikethrough — currently only bold/italic exist as separate command kinds)
- [ ] `updateParagraphStyle`
- [ ] `deleteBlocks`
- [ ] `moveBlocks`
- [ ] `createParagraphBullets`
- [ ] `deleteParagraphBullets`
- [ ] `updateImageProperties`
- [ ] `insertToc` (real Word TOC field via `TablesOfContents.Add`)

## Excel (`ExcelAiAddIn/ExcelTools.cs`)

### Top-level tools

- [x] `get_workbook_context`
- [x] `read_range`
- [x] `read_cells`
- [ ] `read_formats`
- [ ] `read_sheet_features`
- [ ] `find_cells`
- [ ] `select_range`
- [ ] `trace_precedents`
- [ ] `trace_dependents`
- [ ] `load_guide` (genoffice-internal prompt-management mechanism for its 65-op DSL — revisit only if the op count here grows large enough to need it)
- [x] `propose_operations` (gateway tool — see operation kinds below)

### `propose_operations` operation kinds (9 of 65 implemented)

**writing:**
- [x] `set_cell`
- [x] `set_formula`
- [ ] `clear_cell`
- [x] `set_range`
- [ ] `clear_range`

**formatting:**
- [x] `format_range` (bold/italic/numberFormat/fillColor only — genoffice's is richer, e.g. borders)

**layout:**
- [ ] `sort_range`
- [ ] `merge_cells`
- [ ] `unmerge_cells`
- [ ] `set_row_height`
- [ ] `set_col_width`
- [ ] `set_rows_hidden`
- [ ] `set_cols_hidden`
- [ ] `set_freeze`
- [ ] `set_page_setup`

**structure:**
- [x] `insert_rows`
- [x] `delete_rows`
- [x] `insert_cols`
- [x] `delete_cols`
- [ ] `add_sheet`
- [ ] `delete_sheet`
- [ ] `duplicate_sheet`
- [ ] `set_sheet_hidden`
- [ ] `move_sheet`
- [ ] `protect_sheet`
- [ ] `rename_sheet`

**charts:**
- [x] `add_chart` (basic: dataRange/chartType/title only — genoffice's is richer)
- [ ] `edit_chart`
- [ ] `delete_visual`
- [ ] `add_sparkline` (the one item the original feasibility report flagged as a categorical Office.js-only gap VSTO uniquely closes)
- [ ] `add_shape`
- [ ] `edit_shape`
- [ ] `add_image`

**table:**
- [ ] `add_table`
- [ ] `add_table_row`
- [ ] `add_table_column`
- [ ] `delete_table_row`
- [ ] `delete_table_column`
- [ ] `delete_table`

**pivot:**
- [ ] `add_pivot`
- [ ] `refresh_pivot`

**data:**
- [ ] `set_hyperlink`
- [ ] `set_filter`
- [ ] `clear_filter`
- [ ] `set_filter_criteria`
- [ ] `add_conditional_format`
- [ ] `clear_conditional_formats`
- [ ] `set_data_validation`
- [ ] `set_note`
- [ ] `add_defined_name`
- [ ] `delete_defined_name`

## PowerPoint (`PowerPointAiAddIn/PowerPointTools.cs`)

- [x] `get_deck_context`
- [x] `read_slide`
- [x] `set_element_text`
- [x] `set_element_style` (bold/italic/fontSize/color only)
- [ ] `set_element_fill` (missing from `set_element_style`)
- [ ] `set_element_stroke` (missing from `set_element_style`)
- [x] `set_element_transform`
- [x] `add_text_box`
- [x] `add_shape` (rectangle/oval/roundRect only — genoffice supports the full OOXML preset-geometry set)
- [x] `delete_element`
- [ ] `add_slide` (currently no way for the AI to add a new slide at all)
- [ ] `add_chart` (headline VSTO-vs-Office.js justification from the original feasibility report — never built)
- [ ] `edit_chart` (same)
- [ ] `add_smartart` (same headline justification — never built)
- [ ] `add_table`
- [ ] `edit_table_cell`
- [ ] `edit_table_structure`
- [ ] `edit_table_style`
- [ ] `set_slide_background`
- [ ] `ungroup_element`
- [ ] `crop_image` (operates on an existing embedded picture, no internet needed)
- [ ] `set_picture_opacity` (same)
- [ ] `replace_image` (local-file variant — no AI generation needed)
