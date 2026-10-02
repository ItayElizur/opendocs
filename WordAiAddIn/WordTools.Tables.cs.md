# WordTools.Tables.cs

## `EditTable` - `set_style`, `borders` field

Post-hoc fix (2026-08-24, user-reported): this branch had
no border support at all - "borders" is a real field on
updateParagraphStyle elsewhere in this file, and the model
(reasonably, given that precedent) called edit_table with
the same field name expecting the same effect. Since
set_style never checked for it, the call silently did
nothing - a real gap, not a user error. table.Borders
mirrors the Word.Border collection updateParagraphStyle
already uses for paragraph borders, applied here at the
whole-table level (outside + inside edges).

## `EditTable` - `set_style`, border sides enumeration

Post-hoc fix (2026-08-24, user-reported): table.Borders
is not just the 6 grid sides - it also includes
wdBorderDiagonalDown/wdBorderDiagonalUp (the rare
cell-split diagonal lines), so the blind foreach over
the whole collection turned those on too, producing
crisscrossing diagonals across every cell. Enumerate
only the real table grid sides explicitly.

## `EditTable` - `set_shading` case

Post-hoc addition (2026-08-24, user-requested): fills
cell background color at cell/row/col/whole-table
scope. Word.Cell.Shading.BackgroundPatternColor is the
same property/pattern updateParagraphStyle's
shadingFill already uses on paragraphs elsewhere in
this file - applied per-cell here since Word tables
have no single "shade this row" API, only per-cell
shading (matches PowerPoint's own EditTableStyle,
which does the identical per-cell loop for its
shadingColor field).

## Trailing note (end of file, SmartArtLayouts.ByName)

PP-23 Task 4: ported from PowerPointTools.SmartArtLayouts.ByName /
ResolveSmartArtLayout verbatim - same seven keys, same
two-distinct-errors design (unknown key vs. valid-key-but-not-in-
this-install's-gallery). SmartArt is the Office-shared object
model, not PowerPoint-specific - Application.SmartArtLayouts
resolves identically against this add-in's own ThisAddIn.
