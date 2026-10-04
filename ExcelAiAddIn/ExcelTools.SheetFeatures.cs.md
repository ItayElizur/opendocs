# ExcelTools.SheetFeatures.cs

## `FindDataTables`

Splits the sheet's used range into separate contiguous data blocks -
BFS flood-fill over non-empty cells, 4-directional, no gap tolerance
(a single fully-blank row or column separates two tables). Ported
from nlputils' ExcelTableParser.find_data_tables, simplified for COM:
reads Value2 in one bulk call instead of per-cell (a per-cell COM
round-trip would make flood-fill prohibitively slow), and doesn't
special-case merged cells - Value2 already reports a merge's
non-anchor cells as empty, which is enough for the common case
(merge's text lives in its top-left cell) but means a merge wider
than the flood fill's reach can still fragment a table. A sheet with
scattered blank cells inside one logical table will similarly
over-fragment into several regions - acceptable for now since the
goal is finding stacked/side-by-side tables, not validating layout.
