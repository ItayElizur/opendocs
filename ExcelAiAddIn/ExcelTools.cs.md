# ExcelTools.cs

## Shape-name lookup (near `AlwaysAllowedTools`)

Shape-name lookup now lives in OfficeAi.Shared.ShapeTypes (Phase 0) -
union of this map and PowerPoint's near-identical copy. PP-16:
mirrors EXCEL_SHAPE_TYPES / add_shape's shapeType enum in
ExcelAiAddIn/web-src/entry.ts exactly, plus the separately-handled
"textbox". Edit both together.

## Chart-type vocabulary (near `AlwaysAllowedTools`)

PP-15: chart-type vocabulary for BOTH add_chart and edit_chart now
lives in OfficeAi.Shared.ChartTypes, shared with Word and PowerPoint.
The cross-referencing this comment used to describe by hand (and the
PptChartTypeMap "bar" bug it records) is what motivated sharing it.

## `Execute` - editing-mode gating

Excel has no add_comment-equivalent tool yet, so Comment Only
mode allows no mutating tools at all (documented gap - see
Task 16 brief). Track Changes mode currently behaves the
same as Full Autonomy for gating purposes: Excel's
track-changes equivalent (Workbook.HighlightChangesOnScreen /
shared-workbook change tracking) is more limited than
Word's TrackRevisions and is out of scope for this task, so
there is deliberately no COM call wired up for it here.
