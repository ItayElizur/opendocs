# ExcelTools.Charts.cs

## `ApplyChartCategoryRange`

Returns a description including the created chart's name (PP-15
Task 4), so a follow-up edit_chart in the same batch can address it
without guessing Excel's auto-assigned "Chart 1"-style name.
Excel's SetSourceData auto-detects category (x-axis) labels only when
the leftmost column/top row of the bound range is text - if it's
numeric, Excel can't tell it apart from another value series and the
chart falls back to a plain 1,2,3... index. This forces every
series' XValues to an explicit range so the model can put a numeric
column (dates, ids, years) on the x-axis on purpose.

## `EditChartExcel` - legend unmatched-value handling

PP-21 Task 2 Step 5: was a terminal-else-to-bottom - any
unmatched value (a model could plausibly send anything not
in this exact set) silently moved the legend to the bottom
instead of erroring, the identical defect PP-21 fixes on
the PowerPoint side.

## `AddSparkline` - return value

Returns a description including the target address (PP-18 Task 3
Step 5) - this PIA's SparklineGroup exposes no separate stable id
beyond the cells it occupies, so the target address IS the
addressing handle for a later edit (there is no delete_sparkline/
edit_sparkline operation yet to consume it, but it is at least
visible in the transcript for the user/model to reason about).
