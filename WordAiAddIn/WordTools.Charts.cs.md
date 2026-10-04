# WordTools.Charts.cs

## `WriteChartData` - dataWorkbook fetched inside the retry lambda

Post-hoc fix (2026-08-24, code-review finding while adding
diagnostics): chart.ChartData.Workbook was fetched OUTSIDE the
ComRetry.Run-protected block, so if THIS specific call is
the flaky one (plausible under the "OLE server not fully live
yet" hypothesis - it is the very first COM call that opens the
embedded object), the retry wrapper never got a chance to help
at all. Moved inside the lambda so every attempt re-opens it
fresh; declared here (nullable) so `finally` can still clean up
whichever attempt actually succeeded.

## `WriteChartData` - 120ms settle delay

Post-hoc fix (2026-08-24, user-reported the RPC failure
recurring even after the first fix): a brief settle
delay immediately after the embedded OLE workbook is
opened, before the first COM call against it. This is a
documented mitigation for this exact class of embedded-
chart-data-workbook flakiness - the automation surface
is not always fully live the instant ChartData.Workbook
returns. Cheap (one UI-thread sleep) relative to the
cost of a failed/retried chart creation.

## `WriteChartData` - Cells.Clear() before writing

Confirmed repro: a brand-new chart's embedded workbook comes
pre-seeded by Word/Office with placeholder sample data (a
default chart template, commonly 4 categories x 3 series).
Without clearing it first, only the cells the NEW data
actually occupies get overwritten - any leftover placeholder
cells beyond that extent stay in the sheet and get plotted
alongside the real data, producing phantom extra
categories/series the user never asked for.

## `WriteChartData` - SetSourceData takes a string, not a Range

ACTUAL ROOT CAUSE (2026-08-24, confirmed via .NET
reflection against the real referenced
Microsoft.Office.Interop.Word.dll, not a guess):
Word.Chart.SetSourceData's real signature is
SetSourceData(String Source, Object PlotBy) - the
first parameter is a STRING, not a Range at all. Every
prior attempt (round 1's sheet.Range(topLeft,
bottomRight), round 2's reused writeRange, and this
round's sheet.Range[a1Range]) was passing a Range COM
object where the method actually expects a string -
"Could not convert argument 0" was ALWAYS this type
mismatch, not a marshaling-path quirk. The correct
call passes a plain "SheetName!A1:B4"-style reference
string - no Range object needed at all.

## `WriteChartData` - finally block does not mask the real exception

ROOT CAUSE FOUND (2026-08-24, via DebugLog): this cleanup
previously had no catch of its own - when SetSourceData
failed above (see the real bug this block is next to), the
chart/embedded-workbook was left in a state where
dataWorkbook.Close() ALSO threw (a real, observed
RPC_E_DISCONNECTED). In C#, an exception thrown from a
`finally` block while another exception is already
propagating from the `try` block REPLACES it - so the
user only ever saw this cleanup-time exception
("The object invoked has disconnected from its clients"),
never the real SetSourceData ArgumentException that caused
it. This is exactly why two prior rounds of fixes,
diagnosing from the user's reported error text alone,
chased the wrong theory. Cleanup failures are now caught
and logged here instead of being allowed to propagate and
mask whatever real exception is already in flight.

## `ListChartShapes`

dynamic: Word's chart object model (Shapes.AddChart2 / Chart / SeriesCollection) mirrors
Excel/PowerPoint's shared chart engine; using dynamic avoids pinning down the exact
Interop type names for this spike and lets any signature mismatch surface immediately at
runtime instead of guessing overloads at compile time.

PP-9: create-or-edit against an explicit list of ALL charts (inline
first, then floating - see Task 4 Step 4), addressed by chartIndex,
with real categories/named multi-series/chart-type support ported
from PowerPointTools.AddChartPpt.
Every chart shape, inline first then floating, in that fixed order
so chartIndex is predictable across calls (PP-9 Task 4 Step 4).
Shared by EditChart and ReadChart so both address charts identically.
internal (not private): post-hoc fix (2026-08-24, user-reported)
needs this same addressing from TaskPaneHost.OnSelectionChanged, so
a selected chart shape can be reported with the SAME chartIndex
edit_chart/read_chart would use, rather than a second, possibly
drifting copy of this resolution logic.

## `ReadChart`

Lets the model inspect an existing chart's current title/type/
categories/series before deciding what to change via edit_chart -
without this, an incremental edit (e.g. "remove one category") has
no way to know what the other categories/series currently are,
since edit_chart REPLACES the whole dataset rather than patching
it. Reads from the chart's embedded workbook (the same object
WriteChartData writes to) via the same Cells/UsedRange/.Value2
pattern already proven working by the write side, rather than the
Series.Values/.XValues COM properties directly (whose exact
marshaled array shape in this dynamic context is not something
this environment can verify without a live Word session).
