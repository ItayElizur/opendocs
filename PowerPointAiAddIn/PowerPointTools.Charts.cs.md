# PowerPointTools.Charts.cs

## `AddChartPpt` - embedded workbook open/close

Chart data lives in an embedded Excel workbook - open, write the grid,
close, and RELEASE explicitly so no hidden Excel host process leaks.
Post-hoc fix (2026-08-24, ported from Word's identical fix,
found via a real repro's DebugLog): moved inside the retry
lambda below so a flaky OPEN, not just a flaky subsequent
call, also gets retried.

## `AddChartPpt` - settle delay after opening the embedded workbook

Post-hoc fix (2026-08-24, user-reported the RPC failure
recurring even after the first fix - ported from
Word's identical fix): a brief settle delay right after
the embedded OLE workbook opens, before the first COM
call against it - the automation surface is not always
fully live the instant ChartData.Workbook returns.

## `AddChartPpt` - clearing pre-seeded placeholder data

Confirmed repro (Word's identical port of this same pattern,
PP-9): a brand-new chart's embedded workbook comes pre-seeded
by Office with placeholder sample data (a default chart
template, commonly 4 categories x 3 series). Without
clearing it first, only the cells the NEW data actually
occupies get overwritten - any leftover placeholder cells
stay in the sheet and get plotted alongside the real data.

## `AddChartPpt` - SetSourceData's real signature

ACTUAL ROOT CAUSE (2026-08-24, confirmed via .NET
reflection against the real referenced
Microsoft.Office.Interop.PowerPoint.dll, not a guess -
same finding as Word's identical code):
PowerPoint.Chart.SetSourceData's real signature is
SetSourceData(String Source, Object PlotBy) - the
first parameter is a STRING, not a Range. Every prior
attempt (a reused writeRange, then a Range built from
an A1 string) was passing a Range COM object where the
method actually expects a plain "SheetName!A1:B4"
reference string - no Range object needed at all.

## `AddChartPpt` - finally block does not let cleanup mask the real exception

ROOT CAUSE FOUND (2026-08-24): this cleanup had no catch of
its own - when SetSourceData failed above, the chart/
embedded-workbook was left in a state where
dataWorkbook.Close() ALSO threw a second, unrelated
exception, which (per C# finally semantics) REPLACED the
real SetSourceData exception before it ever reached the
caller. Cleanup failures are now caught and swallowed
(not re-thrown) so they can never mask a real exception
already in flight.

## `PptLegendPositions`

PP-21: legendPos's natural names (a model will say "right", not the
genoffice-ism "r") plus the original short aliases for back-compat.
xlLegendPositionCorner was considered and dropped - its code could not
be verified against a live Office install (no interactive GUI access
in this environment), and guessing it wrong would just replace one
silent-wrong-result bug with another.

## End-of-file note (orphaned - SmartArt gallery verification status)

Verified against the standard English display names for PowerPoint's built-in SmartArt
layout gallery. Live cross-check against Application.SmartArtLayouts on this machine's
Office install (plan Task 6 Step 1) requires interactive Office GUI access that was not
available in this environment - remains a manual follow-up for a human with GUI access.
