# PowerPointTools.Tables.cs

## End-of-file historical notes (chart-type vocabulary / COM retry)

PP-21: chart-type vocabulary now lives in OfficeAi.Shared.ChartTypes.
This file's copy previously mapped "bar" to 51 (xlColumnClustered)
instead of 57 - a silent wrong result where even a *successful*
chartType:'bar' produced a column chart, with "barStacked" identically
wrong. That bug is precisely why the table is now single-source.

Transient-COM retry (the embedded chart-data workbook's OLE server
intermittently refuses rapid calls) now lives in
OfficeAi.Shared.ComRetry, shared with Word.
