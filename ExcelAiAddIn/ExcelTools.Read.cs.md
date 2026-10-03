# ExcelTools.Read.cs

## `ReadFormats` - 200-cell cap

PP-13 Task 3: measured ~1.1s for a 200-cell read with the widened
per-cell property set (12 COM reads/cell vs. the previous 4) on this
dev machine - comfortably under the ~2s budget the plan flagged, so
the 200-cell cap is kept as-is rather than lowered or split into a
properties?:string[] filter.
