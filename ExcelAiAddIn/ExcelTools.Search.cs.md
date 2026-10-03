# ExcelTools.Search.cs

## `ResolveSheetsToSearch`

Shared by find_cells and propose_operations' find_replace: which
sheets a call should touch. sheetId names one specific sheet;
otherwise allSheets picks every sheet in the workbook (Ctrl+F's
"Within: Workbook") or, by default, just the active sheet (Ctrl+F's
default "Within: Sheet") - a narrower default than this tool used
to have (previously: omitting sheetId meant the WHOLE workbook).
