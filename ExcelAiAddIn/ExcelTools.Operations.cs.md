# ExcelTools.Operations.cs

## `RequiredFields`

PP-5: mirrors EXCEL_OPS's `required` arrays in
ExcelAiAddIn/web-src/entry.ts exactly (minus "kind" itself, which is
validated separately in ProposeOperations before this table is
consulted) - the two must be edited together. Kinds needing no
field beyond "kind" (set_page_setup, delete_sheet, duplicate_sheet,
refresh_pivot, clear_filter, clear_conditional_formats) have no
entry here, matching WordTools.cs's RequiredFields convention. This
is the actual guarantee: the TS schema is documentation the model
reads, not a validator that runs (not every provider enforces
oneOf/const, and Excel's grouped-variant collapsed branch carries
no per-kind structure at all) - this precheck is what turns a
missing field into a specific, per-operation error instead of a
raw COM/NullReference exception.
