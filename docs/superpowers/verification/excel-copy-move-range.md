# Verification — `copy_range` / `move_range` (2026-09-16)

## Four-way parity cross-check (PP-5's manual-audit contract, repeated by hand)

| Location | `copy_range` | `move_range` |
|---|---|---|
| `RequiredFields` (`ExcelTools.Operations.cs:37-38`) | `["sourceRange", "targetCell"]` | `["sourceRange", "targetCell"]` |
| `KnownOperationKinds` (`ExcelTools.Operations.cs:83`) | present | present |
| `switch(kind)` (`ExcelTools.Operations.cs:170-171`) | `CopyOrMoveRange(op, cut: false)` | `CopyOrMoveRange(op, cut: true)` |
| `EXCEL_OPS` (`entry.ts:210-227`) | `required: ['sourceRange', 'targetCell']` | `required: ['sourceRange', 'targetCell']` |

Each kind appears exactly once in all four locations; required-field sets match exactly (`sourceRange`, `targetCell`; `targetSheetId` optional in all four, consistent with `add_pivot`'s own entry). Diff: empty.

## Automated checks (this environment, no live Office session)

- `MSBuild ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug` — **clean.** Real compile-time confirmation that `Range.Copy(Destination:)`, `Range.Cut(Destination:)`, `Sheets[...]`, `.Rows.Count`/`.Columns.Count` all resolve against the referenced Excel PIA (`Microsoft.Office.Interop.Excel`).
- `tsc --noEmit` (ExcelAiAddIn) — clean after the `EXCEL_OPS` addition.
- esbuild bundle rebuild — clean, `web/bundle.js` regenerated.
- `dotnet test OfficeAi.Shared.Tests` — 136/136 passed, unaffected (no shared C# files touched by this feature).

## Deferred to real manual Excel testing (not possible from this environment)

- Whether `Range.Copy(Destination:)`/`Range.Cut(Destination:)` actually marshal correctly at runtime from this PIA in this project — compiling clean only confirms the signatures resolve, not that the COM call succeeds (see the plan's Risk #1: this is genuinely unexplored territory in this codebase, with a documented `Value2`-based fallback if it fails).
- Whether formatting, merged cells, and formulas survive the round trip as expected.
- Cross-sheet `targetSheetId` behavior.
- `move_range`'s cross-workbook formula-reference auto-update on Cut.
- The 2000-cell cap's suitability — unmeasured for a Copy/Cut-based op (see code comment in `CopyOrMoveRange`).

**Suggested first manual test:** on a fresh sheet, `copy_range {sourceRange:"A1:B3", targetCell:"D1"}` with a mix of a formula, a bold cell, and a number-formatted cell in the source block; then `move_range {sourceRange:"A1:B3", targetCell:"F1"}` and confirm A1:B3 is empty afterward with no separate clear needed.
