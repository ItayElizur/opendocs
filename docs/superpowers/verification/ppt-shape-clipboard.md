# Verification — `duplicate_element` / `copy_element` / `move_element` / `copy_element_style` (2026-09-16)

## Registration cross-check (4 tools × 4 places)

| Tool | `Execute` switch | `MUTATION_TOOLS` | `POWERPOINT_TOOL_DISPLAY` | `systemPrompt` |
|---|---|---|---|---|
| `duplicate_element` | `PowerPointTools.cs:93` | `entry.ts:235` | `entry.ts:674` | mentioned |
| `copy_element` | `PowerPointTools.cs:94` | `entry.ts:254` | `entry.ts:678` | mentioned |
| `move_element` | `PowerPointTools.cs:95` | `entry.ts:275` | `entry.ts:682` | mentioned |
| `copy_element_style` | `PowerPointTools.cs:96` | `entry.ts:425` | `entry.ts:730` | mentioned |

All 4 present in all 4 locations. No `AlwaysAllowedTools` entry needed — none of the four is a read-only tool.

## Automated checks (this environment, no live Office session)

- `MSBuild PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug` — **clean.** Confirms:
  - `Shape.Duplicate()` → `ShapeRange` → 1-based indexer (`duplicate_element`).
  - All 5 reconstruction paths' Interop calls (`HasTable`/`HasChart`/dynamic `HasSmartArt`, `AutoShapeType`, `Table.Cell(r,c)`, `Chart.SeriesCollection()`/`XValues`/`Values`/`Name`, `SmartArt.Layout.Name`/`Nodes`) resolve against the referenced PowerPoint PIA.
  - The synthetic-JSON reuse of `AddTextBox`/`AddShape`/`AddTable`/`AddChartPpt`/`AddSmartArt` compiles (`PowerPointTools.CrossSlide.cs`'s `BuildJson` + `CallAddAndGetNewShape`).
  - The 3 shared copy helpers (`CopyTextFormatting`/`CopyFillFormatting`/`CopyStrokeFormatting`, `PowerPointTools.FormatPainter.cs`) compile against `Font`/`Fill`/`Line`/`TextRange.Characters`/`.Paragraphs`.
  - Both new `.csproj` `<Compile Include>` entries are correct.
- `tsc --noEmit` (PowerPointAiAddIn) — clean.
- esbuild bundle rebuild — clean.
- `dotnet test OfficeAi.Shared.Tests` — 136/136 passed, unaffected (zero `OfficeAi.Shared` changes — all reverse-lookup maps are local to `PowerPointAiAddIn`, by design).

## Deferred to real manual PowerPoint testing (not possible from this environment)

Per the plan's Risks section, in priority order:

1. **Chart cross-slide copy/move** — highest risk. Whether `SeriesCollection().Item(i).XValues`/`.Values` reliably marshal as usable `object[]` arrays, and round-trip correctly for a chart built by this same tool's own `SetSourceData` call, is genuinely unverified. **Suggested first test.**
2. **`Shape.Duplicate()`'s exact positioning/indexing** — `duplicate_element`'s design (position from the original shape, index via `ZOrderPosition`) is chosen to tolerate either possible real behavior, but the actual behavior itself is unconfirmed.
3. **Mixed-formatting `TextRange` reads** in `CopyTextFormatting` (`Characters(1,1)`/`Paragraphs(1,1)`) — PowerPoint's sentinel convention (if any) for a non-uniform range is unconfirmed and distinct from Word's.
4. **`copy_element_style` on a `HasTextFrame`-true, zero-length-text shape** — the `Characters(1,1)` call is gated behind `TextFrame.HasText == msoTrue`, but this exact guard combination is untested.
5. **Full per-kind matrix**: each of the 5 supported cross-slide kinds (copy and move), each unsupported-kind error (group/picture/OLE/media/unmapped-autoshape-preset/non-curated-SmartArt-layout), `targetSlideIndex` validation, and orphan-shape cleanup on an induced mid-reconstruction failure.

**Suggested first manual test sequence**: (a) `duplicate_element` on one shape of each kind (text box, autoshape, picture, table, chart, SmartArt, group, line) — confirms the "no restriction" property. (b) `copy_element`/`move_element` for a simple text box and autoshape across two slides — lowest-risk of the 5 reconstructable kinds. (c) `copy_element`/`move_element` for a chart — the confirmed highest-risk path. (d) `copy_element_style` from a bold/colored text box onto two other shapes.
