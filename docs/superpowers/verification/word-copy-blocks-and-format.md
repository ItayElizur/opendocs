# Verification — `copyBlocks` / `copyFormat` (2026-09-16)

## Five-place parity cross-check (mirrors the `apply_commands` contract stated at `entry.ts:13-15`)

| Location | `copyBlocks` | `copyFormat` |
|---|---|---|
| `RequiredFields` (`WordTools.Commands.cs:37-38`) | `["target", "afterBlockIndex"]` | `["sourceBlockIndex", "target"]` |
| `NonNullFields` | none needed (no ambiguous-null field) | none needed |
| `KnownCommandKinds` (`WordTools.Commands.cs:64`) | present | present |
| `ApplyCommands` switch (`WordTools.Commands.cs:162-167`) | `CopyBlocksCmd(cmd)` | `CopyFormatCmd(cmd)` |
| `WORD_COMMAND_SCHEMAS` (`entry.ts:232-251`) | `required: ['kind', 'target', 'afterBlockIndex']` | `required: ['kind', 'sourceBlockIndex', 'target']` |

Each kind appears exactly once in all five locations; required-field sets match exactly. Diff: empty.

## Automated checks (this environment, no live Office session)

- `MSBuild WordAiAddIn.csproj -t:Build -p:Configuration=Debug` — **clean.** Confirms `WordOpenXML`/`InsertXML` (`copyBlocks`) and all ~18 read-side Font/ParagraphFormat/Shading/Border property names + the two new `ApplyTextStyle`/`ApplyParagraphStyle` helper signatures (`copyFormat`) resolve against the referenced Word PIA.
- `tsc --noEmit` (WordAiAddIn) — clean after the `WORD_COMMAND_SCHEMAS` and system-prompt additions.
- esbuild bundle rebuild — clean.
- `dotnet test OfficeAi.Shared.Tests` — 136/136 passed, unaffected (no shared C# files touched; `UpdateTextStyle`/`UpdateParagraphStyle`'s refactor is entirely within `WordAiAddIn`).

## Deferred to real manual Word testing (not possible from this environment)

- **`copyBlocks`**: single/multiple paragraphs (verify ascending-order output), `afterBlockIndex: -1`, mid-document `afterBlockIndex`, `afterBlockIndex` equal to one of the copied paragraphs' own indices (the deliberately-allowed case), zero-match `target` (confirm the error), formatting survival through the copy.
- **`copyFormat`**: single/multiple targets; a source with **uniform** formatting (baseline correctness); a source with **mixed** character formatting within it — this is the one genuinely open question (see below); a source with mixed border sides (e.g. bottom-only) to confirm the per-side copy is faithful and not collapsed to all-on/all-off; out-of-range `sourceBlockIndex`; `target` including the source paragraph's own index (confirm harmless no-op).
- **`UpdateTextStyle`/`UpdateParagraphStyle` regression check**: both were refactored to route through the two new `Apply*` helpers. Re-verify a representative call of each (e.g. an existing `updateTextStyle`/`updateParagraphStyle` command from a prior manual test session) still produces byte-identical behavior — same fields applied, same `highlight`/`link` handling, same `borders` on/off loop (untouched).

## Open risk carried into manual testing (from the plan's Risk #1)

`copyFormat` reads Font/ParagraphFormat properties directly from a whole-paragraph `Range`/`ParagraphFormat`, which can legitimately span non-uniform formatting (e.g. a paragraph that's half-bold). Word Interop's documented convention for this is a "mixed value" sentinel (`wdUndefined`, `9999999`) for many int/enum-typed properties, and the current read code (`bold = srcRange.Font.Bold == -1`, etc.) does not special-case that sentinel — a mixed-bold source paragraph would currently read as `bold = false` (sentinel doesn't equal `-1`), not "leave bold alone." This is untested against real Word from this environment. **Suggested first manual test**: apply `copyFormat` from a paragraph with mixed bold within itself, and confirm the result is reasonable (or note exactly what happens, to decide whether a mixed-value guard is needed).
