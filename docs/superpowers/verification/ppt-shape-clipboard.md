# Verification — `duplicate_element` / `copy_element` / `move_element` / `copy_element_style`

## History

This started (2026-09-16) as a property-based reconstruction approach for
`copy_element`/`move_element`: read a source shape's properties and manually
rebuild an equivalent on the destination slide, avoiding the OS clipboard
entirely (matching this codebase's rule for Word). After several rounds of
live testing (2026-09-22 through 2026-09-23) that approach was abandoned:
SmartArt has no COM API to read its layout's own placeholder structure,
PowerPoint refuses to `Group()` a SmartArt with any other shape, table cell
shading/borders/merged cells aren't fully readable through this PIA, and
gradient fills don't round-trip cleanly through `GradientStops`. Each fix
closed one gap and surfaced another.

**Current approach (2026-09-23):** `copy_element`/`move_element` use
PowerPoint's own native `Shape.Copy()` + `Slide.Shapes.Paste()`, which
reproduces the exact underlying OOXML the same way Ctrl+C/Ctrl+V does — no
shape kind can fail to support this. This **does use the real Windows
clipboard**, an explicit, deliberate exception to this codebase's otherwise-
universal "never the clipboard" rule; the clobbering/racing risk is
mitigated (not eliminated) by saving and restoring the clipboard's prior
contents around the operation. `duplicate_element` (`Shape.Duplicate()`) and
`copy_element_style` (`Shape.PickUp()`/`Apply()`) were never part of the
reconstruction problem and are unaffected by any of this - both are native,
non-clipboard mechanisms.

## Registration cross-check (4 tools × 4 places)

| Tool | `Execute` switch | `MUTATION_TOOLS` | `POWERPOINT_TOOL_DISPLAY` | `systemPrompt` |
|---|---|---|---|---|
| `duplicate_element` | `PowerPointTools.cs:93` | `entry.ts:235` | `entry.ts:682` | mentioned |
| `copy_element` | `PowerPointTools.cs:94` | `entry.ts:255` | `entry.ts:686` | mentioned |
| `move_element` | `PowerPointTools.cs:95` | `entry.ts:274` | `entry.ts:690` | mentioned |
| `copy_element_style` | `PowerPointTools.cs:96` | `entry.ts:424` | `entry.ts:738` | mentioned |

All 4 present in all 4 locations. No `AlwaysAllowedTools` entry needed — none of the four is a read-only tool.

## Automated checks (this environment, no live Office session)

- `MSBuild PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug` — **clean.** Confirms:
  - `Shape.Duplicate()` → `ShapeRange` → 1-based indexer (`duplicate_element`).
  - `Shape.Copy()` / `Slide.Shapes.Paste()` → `ShapeRange` → 1-based indexer, and `System.Windows.Forms.Clipboard.GetDataObject()`/`SetDataObject()` resolve (`copy_element`/`move_element`, `PowerPointTools.CrossSlide.cs`).
  - `Shape.PickUp()`/`Shape.Apply()` (no-arg format painter) and every hand-rolled `Copy*Formatting` helper (`Font`/`Font2`/`Fill`/`FillFormat.GradientStops`/`Line`/`ThreeDFormat`/`TextRange.Characters`/`.Paragraphs`) compile against the referenced PIAs (`PowerPointTools.FormatPainter.cs`, used by `copy_element_style`).
  - Both `.csproj` `<Compile Include>` entries are correct.
- `tsc --noEmit` (PowerPointAiAddIn) — clean.
- esbuild bundle rebuild — clean.
- Zero `OfficeAi.Shared` changes across this whole history — the shared test suite is unaffected by construction, not just by re-running it.

## Live-tested (via the mock server's `FORCE_TOOL:` harness, across many rounds)

- `duplicate_element` — works for every shape kind, auto-suffixes on name collision.
- `copy_element`/`move_element` — confirmed working post-native-Copy/Paste-switch for: tables with merged cells and custom colors/borders, gradient-filled shapes, groups nesting SmartArt, cross-slide moves. Also fixed live: a slow clipboard restore (`SetDataObject`'s `copy` flag was eagerly flushing every format - see `PowerPointTools.CrossSlide.cs`'s `RestoreClipboard`), and a cryptic raw COM error copying a completely empty shape (no pasteable clipboard payload - see `CopyPasteShape`'s catch).
- `copy_element_style` — cross-slide targets, per-run text formatting, and text effects (outline/strikethrough/glow/reflection/shadow/soft-edge/bevel) confirmed; also fixed live: reflection/glow effects coming out "always set" regardless of source (`ReflectionFormat`/`GlowFormat` have no `Visible` gate - both are now only touched when the source affirmatively has a real, non-default effect).

## Deferred / open questions

1. **`copy_element_style`'s hand-rolled Fill/Stroke/Rotation/Adjustments copies may now be redundant** with `CopyViaPickUpApply` (PowerPoint's native format painter, called first in the same pipeline) - flagged in PR #14's description as a possible future cleanup, not done here since it can't be verified without a live A/B test of PickUp/Apply alone vs. the current layered approach.
2. **Clipboard save/restore in a real (non-mock-harness) PowerPoint session** — not yet manually confirmed that restoring the user's own clipboard content is visually/functionally seamless (e.g. immediately available to a manual Ctrl+V right after the tool call).
