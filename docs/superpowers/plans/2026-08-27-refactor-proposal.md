# Refactor Proposal — officeoffice

**Status:** Proposal only — nothing in this document has been implemented. Written in response to a request to assess how manageable the codebase is to review/maintain and suggest a refactor direction.

**Basis:** A concrete audit of the current tree (see the conversation this originated from) — file sizes, duplication check (`grep -c` across the three `*Tools.cs` files), test coverage check (`OfficeAi.Shared.Tests` contents), and a read-through of internal method ordering in `WordTools.cs`.

---

## Where things actually stand

- `WordTools.cs` 2,602 lines / `ExcelTools.cs` 2,266 / `PowerPointTools.cs` 1,804 — each a single `static class` with one dispatcher `switch` and dozens of private methods. No `#region`/section markers in any of them.
- Hex-color conversion and shape/chart-type maps are independently reimplemented in all three files rather than living once in `OfficeAi.Shared` (confirmed via `grep -c "HexTo\|ShapeTypeMap\|ChartTypeMap"` — 14/27/19 hits respectively). The tool-surface audit doc already attributes at least one real feature-parity gap between apps to this.
- `OfficeAi.Shared.Tests` covers `ChatStore`/`DocSettingsStore` only. The ~6,600 lines of COM-integration logic across the three `Tools.cs` files — the highest-risk code in the project — has zero automated tests. The only verification is the mock-server `FORCE_TOOL` harness or manual testing in real Office apps.
- Each app's `entry.ts` invented its own shape for declaring tools: Word is one flat array; PowerPoint splits `READER_TOOLS`/`MUTATION_TOOLS`; Excel layers `ALL_TOOLS` + `EXCEL_OPS` + a `DETAILED_KINDS` schema-size mechanism.
- `docs/superpowers/` holds 37 plan files + 25 verification files, on top of `STATUS.md` and `ai-tool-surface.md`. This has already gone stale once — `tool-surface-todo.md` had to be retired because it actively misled about what was implemented.
- On the plus side: the shared web layer (`agent-core`/`ai-provider`/`app-shell`) is already properly factored (11 files, ~190 lines average), and inline comments consistently explain *why* a piece of code exists, dated and attributed — worth preserving, not "fixing."

---

## Proposed changes, in the order I'd actually do them

### Phase 0 — before touching any code — **DONE (2026-08-27)**
Get the three `Tools.cs` files under *some* test coverage before restructuring them, even a thin one. Refactoring 6,600 untested lines of COM logic on reasoning and compilation alone (as this session's perf fixes had to be) is the riskiest possible time to also start moving code around. Concretely: pick the pure, non-COM logic already embedded in these files (JSON parsing/validation like `ValidateRequired`/`ValidateKnownFields`, the color/shape/chart-type maps once extracted, the paragraph-index-resolution logic) and write unit tests for *that* first, since it's the part that doesn't need a live Office instance. This buys a small but real regression net before Phase 2/3 start moving things.

Implemented in full per `docs/superpowers/plans/2026-08-27-phase0-test-seam.md` - `dotnet test` 23 → 90 passed. One correction worth flagging: paragraph-index-resolution logic turned out to be COM-bound (holds `Word.Paragraph` state) and could not be extracted after all - deferred to Phase 5. The color/shape-type maps *were* extracted, but not as originally envisioned - an embedded Office interop type cannot cross an assembly boundary as a generic type argument (`CS1769`, found via a spike), so both are `Dictionary<string, int>` with each app casting at its call site, not `Dictionary<string, TheirEnum>`. Two tool-facing bugs were fixed along the way (hex-color validation, Word's null required-field handling) - see `docs/ai-tool-surface.md`'s 2026-08-27 note.

### Phase 1 — quick wins (low risk, do first, no behavior change)
- Add `#region`/section-header comments to the three `Tools.cs` files, grouped by the tool grouping that already roughly exists in the method ordering (read tools → gateway dispatcher → per-command handlers → media/misc). Pure navigation aid, zero functional risk.
- Archive completed plan docs under `docs/superpowers/plans/` into a `docs/superpowers/plans/archive/` (or similar) subfolder, keeping only currently-relevant ones at the top level, and add one line to `STATUS.md` pointing at `ai-tool-surface.md` as the canonical current-state doc. Prevents the `tool-surface-todo.md` staleness problem from recurring.

### Phase 2 — de-duplicate the cross-app C# helpers
Extract into `OfficeAi.Shared` (already exists, already referenced by all three add-ins):
- Hex-color → native-color conversion (`HexToWdColor`/`HexToOleColor`/PowerPoint's equivalent).
- Shape-type name → native enum maps (already flagged in-repo as "ported, not extracted" for PowerPoint's copy of Excel's `ShapeTypeMap").
- Chart-type name → native enum maps.

Each app's native color/shape/chart enum types differ (`Word.WdColor` vs `Excel`'s OLE color vs PowerPoint's `MsoAutoShapeType`), so this isn't a single shared function — it's shared *string-key tables* with each app supplying its own native-enum values, or small per-app adapter functions calling into one shared parsing/validation routine. Worth scoping carefully per-map rather than forcing a single generic abstraction that doesn't fit all three.

### Phase 3 — split the giant files (structure only, no logic change)
Once Phase 1's section markers exist, turn each `*Tools.cs` into a C# `partial class` split across a few files along the same grouping — e.g. for Word: `WordTools.cs` (dispatcher + shared helpers), `WordTools.Content.cs` (read/insert/replace blocks, find_text, get_headings), `WordTools.Commands.cs` (`apply_commands` and every command handler, including `ResolveTargetParagraphs`), `WordTools.Charts.cs`, `WordTools.Tables.cs`, `WordTools.SmartArt.cs`. Same pattern for Excel/PowerPoint. This is mechanical (move methods, no logic changes) but still needs a full rebuild + the mock-server smoke pass per app afterward, since a missed `private` visibility or an accidentally-duplicated helper name would be a real compile break, not silent.

### Phase 4 — unify the `entry.ts` tool-declaration shape
Design one shared shape (in `shared/web-src`, alongside `app-shell`) that all three apps' `entry.ts` files build their tool list with — something like a common `defineTool(name, description, schema)` and a shared `readOnlyTools`-computation helper, so Excel's `DETAILED_KINDS`-style schema-size tradeoff becomes a documented, reusable pattern instead of a bespoke one-off. This is the most disruptive phase (touches every tool declaration in every app) — do it last, after the other phases have already reduced how much is moving at once.

### Phase 5 (stretch, not a near-term ask) — a real testability seam for the COM logic
The deeper fix behind Phase 0's thin coverage: introduce a thin interface between each `Tools.cs` and the raw `Microsoft.Office.Interop.*` calls (even just wrapping the handful of collection-indexing/Find/GoTo operations that have already caused two rounds of real bugs this session), so the *logic* (which paragraph matched, what the resolved index is) can be unit-tested against a fake implementation without a live Word/Excel/PowerPoint process. This is a genuine architectural change, not a quick win — flagging it as a direction, not proposing it as part of this pass.

---

## What I would *not* do
- Don't collapse the three add-ins into one shared codebase beyond what already exists in `shared/` — Word/Excel/PowerPoint's object models are different enough (confirmed repeatedly this session: Word's `Paragraphs`, Excel's `Range`, PowerPoint's `Shapes`/`TextRange` all have distinct quirks) that forcing one abstraction over all three COM models would trade real duplication for a leakier, harder-to-reason-about one.
- Don't attempt Phase 3 (file splitting) and Phase 4 (`entry.ts` unification) in the same pass — each touches a lot of surface area; overlapping them makes a regression much harder to bisect.
- Don't delete `docs/superpowers/plans/`'s historical entries — archive, don't discard. They're a genuinely useful audit trail of *why* something is the way it is, which this session's own investigations leaned on repeatedly.

---

## Suggested order
Phase 0 → Phase 1 → Phase 2 → Phase 3 → Phase 4 → (Phase 5 as a separate future decision). Each phase should land as its own reviewable change, verified by a full rebuild of all three add-ins plus a mock-server smoke pass, before starting the next.
