# AI Tool Surface — Changelog

This is the historical record for `docs/ai-tool-surface.md`. It exists so the current
document can stay current-state-first: every entry below is a dated, point-in-time note
that has already been folded into the tables and prose in the main doc (or, for the
schema-vs-implementation audit, superseded by later fixes). **Nothing here should be
read as the current state of the tool surface — always check the main doc's tables
first.** Entries are kept verbatim (including line numbers, bug descriptions, and
"NOT VERIFIED" caveats that were true when written) and ordered oldest-first.

> **2026-09-27 note:** this changelog was split out of `docs/ai-tool-surface.md` when
> that document was reorganized (docs restructuring pass). At that time every entry
> below had already been folded into the main doc's current-state tables — Word's
> `apply_commands` kinds, PowerPoint's full ~49-tool surface (the doc previously only
> described a stale ~23/35-tool snapshot), Excel's 54 `propose_operations` kinds, and
> the schema-vs-implementation audit's findings (all of which turned out to have already
> been fixed by their owning PP item — PP-5/PP-9/PP-12 for Word, PP-13 through PP-18 for
> Excel, PP-21/PP-22 for PowerPoint). Where a table below and the main doc's current
> tables disagree, **the main doc wins** — this file is history, not a second source of
> truth. Some entries below predate `docs/ai-tool-surface.md`'s Outlook section and
> pre-2026-09-19 PowerPoint tool count and were superseded by later work that was never
> itself written up as a dated entry here (the tool surface grew faster than the
> changelog in a few places) — the main doc's tables reflect the actual current source,
> verified directly against `entry.ts`/`*Tools.cs` as of this restructuring.

---

### 2026-08-23 — Schema-vs-implementation audit (original)

> Triggered by a real bug: Word's `edit_chart` let the model set a chart title and one
> series' numeric values, but had no `categories` parameter at all — the model could never
> label a chart's axis categories or name its series, even though the *schema and its
> description* didn't oversell this (they only ever mentioned title+values). That specific
> case was a narrow-but-honest tool, not a schema/handler mismatch. The question this
> section answered is broader: **across every tool/op in all three add-ins, does the JSON
> schema advertised to the LLM (`entry.ts`) actually match what the C# handler
> (`*Tools.cs`) reads and does?** Verified by reading every schema definition against its
> handler body directly, not by inference.
>
> Two bug classes turned up, ranked by how badly they mislead the model:
>
> - **Silent no-op with false success** — the tool accepts a parameter, does nothing
>   useful with an out-of-range value, and still reports success. Worse than an error,
>   because the model has no signal to retry or ask the user.
> - **Undocumented schema** — the wire-level JSON Schema for a parameter is just
>   `{type: 'object'}` or `{type: 'string'}` with no enum/field list; the *real* contract
>   lives only in a free-text description (or nowhere). The model can only guess valid
>   values, and any guess outside the handler's recognized set falls into one of the
>   silent-no-op cases above.
>
> **Word (`WordTools.cs` / `entry.ts`)**
>
> | # | Tool / command | Issue | Class | File:line (schema / handler) | Status as of this changelog entry |
> |---|---|---|---|---|---|
> | 1 | `apply_commands` (whole tool) | The wire schema for `commands` items is just `{type:'object'}` — no `kind` field, no per-kind shape; everything is prose-only in the description. Consequence: `cmd.GetProperty("kind")` sits **outside** the per-command try/catch, so one malformed command (missing `kind`) throws and aborts the **entire remaining batch** with a generic error — while any commands already applied earlier in the same batch stay applied (no rollback). | Undocumented schema → robustness gap | entry.ts (commands: array<object>) / WordTools.cs:210-268 | **Fixed by PP-5** (structural per-kind schema + `kind` moved inside the per-command try/catch) and **PP-12** (numbered result lines + summary header). |
> | 2 | `updateTextStyle` | Schema's `style` field is `{type:'object'}` with zero enumerated keys. Concretely: **`highlight` is not implemented** (9 of genoffice's 10 fields), and because nothing enumerates valid keys, requesting `highlight` (or any hallucinated key) is silently ignored. | Silent no-op + false success | entry.ts (style: object) / WordTools.cs:236-238, 379-416 | **Fixed by PP-12** — `highlight` implemented (16-entry `WdColorIndex` palette), and `ValidateKnownFields` now errors by name on any unrecognized field instead of silently no-opping. |
> | 3 | `createParagraphBullets` (`bulletPreset`) | Handler only ever checks `bulletPreset.StartsWith("NUMBERED")` — every other value collapses to the same generic `ApplyBulletDefault()`. | Silent no-op | entry.ts / WordTools.cs:248-250, 544-561 | **Fixed by PP-12** — 7 named presets, each with real distinct behavior; an unrecognized preset now throws listing valid names. |
> | 4 | `edit_chart` | No mismatch — schema and description accurately describe the narrow capability (title + one series' values, no categories, no chart-type choice). | Honest but narrow (not a bug) | entry.ts / WordTools.cs:132-169 | **Superseded by PP-9** — `edit_chart` was rewritten with categories, named multi-series, chart-type selection, and multi-chart addressing. |
>
> `updateParagraphStyle` had the same "no enumerated keys in the wire schema" structural
> issue as `updateTextStyle`, but had full 10/10 field parity with genoffice underneath,
> so it worked only because the model happened to send valid keys — a latent version of
> the same risk, not an active bug at the time. PP-5's structural schema work covers it too.
>
> **Excel (`ExcelTools.cs` / `entry.ts`)**
>
> `propose_operations`'s formal schema was `{operations: {type:'array', items:
> {type:'object'}}}` — like Word's `apply_commands`, no per-operation-kind schema existed
> at all; every op's real parameter shape lived only in one long free-text description
> block (`entry.ts:198-224`). PP-5 replaced this with `EXCEL_OPS`, a single source of truth
> for both the JSON Schema and the human-readable description.
>
> | # | Operation | Issue | Class | File:line (schema / handler) | Status |
> |---|---|---|---|---|---|
> | 1 | `add_conditional_format` | The description said `rule: {kind, ...}` and never listed what fields any of its 8 `kind` values needed. | Undocumented schema | entry.ts:222 / ExcelTools.cs:400-476 | **Fixed by PP-5** (`CF_RULE_SCHEMA`, all 8 kinds field-for-field) and **PP-14** (silent-fallback behavior). |
> | 2 | `add_conditional_format` (`kind:"number"`, `operator`) | Any `operator` string other than `greaterThan`/`lessThan`/`equal`/`between` silently became `equal`. | Silent no-op | ExcelTools.cs:388-398 | **Fixed by PP-14** — all 8 comparison operators supported, unknown throws. |
> | 3 | `add_pivot` (`values[].agg`) | Any `agg` value other than `count`/`average`/`max`/`min` silently became `sum`. | Silent no-op | entry.ts:218 / ExcelTools.cs:1202-1211 | Not independently re-verified in this restructuring — flagged here for a future audit pass rather than assumed fixed. |
> | 4 | `add_shape` (`shapeType`) | 26 valid preset names existed in the handler; none were listed in the schema, and an unrecognized name silently became a rectangle. | Undocumented schema + silent no-op | entry.ts:213 / ExcelTools.cs:27-56, 746-773 | **Fixed by PP-16** — enum'd, case-insensitive, throws on unknown. |
> | 5 | `edit_chart` (`chartType`) | The handler supported 6 chart types but only `add_chart`'s narrower 3-type enum was ever documented. | Undocumented schema (capability hidden) | entry.ts:210-211 / ExcelTools.cs:58-66, 665-726 | **Fixed by PP-15** — `add_chart` now uses the same widened `ExcelChartTypeMap`. |
> | 6 | `edit_chart` (data rebinding) | No parameter existed to repoint an existing chart at a different range. | Capability gap | ExcelTools.cs:643-663 | **Fixed by PP-15** — rebinding (`dataRange`+`dataSheet`+`plotBy`) added. |
> | 7 | `set_page_setup` (`scale` + `fitToWidth`/`fitToHeight` together) | Combining them silently discarded `scale`. | Silent no-op | ExcelTools.cs:857-865 | **Fixed by PP-18** — combining now errors instead of silently dropping a value. |
> | 8 | `add_defined_name` / `delete_defined_name` | Only ever touched workbook-scoped names, no sheet-scoped create. | Capability gap | entry.ts:171 / ExcelTools.cs:1124-1135 | **Fixed by PP-17** — `scope:'sheet'|'workbook'` added. |
> | 9 | `delete_table` | `.Unlist()` keeps all the data — the name/description could read as "remove the table and its data." | Misleading description | entry.ts:217 / ExcelTools.cs:1085-1088 | **Fixed by PP-18** — `deleteData`+`shift` params added; default behavior (keep data) unchanged but now accurately described. |
> | 10 | `add_sparkline` (`targetCell` omitted) | Defaulted to the *same* cells as `dataRange`. | Undocumented default | entry.ts:212 / ExcelTools.cs:728-744 | **Fixed by PP-18** — `targetCell` now required. |
> | 11 | `add_conditional_format` (`kind:"text"`) | Always used "contains" — no starts-with/ends-with/not-contains. | Capability gap, undocumented | ExcelTools.cs:418-423 | **Fixed by PP-14** — all 4 `XlContainsOperator` match modes. |
> | 12 | `add_conditional_format` (`kind:"duplicate"`) | Hardcoded to highlight duplicates only. | Capability gap, undocumented | ExcelTools.cs:427-430 | **Fixed by PP-14** — duplicate/unique mode. |
>
> **PowerPoint (`PowerPointTools.cs` / `entry.ts`)**
>
> All tools had exact 1:1 name correspondence between schema and handler; the gaps were
> narrower and more contained than Word's or Excel's:
>
> | # | Tool | Issue | Class | File:line (schema / handler) | Status |
> |---|---|---|---|---|---|
> | 1 | `edit_chart` (`chartType`) | Unrecognized/typo'd value silently ignored, tool still returned "Chart updated." | Silent no-op + false success | PowerPointTools.cs:510-513 | **Fixed by PP-21** — unknown value throws; the underlying `PptChartTypeMap` "bar" → 51 (xlColumnClustered, should be 57/xlBarClustered) mapping bug was also fixed. |
> | 2 | `edit_chart` (`legendPos`) | Bare `{type:'string'}`; handler expected exactly `"none"/"r"/"t"/"l"`, any natural phrasing (`"right"`, `"bottom"`) silently fell into the bottom-position branch. | Undocumented schema + silent no-op | entry.ts:356 / PowerPointTools.cs:519-531 | **Fixed by PP-21** — natural names + short aliases both accepted, unknown throws. |
> | 3 | `add_chart` (`kind`) | Unrecognized value silently defaulted to `"bar"`, no enum declared. | Undocumented schema + silent fallback | PowerPointTools.cs:442 | **Fixed by PP-22** — enum'd from the corrected chart-type map, matches `edit_chart.chartType` exactly. |
> | 4 | `add_smartart` (`layout`) | Unrecognized value silently defaulted to `"list"`, no enum declared. | Undocumented schema + silent fallback | PowerPointTools.cs:564 | **Fixed by PP-22** — throws on unknown key, and a distinct error when a key is valid but the display name isn't found in the install's gallery (possible non-English Office). |
> | 5 | `edit_table_structure` (`kind`) / `edit_table_style` (`borderPreset`) | Documented only in free-text description, no JSON-schema `enum`. | Undocumented schema | entry.ts (both) | **Fixed by PP-22** — both enum'd; `borderPreset` also widened to include `outline`. |
>
> `set_element_style`'s missing `underline`/`align`/`fontFamily` was **not** a
> schema/handler mismatch — both sides consistently omitted them. **Fixed by PP-20**
> (underline/align/fontFamily/shadow/baselineOffset added; strikethrough deliberately
> not — this PIA has no `TextFrame2` on `Shape`, confirmed via a `CS0234` compile
> failure).
>
> **Pattern across all three:** the two gateway tools (`apply_commands` in Word,
> `propose_operations` in Excel) both used a completely untyped `items: {type:'object'}`
> schema for their batched sub-commands. PP-5 replaced both with real structural
> per-kind JSON Schema. The single highest-value fix identified by the original audit —
> replacing free-text parameter descriptions with real per-kind JSON-schema shapes — is
> what PP-5 did, and it's what let every other PP item above enumerate its own closed
> value sets.

### 2026-08-24 — PP-8: `docs/tool-surface-todo.md` retired

That checklist was written against an early snapshot and its implementation counts were
badly out of date (it claimed "9 of 65" Excel operations when essentially all of them
were implemented). Planning off it produced work items for things that already shipped.
It now just points to `docs/ai-tool-surface.md`, which is verified against current
source and carries the tool-by-tool comparison, the full schema audit, and the
project's scope boundary (what is deliberately out of scope).

### 2026-08-24 — PP-5: gateway tool schemas landed

The structural root cause the schema-vs-implementation audit pointed at — both gateway
tools' `commands`/`operations` items being a bare `{type:'object'}` with the entire
per-kind contract living only in prose — is fixed. `apply_commands` and
`propose_operations` now carry real per-kind JSON Schema (`WORD_COMMAND_SCHEMAS` in
`WordAiAddIn/web-src/entry.ts`; `EXCEL_OPS` + `opSchemas`/`opsDescription` in
`ExcelAiAddIn/web-src/entry.ts`), and `kind` parsing in both `ApplyCommands`/
`ProposeOperations` moved inside the per-command try/catch with a required-field
precheck (`WordTools.cs`'s and `ExcelTools.cs`'s `RequiredFields`/`ValidateRequired`) —
so a malformed command now fails only itself, with a specific error naming the missing
field, instead of aborting the whole batch or silently reaching a COM handler.

Cross-checked exhaustively, not sampled: every `case` in `WordTools.cs`'s
`ApplyCommands` switch (12 at the time) had exactly one matching entry in
`WORD_COMMAND_SCHEMAS`, and vice versa; same for `ExcelTools.cs`'s `ProposeOperations`
switch (51 at the time) against `EXCEL_OPS`. Both diffs were empty. Excel's schema uses
a grouped variant (`DETAILED_KINDS` in `entry.ts`) rather than full `oneOf` detail on
all kinds — a full-detail schema measured ~4,870 added tokens (cl100k_base), over the
~4k budget PP-5 set as the threshold; only the 7 highest-ambiguity kinds
(`format_range`, `add_conditional_format`, `add_chart`, `edit_chart`, `add_shape`,
`set_data_validation`, `add_pivot`) get full structural detail. The grouped version
measured ~2,120 added tokens.

### 2026-08-24 — Provider/settings note (PP-0)

The original architecture write-up described each `entry.ts` as hardcoding
`streamOpenAiCompatible` against a local test endpoint, with `onSettingsSave` a stub.
That was accurate before PP-0's shared-app-shell refactor. As of PP-0 landing
(2026-08-24), provider/model/key selection, the settings screen, and the transport all
live once in `shared/web-src/app-shell/` (`getSettings` / `makeTransport` /
`onSettingsSave`, persisted in WebView `localStorage`), shared by all four add-ins
including Outlook. The main doc's Architecture section has been updated in place to
state this directly rather than carrying the old paragraph plus a correction.

### 2026-08-24 — PowerPoint section marked stale (PP-19/PP-24), later fully rewritten

At the time, PowerPoint's tool table predated PP-19 (`delete_slide`/`move_slide`/
`duplicate_slide`) and PP-24 (`set_slide_layout`/`set_slide_transition`/
`add_animation`/`read_animations`/`edit_animation`), plus `duplicate_element`/
`copy_element`/`move_element`/`copy_element_style` (cross-slide shape copy/move via
PowerPoint's native clipboard — an explicit exception to this codebase's usual rule
against using the real Windows clipboard, since it reproduces the exact underlying
OOXML the same way Ctrl+C/Ctrl+V does). Rather than rewrite the table line-by-line at
the time, a pointer note was left directing readers to `docs/superpowers/plans/
2026-08-24-pp24-powerpoint-layout-transitions-animations.md` and
`docs/superpowers/plans/STATUS.md` for current state.

**This restructuring (2026-09-27) replaced the pointer with a full rewrite**, verified
directly against `PowerPointAiAddIn/web-src/entry.ts`: the PowerPoint tool surface has
grown further since PP-24 (list_layouts, read_master_elements, add_master_element,
set_master_element_transform, remove_master_element, read_group, group_element,
set_slide_notes, set_headers_footers — none of which had ever been recorded here as a
dated update) to a current total of 49 tools. See the main doc's PowerPoint section for
the authoritative current table.

### 2026-08-26 — Search/replace + heading outline

A user-reported gap: Word had no read-only search tool at all (only `apply_commands`'
mutating `find_replace`, which also under-reported its replacement count — always 0 or
1 regardless of how many occurrences it actually replaced, now fixed to loop and count
accurately), Excel's `find_cells` had no write-side counterpart, and PowerPoint had
neither search nor find/replace of any kind. Added: Word's `find_text` (read-only
search) and `get_headings` (Navigation-Pane-style heading outline, `[index] H<level>:
text`); Excel's `find_replace` op under `propose_operations` (text-value cells only,
never formulas); PowerPoint's `find_text` (read-only, across shape text + notes) and
`replace_text` (every text-frame shape + notes, NOT table cells or SmartArt node text —
those keep using `edit_table_cell`/`edit_smartart`). Not re-verified against every
section at the time — see the main doc's current tool tables for where these tools
ended up.

### 2026-08-27 — Word `find_text`/`get_headings` performance fix

Both tools' first cut scanned every paragraph via positional `Paragraphs[i]` indexing —
`Paragraphs` is not a real array in Word's COM object model, so each indexed access
re-walks the document from the start, turning a full scan into roughly O(n²)
internally. A user reported this as Word visibly freezing on a large document (the
automation call runs synchronously on Word's own UI thread, so it can't pump its
message loop until the call returns). Fixed: `find_text`'s plain-substring path now
uses Word's native `Range.Find` engine (the one behind Ctrl+F) via a
`wdFindStop`-wrapped `Execute()` loop, costing work proportional to match count, not
document size; `get_headings` now uses `Range.GoTo(wdGoToHeading, wdGoToNext)` — the
same internal heading index that powers the Navigation Pane / "Browse by Heading" — so
it never visits a non-heading paragraph at all. Both still need to report a 0-based
paragraph index matching `read_blocks`/`apply_commands`' convention; resolving that no
longer uses `Paragraphs[i]` either — a shared `ParagraphIndexResolver` marches forward
via the cheap `Paragraph.Next()` chaining method, visiting each paragraph at most once
across a whole call. `find_text`'s `regex:true` path still needs a per-paragraph scan
(Word's Find has no regex mode, only its own more limited wildcard syntax), but via
that same cheap forward chain now, not positional indexing.

### 2026-08-27 — Excel `find_cells`/`find_replace` performance + scope

A related but milder issue than Word's: `find_cells`' plain-substring path (and
`find_replace`, added earlier the same day) read `.Text`/`.Formula` on every single
cell in a sheet's `UsedRange` via a `foreach` loop. Not the same O(n²) bug (`foreach`
over `Range.Cells` is a real enumerator, not positional re-indexing), but still one COM
round-trip per cell regardless of match count, which adds up on a sheet with a large
`UsedRange`. Fixed: the plain-substring path now uses Excel's own native
`Range.Find`/`FindNext` (the engine behind Ctrl+F/Ctrl+H) — `xlValues`/`xlFormulas` are
separate native passes (Excel's `Find` only searches one `LookIn` mode per call),
de-duped by address for `look_in:"both"`. `regex:true` still needs the per-cell scan
(same reason as Word — no regex mode, only wildcards). `find_replace` mirrors this:
native `Find`/`FindNext` locates candidate cells fast, then only the matched cell's
literal text `Value2` is read/replaced directly (same safety scope as before — never a
formula, since a numeric/date/formula cell that merely *displays* a match via its
formatted text has a non-string `Value2` and is skipped). **Also, per user request, a
scope default change**: both tools previously searched the whole workbook when
`sheetId` was omitted; they now default to the **active sheet only** (matching Ctrl+F/
Ctrl+H's default "Within: Sheet"), with a new `allSheets` boolean to opt into
workbook-wide search ("Within: Workbook") — `sheetId` still names one specific sheet
directly, active or not. (This scope default is reflected directly in the main doc's
Excel tool table now.)

### 2026-08-27 — Word: the same positional-indexing bug found in 3 more places, plus a `read_blocks` cap

A broader sweep for the same `Paragraphs[i]`-indexing anti-pattern (per user request)
turned up three more hotspots, all fixed the same way (`Paragraph.Next()` forward
chaining instead of positional indexing):

- **`ResolveTargetParagraphs`** — the shared Target matcher behind `apply_commands`'
  `updateTextStyle`, `updateParagraphStyle`, `deleteBlocks`, `createParagraphBullets`,
  and `deleteParagraphBullets` — scanned every paragraph in the document via positional
  indexing regardless of how narrow the Target filter was. Its return type also changed
  from `List<int>` to `List<(int Index, Word.Paragraph Paragraph)>` — every caller used
  to re-look-up `paragraphs[i + 1]` per match after this function already had the
  paragraph in hand during its scan; now it just hands the object back, removing that
  second round of positional lookups too. All 5 call sites updated to match.
- **`read_blocks`'s plain-`"text"` format** (the default) had no upper cap at all —
  only `format:"html"` was capped, at 100 paragraphs. Added a 1000-paragraph cap for
  text mode (not independently benchmarked the way `read_formats`' 200-cell cap or html
  mode's 100-paragraph cap were — chosen conservatively, documented as such in the
  code) and switched its indexing to the same forward-chaining walk.
- `find_text`'s tool description and Word's system prompt now explicitly tell the
  model that `find_text`'s returned `[index]` is the exact same 0-based paragraph index
  `read_blocks`/`replace_blocks`/`apply_commands`' `Target.blockIndexes` use (no
  translation needed), and to prefer `find_text` over reading a large range blindly.

### 2026-08-27 — Phase 0 complete: shared pure-logic seam, plus two tool-facing bug fixes

Implemented `docs/superpowers/plans/2026-08-27-phase0-test-seam.md` in full (6 tasks, 6
commits, `dotnet test` 23 → 90 passed). No tool *schema* changed. Two deliberate
behavior changes did land, each its own commit:

- **Hex-color validation** (`propose_operations`/`apply_commands`/`set_element_fill`
  etc., any op taking a hex color): a malformed color used to reach the model as a raw
  .NET exception — `"#abc"` threw `ArgumentOutOfRangeException`, `"#GGGGGG"` threw
  `FormatException`, neither naming the bad value or the expected format. Now: 3-digit
  CSS shorthand (`"#abc"` == `"#aabbcc"`) is accepted, and anything else throws a clean
  `ArgumentException` naming the offending value and the expected `#RRGGBB` form.
- **Word's `apply_commands`**: `set_bold`/`set_italic`'s `"value": null` and
  `set_heading`'s `"level": null` used to reach `GetBoolean()`/`GetInt32()` and throw an
  opaque `InvalidOperationException` where the existing "missing required field"
  message belongs. Now rejected with that same clean message. **Deliberately not a
  global rule** — Excel's `set_cell`/`set_range` legitimately use `"value": null` to
  clear a cell, so Excel's `RequiredFields` validation is untouched; rejection is
  opt-in per field via a new `NonNullFields` table, Word-only so far.

Also: a `Dictionary<string, MsoAutoShapeType>` cannot cross an assembly boundary from
`OfficeAi.Shared` into the VSTO app projects — confirmed via `CS1769` ("embedded
interop type" as a generic type argument). Excel's and PowerPoint's shape-type maps
(previously separately duplicated, one with two extra aliases) are now one union table
in `OfficeAi.Shared.ShapeTypes`, carried as `Dictionary<string, int>`, each app casting
to `MsoAutoShapeType` at its own call site — the same split `ColorUtil` already used
for color. Side effect: both apps' "unknown shapeType" error now lists the full union
of valid names, not just each app's own subset.

### 2026-08-27 — Word gains `barStacked`; chart types unified

Pulled forward out of Phase 2 at user request, in two commits.

- **`feat(word)`:** Word's `edit_chart` now accepts **`barStacked`** (`xlBarStacked`,
  58). Word's chart-type map had 7 entries to Excel's and PowerPoint's 8, so Word could
  not draw a stacked bar chart while the other two could. Word's `entry.ts` enum
  honestly advertised only 7, so schema and handler agreed — a genuine **capability
  gap**, not a mismatch. Its `chartType` enum gains the value. Secondary fix:
  `read_chart`'s reverse type-code lookup used the same map, so reading an existing
  stacked bar chart reported *"unrecognized chart type code 58"* instead of
  `barStacked`.
- **`refactor(shared)`:** the three per-app maps are now one
  `OfficeAi.Shared.ChartTypes.ByName`. Sequencing the Word fix first was deliberate — it
  made all three byte-identical, so the extraction itself changed no behavior at all.
  Tests assert the exact `xlChartType` codes (unlike `ShapeTypes`, where raw ints would
  be brittle): a wrong code here is a **silent** wrong result, and that bug has shipped
  here before — PowerPoint's copy mapped `"bar"` to `51`/`xlColumnClustered` instead of
  `57`, so `chartType:'bar'` drew a column chart *and reported success* (this specific
  bug was independently rediscovered and fixed again by PP-21 a few weeks later — see
  the schema audit above). Tests also guard drift against every app's `entry.ts` enum
  in both directions. `dotnet test` 90 → 102.

**Not verified against live Office at the time** — `barStacked` followed the identical
code path as the seven chart types already supported (a `Dictionary` lookup handing an
int to the same COM property), and all three apps built clean, but no live Word
instance was reachable to confirm a stacked bar chart actually rendered.

### 2026-08-27 — Phase 2 duplication cleared; PowerPoint SmartArt parity; `set_bullet`

Three changes, at user request.

**1 — Remaining cross-app duplication united into `OfficeAi.Shared`:** `ComRetry` (was
`TransientComHResults`+`RetryTransientCom`, duplicated Word/PowerPoint — the copies
differed only in whether they took a `label`) and `SmartArtLayouts` (was
`SmartArtLayoutNames`, byte-identical in both). Resolving a layout key to a live object
stays app-side: it walks `Globals.ThisAddIn.Application`, and `Globals` is generated per
project. With the chart-type maps done earlier the same day, the Phase 2 duplication
inventory was empty at this point. `dotnet test` 102 → 114.

**2 — PowerPoint gained `read_smartart` and `edit_smartart`**, reaching parity with
Word (which already had all three). One real adaptation, not a copy: Word finds
SmartArt in a flat document (`InlineShapes`+`Shapes`), while PowerPoint's shapes are
per-slide — so a diagram is addressed by `(slideIndex, smartArtIndex-within-that-slide)`,
matching how every other PowerPoint tool addresses shapes. `read_smartart` is read-only
and is registered in both the client `READER_TOOLS` list and the C#
`AlwaysAllowedTools` gate; registering only one would have the client offer a tool the
server then refuses.

**NOT VERIFIED AGAINST LIVE POWERPOINT at the time** — compiled clean in Debug and
Release, but the SmartArt COM paths (`Nodes.Add`/`Delete`, `.Color`/`.QuickStyle`/
`.Layout` assignment) had never been exercised against a running instance. `set_style`
and `set_layout` resolve against the local install's gallery by English display name,
so a non-English Office fails them with the "not in gallery" error.

**3 — `apply_commands` accepts `set_bullet`.** A model sent that kind and got back only
"unknown command kind" — a dead end. The guess was reasonable: its neighbours are
`set_bold`/`set_italic`/`set_heading`, while the real commands are camelCase
`createParagraphBullets`/`deleteParagraphBullets`. `set_bullet` became an accepted
alias (same `target`, plus `value:true|false` picking which underlying command runs)
and is in the schema so it is discoverable, not merely forgiven. Separately, both
Word's `apply_commands` and Excel's `propose_operations` now list every valid kind when
rejecting an unknown one, so any wrong guess is recoverable on retry — the same defect
class as the malformed-hex-colour error fixed in Phase 0.

### 2026-08-27 — Phases 1+3: tool files split into partial classes

The three `*Tools.cs` files were split by tool area into `partial class` file sets —
Word 2,531 lines → 10 files (2,676 total), Excel 2,183 → 10 (2,310), PowerPoint 1,820 →
10 (1,956). Largest file was `WordTools.Content.cs` at 452 lines — 2 over the ~450
target; left as-is rather than re-splitting for 2 lines. Structure only: no method
body, tool schema, `entry.ts`, or system prompt changed, and `dotnet test` stayed at
114. Each split was verified by an order-independent member-set diff (sorted
declaration list identical before and after), which is what makes an otherwise
unreviewable move-diff trustworthy. All three apps rebuilt clean in both Debug and
Release afterward — Release matters because `deploy/package.ps1` only signs manifests
in that configuration.

Two things worth knowing before adding code to these projects (**still true — see
`CLAUDE.md` at the repo root**):

- The `.csproj` files are **classic format with explicit `<Compile Include>` items** —
  they do not glob `*.cs`. A new file must be added by hand, or it is silently not
  compiled and the error points at a call site elsewhere. `tools/split-partial.py` did
  this split and can do further ones; it refuses to run unless every member is
  assigned to exactly one destination.
- Word's `ListChartShapes`/`ListSmartArtShapes` stayed `internal` (not `private`)
  because `WordAiAddIn/TaskPaneHost.cs` calls them from outside the class — the
  partial-class split keeps that legal since it's still one class, one assembly.

The cross-app duplication this phase might otherwise have surfaced (`ComRetry`,
`SmartArtLayouts`) was already resolved into `OfficeAi.Shared` the same day, before
this split ran — see the Phase 2 entry above.

### 2026-09-17 — Outlook gains color-tag/category tools

Three new tools, `OutlookAiAddIn/OutlookTools.Categories.cs`. `list_color_categories`
(read-only) lists the profile's master `Namespace.Categories` list — name + friendly
color name (the 26 `OlCategoryColor` values, mapped in that file since
`OfficeAi.Shared` has no Outlook PIA reference, same split as `ColorUtil` uses for
RGB). `set_event_categories` sets/clears an `AppointmentItem`'s `Categories` string
(the colored block shown on a calendar event) — Full autonomy only.
`set_category_color` creates a new tag or recolors an existing one via
`Categories.Add`/`Category.Color` — also Full autonomy only. Verified against a live
Outlook client (category enumeration, tag creation, and event `Categories` round-trip)
before wiring into the tool switch; `dotnet test` unaffected (no new pure logic — the
color-name map is a small dictionary, not extracted for unit testing).

### 2026-09-27 — Outlook reschedule tools, `set_event_availability`, `delete_email` permanent tier; undo/redo for Word, PowerPoint, Outlook (and briefly Excel)

Moved verbatim out of `docs/ai-tool-surface.md` when `main` was merged into the current-state restructure on 2026-10-02; the facts that are still true now live in that file's tables. Cross-references inside the quoted blocks ("the dated Update block above", "see below") point at that file's old layout.

> **Update 2026-09-27 (Outlook gains reschedule tools):** two new tools in
> `OutlookAiAddIn/OutlookTools.Calendar.cs`, filling the "no way to move an
> existing event" gap (previously only `draft_event`/`create_event` could
> place a *new* event at a time). **Fully replaced 2026-09-28 by
> `edit_event`/`draft_edit_event`** (see that dated update below) — everything
> described in this block carried over unchanged onto the new tool names.
> `draft_reschedule_event` (Draft tier) sets `Start`/`End` on the resolved
> `AppointmentItem` and `Display(false)`s it unsaved, same shape as
> `draft_event` — the user reviews the moved time and decides whether to
> save/send. `reschedule_event` (Full autonomy only, `SendTierTools`) sets
> `Start`/`End` then `.Save()` (no attendees) or `.Send()` (attendees present)
> to dispatch the reschedule notice, mirroring `create_event`'s
> attendee-present branch. Both required `event_id`, `start`, and `end`, with
> a clean `IsError` naming whichever of `start`/`end` is missing rather than
> defaulting or throwing a raw COM exception.
>
> **Organizer-only, by design.** Both tools first check
> `AppointmentItem.MeetingStatus`. On `olMeetingReceived` (the user is only an
> *attendee* on someone else's meeting, not the organizer) they refuse up
> front with an `IsError` explaining the user has no authority to move it and
> pointing at Outlook's own "Propose New Time" UI — see "Excluded / deferred"
> below for why no tool attempts that path itself. `olNonMeeting`/`olMeeting`
> (the user's own appointment or meeting) go through unchanged.
>
> **Verification status: unproven at runtime**, same caveat as the rest of
> this section (see "Unproven at runtime" below) — no live Outlook was
> available to exercise either tool's COM calls. One piece *was* checked
> concretely rather than assumed: whether the interop PIA exposes a "Propose
> New Time" member at all. .NET reflection against the actually-referenced
> `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA (same verification
> discipline `CreateEvent`'s comments use) confirmed neither
> `AppointmentItem`/`_AppointmentItem` nor `MeetingItem`/`_MeetingItem` expose
> anything with "Propose" (or "Counter"/"NewTime") in its name —
> `_AppointmentItem.Respond` only takes an `OlMeetingResponse`
> (Accept/Decline/Tentative), no counter-proposal overload. "Propose New
> Time" is real Outlook-client functionality, but it isn't reachable through
> the classic COM object model this add-in automates, which is why the
> `olMeetingReceived` case is a hard refusal rather than a half-working
> attempt.
>
> **Update 2026-09-28: fixed a gap in the organizer-authority check, wired
> into undo/redo.** The original check only excluded the exact
> `olMeetingReceived` value; `OlMeetingStatus` actually has 5 values
> (`olNonMeeting`, `olMeeting`, `olMeetingReceived`, `olMeetingCanceled`,
> `olMeetingReceivedAndCanceled`), so an attendee's copy of a meeting the
> organizer has since canceled (`olMeetingReceivedAndCanceled`) — and the
> organizer's own canceled copy (`olMeetingCanceled`) — both fell through to
> the silent `.Save()` path, reporting a `Mutated: true` "Rescheduled" success
> on a canceled or not-actually-yours meeting. Both tools now check for
> either canceled status first (a clean `IsError`: "has been canceled, so
> there's nothing to reschedule"), then the (now also two-valued)
> received-meeting check. Also merged with PR #23 (undo/redo) and #20
> (`set_event_availability`), which this branch predates: `reschedule_event`
> now records a `Start`/`End` snapshot (undo-able) on its `.Save()` branch and
> a barrier (not undo-able) on its `.Send()` branch, mirroring `create_event`'s
> two branches. `draft_reschedule_event` needs no wiring — like `draft_event`,
> it never saves/sends anything itself. (Carried over unchanged onto
> `edit_event`/`draft_edit_event` later the same day — see the dated update
> below.)

> **Update 2026-09-27 (Outlook gains `set_event_availability`):** one new
> tool, added to `OutlookAiAddIn/OutlookTools.Categories.cs` right beside
> `set_event_categories`/`set_category_color` since it's the same shape (a
> single-property `AppointmentItem` mutation with a friendly-name map for an
> Outlook-PIA-only enum). Sets `AppointmentItem.BusyStatus` — the calendar's
> "Show As" dropdown — to one of the 5 `OlBusyStatus` values (`olFree`,
> `olTentative`, `olBusy`, `olOutOfOffice`, `olWorkingElsewhere`; names
> confirmed via .NET reflection against the referenced
> `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA, not assumed) via a
> case-insensitive friendly-name lookup (`free`/`tentative`/`busy`/`out of
> office`/`working elsewhere`); an unrecognized string returns a clean
> `IsError` result listing the valid names rather than throwing, unlike
> `set_category_color`'s `ParseColor` (which throws and relies on the outer
> `ExecuteAsync` catch). Unrelated to `Categories`/color tags — a distinct
> `AppointmentItem` property entirely. Placed in **Draft tier**
> (`DraftTierTools` + `entry.ts`'s `commentOnlyExtraTools`), not Full
> autonomy: it's purely local (`.BusyStatus` + `.Save()`, no `.Send()`),
> the same risk profile that put `set_event_categories`/`set_category_color`
> in Draft tier despite the "Mutating tools" table heading below still
> reading "Full autonomy only" (that heading predates the 2026-09-19
> four-tier gate for those two tools and was already stale before this
> change — not fixed here to keep this update focused). **Update
> 2026-09-28: verified against a live Outlook client** — all 5
> `OlBusyStatus` values round-trip correctly via `.BusyStatus` + `.Save()`,
> case/spacing-tolerant parsing confirmed, invalid values return a clean
> `IsError` instead of throwing, and saving a `BusyStatus` change on a
> meeting the user organizes (has attendees) does not prompt or notify
> attendees. Also wired into the undo/redo stack the same day (see that
> section below) after merging with PR #23.

> **Update 2026-09-27 (delete_email's permanent path split to its own tier):**
> another instance of the same class of gap closed by `aeae77c`/`b3fc5d2`
> (category tools left one gate too low) — here `delete_email` as a whole sat
> in `DraftTierTools`, but the tool bundles two risk classes under one name:
> `permanent: false` (default) just moves the message to Deleted Items, fully
> reversible, correctly Draft-tier; `permanent: true` additionally calls
> `.Delete()` from there, irreversible from within Outlook ("may still be
> server-recoverable" per its own result text — not a claim it can be undone
> here), which belongs at Full autonomy alongside `send_email` and friends,
> not one gate below it. `OutlookTools.cs`'s `ExecuteAsync` tier check is
> otherwise purely name-based (`AlwaysAllowedTools`/`DraftTierTools`/
> `ApprovalTierTools`/`SendTierTools`, checked before `input` is inspected at
> all) — added one narrow, input-aware special case immediately after the
> name-based check: if `name == "delete_email"` and `input.permanent == true`
> and the caller's mode is below Full autonomy, block with a message naming
> Full autonomy specifically (not delete_email in general — the tool is still
> fine at Draft only for the non-permanent path). Deliberately not
> generalized into a per-argument gating system for every tool; `delete_email`
> stays in `DraftTierTools` and in `entry.ts`'s `commentOnlyExtraTools`
> unchanged, since the tool overall is still reachable from Draft only.
> `entry.ts`'s `delete_email` description now says plainly that
> `permanent: true` requires Full autonomy, so the model doesn't attempt it
> needlessly at a lower tier and get a confusing runtime block. No existing
> automated test covers Outlook's tier-gate logic (it lives in the VSTO
> project, not `OfficeAi.Shared`); verified by code review plus
> `OutlookAiAddIn` MSBuild.

> **Update 2026-09-27 (undo/redo tooling — Word, PowerPoint; a custom
> undo/redo stack for Outlook; Excel's removed 2026-09-28):** added
> `undo_last_action`/`redo_last_action` to Word and (reversing this
> document's own earlier "no undo/redo tool" conclusion for it) PowerPoint,
> plus a custom stack of the assistant's own actions for Outlook. **None of this is runtime-verified
> against a live Office/Outlook client** — no live instance was reachable;
> every mechanism below was instead confirmed by .NET reflection directly
> against the exact PIA versions these projects reference (15.0.0.0), the
> same standard of evidence this document already applies elsewhere (e.g.
> the chart-type-map and SmartArt-parity updates above), and by a clean
> Debug build of all four touched projects plus their esbuild bundles.
>
> - **Word** (`WordAiAddIn/WordTools.History.cs`): the initial plan for this
>   feature assumed `Application.Undo()`/`Application.Redo()` exist — they do
>   **not** (reflection against the referenced PIA finds no such method on
>   `_Application`/`ApplicationClass` at all, only an unrelated `UndoRecord`
>   property for grouping automation edits into one user-visible undo entry).
>   The real, callable methods are one level down, on `Document`:
>   `Document.Undo(ref object Times)` / `Document.Redo(ref object Times)`,
>   both returning `bool`. The `Times` parameter is COM-optional (confirmed
>   via `ParameterInfo.IsOptional`), so C# calls them with no arguments -
>   `ActiveDoc.Undo()` / `ActiveDoc.Redo()` - exactly like real-world Word
>   automation code. The `bool` return gives an honest "Undid the last
>   action." vs. "Nothing to undo." (same for redo) instead of always
>   claiming success. Gated the same as every other Word mutating tool:
>   blocked in Read Only/Comment Only, available from Track Changes upward.
> - **Excel — removed (2026-09-28).** The first version dispatched
>   `Application.CommandBars.ExecuteMso("Undo"/"Redo")` (Excel has no
>   `Application.Redo()`, and `Application.Undo()` returns `void`). That
>   can't work for the assistant's edits: every write this add-in makes goes
>   through the object model (`propose_operations`, find/replace, and so on),
>   and Excel clears its whole undo stack on any object-model change. So
>   right after an AI edit there is nothing to undo, and later the tool would
>   undo the *user's* next manual edit instead. `Application.OnUndo` only
>   takes a VBA macro name, so it isn't usable from a C# add-in. A snapshot-
>   based custom undo was considered and not pursued; the tools were removed.
> - **PowerPoint** (`PowerPointAiAddIn/PowerPointTools.History.cs`) —
>   **reversing this document's earlier conclusion that PowerPoint should be
>   excluded entirely.** That conclusion was correct that no
>   `Application.Undo()`/`Redo()` method exists anywhere in the PowerPoint
>   interop surface (still true, reconfirmed here), but missed a distinct
>   mechanism: `Application.CommandBars.ExecuteMso(string)` /
>   `GetEnabledMso(string)` — the generic Office 2007+ ribbon-command
>   dispatch API, exposed via `Microsoft.Office.Core.CommandBars` (the
>   "Office" PIA reference already in every one of these four projects),
>   which every main Office host's `Application.CommandBars` returns,
>   PowerPoint included. Confirmed real and callable by reflection: the
>   `CommandBars` property on PowerPoint's `_Application` returns
>   `Microsoft.Office.Core.CommandBars`, and that type's `_CommandBars`
>   interface declares both `Void ExecuteMso(String)` and `Boolean
>   GetEnabledMso(String)`. `ExecuteMso("Undo"/"Redo")` dispatches by the
>   same ribbon-command ID the real Undo/Redo buttons and Ctrl+Z/Ctrl+Y use
>   internally, so it works despite there being no direct method to call.
>   `GetEnabledMso` is checked first (mirroring Excel above) since
>   `ExecuteMso` itself returns nothing, so without that check every call
>   would have to claim success unconditionally. See
>   `PowerPointTools.Master.cs`'s `RemoveMasterElement` (2026-09-22 incident,
>   in the PowerPoint section below) for why this repo cares: a Slide Master
>   placeholder deletion had no code path at all to undo at the time, one
>   direct consequence of `Application.Undo()` genuinely not existing —
>   `ExecuteMso` is the mechanism that generally closes that "no way back"
>   gap going forward (it does not retroactively change that method's own
>   still-valid placeholder refusal, which stands for an unrelated reason —
>   see that method's comment). Gated the same as every other PowerPoint
>   mutating tool (Track Changes upward).
> - **Outlook** (`OutlookAiAddIn/OutlookTools.Undo.cs`, stack logic in
>   `OfficeAi.Shared/ActionHistory.cs`) — a **custom** undo/redo stack of
>   the assistant's own actions, not a native-undo wrapper. Native Undo via
>   `Explorer.CommandBars.ExecuteMso("Undo")` was tried and tested by hand
>   on 2026-09-28. It's a single slot tied to the window that toggles
>   undo/redo, it doesn't see object-model changes (a `move_email` followed by
>   the ribbon Undo did nothing), and after one use both our tool and the
>   ribbon button failed with "The operation cannot be performed because the
>   message has changed." Outlook has no API to put an object-model change
>   onto that slot.
>
>   How it works: an in-memory stack per mailbox chat (keyed like
>   `ModeByMailbox`, capped at 50, lost on restart). Each mutating handler
>   records one entry after its change succeeds:
>   - **Property snapshots** (before/after values, restored and `Save()`d):
>     `mark_email_read/unread` (`UnRead`), `flag_email_important` (full
>     `Importance`, so Low is preserved), `set_event_categories`,
>     `set_event_availability` (`BusyStatus`), `edit_event` (added 2026-09-28,
>     replacing `reschedule_event`) whenever the call doesn't send/barrier
>     (see below) — `Start`/`End`, `Subject`, `Body`, `Location`, whichever
>     subset was actually touched that call, via a dynamically-built props
>     list (replacing the old fixed `{Start,End}` `RescheduleProps` constant) —
>     including an occurrence-level edit via `occurrence_date`:
>     `RecurrencePattern.GetOccurrence()` returns a real item with its own
>     `EntryID` once saved, so this same mechanism covers it with no new entry
>     type — `set_reminder`, and `update_task` (task fields as a group, or the
>     flagged-mail fields).
>   - **`set_email_reminder`**: if the message wasn't flagged before, undo
>     calls `ClearTaskFlag()`, and redo calls `MarkAsTask` again.
>   - **Moves**: `move_email`, non-permanent `delete_email`, and `cancel_event`
>     on a plain appointment or an already-canceled event (no one to notify
>     either way, so canceling it is just a move to Deleted Items). The
>     folder is resolved by its own EntryID/StoreID via
>     `Namespace.GetFolderFromID`, and the new EntryID after each move is
>     rewritten into every entry for that item.
>   - **Created items**: `create_task` and `create_event` without
>     attendees. Undo moves the item to Deleted Items (recoverable); redo
>     moves it back.
>   - **`set_category_color`**: undo restores the old color, or removes a tag
>     the assistant created.
>   - **Barriers**: `send_*`, `create_event` with attendees, `edit_event`
>     and `cancel_event` on a still-active meeting the user organizes (both
>     send a notice to attendees), `accept/decline/tentative_meeting`, `delete_email
>     permanent:true`, — regardless of plain-appointment vs. meeting — any
>     occurrence-level cancellation via `cancel_event`'s `occurrence_date`
>     (added 2026-09-28: `RecurrencePattern.Exceptions`/`Exception` are
>     read-only via COM, confirmed by reflection, so a deleted occurrence has
>     no API to reverse it), and — again regardless of plain-appointment vs.
>     meeting — `edit_event` on the **whole series** whenever it includes a
>     time change and `event_id` resolves to a genuinely recurring master
>     (confirmed live 2026-09-28: Outlook rejects setting `Start`/`End`
>     directly on a recurring master at all, so this branch writes to
>     `RecurrencePattern` fields instead, which the undo mechanism can't yet
>     read/write). This is the one case where an otherwise-undo-able-looking
>     `edit_event` call (no attendees, no `occurrence_date`) is still a
>     barrier — only true non-recurring events, or a single occurrence, stay
>     undo-able via the dynamically-built props snapshot. `edit_event` is also
>     a barrier any time it results in a `.Send()` — including converting a
>     plain event into a meeting by adding `required_attendees`/
>     `optional_attendees` for the first time (added 2026-09-28, unverified at
>     runtime — see the dated update below). Undo stops at a barrier instead
>     of reaching past it.
>   - **Known gap — `set_event_categories`**: the snapshot only covers the
>     appointment's own `Categories` string. If the assigned name wasn't
>     already in the mailbox's master category list, Outlook auto-adds it
>     on `Save()` with an arbitrary color (see the mutating-tools table
>     below); undo restores the appointment but does not remove that
>     auto-created master category entry, which is a permanent side effect
>     undo can't see or reverse.
>
>   Before reversing, each entry checks that the item still holds what the
>   assistant left there. If it was changed since (by the user or anything
>   else), undo refuses rather than overwrite it. A failed or refused entry
>   is dropped, never retried. `redo_last_action` re-applies undone entries
>   until the next new action. Both tools are gated at Draft tier. **Not
>   verified against a live Outlook client.**

### 2026-09-28 — Outlook cancel tools, recurring-series support; `reschedule_event` replaced by `edit_event`

Moved verbatim out of `docs/ai-tool-surface.md` when `main` was merged into the current-state restructure on 2026-10-02; the facts that are still true now live in that file's tables. Cross-references inside the quoted blocks ("the dated Update block above", "see below") point at that file's old layout.

> **Update 2026-09-28 (Outlook gains cancel tools):** two more tools in
> `OutlookAiAddIn/OutlookTools.Calendar.cs`, right beside `edit_event`/
> `draft_edit_event` and reusing their `IsCanceledMeeting`/
> `IsReceivedMeeting` organizer-authority checks rather than duplicating them.
> `draft_cancel_event` (Draft tier) opens the item unchanged either way — for
> a meeting the user organizes, they cancel it themselves via Outlook's own
> Cancel Meeting/Send Cancellation buttons; for a plain appointment, via
> Delete. Refuses (clean `IsError`) on an already-canceled event (points at
> `cancel_event` instead — see below) or a meeting the user only attends
> (points at `decline_meeting`, not "Propose New Time" — canceling isn't
> something an attendee can do to someone else's meeting at all, unlike
> rescheduling where Outlook at least has a UI-level counter-proposal feature
> this add-in can't reach). (Originally set `MeetingStatus = olMeetingCanceled`
> unsaved before `Display(false)`, mirroring `draft_edit_event`'s (then
> `draft_reschedule_event`'s) unsaved `Start`/`End` — **removed the same day** after live testing showed
> Outlook persists that change when the Inspector closes even without the
> user clicking "Send Cancellation," silently canceling the meeting locally
> with attendees never notified. Confirmed live, fixed same-day by never
> mutating the item in `draft_cancel_event` at all.)
>
> `cancel_event` (Full autonomy only, `SendTierTools`) does it immediately:
> sets `olMeetingCanceled` and `.Send()`s the cancellation notice for an
> organized meeting, or skips straight to removal for a plain appointment;
> either way the item is then moved to Deleted Items (recoverable there),
> same as `delete_email`'s non-permanent path, not permanently deleted.
> Still refuses on a meeting the user only attends and isn't canceled yet
> (points at `decline_meeting`) — but **an already-canceled event (either the
> organizer's own `olMeetingCanceled` copy or an attendee's stale
> `olMeetingReceivedAndCanceled` one) is treated as cleanup, not refused**:
> there's nothing new to notify anyone of, so it just moves straight to
> Deleted Items, same as a plain appointment. Added the same day as the
> tools themselves, once live testing surfaced there was otherwise no way to
> dismiss a canceled event at all (not even the one this branch's own
> `draft_cancel_event` bug, above, could accidentally create).
>
> **Undo/redo:** `cancel_event` on a plain appointment, or on an
> already-canceled event, records a move (undo moves it back out of Deleted
> Items), matching `delete_email`'s pattern. `cancel_event` on a still-active
> organized meeting is a barrier, not a move snapshot — the cancellation
> notice already went out, so restoring the calendar entry would only
> half-undo the action and misleadingly imply it was fully reversed.
> `draft_cancel_event` needs no wiring — it never saves/sends.
>
> **Verification status: unproven at runtime**, same caveat as `edit_event`
> above — no live Outlook was available. One specific sequence is new and
> untested even relative to that PR: `cancel_event`'s organized-meeting branch
> calls `.Move()` right after `.Send()` on the same item, which no existing
> tool in this add-in does (every other `.Send()` call — `create_event`'s
> invite branch, `edit_event`'s notify branch — only reads properties
> off the item afterward, never mutates its folder). Whether Outlook still
> allows relocating a just-canceled-and-sent appointment is unconfirmed; the
> code fails soft (catches the exception, reports the notice went out
> regardless and tells the user to delete manually if the move didn't take) —
> see the manual test steps for this specifically.

> **Update 2026-09-28 (Outlook gains recurring-series support: occurrence
> targeting, recurrence creation):** two additions to
> `OutlookAiAddIn/OutlookTools.Calendar.cs` and one to
> `OutlookAiAddIn/OutlookTools.Compose.cs`, closing the "recurring occurrences
> share one EntryID" gap flagged in "Structural fragility" below (see that
> section's own correction, same date) and the "no way to create a repeating
> series" gap.
>
> **1 — Occurrence targeting.** `edit_event`/`draft_edit_event`/
> `cancel_event`/`draft_cancel_event` all gained an optional `occurrence_date`
> string param. A new `ResolveOccurrenceTarget` helper resolves it via
> `RecurrencePattern.GetOccurrence(DateTime)` — confirmed present and callable
> via .NET reflection against the referenced PIA, the same evidence standard
> this document uses elsewhere — into one specific occurrence, instead of
> always acting on the recurring master. Omitting `occurrence_date` still acts
> on the whole series (or a non-recurring event), unchanged from before this
> feature.
>
> **2 — Recurrence creation.** `create_event`/`draft_event` gained an optional
> `recurrence` object supporting all 6 `OlRecurrenceType` values (`daily`/
> `weekly`/`monthly`/`monthlyNth`/`yearly`/`yearlyNth`), validated by a new
> pure, unit-tested `OfficeAi.Shared.RecurrenceValidator` (23 tests) before any
> COM call — every validation failure names the specific field and the fix
> (e.g. `"recurrence.days_of_week is required for type \"weekly\"..."`, not a
> generic "invalid recurrence"). Applied via `AppointmentItem.GetRecurrencePattern()`
> before `.Save()`/`.Send()`, after attendee/`MeetingStatus` setup — **that
> ordering is a design choice, not yet verified against live Outlook** (see
> "Unproven at runtime" below). Series creation's own undo/redo is unchanged
> from `create_event`'s existing non-recurring behavior (a `RecordCreated`
> move-to-Deleted-Items entry on the master) — deleting a recurring master
> deletes the whole series, so no new undo mechanism was needed here either.
>
> **3 — Undo/redo asymmetry (the most important nuance of this feature).**
> Occurrence-level **reschedule** stays undo-able, and needed no new
> undo-entry type: `RecurrencePattern.GetOccurrence()`'s returned item becomes
> a real, independently-resolvable item with its own `EntryID` once saved, so
> the existing generic `SnapshotEntry`/`RecordSnapshot`/`ItemEntryIdOf`
> mechanism already works on it exactly like any other item. Occurrence-level
> **cancellation is always a barrier, never undo-able** — plain appointment or
> meeting occurrence, no exception either way. This is a deliberate asymmetry
> from whole-event cancellation (`cancel_event` without `occurrence_date` on a
> plain appointment stays undo-able, per the existing "Undo/redo tools"
> section). The reason: `RecurrencePattern.Exceptions`/`Exception` — the COM
> objects Outlook uses to track per-occurrence deletions — are entirely
> read-only. Confirmed via .NET reflection against the referenced PIA:
> `Exception` exposes only getters (`AppointmentItem`, `Deleted`,
> `OriginalDate`, `ItemProperties`) and no `Delete`/`Remove`/`Add` method
> anywhere, so there is no API to reverse a deleted occurrence once it's gone.
>
> **4 — A suspected bug fixed by reasoning, not by a live-confirmed observation
> (organizer-authority checks against the wrong item).** The four
> occurrence-aware methods originally checked `appt.MeetingStatus`/
> `IsCanceledMeeting(appt)`/`IsReceivedMeeting(appt)`, where `appt` is the
> resolved occurrence or master. A live reschedule attempt against what was
> believed at the time to be a recurring meeting occurrence failed with
> `COMException "Cannot save this item."` Reasoning about `GetOccurrence()`'s
> known-unreliable behavior for this property — rather than a live-confirmed
> observation on a verified meeting occurrence — identified a likely bug:
> `GetOccurrence()`'s returned occurrence item does **not** reliably report
> `MeetingStatus == olMeeting` even when it genuinely belongs to a recurring
> meeting with attendees. If so, this would cause `edit_event` on a
> meeting occurrence to wrongly take the `.Save()` branch — crashing with
> that same `COMException` — instead of `.Send()`, and would cause
> `cancel_event` on a meeting occurrence to silently delete it without ever
> notifying attendees. Fixed by checking
> `master.MeetingStatus`/`IsCanceledMeeting(master)`/`IsReceivedMeeting(master)`
> instead (the master is always reliable) — the actual mutation still targets
> `appt` (the occurrence or master, whichever `ResolveOccurrenceTarget`
> resolved). After applying this fix and testing again, the user confirmed
> the test event had actually been a plain (non-meeting) appointment all
> along — so the original failure had a different, unrelated cause (see
> point 5 below). **Caveat: this fix has never actually been exercised against
> a genuine meeting occurrence with attendees** — every live test run against
> this feature used a plain (non-meeting) recurring series. The code and its
> review are sound, but the specific scenario it was written to fix is
> unverified.
>
> **5 — A second real bug found and fixed (false-negative `.Save()`/
> `.Delete()` failures).** Outlook can throw `COMException "Cannot save this
> item."` from an occurrence's `.Save()` (in `EditEvent`, formerly `RescheduleEvent`) or `.Delete()`
> (in `CancelEvent`, both the plain-appointment and meeting-occurrence
> branches) **even when the mutation already persisted** — confirmed live: a
> reschedule call that reported this exception had, per a follow-up
> `list_events` call, actually moved the occurrence. Property setters on
> `AppointmentItem` write immediately via RPC, independent of `.Save()`/
> `.Delete()`'s own finalize step, which can fail on its own. Fixed by
> wrapping these calls in try/catch and re-verifying the real outcome (via
> `ResolveOccurrenceTarget`) before deciding what to report, instead of
> trusting the exception at face value.
>
> **6 — A third real, confirmed-live constraint: Outlook rejects reordering
> occurrences relative to each other.** Confirmed by reproducing the exact
> error manually in Outlook's own UI: *"Cannot reschedule an occurrence of the
> recurring appointment ... if it skips over a later occurrence of the same
> appointment"* — and separately confirmed that two occurrences of the same
> series also can't share the same calendar day. Both surface from automation
> only as the same generic `COMException "Cannot save this item."` — confirmed
> no `InnerException` carries the specific reason, checked via debug log
> across every occurrence of the failure. A new proactive check,
> `CheckOccurrenceReorderCollision`, now runs before `.Save()` is ever
> attempted (in both `EditEvent` and `DraftEditEvent`, formerly `RescheduleEvent`/
> `DraftRescheduleEvent`, occurrence-level only): it queries the Calendar folder for other occurrences
> of the same series between the occurrence's original date and the target
> date, and if any exist, returns an exact error naming the valid range
> instead of relying on Outlook's generic exception. This is Outlook-level
> behavior, not a limitation of this add-in's own code.
>
> **Tier placement unchanged.** `create_event`/`edit_event`/
> `cancel_event` stay Full-autonomy-only (`SendTierTools`); `draft_*` stay
> Draft-tier. `occurrence_date`/`recurrence` are additional scope on existing
> tools, not new risk categories — no edits to `DraftTierTools`/`SendTierTools`.

> **Update 2026-09-28 (`reschedule_event`/`draft_reschedule_event` fully
> replaced by `edit_event`/`draft_edit_event`):** a **full replacement, not an
> alias** — the retired tools are removed entirely from `entry.ts` and the C#
> switch in `OutlookAiAddIn/OutlookTools.cs`; `edit_event`/`draft_edit_event`
> are the only way to change an event's time (or anything else about it) going
> forward. Same tier placement as before (`edit_event` in `SendTierTools`,
> `draft_edit_event` in `DraftTierTools`/`entry.ts`'s `commentOnlyExtraTools`).
> Generalizes reschedule into a full edit: beyond `start`/`end` (still required
> together if either is given), a single call can now also change `subject`,
> `body`, `location`, and/or replace `required_attendees`/`optional_attendees`
> wholesale (not a diff/merge — the caller supplies the full new list each
> time, reading the current one first via `get_event` if it needs to preserve
> someone) — at least one of the seven optional fields must be given, else a
> clean `IsError` naming all seven. Attendee edits are whole-series/
> non-recurring only (`occurrence_date` + an attendee field together is a
> clean `IsError`); `subject`/`body`/`location`/`start`/`end` remain editable
> at the occurrence level, same as reschedule always was. Every mechanism
> `reschedule_event`/`draft_reschedule_event` had — `ResolveOccurrenceTarget`
> occurrence targeting, the whole-series `RecurrencePattern.PatternStartDate`/
> `StartTime`/`EndTime` fix, the false-negative `.Save()` retry via
> `ResolveOccurrenceTarget` re-verification, `CheckOccurrenceReorderCollision`,
> and the `IsCanceledMeeting`/`IsReceivedMeeting` organizer-authority checks —
> is reused unchanged by `edit_event`/`draft_edit_event`, not duplicated (see
> the "Occurrence targeting" / "whole-series `RecurrencePattern`" / "false-
> negative `.Save()`" / "reordering collision" points in the dated update
> above, now describing `edit_event`). New pieces specific to this change:
> a `ReplaceAttendees` helper — per-category, not a blanket clear: it removes
> existing `olRequired`-type `Recipients` entries only if `required_attendees`
> is given, and existing `olOptional`-type entries only if `optional_attendees`
> is given (omitting one leaves that category's existing attendees untouched;
> the organizer recipient is never touched either way), then re-adds via the
> existing `AddAttendees` helper `create_event`/`draft_event` already use, then
> `Recipients.ResolveAll()`; a call flips `MeetingStatus`
> to `olMeeting` and takes the `.Send()` branch (instead of `.Save()`) whenever
> the event is or becomes a meeting, setting `ForceUpdateToAllAttendees = false`
> explicitly first — confirmed via .NET reflection to be Outlook's own default
> for notifying only added/removed attendees, not everyone, but set explicitly
> here so correctness doesn't depend on that default never changing; and the
> fixed `RescheduleProps` (`{"Start","End"}`) undo/redo constant is gone,
> replaced by a dynamically-built props list per call, snapshotting only the
> fields the call actually touched (`DescribeEditEventChanges` builds the
> matching human-readable summary). A whole-series time change stays a barrier
> exactly as before (`RecurrencePattern` fields aren't reachable by
> `SnapshotEntry`); if other fields are *also* changed in that same call, the
> entire call stays a barrier too — no partial-undo of a mixed pattern-field +
> item-property change.
>
> **Unverified at runtime, ranked by consequence of being wrong** (per the
> design spec's risk-ordering, `docs/superpowers/specs/2026-09-28-outlook-edit-event-design.md`):
> 1. **`ReplaceAttendees`'s recipient-clearing** — highest risk, since a wrong
>    organizer-exclusion could drop the user's own organizer entry or fail to
>    actually clear old attendees, feeding directly into a `.Send()` with real
>    people.
> 2. **Converting a plain event into a meeting via `edit_event`** — adding
>    attendees to an *existing*, previously-saved item (unlike `create_event`,
>    which only ever adds attendees to a brand-new one) is an untested
>    combination.
> 3. **`subject`/`body`/`location` edits on a recurring master** — likely fine
>    (none of these are part of `RecurrencePattern`), but unverified.
>
> None of this has been exercised against a live Outlook client yet — same
> caveat as every other addition in this document's "Unproven at runtime"
> section below, which this update also folds into.

### 2026-09-29 — Outlook `tentative_meeting` and draft meeting responses (PR #29 redesign); `list_events` `mailbox` for shared calendars (PR #28 review, partly reversed)

Moved verbatim out of `docs/ai-tool-surface.md` when `main` was merged into the current-state restructure on 2026-10-02; the facts that are still true now live in that file's tables. Cross-references inside the quoted blocks ("the dated Update block above", "see below") point at that file's old layout.

> **Update 2026-09-29 (Outlook gains `tentative_meeting` + draft meeting-response
> tools, plus an optional comment on all three):** `OutlookAiAddIn/OutlookTools.Calendar.cs`'s
> `RespondMeeting` — previously hardcoded to a `bool accept` — was generalized to take
> the actual `OlMeetingResponse` value, so one shared helper now backs `accept_meeting`,
> `decline_meeting`, and the new `tentative_meeting` (same "Automate approvals" tier as
> the other two — it already calls `resp.Send()` to notify the organizer, same rationale
> as the original two). All three also gained an optional `message` parameter: when
> given, it's set as `resp.Body` before `.Send()`, a short comment attached to the
> accept/decline/tentative response. **Unverified live** whether the organizer actually
> sees this text on the delivered response — needs a real received invite to test
> against, not a self-organized item (same category of gap as this document's other
> "Unproven at runtime" entries below). A new `DraftRespondMeeting` helper mirrors
> `RespondMeeting` for Draft tier: same `appt.Respond(response, true, false)` call, but
> ends in `resp.Display(false)` for the user to review and send themselves instead of
> `.Send()`-ing directly — `draft_accept_meeting`/`draft_decline_meeting`/
> `draft_tentative_meeting`, added to `entry.ts`'s `commentOnlyExtraTools` alongside the
> existing draft tools, never record undo/redo (same contract as `draft_event`/
> `draft_edit_event`/`draft_cancel_event` — nothing is saved or sent until the user acts
> on the opened window). `message` pre-fills the same `resp.Body` there, subject to the
> same unverified-live caveat. A second, separate unverified-live risk applies to the
> draft tools specifically: it is **unconfirmed whether `AppointmentItem.Respond()`
> itself commits local calendar changes at call time** — independent of whether
> `.Send()` or `.Display()` is subsequently called — such as replacing the appointment
> with a new EntryID on accept/tentative, or removing/moving the original appointment
> on decline. If Outlook's COM implementation does this, the draft tools' "nothing
> persists until the user acts" contract would not actually hold, even though they
> correctly never call `.Send()`/`RecordIrreversible`/`RecordSnapshot`. Needs a live
> test: open each draft response, close the window without sending, then confirm the
> appointment still exists under the same EntryID, `ResponseStatus` is unchanged, and
> the item is still in the inbox/calendar as before — `draft_decline_meeting` most
> carefully, since decline is the destructive direction.

> **Update 2026-09-29 (PR #29 code review: `draft_respond_meeting` redesign,
> plus smaller `RespondMeeting`/wording fixes):** a code review of the
> tentative_meeting/draft-meeting-response work above (the immediately
> preceding "Update 2026-09-29" block) raised the unverified `Respond()`
> side-effect risk it flagged from a caveat to a Critical finding, and this
> update is the response to it — not just a footnote on the same design, an
> actual redesign.
>
> **The `Respond()` side-effect risk is now the reason a design changed, not
> just a caveat.** The review treated `AppointmentItem.Respond()`'s
> documented/known behavior — committing a real calendar change at call time
> (a new EntryID on accept/tentative, a move to Deleted Items on decline),
> independent of whether the resulting response is ever sent — as
> confirmed-plausible enough that shipping the three draft tools
> (`draft_accept_meeting`/`draft_decline_meeting`/`draft_tentative_meeting`)
> unchanged would mean a "draft" tool could silently alter or destroy a real
> calendar item the instant it's invoked, before the user takes any action at
> all. That's the exact shape of bug `draft_cancel_event` itself hit and fixed
> the same way months earlier (see the "Update 2026-09-28" `draft_cancel_event`
> note above: an unsaved `MeetingStatus` change that Outlook still persisted on
> window-close). The fix follows that same precedent: `draft_respond_meeting`
> (replacing all three retired tools) never calls `Respond()` at all. It just
> opens the original, completely unmodified item via `Display(false)` and lets
> the user pick Accept/Tentative/Decline themselves from Outlook's own native
> ribbon buttons. Since the redesigned tool no longer calls `Respond()` with a
> specific response type, there is no longer a technical reason for three
> separate draft tools — one unified `draft_respond_meeting` replaces them,
> gated the same as they were (`DraftTierTools`/`entry.ts`'s
> `commentOnlyExtraTools`). `message` can no longer be pre-filled into a
> response body (that would require calling `Respond()` to get the
> `MeetingItem`, the exact call this redesign avoids) — it's returned in the
> tool's output text instead, for the user to paste in themselves via
> Outlook's own "Edit response before sending" option. See the
> `draft_respond_meeting` table row below for the mechanics.
>
> **Smaller fixes to `RespondMeeting`/`DraftRespondMeeting` from the same
> review, all in `OutlookAiAddIn/OutlookTools.Calendar.cs`:** a new shared
> `ResolveMeetingAppointment` helper replaces the duplicated
> item-to-`AppointmentItem` resolution logic both methods had. Both methods
> now refuse (`IsError`) up front on a meeting that's already been canceled by
> the organizer, or on an item that isn't a meeting the user was actually
> invited to (an organizer's own `olMeeting` copy, or a plain `olNonMeeting`
> appointment) — via two new shared helpers, `AlreadyCanceledRespondError`/
> `NotInvitedError` — instead of calling `Respond()` unconditionally and
> letting Outlook's own behavior in those cases go unchecked.
> `RespondMeeting`'s subject is now captured *before* `Respond()` runs, since
> `Respond()` replacing the item on accept/tentative (per the risk described
> above) could otherwise leave `appt.Subject` reading a stale reference
> afterward. `RespondMeeting`'s result text now says explicitly when
> `.Send()` fails ("...but the response could not be sent to the organizer")
> instead of unconditionally reporting success — the local `Respond()` had
> already gone through, but the organizer was never notified and any comment
> was lost. Two stale "accept/decline" tool-capability descriptions predating
> `tentative_meeting` (in `entry.ts`'s `trackChangesExtraTools` comment, and
> `OutlookTools.cs`'s comment on the `CommentOnly`/`TrackChanges` tier meaning
> just above `DraftTierTools`) now say "accept/decline/tentatively-respond".
> `draft_cancel_event`'s description, the system prompt's cancel-tools
> sentence, and `ReceivedMeetingCancelError`'s message all pointed an attendee
> who can't cancel someone else's meeting at `decline_meeting` — one tier
> above `draft_cancel_event` itself (Draft only vs. Automate approvals), so a
> Draft-only caller couldn't actually reach that suggestion; all three now
> also mention `draft_respond_meeting` as the Draft-tier-reachable
> alternative.
>
> **Still unverified live, pending the project owner's own test:** whether
> the organizer actually sees the `message`/comment text on the delivered
> response for `accept_meeting`/`decline_meeting`/`tentative_meeting` — this
> redesign didn't touch that mechanism (`resp.Body` set before `.Send()`) or
> add any new evidence toward confirming it; it remains exactly as
> unverified as the immediately preceding "Update 2026-09-29" block already
> described it.

> **Update 2026-09-29 (Outlook `list_events` gains `mailbox` parameter for shared calendars):**
> `list_events` now accepts an optional `mailbox` parameter (email address) to list
> events on a shared calendar instead of the default user's own calendar. When
> provided, `mailbox` is resolved via `Ns.CreateRecipient(mailbox).Resolve()` and
> the calendar folder is opened via `Ns.GetSharedDefaultFolder(recipient,
> olFolderCalendar)`. Resolution failures or folder access errors return an `IsError`
> naming the mailbox and the likely cause (not shared with the user, or needs
> adding via Outlook's own "Open Calendar" first). The identical `Sort`/
> `IncludeRecurrences`/`Restrict` query is reused unchanged for both the default
> calendar and shared calendars. When `mailbox` is given, each event's output gains
> a `calendar_owner: <mailbox>` line. The zero-results message becomes "No events on
> <mailbox>'s calendar between X and Y." when `mailbox` is given (vs. "No events
> between X and Y." when omitted).
>
> **Verification status: Exchange calendar-sharing permission tiers untested.** The
> code path `GetSharedDefaultFolder` should work identically across all permission
> levels (Full Access / Editor / Reviewer / etc.), but behavior has not yet been
> exercised live against a real second mailbox with different permission tiers
> configured. First live use against a shared calendar with limited permissions
> (free-busy only, titles+locations only, or no-access) should verify the exact
> error messages and whether partial-read tiers degrade gracefully or fail outright.

> **Update 2026-09-29 (PR #28 code review: cross-mailbox write refusal, `get_event`
> `store_id`, private-item redaction, resolved display name) — since reversed, see
> the next dated update below:**
> A review of the `list_events` `mailbox` feature above (PR #28) found that nothing
> stopped an `event_id` obtained from someone else's shared calendar from being fed
> into a write tool — `edit_event`, `draft_edit_event`, `cancel_event`,
> `draft_cancel_event`, and `accept_meeting`/`decline_meeting`/`tentative_meeting`'s
> shared `RespondMeeting`/`draft_respond_meeting` all resolved `event_id` via
> `ItemById` with no check on which mailbox store the resolved item actually lived
> in. This was fixed with a new shared helper, `RefuseIfNotOwnStore`, that refused
> any of those six tools on an item outside the caller's own default store, plus a
> redaction of `subject`/`location` to `"(private)"` for `Sensitivity == olPrivate`
> items on the shared-calendar path in `list_events`. **Both of these were reverted
> the same day — see the dated update immediately below for the project owner's
> explicit reasoning and the resulting design.** `get_event`'s `store_id` parameter
> (described in the `get_event` row above) was the one piece of this review kept
> as-is, and is now extended to the calendar-editing tools too.
>
> Separately, same review: `calendar_owner` and the zero-results message now use the
> resolved `Recipient`'s own `.Name` (falling back to the raw `mailbox` string only
> if that's empty) instead of echoing back the raw `mailbox` input verbatim — so
> `mailbox: "dana"` reports back who it actually resolved to, not the ambiguous
> string the caller typed. This part is unaffected by the reversal below.
>
> **Update 2026-09-29 (explicit project-owner reversal: `RefuseIfNotOwnStore` and
> private-item redaction removed; `store_id` added to all calendar-editing tools
> instead):** The project owner explicitly overrode the `RefuseIfNotOwnStore` check
> and the private-item redaction added by the PR #28 review immediately above, on
> this principle: this add-in should not layer its own authorization or
> content-redaction logic on top of Outlook/Exchange's own permission model.
> Whatever a caller's real Exchange sharing permissions would let them do through
> Outlook's own UI on a calendar they've been given access to — view an event
> (including a private one, if their access level exposes it), edit it, cancel it,
> or respond to it — these tools should allow too, with Outlook/Exchange itself
> (not this code) the only thing that can refuse. `RefuseIfNotOwnStore` and its six
> call sites (`EditEvent`, `DraftEditEvent`, `CancelEvent`, `DraftCancelEvent`,
> `RespondMeeting`, `DraftRespondMeeting`) were deleted outright from
> `OutlookAiAddIn/OutlookTools.Calendar.cs`; `QueryCalendarItems`' `isPrivate`
> redaction of `subject`/`location` was deleted too, so shared-calendar rows from
> `list_events` now show whatever Outlook itself resolves for a private item, same
> as any other.
>
> In their place, `edit_event`, `cancel_event`, `accept_meeting`, `decline_meeting`,
> `tentative_meeting`, `set_event_categories`, `set_event_availability`, and their
> `draft_` counterparts (8 tools total) all gained the same optional `store_id`
> parameter `get_event` already had, passed straight through to `ItemById`/
> `GetItemFromID` — purely a lookup aid for resolving an `event_id` outside the
> caller's own default store (as returned by `list_events`' `store_id` field for a
> shared-calendar event), with zero authorization logic attached. If the caller
> lacks real Exchange permission for the action, that now surfaces as whatever
> COMException `.Save()`/`.Send()`/etc. naturally throws, caught by each tool's
> existing generic exception handling (or the outer `ExecuteAsync` catch, for a
> path with no local try/catch) — not as a custom pre-check message.
>
> **Verification status: unverified.** This reversal has not been exercised live
> against a real second mailbox with restricted permissions — it is not yet known
> what error (if any) actually surfaces from Outlook when a caller genuinely lacks
> permission for a cross-mailbox write (e.g. attempting `edit_event` with a
> `store_id` from a calendar where the caller only has Reviewer access). It's
> assumed to come back as a COMException from `.Save()`/`.Send()`, per this design's
> own reasoning, but that assumption itself is untested, same as the underlying
> calendar-sharing permission-tier behavior called out in the block above.

### 2026-09-30 — Shared-calendar `list_events` moved onto EWS; `search_contacts` queries Contacts and Directory separately

Moved verbatim out of `docs/ai-tool-surface.md` when `main` was merged into the current-state restructure on 2026-10-02; the facts that are still true now live in that file's tables. Cross-references inside the quoted blocks ("the dated Update block above", "see below") point at that file's old layout.

> **Update 2026-09-30 (shared-calendar path rewritten onto EWS - was freezing Outlook):**
> The `GetSharedDefaultFolder` + `Items.Sort/IncludeRecurrences/Restrict/foreach` COM
> enumeration above froze Outlook - confirmed live by the project owner, even for a
> single-day range. Root cause: that folder is normally not cached offline the way the
> caller's own default calendar is, so per-property reads during enumeration could mean
> a live, blocking round trip to Exchange for every property of every event, all on
> Outlook's own UI thread (Outlook COM objects are STA-bound, unlike this add-in's EWS
> calls). `list_events`' shared-calendar path now queries EWS's `FindAppointments` +
> `CalendarView` instead (server-side date-range filtering and recurrence expansion,
> off the UI thread via `Task.Run` - the same pattern `search_contacts` already uses for
> the same reason). `GetSharedDefaultFolder` is still called exactly once per
> `list_events` call, but only to read `.Store.StoreID` for `get_event`'s `store_id`
> parameter - it never touches `.Items`. Output format is unchanged (`list_events`'
> shared-calendar path now produces the same per-event text shape via the new
> `SharedCalendarEventFormat.Format` that `QueryCalendarItems` already produces
> inline for the own-calendar path — the two paths intentionally use separate code,
> not a shared call, per the design doc's rationale). See
> `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md` for the
> full design, including the one assumption this fix rests on that still needs live
> confirmation (whether `GetSharedDefaultFolder` alone, independent of enumeration, was
> ever part of the freeze). One accepted, deliberate difference: the EWS path's
> `response`/`meeting_status` values are EWS's own label names (e.g. `Accept`,
> `Meeting`, `Cancelled`) rather than the COM path's `Ol*` enum names (e.g.
> `olResponseAccepted`, `olMeeting`) — both are just human/LLM-readable text that
> nothing parses, so this is not a bug.

From the old Outlook section's EWS carve-out bullet:

> Contact resolution calls **EWS `ResolveName` twice, `ContactsOnly` then `DirectoryOnly` (both `returnContactDetails: true`), merging both result sets** (EWS Managed API 2.2, `Microsoft.Exchange.WebServices` 2.2.0) — **fixed 2026-09-30**: the original single-call `ContactsThenDirectory` short-circuited on any Contacts-folder hit (whose ANR only matches `DisplayName`) and never reached the Directory/GAL phase (whose ANR does cover given name/surname), so a query could resolve only against display names. Querying both locations unconditionally and merging (`ContactSearchFormat.Format` already dedupes by email/name) fixes that at the cost of one extra EWS round trip.

### 2026-10-02 — `main` merged into the current-state restructure; historical Outlook sentences moved here

`docs/ai-tool-surface.md`'s Outlook section was rewritten against current source
(`OutlookAiAddIn/web-src/entry.ts`, `OutlookTools*.cs`) instead of picking a merge side.
Sentences that described past states rather than the current one were removed from it;
they're kept here:

- **Four tiers (2026-09-19).** Outlook began repurposing all four `EditingMode` slots
  (Read only / Draft only / Automate approvals / Full autonomy). Before that it used only
  `ReadOnly`/`FullAutonomy`, and the color-tag/category tools (see 2026-09-17) and the
  other mutating tools were Full autonomy only. `OutlookTools.Ews.cs` was split out of
  `OutlookEws.cs` the same day, and `find_meeting_slots` began reading the work week from
  EWS `GetUserAvailability` (it previously defaulted to "today→Thursday, or next week if
  Fri/Sat").
- **`search_contacts` before EWS.** The pre-2026-09 COM implementation was a recursive
  multi-store contact-folder crawl that froze and then crashed Outlook; it was removed.
- **`draft_cancel_event`'s first version** set `MeetingStatus = olMeetingCanceled`
  unsaved before displaying the item. It was removed the same day (2026-09-28) after live
  testing showed the change persists on window-close without an explicit Send. Its
  attendee-only refusal also pointed only at `decline_meeting` (unreachable from Draft
  only) until the PR #29 review (2026-09-29) added `draft_respond_meeting`.
- **Structural-fragility correction (2026-09-28).** Verbatim from the old section:

> **Corrected 2026-09-28:** this section previously claimed `reschedule_event` /
> `draft_reschedule_event` / `cancel_event` / `draft_cancel_event` were bound by the same
> limitation, documenting it as though it were a fundamental COM constraint. That was
> wrong: those four tools (the first two since fully replaced, same day, by `edit_event`/
> `draft_edit_event` — see the dated update near the top of this document) took an
> optional `occurrence_date`, resolved via a new `ResolveOccurrenceTarget` helper against
> `RecurrencePattern.GetOccurrence(DateTime)` — a real, callable member, confirmed via
> .NET reflection against the referenced PIA rather than assumed absent — so they can
> target one specific occurrence precisely. The shared-`EntryID` limitation still applies
> to `edit_event`/`cancel_event`/`draft_edit_event`/`draft_cancel_event` only when
> `occurrence_date` is omitted (they then act on the whole series, exactly as before this
> feature). See the "Update 2026-09-28 (Outlook gains recurring-series support...)" block
> near the top of this document for full detail, including a since-discovered undo/redo
> asymmetry between occurrence-level edit (undo-able) and occurrence-level cancellation
> (always a barrier).

- **Unproven-at-runtime item resolved (2026-09-29).** Verbatim from the old section:

> - **RESOLVED 2026-09-29 (was: `draft_accept_meeting`/`draft_decline_meeting`/
>   `draft_tentative_meeting`'s "nothing persists until the user acts" claim rested on
>   an unverified assumption about `AppointmentItem.Respond()` itself).** A PR #29 code
>   review treated this gap as confirmed-plausible enough to be a Critical finding
>   rather than leave it as an open question — `Respond()` is documented/known to
>   commit a real calendar change at call time (a new EntryID on accept/tentative, a
>   move to Deleted Items on decline) independent of whether `.Send()`/`.Display()` is
>   subsequently called, which would have silently violated the draft contract. Fixed
>   by redesign, not by verification: the replacement `draft_respond_meeting` never
>   calls `Respond()` at all, so this class of risk no longer applies to the draft
>   tool — see the "Update 2026-09-29 (PR #29 code review...)" block above. The
>   underlying question (does `Respond()` itself commit changes at call time) remains
>   formally unconfirmed, but it's now moot for the draft path specifically, since
>   nothing in `draft_respond_meeting` ever calls it.

- **Brief-vs-source note.** The merge brief expected `draft_accept_meeting`/
  `draft_decline_meeting`/`draft_tentative_meeting`. They don't exist in current source;
  PR #29 replaced them with the single `draft_respond_meeting`.

Also folded in during the merge, for Word and PowerPoint (PR #33 / commits `3a8835b`,
`9461e40`): each mutating tool call is now exactly one undo step (Word wraps every
non-read tool call except undo/redo in `UndoRecord.StartCustomRecord`/`EndCustomRecord`;
PowerPoint calls `StartNewUndoEntry()` before every non-read call, undo/redo included,
on purpose). Word's `edit_table` and PowerPoint's `edit_table_cell`/
`edit_table_structure` now state that row/column index 0 is simply the first physical
row/column, header row included. Both apps' tool tables gained the
`undo_last_action`/`redo_last_action` rows (Word 17 → 19 tools, PowerPoint 49 → 51).

---

### 2026-10-02 — Corrected stale attendee-resolution description

The main doc's `draft_event`/`edit_event`/`create_event` rows described attendee
resolution as `Recipients.Add(...)` followed by the collection-level
`Recipients.ResolveAll()`. That stopped being true before this doc was last verified
against source: `OutlookAiAddIn/OutlookTools.Compose.cs`'s `AddAttendees` resolves each
added `Recipient` individually via `r.Resolve()` right after `Recipients.Add(...)` —
`ResolveAll()` was confirmed live to not reliably resolve these at all (`Resolved`
stayed `False`, `Address` stayed empty, and a `.Send()` using it never actually
delivered). `edit_event`'s `ReplaceAttendees` reuses the same per-recipient `AddAttendees`
helper, so it was never calling `ResolveAll()` either. Also undocumented until now: all
four attendee-touching tools (`create_event`/`draft_event`/`edit_event`/
`draft_edit_event`) append a trailing note to their result text naming any address that
failed to resolve, via the shared `FormatUnresolvedAttendeesNote`. Fixed in the main
doc's tables directly; see `OutlookAiAddIn/OutlookTools.Compose.cs.md` for the full
rationale behind resolving per-recipient instead of via `ResolveAll()`.

## See also

- `docs/superpowers/plans/STATUS.md` — the full PP/FT plan execution log (all ~25
  plans across Phases 0-6), including which manual-verification items are still
  outstanding.
- `docs/superpowers/verification/` — per-plan verification notes.
- `docs/ai-tool-surface.md` — the current-state reference this file is history for.
