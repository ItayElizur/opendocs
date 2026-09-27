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

---

## See also

- `docs/superpowers/plans/STATUS.md` — the full PP/FT plan execution log (all ~25
  plans across Phases 0-6), including which manual-verification items are still
  outstanding.
- `docs/superpowers/verification/` — per-plan verification notes.
- `docs/ai-tool-surface.md` — the current-state reference this file is history for.
