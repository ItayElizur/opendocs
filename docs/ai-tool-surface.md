# AI Tool Surface Reference — officeoffice

This document catalogs every mechanism by which the LLM (the AI assistant embedded in
the Word/Excel/PowerPoint VSTO add-ins) can read or mutate a real Microsoft Office
document in this repo, and compares it against the equivalent surface already
documented for `genoffice` (`C:\dev\genoffice\docs\ai-tool-surface.md`, a from-scratch
web-based Office clone suite that this project ports its tool design from).

> **Note on `docs/tool-surface-todo.md`**: that checklist has been retired (see
> PP-8, `docs/superpowers/plans/2026-08-23-pp08-retire-stale-todo.md`, 2026-08-24)
> and now just points here. It was written against an earlier snapshot and marked
> many command/operation kinds as `[ ]` unimplemented that are, as of the current
> `main` (commit `1a478c5`), fully implemented. This document was produced by
> reading the actual current `WordTools.cs`, `ExcelTools.cs`, and
> `PowerPointTools.cs` end-to-end, not by trusting that file.

> **Update 2026-08-26 (search/replace + heading outline):** a user-reported gap -
> Word had no read-only search tool at all (only `apply_commands`' mutating
> `find_replace`, which also under-reported its replacement count - always 0 or 1
> regardless of how many occurrences it actually replaced, now fixed to loop and
> count accurately), Excel's `find_cells` had no write-side counterpart, and
> PowerPoint had neither search nor find/replace of any kind. Added: Word's
> `find_text` (read-only search) and `get_headings` (Navigation-Pane-style
> heading outline, `[index] H<level>: text`); Excel's `find_replace` op under
> `propose_operations` (text-value cells only, never formulas); PowerPoint's
> `find_text` (read-only, across shape text + notes) and `replace_text` (every
> text-frame shape + notes, NOT table cells or SmartArt node text - those keep
> using `edit_table_cell`/`edit_smartart`). Not re-verified against every section
> below - see each app's tool table for the pre-existing state this fixes.
>
> **Update 2026-08-27 (Word find_text/get_headings performance fix):** both
> tools' first cut scanned every paragraph via positional `Paragraphs[i]`
> indexing - `Paragraphs` is not a real array in Word's COM object model, so
> each indexed access re-walks the document from the start, turning a full
> scan into roughly O(n²) internally. A user reported this as Word visibly
> freezing on a large document (the automation call runs synchronously on
> Word's own UI thread, so it can't pump its message loop until the call
> returns). Fixed: `find_text`'s plain-substring path now uses Word's native
> `Range.Find` engine (the one behind Ctrl+F) via a `wdFindStop`-wrapped
> `Execute()` loop, costing work proportional to match count, not document
> size; `get_headings` now uses `Range.GoTo(wdGoToHeading, wdGoToNext)` - the
> same internal heading index that powers the Navigation Pane / "Browse by
> Heading" - so it never visits a non-heading paragraph at all. Both still
> need to report a 0-based paragraph index matching `read_blocks`/
> `apply_commands`' convention; resolving that no longer uses `Paragraphs[i]`
> either - a shared `ParagraphIndexResolver` marches forward via the cheap
> `Paragraph.Next()` chaining method, visiting each paragraph at most once
> across a whole call. `find_text`'s `regex:true` path still needs a
> per-paragraph scan (Word's Find has no regex mode, only its own more
> limited wildcard syntax), but via that same cheap forward chain now, not
> positional indexing.
>
> **Update 2026-08-27 (Excel find_cells/find_replace performance + scope):**
> a related but milder issue than Word's - `find_cells`' plain-substring path
> (and `find_replace`, added earlier the same day) read `.Text`/`.Formula` on
> every single cell in a sheet's `UsedRange` via a `foreach` loop. Not the
> same O(n²) bug (`foreach` over `Range.Cells` is a real enumerator, not
> positional re-indexing), but still one COM round-trip per cell regardless
> of match count, which adds up on a sheet with a large `UsedRange`. Fixed:
> the plain-substring path now uses Excel's own native `Range.Find`/
> `FindNext` (the engine behind Ctrl+F/Ctrl+H) - `xlValues`/`xlFormulas` are
> separate native passes (Excel's `Find` only searches one `LookIn` mode per
> call), de-duped by address for `look_in:"both"`. `regex:true` still needs
> the per-cell scan (same reason as Word - no regex mode, only wildcards).
> `find_replace` mirrors this: native `Find`/`FindNext` locates candidate
> cells fast, then only the matched cell's literal text `Value2` is read/
> replaced directly (same safety scope as before - never a formula, since a
> numeric/date/formula cell that merely *displays* a match via its formatted
> text has a non-string `Value2` and is skipped). **Also, per user request, a
> scope default change**: both tools previously searched the whole workbook
> when `sheetId` was omitted; they now default to the **active sheet only**
> (matching Ctrl+F/Ctrl+H's default "Within: Sheet"), with a new `allSheets`
> boolean to opt into workbook-wide search ("Within: Workbook") - `sheetId`
> still names one specific sheet directly, active or not.
>
> **Update 2026-08-27 (Word: the same positional-indexing bug found in 3 more
> places, plus a read_blocks cap):** a broader sweep for the same
> `Paragraphs[i]`-indexing anti-pattern (per user request) turned up three
> more hotspots, all now fixed the same way (`Paragraph.Next()` forward
> chaining instead of positional indexing):
> - **`ResolveTargetParagraphs`** - the shared Target matcher behind
>   `apply_commands`' `updateTextStyle`, `updateParagraphStyle`,
>   `deleteBlocks`, `createParagraphBullets`, and `deleteParagraphBullets` -
>   scanned every paragraph in the document via positional indexing
>   regardless of how narrow the Target filter was. Its return type also
>   changed from `List<int>` to `List<(int Index, Word.Paragraph Paragraph)>`
>   - every caller used to re-look-up `paragraphs[i + 1]` per match after
>   this function already had the paragraph in hand during its scan; now it
>   just hands the object back, removing that second round of positional
>   lookups too. All 5 call sites updated to match.
> - **`read_blocks`'s plain-`"text"` format** (the default) had no upper cap
>   at all - only `format:"html"` was capped, at 100 paragraphs. Added a
>   1000-paragraph cap for text mode (not independently benchmarked the way
>   `read_formats`' 200-cell cap or html mode's 100-paragraph cap were -
>   chosen conservatively, documented as such in the code) and switched its
>   indexing to the same forward-chaining walk.
> - `find_text`'s tool description and Word's system prompt now explicitly
>   tell the model that `find_text`'s returned `[index]` is the exact same
>   0-based paragraph index `read_blocks`/`replace_blocks`/
>   `apply_commands`' `Target.blockIndexes` use (no translation needed), and
>   to prefer `find_text` over reading a large range blindly.
>
> **Update 2026-08-27 (Phase 0 complete - shared pure-logic seam, plus two
> tool-facing bug fixes):** implemented
> `docs/superpowers/plans/2026-08-27-phase0-test-seam.md` in full (6 tasks,
> 6 commits, `dotnet test` 23 -> 90 passed). No tool *schema* changed. Two
> deliberate behavior changes did land, each its own commit:
> - **Hex-color validation** (`propose_operations`/`apply_commands`/
>   `set_element_fill` etc., any op taking a hex color): a malformed color
>   used to reach the model as a raw .NET exception - `"#abc"` threw
>   `ArgumentOutOfRangeException`, `"#GGGGGG"` threw `FormatException`,
>   neither naming the bad value or the expected format. Now: 3-digit CSS
>   shorthand (`"#abc"` == `"#aabbcc"`) is accepted, and anything else throws
>   a clean `ArgumentException` naming the offending value and the expected
>   `#RRGGBB` form.
> - **Word's `apply_commands`**: `set_bold`/`set_italic`'s `"value": null`
>   and `set_heading`'s `"level": null` used to reach `GetBoolean()`/
>   `GetInt32()` and throw an opaque `InvalidOperationException` where the
>   existing "missing required field" message belongs. Now rejected with
>   that same clean message. **Deliberately not a global rule** - Excel's
>   `set_cell`/`set_range` legitimately use `"value": null` to clear a cell,
>   so Excel's `RequiredFields` validation is untouched; rejection is
>   opt-in per field via a new `NonNullFields` table, Word-only so far.
>
> Also: a `Dictionary<string, MsoAutoShapeType>` cannot cross an assembly
> boundary from `OfficeAi.Shared` into the VSTO app projects - confirmed via
> `CS1769` ("embedded interop type" as a generic type argument). Excel's and
> PowerPoint's shape-type maps (previously separately duplicated, one with
> two extra aliases) are now one union table in `OfficeAi.Shared.ShapeTypes`,
> carried as `Dictionary<string, int>`, each app casting to
> `MsoAutoShapeType` at its own call site - the same split `ColorUtil`
> already used for color. Side effect: both apps' "unknown shapeType" error
> now lists the full union of valid names, not just each app's own subset.

> **Update 2026-08-27 (Word gains `barStacked`; chart types unified):** pulled
> forward out of Phase 2 at user request, in two commits.
> - **`feat(word)`:** Word's `edit_chart` now accepts **`barStacked`**
>   (`xlBarStacked`, 58). Word's chart-type map had 7 entries to Excel's and
>   PowerPoint's 8, so Word could not draw a stacked bar chart while the other
>   two could. Word's `entry.ts` enum honestly advertised only 7, so schema and
>   handler agreed — a genuine **capability gap**, not a mismatch. Its
>   `chartType` enum gains the value. Secondary fix: `read_chart`'s reverse
>   type-code lookup used the same map, so reading an existing stacked bar
>   chart reported *"unrecognized chart type code 58"* instead of `barStacked`.
> - **`refactor(shared)`:** the three per-app maps are now one
>   `OfficeAi.Shared.ChartTypes.ByName`. Sequencing the Word fix first was
>   deliberate — it made all three byte-identical, so the extraction itself
>   changed no behavior at all. Tests assert the exact `xlChartType` codes
>   (unlike `ShapeTypes`, where raw ints would be brittle): a wrong code here
>   is a **silent** wrong result, and that bug has shipped here before —
>   PowerPoint's copy mapped `"bar"` to `51`/`xlColumnClustered` instead of
>   `57`, so `chartType:'bar'` drew a column chart *and reported success*.
>   Tests also guard drift against every app's `entry.ts` enum in both
>   directions. `dotnet test` 90 → 102.
>
> **Not verified against live Office** — `barStacked` follows the identical
> code path as the seven chart types already supported (a `Dictionary` lookup
> handing an int to the same COM property), and all three apps build clean,
> but no live Word instance was reachable to confirm a stacked bar chart
> actually renders. Worth one manual check.

## Architecture

officeoffice drives the **real desktop Office applications** via VSTO + COM interop
(`Microsoft.Office.Interop.{Word,Excel,PowerPoint}`), unlike genoffice's from-scratch
web renderers. The chat UI runs in a WebView2 page inside a CustomTaskPane; tool calls
cross a `chrome.webview.postMessage` ⇄ `CoreWebView2.PostWebMessageAsJson` JSON bridge
(`OfficeAi.Shared/ToolProtocol.cs`) into C# handlers that call the COM object model
directly — there is no Electron/IPC hop.

`packages/agent-core` and `packages/ai-provider` from genoffice are copied verbatim
into `shared/web-src/{agent-core,ai-provider}` (same `AgentLoop`, same `AgentSkill`
contract, same `maxTurns` default of 8, same multi-provider types including
`genspark`/`anthropic`/`gemini`/`deepseek`/`openai`/`custom`). **However, none of the
three add-ins actually wire up provider selection yet** — each `entry.ts` hardcodes
`streamOpenAiCompatible` against a local test endpoint (`http://127.0.0.1:9000/v1`),
and `onSettingsSave` is a stub ("Not yet wired to the transport/provider config —
deferred"). So the provider-abstraction layer exists but is not live.

There is **no `packages/ai-search` equivalent** in officeoffice: no `web_search`,
`image_search`, `generate_image`, `analyze_media`, or `read_attachment` anywhere in
the repo. This is a deliberate scope decision (the deployment target is air-gapped —
see `add_image`/`replace_image`'s explicit rejection of remote URLs below), not an
oversight.

Each add-in has a governance layer genoffice's docs surface doesn't have in this
form: a shared `EditingMode` enum (`ReadOnly | CommentOnly | TrackChanges |
FullAutonomy`), filtered client-side (which tools are advertised to the model) *and*
re-enforced server-side in each `Tools.Execute()` (mutating tools blocked outright in
Read Only / Comment Only mode, regardless of what the model requests).

---

## Word (`WordAiAddIn/WordTools.cs`)

### Top-level tools (7)

| Tool | Implemented | Notes vs. genoffice |
|---|---|---|
| `get_document_context` | Yes | Much thinner: paragraph/word count + a flat 300-char text preview. No block-indexed list (index\|type\|preview) like genoffice's version. |
| `read_blocks` | Yes, but plain-text only | Paragraph-indexed range read matches genoffice's index shape, but returns plain text (`[i] text` lines) rather than genoffice's restricted-HTML serialization, so returned content loses formatting/structure markup. |
| `insert_content` | Yes, but narrow | Takes plain `text` only — **always appends at the end of the document**. No HTML, no `afterBlockIndex` positioning, no rich content (images, charts, lists) via this tool. |
| `replace_blocks` | Yes, but plain-text only | Same start/end-index-range shape as genoffice (empty text deletes the range), but the replacement is a plain `text` string set via `.Text` — no HTML/rich formatting on the replacement content, unlike genoffice's HTML-parsing version. |
| `apply_commands` | Yes — gateway, see below | |
| `edit_chart` | Yes, but narrow | Combines genoffice's separate `insert_chart`+`edit_chart` into one create-or-edit call, but only sets a title and a **single series'** numeric values — no categories, no chart-type selection, no multi-series support. |
| `add_comment` | Yes | **Not in genoffice's docs surface at all.** Anchors a real Word comment to the first match of given text. Available in every editing mode, including Comment Only. |

No image-insertion tool exists for Word at all (genoffice's docs has `insert_image`).

### `apply_commands` command kinds (12 of 12 genoffice kinds + 4 officeoffice-only aliases — all genuinely implemented)

| Command | Implemented | Notes |
|---|---|---|
| `updateTextStyle` | Yes, but 9/10 fields | bold/italic/underline/strike/sizeHalfPoints/font/color/baselineOffset/link implemented. **`highlight` (text highlight color) is missing** — no handling in `UpdateTextStyle` (WordTools.cs:379-416) and no `highlight` property in the `entry.ts` schema. |
| `updateParagraphStyle` | Yes | align/lineSpacing/indentLeft/indentRight/indentFirstLine/spaceBefore/spaceAfter/pageBreakBefore/shadingFill/borders — full parity. |
| `deleteBlocks` | Yes | Same `Target` matcher (nodeType/headingLevel/containsText/blockIndexes/scope). Deleting every paragraph clears content instead, leaving one empty paragraph (explicitly mirrors genoffice's own guard). |
| `moveBlocks` | Yes | Captures moved paragraphs as OOXML snapshots before deleting, reinserts via `InsertXML` — preserves formatting through the move. |
| `createParagraphBullets` / `deleteParagraphBullets` | Yes | Heading paragraphs matched-but-skipped, non-list matched-but-skipped — explicitly mirrors genoffice. |
| `updateImageProperties` | Yes | width/height (proportional scale from current size)/align, targets `doc.InlineShapes` by index. |
| `insertToc` | Yes | Uses Word's **native** `TablesOfContents.Add(UseHeadingStyles: true)` — real, auto-paginating, more direct than genoffice's hand-built TOC field-XML (which has to work around its web renderer not paginating). |
| `set_bold` / `set_italic` | Yes | officeoffice-only convenience aliases for `updateTextStyle`'s bold/italic fields, addressed by paragraph-index range rather than a `Target`. |
| `set_heading` | Yes | officeoffice-only alias, ≈ genoffice's `setHeadingLevel`. |
| `find_replace` | Yes | officeoffice-only alias, ≈ genoffice's `replaceAllText`. |

Every genoffice `apply_commands` kind has a real implementation here — the todo
checklist's claim that these are unimplemented is wrong for the current source.

---

## Excel (`ExcelAiAddIn/ExcelTools.cs`)

### Top-level tools (10)

All 9 read/query tools plus `propose_operations` are implemented and genuinely
functional — full 1:1 parity with genoffice's naming and shape (`get_workbook_context`,
`read_range`, `read_cells`, `select_range`, `read_formats`, `read_sheet_features`,
`find_cells`, `trace_precedents`, `trace_dependents`). `load_guide` has no equivalent
(deliberately out of scope — genoffice's is an internal prompt-budget mechanism for
managing its larger op count in context, not needed at officeoffice's current scale).

Notable native-COM advantage: `find_cells`'s `errors_only` mode uses
`Range.SpecialCells(xlCellTypeFormulas, xlErrors)` — a genuinely native error-cell
scan the code comments call out as the categorical VSTO/COM advantage over
Office.js's wildcard-only `Range.find`.

### `propose_operations` operation kinds — **all named kinds from genoffice's list are implemented**

Every operation kind genoffice documents (writing, formatting, layout, structure,
charts, table, pivot, data — 51 distinct named kinds across those groups) has a real
handler in the `ProposeOperations` switch. This directly contradicts the now-retired
`tool-surface-todo.md`'s "9 of 65 implemented" claim — that count was from an earlier
snapshot; the current source is functionally complete against genoffice's named op
list.

Where officeoffice's version is **narrower** than genoffice's:

| Op | Gap |
|---|---|
| `format_range` | Only `bold`/`italic`/`numberFormat`/`fillColor` (4 properties). Missing: font family, size, font color, underline, strikethrough, align, wrap, rotation, indent, borders — all present in genoffice's version. |
| `add_chart` | Basic path only supports `column`/`line`/`pie` (chartType silently falls back to column for anything else). The richer chart vocabulary (bar/area/doughnut, legend, dataLabels, seriesColors, per-series renaming) only exists on `edit_chart`, not on creation — genoffice supports the richer set on both. |
| `add_image` | **Local file paths only** — remote URLs throw `NotSupportedException` ("air-gapped deployment"). genoffice downloads from `image_search`/`generate_image` results. This is a deliberate scope boundary, not a bug. |
| `add_shape` | 26 named preset types + textbox — narrower than genoffice's "full OOXML preset-geometry set" but still substantial (rect/roundRect/ellipse/triangle/parallelogram/trapezoid/diamond/pentagon/hexagon/octagon/pie/chord/donut/foldedCorner/heart/lightningBolt/sun/moon/cloud/arc/star5/4 arrow directions). |
| `set_data_validation` | `checkbox` kind explicitly rejected — Excel's Data Validation COM API (verified via reflection against the referenced PIA) has no boolean-checkbox validation type; only 8 kinds exist total, none map to it. |

No `dataSource`/provenance-enforcement mechanism exists anywhere (genoffice's slides
app gates chart data-source claims; nothing analogous exists in officeoffice's Excel
or PowerPoint chart tools).

---

## PowerPoint (`PowerPointAiAddIn/PowerPointTools.cs`)

> **Stale as of PP-24 (2026-08-24):** this section predates PP-19 through PP-24 and undercounts the tool list (now 31: the table below plus `delete_slide`/`move_slide`/`duplicate_slide` from PP-19, and `set_slide_layout`/`set_slide_transition`/`add_animation`/`read_animations`/`edit_animation` from PP-24 — slide layout, transitions, and shape animations, none of which existed when this doc was written). Not rewritten line-by-line here; see `docs/superpowers/plans/2026-08-24-pp24-powerpoint-layout-transitions-animations.md` and `docs/superpowers/plans/STATUS.md` for the current, accurate state.

### Tools (23 total, all genuinely implemented) — pre-PP-19/PP-24 snapshot, see note above

| Tool | Implemented | Notes vs. genoffice |
|---|---|---|
| `get_deck_context` | Yes | Per-slide flat text preview (120 chars), no per-element type/id inventory like genoffice's version. |
| `read_slide` | Yes | Shape name + text per element, indexed. |
| `set_element_text` | Yes | Matches. |
| `set_element_style` | Yes (widened, PP-20) | bold/italic/fontSize/color/fontName/underline/shadow/alignment(left/center/right/justify)/baselineOffset(SUPERSCRIPT/SUBSCRIPT/NONE). Reports which properties actually applied instead of a flat "Style updated." **Strikethrough deliberately not implemented**: unlike Excel (whose `TextFrame2`/`TextRange2` newer text model has it), this PowerPoint PIA has no `TextFrame2` member on `Shape` at all (confirmed via a direct `CS0234` compile failure) — no field was added to the schema for it, so nothing silently no-ops. |
| `set_element_transform` | Yes | left/top/width/height/rotation — matches. |
| `add_text_box` | Yes | Matches. |
| `add_shape` | Yes (widened, PP-20) | Now shares Excel's 26-preset `ShapeTypeMap` (ported, not extracted to `OfficeAi.Shared` — see rationale below), plus `rectangle`/`oval` kept as aliases for `rect`/`ellipse` so existing calls keep working. Unrecognized name errors listing valid ones (previously silently became a rectangle). No fill/line params — chain `set_element_fill`/`set_element_stroke` for those. |
| `delete_element` | Yes | Matches. |
| `add_slide` | Yes | Duplicates a source slide's layout, optional text-clear. |
| `set_element_fill` | Yes | Solid fill or none. |
| `set_element_stroke` | Yes | Color/width or remove. |
| `set_slide_background` | Yes | Single color; `slideIndex = -1` applies to all slides — matches genoffice. |
| `ungroup_element` | Yes | Promotes children; explicitly tells the model to re-`read_slide` for fresh indices afterward. |
| `add_table` / `edit_table_cell` / `edit_table_structure` / `edit_table_style` | Yes | Native `Shapes.AddTable`; structure supports insert/delete row+col; style supports firstRow/bandRow/shading/borders. Reasonably close to genoffice's parity. |
| `add_chart` / `edit_chart` | Yes | Writes real data into the chart's embedded Excel workbook (`ChartData.Workbook`), explicitly closes/releases the COM object to avoid a leaked hidden Excel process. Supports bar/barStacked/line/area/pie/doughnut, legend, dataLabels, gridlines. |
| `add_smartart` | Yes | Maps 7 layout keys (list/process/cycle/hierarchy/pyramid/matrix/venn) to native SmartArt layouts by display name; flat item list only (matches genoffice's own flat-list scope, per source comment). |
| `crop_image` | Yes | Fractional crop against current on-slide size (documented imprecision under repeated crops — no reliable "natural size" once already resized in classic Interop). |
| `replace_image` | Yes, narrow | **Local file path only** (same air-gapped constraint as Excel's `add_image`) — no AI-generation pairing since `generate_image` doesn't exist here. |
| `set_picture_opacity` | Yes | Via `Fill.Transparency`. |

### Missing entirely (confirmed absent from both the C# switch and the advertised tool list)

- `delete_slide` — no way for the AI to remove a slide.
- `execute_slide_script` — no scripting DSL; every multi-property/multi-element edit
  must go tool-by-tool (`set_element_transform` one shape at a time), rather than
  genoffice's atomic AST-interpreted batch script.
- The entire deck-generation pipeline: `ask_clarification`, `plan_deck`,
  `generate_deck`, `regenerate_slide`, `save_style_template`, `list_style_templates`.
- No automatic post-edit audit/QC pass — genoffice's `auditSlideLayout` (geometric
  overflow/overlap/bounds check after every script run) and `slide-qc.ts` (vision-based
  QC sub-loop after generated pages) have no counterpart; nothing here checks the
  result of an edit automatically.
- No `add_comment`-equivalent for PowerPoint (unlike Word), so Comment Only mode
  currently behaves identically to Read Only — a documented gap in the code comments.

### Structural fragility note

Shapes are addressed by **positional index** (`slideIndex`, `shapeIndex` into
`slide.Shapes`) rather than a stable id — indices shift whenever shapes are
added/removed/reordered, unlike genoffice's `sourceId`-based addressing. The model
is instructed to re-read the slide after structural changes (e.g. `ungroup_element`'s
result text says so explicitly), but there's no protection against acting on a stale
index otherwise.

---

## Explicitly out of scope everywhere (per project scope, not gaps)

> This scope boundary originated in the project's original feasibility report and
> the toolset-port plan's Global Constraints. It was previously stated only in
> `docs/tool-surface-todo.md`'s header before that file was retired (see PP-8,
> `docs/superpowers/plans/2026-08-23-pp08-retire-stale-todo.md`); this is now its
> canonical location.

- `web_search`, `image_search`, `generate_image`, `analyze_media`, `read_attachment`
  — no `ai-search` equivalent; air-gapped deployment target.
- The PDF app and the Markdown app have no officeoffice counterpart (Markdown's
  scope is folded into Word).
- PowerPoint's `execute_slide_script` DSL and entire deck-generation pipeline (see
  above: `ask_clarification`, `plan_deck`, `generate_deck`, `regenerate_slide`,
  `save_style_template`, `list_style_templates`).

**Inconsistency flagged, not resolved here:** the retired `tool-surface-todo.md`
bundled `delete_slide` into this same out-of-scope list, alongside the DSL and
generation pipeline. This document's own "Missing entirely" section above already
treats `delete_slide` separately, and
`docs/superpowers/plans/2026-08-23-pp19-powerpoint-scope-and-delete-slide.md` argues
explicitly that `delete_slide` is a small, clearly-in-scope fix with no dependency on
the DSL (`add_slide` already exists; deleting a slide needs no scripting language).
That plan's Task 1 ships `delete_slide` independently of its Task 2 decision gate on
the larger DSL/generation question. Treat `delete_slide` as in-scope going forward;
this out-of-scope list covers only the DSL, the generation pipeline, and the QC pass.

---

## Summary: what genoffice has that officeoffice doesn't

1. **Live multi-provider selection** — the abstraction is copied in, but no add-in
   actually wires user-selected provider/model/API key to the transport yet; all
   three hardcode a local OpenAI-compatible test endpoint.
2. **Web-sourced content** — no search, no AI image generation, no media analysis,
   no chat-attachment reading, anywhere. Image tools are local-file-only by design.
3. **PowerPoint's scripting DSL and generation pipeline** — no `execute_slide_script`,
   no `generate_deck`/`regenerate_slide`, no automatic QC/audit pass, no `delete_slide`.
4. **Richer per-op parameter sets** in several places (status as of the items below —
   several since fixed by their own PP item, noted inline; this list otherwise
   reflects the original audit and is not re-verified wholesale here): Excel's
   `format_range` (missing ~7 of ~11 style properties genoffice supports), Excel's
   `add_chart` on creation (missing chart-type breadth `edit_chart` has),
   ~~PowerPoint's `set_element_style` (missing underline/align/family)~~ and
   ~~`add_shape` (3 types vs a full preset-geometry set)~~ — **both fixed, PP-20**
   (see the PowerPoint tool table above), Word's `insert_content`
   (plain-text-append-only, no positioning or rich content), `read_blocks`/
   `replace_blocks` (plain text only, no HTML), `updateTextStyle` (missing
   `highlight`), and `edit_chart` (single-series, no categories).
5. **Word has no image-insertion tool at all** (genoffice's docs app does).
6. **`dataSource`/provenance enforcement** on chart and data-bearing content — genoffice
   gates this at the tool layer for slides; officeoffice has no equivalent anywhere.
7. **Block-indexed document context** — genoffice's `get_document_context`/
   `get_deck_context` return structured per-block/per-element inventories; Word's and
   PowerPoint's officeoffice equivalents return flat text previews only.

## Summary: what officeoffice has that genoffice doesn't

1. **Real Word comments** (`add_comment`) — anchored, native, available in every
   editing mode including Comment Only. genoffice's docs surface has no comment tool.
2. **A real native `TablesOfContents` TOC** on `insertToc`, vs. genoffice's hand-built
   TOC field-XML workaround (needed because genoffice's own renderer doesn't
   paginate).
3. **A native error-cell scan** (`find_cells` with `errors_only`, via
   `SpecialCells(xlErrors)`) — a categorical COM-vs-Office.js/web advantage called out
   explicitly in the source.
4. **Server-enforced editing modes** (Read Only / Comment Only / Track Changes / Full
   Autonomy) as a first-class, uniformly-applied gate across all three apps' tool
   dispatch — genoffice's docs app has Track-Changes-aware writes but no equivalently
   formal, uniform mode-gating system across its apps.
5. **Live selection-push into context** (Word) — `WindowSelectionChange` pushes the
   user's current selection text into `buildContext()` automatically on every turn,
   driven by a real Office event rather than app-side selection-range plumbing.
6. **`add_pivot` with calculated fields** — Excel's pivot op supports
   `PivotTable.CalculatedFields().Add(name, formula)` for formula-derived pivot
   values, in addition to row/column/page/data fields.

---

## Schema-vs-implementation audit

> **Update 2026-08-24 (PP-5 landed):** the structural root cause this whole section
> points at — both gateway tools' `commands`/`operations` items being a bare
> `{type:'object'}` with the entire per-kind contract living only in prose — is fixed.
> `apply_commands` and `propose_operations` now carry real per-kind JSON Schema
> (`WORD_COMMAND_SCHEMAS` in `WordAiAddIn/web-src/entry.ts`; `EXCEL_OPS` +
> `opSchemas`/`opsDescription` in `ExcelAiAddIn/web-src/entry.ts`), and `kind` parsing
> in both `ApplyCommands`/`ProposeOperations` moved inside the per-command try/catch
> with a required-field precheck (`WordTools.cs`'s and `ExcelTools.cs`'s
> `RequiredFields`/`ValidateRequired`) — so a malformed command now fails only itself,
> with a specific error naming the missing field, instead of aborting the whole batch
> (Word finding #1 above) or silently reaching a COM handler.
>
> **Cross-checked exhaustively, not sampled:** every `case` in `WordTools.cs`'s
> `ApplyCommands` switch (12) has exactly one matching entry in `WORD_COMMAND_SCHEMAS`,
> and vice versa. Same for `ExcelTools.cs`'s `ProposeOperations` switch (51) against
> `EXCEL_OPS`. Both diffs are empty. Excel's schema uses a **grouped variant**
> (`DETAILED_KINDS` in `entry.ts`) rather than full `oneOf` detail on all 51 kinds — a
> 51-branch schema measured ~4,870 added tokens (cl100k_base, before/after
> `JSON.stringify(ALL_TOOLS)`), over the ~4k budget PP-5 set as the threshold; only the
> 7 highest-ambiguity kinds (`format_range`, `add_conditional_format`, `add_chart`,
> `edit_chart`, `add_shape`, `set_data_validation`, `add_pivot`) get full structural
> detail, the rest collapse to a `kind` enum + generated prose (still complete — the
> cut is schema-size only, not documentation). The grouped version measured ~2,120
> added tokens.
>
> **What PP-5 did NOT fix, and is not meant to:** every specific finding below —
> `highlight` unimplemented, `bulletPreset` collapsing to two effective values,
> conditional-format's silent-fallback operators/kinds, `add_shape`'s undocumented
> preset names, chart-type gaps, etc. — is a **capability** or **silent-fallback**
> defect in the handler itself, not a schema-structure problem, and is owned by its own
> PP item (PP-9, PP-12, PP-13, PP-14, PP-15, PP-16, PP-21, PP-22 respectively — see
> `docs/superpowers/plans/2026-08-23-pp-index.md`). The tables below are left exactly as
> written at the time of the original audit (a dated record), not updated in place, so
> whoever implements those items has the original evidence rather than a paraphrase.
> The schema **now correctly states the narrow truth** for each of these (e.g.
> `bulletPreset`'s schema enum is `['BULLET','NUMBERED']` with a note that anything else
> collapses to BULLET — matching the handler exactly, not overselling it) — the
> resulting behavior is unchanged until the owning PP item widens it.

Triggered by a real bug: Word's `edit_chart` lets the model set a chart title and one
series' numeric values, but has no `categories` parameter at all — the model can never
label a chart's axis categories or name its series, even though the *schema and its
description* don't oversell this (they only ever mention title+values). That specific
case is a narrow-but-honest tool, not a schema/handler mismatch. The question this
section answers is broader: **across every tool/op in all three add-ins, does the JSON
schema advertised to the LLM (`entry.ts`) actually match what the C# handler
(`*Tools.cs`) reads and does?** Verified by reading every schema definition against its
handler body directly, not by inference.

Two bug classes turned up, ranked by how badly they mislead the model:

- **Silent no-op with false success** — the tool accepts a parameter, does nothing
  useful with an out-of-range value, and still reports success. Worse than an error,
  because the model has no signal to retry or ask the user.
- **Undocumented schema** — the wire-level JSON Schema for a parameter is just
  `{type: 'object'}` or `{type: 'string'}` with no enum/field list; the *real* contract
  lives only in a free-text description (or nowhere). The model can only guess valid
  values, and any guess outside the handler's recognized set falls into one of the
  silent-no-op cases above.

### Word (`WordTools.cs` / `entry.ts`)

| # | Tool / command | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `apply_commands` (whole tool) | The wire schema for `commands` items is just `{type:'object'}` — no `kind` field, no per-kind shape; everything is prose-only in the description. Consequence: `cmd.GetProperty("kind")` sits **outside** the per-command try/catch, so one malformed command (missing `kind`) throws and aborts the **entire remaining batch** with a generic error — while any commands already applied earlier in the same batch stay applied (no rollback). The tool's `IsError:true` result can under-report what actually changed. | Undocumented schema → robustness gap | entry.ts (commands: array<object>) / WordTools.cs:210-268 |
| 2 | `updateTextStyle` | Schema's `style` field is `{type:'object'}` with zero enumerated keys — valid keys only exist in prose. Concretely: **`highlight` is not implemented** (9 of genoffice's 10 fields), and because nothing enumerates valid keys, requesting `highlight` (or any hallucinated key) is silently ignored — `fields.Contains(x)` never matches, nothing throws, and the tool still returns `"updateTextStyle: ok"`. | Silent no-op + false success | entry.ts (style: object) / WordTools.cs:236-238, 379-416 |
| 3 | `createParagraphBullets` (`bulletPreset`) | Schema implies distinct named presets (mirroring genoffice's `BULLET_DISC_CIRCLE_SQUARE`-style names). Handler only ever checks `bulletPreset.StartsWith("NUMBERED")` — every other value, including a correctly-formed preset name, collapses to the same generic `ApplyBulletDefault()`. The model can ask for a specific bullet style and be silently downgraded. | Silent no-op | entry.ts / WordTools.cs:248-250, 544-561 |
| 4 | `edit_chart` | No mismatch — schema and description accurately describe the narrow capability (title + one series' values, no categories, no chart-type choice). Listed here for completeness since it's the finding that prompted this audit. | Honest but narrow (not a bug) | entry.ts / WordTools.cs:132-169 |

`updateParagraphStyle` has the same "no enumerated keys in the wire schema" structural
issue as `updateTextStyle`, but has full 10/10 field parity with genoffice underneath,
so it currently works only because the model happens to send valid keys — a latent
version of the same risk, not an active bug today.

### Excel (`ExcelTools.cs` / `entry.ts`)

`propose_operations`'s formal schema is `{operations: {type:'array', items:
{type:'object'}}}` — like Word's `apply_commands`, **no per-operation-kind schema
exists at all**; every op's real parameter shape lives only in one long free-text
description block (`entry.ts:198-224`). That description is the de facto schema for
everything below.

| # | Operation | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `add_conditional_format` | **Worst gap in the file.** The description says `rule: {kind, ...}` and never lists what fields any of its 8 `kind` values need. Actual per-kind requirements: `number`→`operator,value,value2`; `text`→`text`; `top10`→`rank,percent,bottom,format.{bold,fontColor,fillColor}`; `formula`→`formula`; `colorScale`→`minColor,midColor,maxColor`; `dataBar`→`color` — none named in the schema. The model must guess field names by convention. | Undocumented schema | entry.ts:222 / ExcelTools.cs:400-476 |
| 2 | `add_conditional_format` (`kind:"number"`, `operator`) | Any `operator` string other than `greaterThan`/`lessThan`/`equal`/`between` silently becomes `equal` — no error. | Silent no-op | ExcelTools.cs:388-398 |
| 3 | `add_pivot` (`values[].agg`) | Any `agg` value other than `count`/`average`/`max`/`min` (e.g. a typo like `"avg"`) silently becomes `sum`, unstated in the description. | Silent no-op | entry.ts:218 / ExcelTools.cs:1202-1211 |
| 4 | `add_shape` (`shapeType`) | 26 valid preset names + `"textbox"` exist in the handler; **none are listed in the schema or description**, and an unrecognized name silently becomes a plain rectangle. | Undocumented schema + silent no-op | entry.ts:213 / ExcelTools.cs:27-56, 746-773 |
| 5 | `edit_chart` (`chartType`) | The handler supports 6 chart types (column/bar/line/area/pie/doughnut, `ExcelChartTypeMap`), but only `add_chart`'s narrower 3-type enum (column/line/pie) is ever documented — the model has no way to discover `edit_chart` can do bar/area/doughnut at all. | Undocumented schema (capability hidden, not broken) | entry.ts:210-211 / ExcelTools.cs:58-66, 665-726 |
| 6 | `edit_chart` (data rebinding) | No parameter exists to repoint an existing chart at a different range — the direct Excel analog to Word's `edit_chart` category gap. Less severe than Word's because `add_chart` binds live to a sheet range via `SetSourceData`, so categories/series update automatically when the model edits the underlying cells with `set_cell`/`set_range` — only *rebinding to a new range* is actually missing. | Capability gap (not schema/handler mismatch) | ExcelTools.cs:643-663 |
| 7 | `set_page_setup` (`scale` + `fitToWidth`/`fitToHeight` together) | If both are set in one op, the `fitToWidth` branch runs after `scale` and unconditionally sets `Zoom = false`, silently discarding the `scale` value — Excel's own UI treats these as mutually exclusive, but nothing in the tool description warns the model. | Silent no-op | ExcelTools.cs:857-865 |
| 8 | `add_defined_name` / `delete_defined_name` | Only ever touch workbook-scoped names (no `sheet?` param) — yet `read_sheet_features` reports sheet-scoped defined names as something that can exist. The model can discover a sheet-scoped name but has no op to create one. Schema and handler agree with each other, so not a mismatch, but a real read/write asymmetry. | Capability gap | entry.ts:171 / ExcelTools.cs:1124-1135 |
| 9 | `delete_table` | Calls `.Unlist()` — converts the table back to a plain range **and keeps all the data**. The name/description could easily be read as "remove the table and its data." | Misleading description | entry.ts:217 / ExcelTools.cs:1085-1088 |
| 10 | `add_sparkline` (`targetCell` omitted) | Defaults to the *same* cells as `dataRange` — the sparkline draws inside the very cells holding its own source data. Not mentioned in the description. | Undocumented default | entry.ts:212 / ExcelTools.cs:728-744 |
| 11 | `add_conditional_format` (`kind:"text"`) | Always uses "contains" (`xlContains`) — no "starts with"/"ends with"/"not contains", despite Excel natively offering those and the schema giving no indication only one mode exists. | Capability gap, undocumented | ExcelTools.cs:418-423 |
| 12 | `add_conditional_format` (`kind:"duplicate"`) | Hardcoded to highlight duplicates only (`xlDuplicate`) — no way to flip to "highlight uniques," and nothing documents the restriction. | Capability gap, undocumented | ExcelTools.cs:427-430 |

### PowerPoint (`PowerPointTools.cs` / `entry.ts`)

Best-behaved of the three: all 23 tools have exact 1:1 name correspondence between
schema and handler, and **`add_chart` here does not have Word's bug** — `categories`
and `series[].name`/`series[].values` are all genuinely read and written into the
chart's embedded workbook. The gaps that do exist are narrower and more contained:

| # | Tool | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `edit_chart` (`chartType`) | Handler only applies the change if the value is found in `PptChartTypeMap` — an unrecognized/typo'd value is silently ignored, and the tool still returns `"Chart updated."` (success). | Silent no-op + false success | PowerPointTools.cs:510-513 |
| 2 | `edit_chart` (`legendPos`) | Schema is bare `{type:'string'}` with no enum or valid-value guidance. Handler expects exactly `"none"`/`"r"`/`"t"`/`"l"` — any natural-language guess a model would plausibly send (`"right"`, `"bottom"`, `"top"`, `"left"`) silently falls into the bottom-position branch. | Undocumented schema + silent no-op | entry.ts:356 / PowerPointTools.cs:519-531 |
| 3 | `add_chart` (`kind`) | Unrecognized value silently defaults to `"bar"`, no enum declared in the schema (unlike `add_shape.shapeType`, which does declare one). | Undocumented schema + silent fallback | PowerPointTools.cs:442 |
| 4 | `add_smartart` (`layout`) | Unrecognized value silently defaults to `"list"` ("Basic Block List"), no enum declared. | Undocumented schema + silent fallback | PowerPointTools.cs:564 |
| 5 | `edit_table_structure` (`kind`) / `edit_table_style` (`borderPreset`) | Documented only in free-text description, no JSON-schema `enum`, inconsistent with `add_shape`'s stricter pattern. | Undocumented schema | entry.ts (both) |

`set_element_style`'s missing `underline`/`align`/`fontFamily` is **not** a
schema/handler mismatch — both sides consistently omit them, so it's a smaller feature
set rather than a case of the model being told it can do something it can't.

### Pattern across all three

The two gateway tools (`apply_commands` in Word, `propose_operations` in Excel) both
use a completely untyped `items: {type:'object'}` schema for their batched sub-commands
— the entire per-kind parameter contract lives in prose. This is *architecturally*
consistent with genoffice's own equivalents (which have the same untyped-envelope
design), but genoffice's command/op TypeScript interfaces are internally documented and
tightly matched to their handlers (verified in the original genoffice audit); here, the
prose-only contract combined with several silent-fallback branches in the handlers is
what turns "vague schema" into "the model can silently fail and be told it succeeded."
The single highest-value fix across all three add-ins would be replacing free-text
parameter descriptions with real per-kind JSON-schema shapes (at least `enum` arrays for
every closed set of string values) — that alone would have caught 8 of the 17 issues
above before they could reach the handler.
