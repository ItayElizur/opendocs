# AI Add-in Product Plan — Bugs, Reliability Gaps & Feature Parity

**What this is:** a product-level plan, not an implementation plan. Each item below
states the problem, the evidence it's real, its user impact, and the desired outcome —
deliberately without prescribing the code-level fix. Feed each item to a model
individually to produce a task-by-task implementation plan in the style of
`2026-08-22-word-tools-completion.md` / `2026-08-22-excel-tools-completion.md` /
`2026-08-22-powerpoint-tools-completion.md` / `2026-08-22-addin-ux-fixes.md`.

**Status:** those per-item implementation plans now exist — one per PP item, indexed in
`2026-08-23-pp-index.md` (suggested order, cross-plan dependencies, and the two items
that need a product-owner decision before any code). Every claim below was re-verified
against `main` while writing them; PP-6's scope narrowed as a result, and the plan for
it records what was actually found.

**Source:** a conversation-long audit comparing `officeoffice` (this repo, real
Word/Excel/PowerPoint via VSTO+COM) against `genoffice` (`C:\dev\genoffice`, a
from-scratch web-based Office clone this project ports its tool design from), plus
direct source verification of `WordTools.cs`, `ExcelTools.cs`, `PowerPointTools.cs`,
each app's `entry.ts`, and the shared `agent-core`/`ai-provider`/`chat-ui` packages.
Full tool-by-tool detail lives in `docs/ai-tool-surface.md` (comparison + full schema
audit) — this document is the prioritized, product-framed distillation of it plus the
live-testing bugs found afterward.

**Suggested priority key:** P0 = broken/misleading behavior a user will hit in normal
use; P1 = real capability gap with a workaround; P2 = polish/parity/housekeeping.

---

## Part 1 — Cross-cutting (shared infra, affects all three add-ins)

### PP-1: Task pane doesn't appear in any document window except the one open at add-in startup

**Priority:** P0

**Problem:** `ThisAddIn.cs` in all three add-ins calls `CustomTaskPanes.Add(control,
title)` exactly once, in `Startup`. VSTO binds that pane to whichever window is active
at that moment. Word/Excel/PowerPoint (2013+) give every open document its own
top-level window even within one running process. No code anywhere listens for
`WindowActivate`/`NewWindow`/`DocumentOpen`/`NewDocument` to create a pane for
subsequently-opened windows.

**Evidence:** `WordAiAddIn/ThisAddIn.cs:9-39`, identically in
`ExcelAiAddIn/ThisAddIn.cs` and `PowerPointAiAddIn/ThisAddIn.cs`; confirmed by grep
that no window/document-open event is wired anywhere in the three add-ins.

**User impact:** exactly the reported symptom — the add-in "only opens on a new
document." Any document opened after the app's first window (via File > Open, or a
second file while the app is already running) has no visible way to reach the AI panel
at all; the ribbon toggle button silently does nothing because it's toggling a pane
object that isn't attached to the window the user is looking at.

**Desired outcome:** the AI panel is reachable from every document window, not just
the first one that existed at add-in startup — regardless of how many documents are
opened, in what order, in the same running instance.

**Open questions for the implementer/product owner:** should each window get its own
independent chat session/history (matching `ChatStore.ChatIdForFile` already being
per-document), or should one pane instance be reparented across windows? The former is
more consistent with the existing per-document chat-history design
(`TaskPaneHost.cs:49-62`).

---

### PP-2: Tool-call steps render after the final assistant response instead of before it

**Priority:** P0

**Problem:** `entry.ts`'s `onSend` handler calls `ui.beginAssistantMessage()`
immediately, before the run starts, creating one bubble element that every turn's
`onText` overwrites in place. Tool-call UI groups are created lazily, only when a tool
actually starts. Since DOM order just follows append order, the (early-created,
later-overwritten-with-final-text) assistant bubble ends up positioned **above** the
tool-call group that causally produced the information in that final text.

**Evidence:** `WordAiAddIn/web-src/entry.ts:289-296` (`beginAssistantMessage()` at send
time) vs. `:331-333` (`beginToolGroup()` lazily on first tool call); identical pattern
confirmed in `ExcelAiAddIn/web-src/entry.ts:277-280`/`:319` and
`PowerPointAiAddIn/web-src/entry.ts:432-435`/`:474`.

**User impact:** the chat transcript reads backwards — the reader sees the AI's
conclusion before the work that produced it, undermining trust/legibility of what the
AI actually did (especially in Track Changes / Comment-only modes where seeing "what
happened" in order matters).

**Desired outcome:** tool-call steps render in their true chronological position —
before the response text that depended on their results — for every turn of a run,
including runs with multiple tool-call rounds.

---

### PP-3: Tool call results are never shown anywhere in the UI

**Priority:** P1

**Problem:** `ToolStepHandle.complete()` in the shared chat UI accepts `{ output,
isError, mutated }` but only ever reads `isError` (✓/✗ icon) and `mutated` ("Applied"
tag). The `output` field — the actual text a tool returned — is silently discarded.

**Evidence:** `shared/chat-ui/chat-ui.ts:346-359`.

**User impact:** a user has no way to inspect what a read tool (`get_document_context`,
`read_blocks`, `read_range`, `get_deck_context`, etc.) actually found, even though the
model receives and acts on that data — every tool call in the transcript looks
identical (a name, some input, a checkmark), with no way to verify or debug what
happened. Directly caused the "get_document_context returns empty" report — the tool
may be working correctly and the user simply has no way to see its output.

**Desired outcome:** a user can expand/inspect any tool call's actual returned output
in the transcript, not just its success/failure state.

---

### PP-4: Runs silently end with an opaque "(no text)" reply

**Priority:** P0

**Problem:** `MAX_TOKENS` is hardcoded to 1024 for every turn, including the model's
terminal (answer-writing) turn. When a model spends its entire budget on
reasoning/planning before emitting visible text — more likely on demanding prompts
("summarize the *key points*") than blunter ones ("summarize the content") — the
provider returns `finish_reason: length` with zero content.
`shared/web-src/ai-provider/stream.ts` deliberately treats `length` as a *normal*,
non-error stop reason (not the "abnormal finish" error path), so this surfaces as a
successful empty turn. `agent-core/loop.ts` already detects this precisely
(`result.truncated: true` on `onDone`, `loop.ts:528`) — but `entry.ts`'s `onDone`
handler never looks at `result.truncated`, and just shows a hardcoded, unlocalized
`"(no text)"`.

**Evidence:** `WordAiAddIn/web-src/entry.ts:131` (`MAX_TOKENS = 1024`), same in
`ExcelAiAddIn/web-src/entry.ts:115` and `PowerPointAiAddIn/web-src/entry.ts:116`;
`shared/web-src/ai-provider/stream.ts:819-826` (finish_reason=length excluded from the
error path); `shared/web-src/agent-core/loop.ts:509-530` (truncated flag computed);
`WordAiAddIn/web-src/entry.ts:347-349` (flag ignored, `result.text || '(no text)'`).

**User impact:** the reported bug exactly — some phrasings of an otherwise-reasonable
request produce a dead-end "(no text)" reply with no explanation and no recovery path,
while a differently-worded version of the same request works fine. Confusing and
non-obviously prompt-sensitive.

**Desired outcome:** a turn that got cut off by the token limit is never presented to
the user as an unexplained empty reply — either it's given enough budget to finish, or
the truncation is surfaced honestly with a path to continue/retry, matching how
`TURN_LIMIT_NOTE` already handles the analogous max-*turns* case.

---

### PP-5: The two gateway tools have no structural per-command JSON schema

**Priority:** P1 (root cause underlying several P0/P1 items in Parts 2-3)

**Problem:** Word's `apply_commands` and Excel's `propose_operations` both advertise
their batched sub-commands as bare `items: {type: 'object'}` — there is no JSON Schema
describing what fields any given `kind` needs. The entire contract lives in one
free-text `description` string. The model can only guess field names for anything not
explicitly spelled out in that prose, and nothing structurally prevents a malformed or
field-incomplete command from reaching the C# handler.

**Evidence:** `WordAiAddIn/web-src/entry.ts:225-239` (`commands: {type:'array',
items:{type:'object'}}`); `ExcelAiAddIn/web-src/entry.ts:198-225` (`operations` same
shape, ~50 op kinds documented only in prose).

**User impact:** indirect but broad — this is the mechanism that enables most of the
"silent no-op" and "undocumented capability" findings cataloged in Parts 2 and 3 below
(worst example: Excel's `add_conditional_format`, whose 8 rule kinds have zero
documented per-kind fields). Fixing this one thing would have caught roughly half of
all schema-reliability findings in this plan before they ever reached a handler.

**Desired outcome:** every command/operation kind has a real, enumerable JSON Schema
(at minimum: required fields per `kind`, and `enum` arrays for every closed set of
string values) discoverable by the model at the schema level, not just in prose.

---

### PP-6: Multi-provider/model selection exists in the UI but isn't wired to anything

**Priority:** P1

**Problem:** `packages/agent-core`/`ai-provider` support `genspark`/`anthropic`/
`gemini`/`deepseek`/`openai`/`custom`, and the Settings panel lets a user enter a base
URL/API key/model — but all three add-ins hardcode `streamOpenAiCompatible` against a
local test endpoint (`http://127.0.0.1:9000/v1`) regardless of what's entered, and
`onSettingsSave` doesn't feed the transport at all.

**Evidence:** `WordAiAddIn/web-src/entry.ts:115-116` (default settings, local
endpoint), `:306-315` (`onSettingsSave` only persists to localStorage and toggles TLS
bypass — never touches the transport), `:133-156` (`makeTransport()` always builds
`streamOpenAiCompatible` off `currentSettings`, so this technically *would* pick up
saved settings for `baseUrl`/`apiKey`/`model` — confirm at implementation time whether
this is actually a live gap or was already resolved since the audit; flagged here as
"verify current wiring before scoping").

**User impact:** if genuinely unwired, a deployed add-in cannot point at a real
provider (Anthropic/OpenAI/etc.) — only a locally-hosted OpenAI-compatible endpoint —
which blocks any real-world (non-test) deployment.

**Desired outcome:** a user can pick and successfully use one of the supported
providers (or `custom`) from Settings, and it takes effect on the next request.

---

### PP-7: No `dataSource`/provenance guardrail on AI-authored charts or data content

**Priority:** P2

**Problem:** genoffice's slides app gates chart/data-bearing content behind a
`dataSource` enum (`user`/`document`/`search`/`sample`), rejecting `'search'` unless a
`web_search` actually occurred in-conversation. officeoffice has no equivalent
anywhere — Excel and PowerPoint chart/data tools accept any values with no claim about
where they came from.

**Evidence:** none in officeoffice by design (absence confirmed); genoffice's version
documented in `C:\dev\genoffice\docs\ai-tool-surface.md` (slides section).

**User impact:** lower risk here than in genoffice, since officeoffice has no web
search at all (air-gapped) — so the specific "fabricated-as-if-searched" failure mode
genoffice guards against mostly can't occur. Still relevant for "the model invented
numbers instead of reading them from the sheet/deck."

**Desired outcome:** decide whether this guardrail is worth porting given the
air-gapped scope, or whether it's a genoffice-specific mitigation for a risk this repo
doesn't share. **Recommend product-owner call before scoping an implementation.**

---

### PP-8: `docs/tool-surface-todo.md` is stale and actively misleading

**Priority:** P2 (housekeeping)

**Problem:** that checklist was written against an earlier snapshot and marks many
command/operation kinds `[ ]` unimplemented that are, as of current `main`, fully
implemented (e.g. it claims Excel is "9 of 65 operations implemented" when the real
number is ~all of them).

**Evidence:** direct source audit in `docs/ai-tool-surface.md` (see its top note).

**Desired outcome:** retire the file or replace its content with a pointer to
`docs/ai-tool-surface.md`, which is verified against current source. Trivial, but
worth doing before anyone else plans work off the stale one.

---

## Part 2 — Word (`WordAiAddIn`)

### PP-9: Word's chart tool can't set axis categories or name series

**Priority:** P1

**Problem:** `edit_chart` is Word's only chart tool (creates-or-edits). It accepts
`title` and `values: number[]` and nothing else — no `categories`, no per-series
naming, single series only, chart type is hardcoded to column. The schema is at least
*honest* about this (doesn't claim more than it does) — this is a capability gap, not
a silent-failure bug.

**Evidence:** `WordAiAddIn/web-src/entry.ts:193-205`, `WordTools.cs:132-169`.

**User impact:** the AI cannot build a chart with labeled categories or named series in
Word — every AI-generated chart looks like generic unlabeled bars, unlike genoffice's
docs app (which supports categories/multi-series/bar-line-pie) and unlike
officeoffice's own PowerPoint `add_chart` (which does support all of this correctly).

**Desired outcome:** Word chart authoring reaches parity with what PowerPoint's
`add_chart`/`edit_chart` already do correctly in this same repo — categories, named
multi-series, and a real chart-type choice.

---

### PP-10: Word's content tools are plain-text only — no rich formatting round-trip, no positional insert

**Priority:** P1

**Problem:** `insert_content` always appends plain text at the very end of the
document (no `afterBlockIndex`, no HTML/rich content — no images, lists, or charts via
this tool). `read_blocks` and `replace_blocks` also work in plain text only, unlike
genoffice's HTML-based versions, which preserve/allow formatting.

**Evidence:** `WordTools.cs:111-126` (`InsertContent`), `:171-208` (`ReadBlocks`/
`ReplaceBlocks`); schemas at `entry.ts:184-223`.

**User impact:** the AI can't insert or rewrite formatted content anywhere but the
document's end, and any read-then-rewrite round trip loses formatting. This is the
single biggest capability gap between Word's document-editing tools and genoffice's
docs app.

**Desired outcome:** the model can insert and replace content at arbitrary positions
with basic rich formatting (at minimum: bold/italic/lists/headings) preserved through
read → edit → write, matching genoffice's restricted-HTML approach or an equivalent.

---

### PP-11: Word has no image-insertion tool at all

**Priority:** P1

**Problem:** confirmed absent — no tool anywhere in `WordTools.cs`/`entry.ts` lets the
model place an image into a Word document (local file or otherwise).

**Evidence:** full-file read of `WordTools.cs` and `entry.ts`, no image tool present.

**User impact:** every other app in this repo (Excel via `add_image`, PowerPoint via
`insert_web_image`/`replace_image`) can place images; Word cannot, at all — not even
from a local file path, which is the pattern Excel/PowerPoint already use for the
air-gapped constraint.

**Desired outcome:** Word gets a local-file image-insertion tool, following the same
air-gapped (`local file path only, no remote URLs`) pattern already established in
Excel's `add_image` and PowerPoint's `replace_image`.

---

### PP-12: Word `apply_commands` reliability — missing `highlight`, decorative `bulletPreset`, no-rollback batch abort

**Priority:** P1

**Problem, three related findings:**
1. `updateTextStyle` implements 9 of genoffice's 10 style fields — `highlight` (text
   highlight color) is silently unsupported. Because the schema doesn't enumerate
   valid style keys at all (see PP-5), requesting it fails silently with a false
   `"ok"` result rather than an error.
2. `createParagraphBullets`'s `bulletPreset` parameter only ever checks for a
   `NUMBERED*` prefix — any other preset name (including well-formed ones) silently
   collapses to the same generic bullet style, with no error.
3. `apply_commands` parses each command's `kind` **outside** the per-command
   try/catch. One malformed command (missing `kind`) throws and aborts the entire
   remaining batch with a generic error, while commands already applied earlier in
   that same batch stay applied — no rollback, and the reported error understates what
   actually changed.

**Evidence:** `WordTools.cs:210-268` (batch loop + `kind` parse position), `:379-416`
(`UpdateTextStyle`, no `highlight` handling), `:544-561` (`CreateParagraphBullets`,
`StartsWith("NUMBERED")` check only).

**User impact:** a request to highlight text or apply a specific bullet style
silently does the wrong (or no) thing while reporting success; a single bad command in
a larger batch can leave the document in a half-edited state with a confusing error.

**Desired outcome:** either implement `highlight` and real per-preset bullet styling,
or have the tool report accurately when it can't; a malformed command in a batch fails
that command only, doesn't abort/lose track of the rest.

---

## Part 3 — Excel (`ExcelAiAddIn`)

### PP-13: `format_range` supports far fewer style properties than genoffice's equivalent

**Priority:** P1

**Problem:** `format_range` only handles `bold`/`italic`/`numberFormat`/`fillColor`
(4 properties). genoffice's `format_range` additionally covers font family/size/font
color, underline, strikethrough, alignment, wrap, rotation, indent, and borders.

**Evidence:** `ExcelTools.cs:597-611` (`FormatRange`, exactly 4 properties handled).

**User impact:** the AI can color and bold/italicize cells and set number formats, but
can't change fonts, alignment, borders, or text wrapping — common formatting requests
("center this and add borders", "wrap the text in this column") aren't achievable via
this tool at all.

**Desired outcome:** `format_range` covers the same property set genoffice's version
does, via the real Excel `Range.Font`/`.HorizontalAlignment`/`.WrapText`/`.Borders`
COM properties (all directly available, no COM limitation blocking this).

---

### PP-14: Conditional formatting tool has undocumented per-kind fields and several silent-fallback/narrow behaviors

**Priority:** P1

**Problem, several related findings on `add_conditional_format`:**
1. The description says `rule: {kind, ...}` and never lists what fields any of its 8
   `kind` values (`number`/`text`/`top10`/`formula`/`colorScale`/`dataBar`/etc.)
   actually need — the model must guess field names by convention (this is the worst
   single instance of PP-5's root cause).
2. `kind: "number"`'s `operator` silently becomes `"equal"` for any unrecognized
   value, no error.
3. `kind: "text"` always uses "contains" — no "starts with"/"ends with"/"not
   contains," despite Excel natively supporting all of these, and nothing documents
   the restriction.
4. `kind: "duplicate"` is hardcoded to highlight duplicates only — no way to flip to
   highlight uniques, undocumented.

**Evidence:** `entry.ts:222`, `ExcelTools.cs:400-476` (kind dispatch), `:388-398`
(`MapCfOperator` silent default), `:418-423` (text kind hardcoded to `xlContains`),
`:427-430` (duplicate kind hardcoded to `xlDuplicate`).

**User impact:** conditional formatting is one of the more commonly-requested
spreadsheet operations; right now the model has to guess field shapes for most of its
variants and several requests ("highlight cells NOT containing X", "highlight unique
values") are silently impossible without any indication why.

**Desired outcome:** every `add_conditional_format` rule kind has documented,
schema-discoverable fields; the text-match and duplicate/unique modes cover what Excel
natively supports.

---

### PP-15: Chart tools — narrower type support on creation, no rebinding, undocumented `edit_chart` breadth

**Priority:** P1

**Problem, three related findings:**
1. `add_chart`'s `chartType` only reliably supports `column`/`line`/`pie` — anything
   else silently falls back to column, with no error.
2. `edit_chart` actually supports 6 chart types (column/bar/line/area/pie/doughnut,
   per `ExcelChartTypeMap`), but only `add_chart`'s narrower 3-type list is ever
   documented anywhere — the model has no way to discover `edit_chart` can do
   bar/area/doughnut.
3. There's no way to repoint an existing chart at a different/extended data range —
   `edit_chart` can only change cosmetic properties (colors, title, legend, series
   names, data-label mode), never rebind `SetSourceData`.

**Evidence:** `entry.ts:210-211`, `ExcelTools.cs:58-66` (`ExcelChartTypeMap`, 6
entries), `:643-663` (`AddChart`, add-path fallback), `:665-726` (`EditChartExcel`).

**User impact:** requesting a bar/area/doughnut chart via natural language silently
produces a column chart instead; a chart can never be "updated to reflect the new
data range" without deleting and recreating it.

**Desired outcome:** the full 6-type vocabulary is available (and documented) on chart
creation, not just editing; an existing chart can be rebound to a new range.

---

### PP-16: `add_shape`'s 26 valid preset names are entirely undocumented

**Priority:** P2

**Problem:** the handler supports 26 named shape presets plus `"textbox"`
(`ShapeTypeMap`), but neither the schema nor the description lists any of them — an
unrecognized name silently becomes a plain rectangle, with no error.

**Evidence:** `entry.ts:213`, `ExcelTools.cs:27-56` (the 26-entry map), `:746-773`
(silent fallback).

**User impact:** a request for a specific shape ("add a star", "add an arrow pointing
right") has roughly a 1-in-27 chance of the model happening to guess a recognized
name; otherwise it silently becomes a rectangle.

**Desired outcome:** the shape-name vocabulary is documented (ideally as a real
`enum`, per PP-5) so requests for specific shapes reliably work.

---

### PP-17: Defined names are workbook-scoped only — no sheet-scoped names

**Priority:** P2

**Problem:** `add_defined_name`/`delete_defined_name` only ever touch
`ActiveWorkbook.Names` — there's no `sheet?` parameter — yet `read_sheet_features`
already reports sheet-scoped defined names as something that can exist. The model can
discover a sheet-scoped name but has no operation to create one.

**Evidence:** `entry.ts:171` (read_sheet_features description), `ExcelTools.cs:1124-
1135` (workbook-only `Add`).

**User impact:** narrow but real — any workflow relying on sheet-scoped named ranges
(common in templated financial models) can't be authored by the AI.

**Desired outcome:** `add_defined_name`/`delete_defined_name` accept an optional
sheet scope, matching what `read_sheet_features` already reports.

---

### PP-18: Assorted smaller reliability gaps — `set_page_setup`, `delete_table`, `add_sparkline`

**Priority:** P2

**Problem, three small independent findings:**
1. `set_page_setup`: combining `scale` and `fitToWidth`/`fitToHeight` in one call
   silently discards `scale` (the `fitToWidth` branch runs after and unconditionally
   sets `Zoom = false`) — Excel's own UI treats these as mutually exclusive, but the
   model is never told, so nothing stops it from combining them and getting a
   silently-wrong result.
2. `delete_table` calls `.Unlist()`, which converts the table back to a plain range
   and **keeps all the data** — the name/description reads as "remove the table and
   its data," which isn't what happens.
3. `add_sparkline`: if `targetCell` is omitted, it defaults to the *same* cells as
   `dataRange` — the sparkline draws inside the very cells holding its own source
   data — undocumented.

**Evidence:** `ExcelTools.cs:857-865` (page setup conflict), `:1085-1088`
(`delete_table`), `:728-744` (`add_sparkline` default).

**User impact:** each is a plausible-but-wrong outcome from a reasonable request, with
no error to signal it went wrong.

**Desired outcome:** page setup rejects/warns on the conflicting combination;
`delete_table`'s behavior either matches its name or is renamed/documented
accurately; `add_sparkline` requires or clearly documents its default target.

---

## Part 4 — PowerPoint (`PowerPointAiAddIn`)

### PP-19: No scripting DSL, no deck-generation pipeline, no `delete_slide`, no automatic post-edit QC

**Priority:** P2 (large scope — confirm appetite before planning implementation)

**Problem:** genoffice's slides app has `execute_slide_script` (a sandboxed AST-walked
scripting DSL for atomic multi-element edits), a full deck-generation pipeline
(`ask_clarification`/`plan_deck`/`generate_deck`/`regenerate_slide`/
`save_style_template`/`list_style_templates`), and an automatic geometric + vision-based
QC pass after every generated page. officeoffice's PowerPoint add-in has none of this —
confirmed absent by full-file read. This is explicitly out-of-scope per this project's
existing planning docs, not an oversight.

**Evidence:** grep of `PowerPointTools.cs`/`entry.ts`, none present; scope boundary
stated in `docs/tool-surface-todo.md`'s header (even though that file is stale on
implementation counts, its *scope* note here is accurate) and reflected in
`docs/ai-tool-surface.md`'s "explicitly out of scope" section.

**User impact:** any multi-property, multi-element edit must go tool-by-tool
(`set_element_transform` one shape at a time) rather than atomically; there is no
"generate me a deck about X" capability at all; a slide, once added, can never be
removed by the AI; nothing automatically catches overflow/overlap/misalignment after
an AI edit.

**Desired outcome:** **product-owner decision required** — this is the single largest
scope item in this plan by far (genoffice's version is ~3,500 lines of DSL +
interpreter + generation pipeline + QC). Recommend treating this as its own
initiative with an explicit scope/feasibility pass, not folded into a general
"fix PowerPoint" implementation plan. At minimum, `delete_slide` (missing for no
apparent reason, unlike `add_slide` which exists) is a small, clearly-in-scope fix
that could be pulled out and done independently of the larger DSL/generation
question.

---

### PP-20: `set_element_style` and `add_shape` are narrower than genoffice — but honestly so

**Priority:** P2

**Problem:** `set_element_style` covers bold/italic/fontSize/color only — missing
underline/align/font family, which genoffice's equivalent has. `add_shape` supports
only 3 preset types (rectangle/oval/roundRect) vs. genoffice's full OOXML
preset-geometry gallery — notably narrower than **Excel's own** 26-type shape map in
this same repo, which wasn't reused here. Both sides (schema and handler) agree with
each other in both cases — this is a real capability gap, not a schema/handler
mismatch.

**Evidence:** `PowerPointTools.cs:158-173` (`set_element_style`), `:198-213`
(`add_shape`, 3 types); `ExcelTools.cs:27-56` (the 26-type map that exists elsewhere in
this repo).

**User impact:** text formatting requests involving underline/alignment/font-family
silently have no effect via this tool; shape requests outside rect/oval/roundRect
can't be fulfilled.

**Desired outcome:** `set_element_style` reaches the same field coverage as Word's
`updateTextStyle`; `add_shape` reuses (or ports) Excel's existing 26-type
`ShapeTypeMap` rather than maintaining a separate, narrower one.

---

### PP-21: `edit_chart`'s `chartType`/`legendPos` silently no-op with false success

**Priority:** P1

**Problem:** `chartType` only applies if the value is found in `PptChartTypeMap` — an
unrecognized/typo'd value is silently ignored, and the tool still returns "Chart
updated." `legendPos`'s schema is bare `{type:'string'}` with no valid-value guidance;
the handler expects exactly `"none"`/`"r"`/`"t"`/`"l"`, so any natural phrasing a model
would plausibly send (`"right"`, `"bottom"`) silently falls into the bottom-position
branch.

**Evidence:** `PowerPointTools.cs:510-513` (chartType), `:519-531` (legendPos);
`entry.ts:356` (undocumented schema).

**User impact:** "change this to a bar chart" or "move the legend to the right" can
silently do nothing or do the wrong thing, while the AI reports success either way.

**Desired outcome:** both parameters have documented `enum` values matching what the
handler actually accepts, and an out-of-range value is rejected with an error the
model can react to, not silently swallowed.

---

### PP-22: `add_chart.kind` / `add_smartart.layout` — undocumented enums, silent fallback

**Priority:** P2

**Problem:** unrecognized `add_chart.kind` silently defaults to `"bar"`; unrecognized
`add_smartart.layout` silently defaults to `"list"` ("Basic Block List"). Neither has
a declared `enum` in its schema (unlike `add_shape.shapeType`, which does declare
one, showing the pattern is already known/available in this codebase).
`edit_table_structure.kind` and `edit_table_style.borderPreset` have the same
undocumented-enum gap, lower severity since their value sets are smaller/easier to
guess correctly.

**Evidence:** `PowerPointTools.cs:442` (add_chart fallback), `:564` (add_smartart
fallback); `entry.ts` (all four, no enum declared).

**User impact:** requesting an unusual chart kind or SmartArt layout produces a
plausible-looking but wrong result with no error.

**Desired outcome:** all four parameters get declared `enum` schemas, following the
existing `add_shape.shapeType` pattern already in this codebase.

---

## Summary table

| ID | Area | Title | Priority |
|---|---|---|---|
| PP-1 | Cross-cutting | Task pane doesn't appear in secondary document windows | P0 |
| PP-2 | Cross-cutting | Tool calls render after the final response, not before | P0 |
| PP-3 | Cross-cutting | Tool output never shown in the UI | P1 |
| PP-4 | Cross-cutting | Silent "(no text)" replies from token-limit truncation | P0 |
| PP-5 | Cross-cutting | Gateway tools lack structural per-command schemas | P1 |
| PP-6 | Cross-cutting | Provider/model selection not wired to the transport | P1 |
| PP-7 | Cross-cutting | No dataSource/provenance guardrail (confirm scope) | P2 |
| PP-8 | Cross-cutting | Stale `tool-surface-todo.md` checklist | P2 |
| PP-9 | Word | Chart tool can't set categories/series names | P1 |
| PP-10 | Word | Content tools are plain-text-only, append-only insert | P1 |
| PP-11 | Word | No image-insertion tool at all | P1 |
| PP-12 | Word | Missing `highlight`, decorative `bulletPreset`, no-rollback batch abort | P1 |
| PP-13 | Excel | `format_range` missing most style properties | P1 |
| PP-14 | Excel | Conditional formatting: undocumented fields, silent fallbacks | P1 |
| PP-15 | Excel | Chart type/rebinding gaps | P1 |
| PP-16 | Excel | `add_shape`'s 26 types undocumented | P2 |
| PP-17 | Excel | Defined names workbook-scope only | P2 |
| PP-18 | Excel | `set_page_setup`/`delete_table`/`add_sparkline` small gaps | P2 |
| PP-19 | PowerPoint | No scripting DSL / deck generation / `delete_slide` / QC | P2 (confirm scope) |
| PP-20 | PowerPoint | `set_element_style`/`add_shape` narrower than genoffice | P2 |
| PP-21 | PowerPoint | `edit_chart` chartType/legendPos silent no-op | P1 |
| PP-22 | PowerPoint | `add_chart.kind`/`add_smartart.layout` undocumented enums | P2 |
