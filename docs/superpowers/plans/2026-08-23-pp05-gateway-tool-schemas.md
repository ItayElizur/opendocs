# PP-5: Structural Per-Command JSON Schemas for the Gateway Tools — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-5 (P1 — root cause under several PP-12/14/15/16/21/22 findings).

**Goal:** Give Word's `apply_commands` and Excel's `propose_operations` real, machine-readable schemas — per-`kind` required fields and `enum` arrays for every closed string set — so the model discovers the contract structurally instead of guessing field names from a prose paragraph, and so a malformed command is rejected with a specific error instead of reaching a COM handler.

**Architecture:** Both gateways currently declare `items: { type: 'object' }` — no structure at all (`WordAiAddIn/web-src/entry.ts:239`, `ExcelAiAddIn/web-src/entry.ts:225`). The entire contract is one description string: Word's is ~10 lines (`:225-238`), Excel's is ~30 lines covering ~50 operation kinds (`:198-224`).

The right JSON Schema construct is a **discriminated union**: `commands: { type: 'array', items: { oneOf: [ ...one schema per kind... ] } }`, each branch pinning `kind: { const: 'X' }` plus that kind's own `properties`/`required`. Two practical caveats drive the design:

- Not every provider enforces `oneOf` in tool schemas, and some flatten it. So the schema is treated as *documentation the model reads*, not as a validator that will run — which means **Task 4's runtime validation in C# is not optional**, it is where the actual guarantee lives.
- Excel has ~50 kinds. A hand-written 50-branch literal in `entry.ts` would be unmaintainable and would drift from the handler. Task 2 therefore builds the schema from a single declarative table that also feeds the description text, so both are generated from one source.

This plan delivers the mechanism plus the full Word surface, the mechanism plus the full Excel surface, and a runtime validator. The individual capability items (PP-12, PP-14, PP-15, PP-16, PP-22) then only need to add or correct entries in these tables.

**Tech Stack:** TypeScript (`entry.ts` per app), C# 7.3 / .NET Framework 4.8 (`WordTools.cs`, `ExcelTools.cs`).

## Global Constraints

- No behavior change to any handler in this plan except the new validation path and its error text. Adding capability comes from the other PP items.
- The schema must describe **what the handler actually does today**, not what it should do. Where the handler is narrower than the prose (e.g. Excel `add_chart` reliably supports only column/line/pie — `ExcelAiAddIn/ExcelTools.cs:653-657`), the enum states the narrow truth. Widening is PP-15's job; this plan must not silently paper over a gap.
- Keep the human-readable `description` string too. Some providers weight prose heavily and it costs little. Task 2 generates it from the same table so it cannot drift.
- C# 7.3 / .NET 4.8 only in `.cs` — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Rebuild bundles + MSBuild after each `entry.ts` change (command in `2026-08-23-pp02-tool-steps-chronological-order.md`'s Global Constraints).
- Watch schema size. A 50-branch `oneOf` inflates every request's token cost. Task 5 measures it and, if it is material, falls back to the "grouped kinds" variant described there.

---

### Task 1: Word — discriminated-union schema for `apply_commands`

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `const WORD_COMMAND_SCHEMAS` — an array of per-kind JSON Schema objects, consumed by the `apply_commands` tool definition and (Task 4) mirrored by the C# validator's required-field table.

- [ ] **Step 1: Define the shared `Target` schema once**

```ts
const TARGET_SCHEMA = {
  type: 'object',
  description: 'Selects paragraphs. Fields are AND-combined; at least one of nodeType/containsText/blockIndexes is required.',
  properties: {
    nodeType: { type: 'string', enum: ['heading', 'paragraph', 'listItem'] },
    headingLevel: { type: 'number', minimum: 1, maximum: 6 },
    containsText: { type: 'string' },
    matchCase: { type: 'boolean' },
    blockIndexes: { type: 'array', items: { type: 'number' }, description: '0-based paragraph indices' },
    scope: { type: 'string', enum: ['document', 'selection'] },
  },
} as const
```

This restates the contract `ResolveTargetParagraphs` actually enforces (`WordAiAddIn/WordTools.cs:307-377`), including its "at least one of" rejection.

- [ ] **Step 2: One branch per kind**

Enumerate all 12 kinds the handler's `switch` implements (`WordAiAddIn/WordTools.cs:220-266`): `set_bold`, `set_italic`, `set_heading`, `find_replace`, `updateTextStyle`, `updateParagraphStyle`, `deleteBlocks`, `moveBlocks`, `createParagraphBullets`, `deleteParagraphBullets`, `updateImageProperties`, `insertToc`. Each branch:

```ts
{ type: 'object', properties: { kind: { const: 'set_bold' }, startIndex: { type: 'number' }, endIndex: { type: 'number' }, value: { type: 'boolean' } }, required: ['kind', 'startIndex', 'endIndex', 'value'] }
```

Derive every `required` list from the handler's `GetProperty` calls (which throw when absent) versus `TryGetProperty` calls (which are optional) — read each private method rather than trusting the current description. Notably:
- `updateTextStyle`/`updateParagraphStyle` require `target`, `style`, `fields` (all three are `GetProperty` at `WordTools.cs:381-384`).
- `style`'s own keys get an explicit `properties` block enumerating exactly what `UpdateTextStyle` handles (`WordTools.cs:390-416`): `bold`, `italic`, `underline`, `strike`, `sizeHalfPoints`, `font`, `color`, `baselineOffset` (`enum: ['SUPERSCRIPT','SUBSCRIPT','NONE']`), `link` (`{ url }`). **`highlight` is deliberately absent** — the handler does not implement it (that is PP-12); listing it would be the exact false advertisement this plan exists to remove.
- `fields` becomes `{ type: 'array', items: { type: 'string', enum: [ ...the same key list... ] } }`.
- `createParagraphBullets.bulletPreset` gets an enum of only what `CreateParagraphBullets` distinguishes today — `['BULLET', 'NUMBERED']` — plus a description saying any other value is treated as `BULLET` (`WordTools.cs:547-548`). PP-12 widens this.
- `moveBlocks` requires `blockIndexes` + `afterBlockIndex`; `insertToc` requires `afterBlockIndex`; both document `-1` as "start of document".
- `updateImageProperties` requires `imageIndex`, `properties`, `fields`, with `properties` enumerated from `UpdateImageProperties` (`WordTools.cs:573-613`) including its alignment enum.

- [ ] **Step 3: Wire it into the tool definition**

```ts
inputSchema: {
  type: 'object',
  properties: { commands: { type: 'array', items: { oneOf: WORD_COMMAND_SCHEMAS } } },
  required: ['commands'],
}
```

Keep a shortened description that says "each command is one of the kinds below; see the schema for per-kind fields" rather than repeating every field in prose.

**Verification:** bundle rebuilds; `MSBuild WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds; a manual run asking for a bulleted list still works end to end.

---

### Task 2: Excel — table-driven schema generation for `propose_operations`

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `const EXCEL_OPS: OpSpec[]` and `function opSchemas(ops: OpSpec[])` + `function opsDescription(ops: OpSpec[])` — one source of truth feeding both the schema and the prose.

- [ ] **Step 1: The spec type**

```ts
interface OpSpec {
  kind: string
  group: 'Writing' | 'Formatting' | 'Layout' | 'Structure' | 'Charts/visuals' | 'Tables' | 'Pivot' | 'Data'
  /** JSON Schema properties, excluding `kind` and the shared optional `sheet` */
  props: Record<string, unknown>
  required?: string[]
  note?: string
}
```

- [ ] **Step 2: Populate one entry per kind implemented in `ProposeOperations`'s switch** (`ExcelAiAddIn/ExcelTools.cs:483` onward — enumerate every `case`, do not work from the description string). For each, read the corresponding private method and record `GetProperty` → required, `TryGetProperty` → optional.

- [ ] **Step 3: Enums, from the handlers' actual dispatch**
- `format_range`: exactly `bold`, `italic`, `numberFormat`, `fillColor` (`ExcelTools.cs:597-611`) — nothing more until PP-13 lands.
- `add_chart.chartType`: `enum: ['column','line','pie']` with a note that other values become column (`ExcelTools.cs:653-657`). PP-15 widens.
- `edit_chart.chartType`: `enum: ['column','bar','line','area','pie','doughnut']` from `ExcelChartTypeMap` (`ExcelTools.cs:58-66`); `legend`: `['none','right','top','left','bottom']`; `dataLabels`: `['none','value','percent']`.
- `add_shape.shapeType`: all 26 keys of `ShapeTypeMap` (`ExcelTools.cs:27-56`) plus `'textbox'`. This alone closes PP-16.
- `add_sparkline.type`: `['line','column','stacked']`.
- `set_page_setup.orientation`: `['portrait','landscape']`; `margins`: `['normal','wide','narrow']`.
- `sort_range.order`: `['asc','desc']`.
- `add_pivot` values' `agg`: `['sum','count','average','max','min']`.
- `add_conditional_format.rule`: a nested `oneOf` over its 8 kinds, each with its real fields read from `AddConditionalFormat` (`ExcelTools.cs:400-476`) — `number` needs `operator` (`enum: ['greaterThan','lessThan','equal','between']`, per `MapCfOperator` at `:388-398`) and `value`, `value2` optional; `text` needs `text`; `top10` takes `rank`/`percent`/`bottom`/`format`; `formula` needs `formula`; `colorScale` takes `minColor`/`midColor`/`maxColor`; `dataBar` takes `color`; `blank` and `duplicate` take nothing. This is the single biggest information gain in the plan and most of PP-14's documentation half.
- `set_data_validation.validation`: nested `oneOf` over `['list','listRef','numberBetween','dateBetween','formula']`, keeping the existing "checkbox is NOT supported, will error" note.

- [ ] **Step 4: Generators**

`opSchemas` maps each spec to `{ type:'object', properties: { kind: { const }, sheet: { type:'string' }, ...props }, required: ['kind', ...(required ?? [])] }`. `opsDescription` groups by `group` and emits the same one-line-per-kind prose the current description has, so the human-readable form stays but is generated.

- [ ] **Step 5:** Replace the hand-written description and the `items: {type:'object'}` schema with the generated pair.

**Verification:** bundle rebuilds; Excel project builds; a manual run doing a multi-op batch (set values, format, add a chart) still works.

---

### Task 3: Cross-check the tables against the handlers

**Files:** none modified (audit), then whichever of the above needs correcting

- [ ] **Step 1:** For Word, list every `case` in `ApplyCommands` (`WordTools.cs:220-266`) and every kind in `WORD_COMMAND_SCHEMAS`; the two sets must be identical. Any kind in the handler but not the schema is an undiscoverable capability; any in the schema but not the handler is a false promise.
- [ ] **Step 2:** Same for Excel's `ProposeOperations` switch vs. `EXCEL_OPS`.
- [ ] **Step 3:** Record the comparison result as a short table in `docs/ai-tool-surface.md` so the next audit starts from a verified baseline.

**Verification:** both sets match exactly, or every discrepancy is written down with a linked PP item.

---

### Task 4: Runtime validation in the C# gateways

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `ExcelAiAddIn/ExcelTools.cs`

**Problem:** the schema is advisory (providers vary in enforcement), so the actual guarantee has to be server-side. Today Word parses `kind` *outside* the per-command try/catch (`WordTools.cs:216-218`), so one command missing `kind` throws out of the whole loop and aborts the rest of the batch — the third finding in PP-12. Excel has the identical shape at `ExcelTools.cs:485`.

- [ ] **Step 1: Move `kind` extraction inside the try, in both files**

```csharp
foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
{
    string kind = null;
    try
    {
        if (!cmd.TryGetProperty("kind", out var kindEl) || kindEl.ValueKind != JsonValueKind.String)
            throw new ArgumentException("Command is missing a string \"kind\" field.");
        kind = kindEl.GetString();
        ...
    }
    catch (Exception ex)
    {
        lines.AppendLine((kind ?? "(unknown kind)") + ": ERROR - " + ex.Message);
        anyError = true;
    }
}
```

A malformed command now fails that command only; the rest of the batch continues, and the result line names it.

- [ ] **Step 2: Required-field precheck**

Add a table mirroring the schema's `required` lists:

```csharp
private static readonly Dictionary<string, string[]> RequiredFields = new Dictionary<string, string[]>
{
    ["set_bold"] = new[] { "startIndex", "endIndex", "value" },
    // ... one entry per kind, matching WORD_COMMAND_SCHEMAS exactly
};

private static void ValidateRequired(string kind, JsonElement cmd)
{
    string[] required;
    if (!RequiredFields.TryGetValue(kind, out required)) return;
    foreach (string f in required)
    {
        if (!cmd.TryGetProperty(f, out _))
            throw new ArgumentException("Command \"" + kind + "\" is missing required field \"" + f + "\".");
    }
}
```

Call it right after `kind` is read. The error text names the kind and the field, so the model can correct itself on the next turn instead of getting a generic COM exception.

- [ ] **Step 3:** Add a comment above each table pointing at its `entry.ts` counterpart and stating that the two must be edited together.

- [ ] **Step 4 (partial-batch honesty):** The existing result already lists per-command outcomes and sets `IsError` when any failed, so a half-applied batch is reported accurately once Step 1 lands. Do **not** attempt transactional rollback — Word/Excel COM offers no batch transaction, and a hand-rolled undo would be worse than the honest report. State this explicitly in a comment so it is a decision, not an omission.

**Verification:**
- [ ] Both projects build.
- [ ] Manual: send `apply_commands` with `[{valid command}, {no kind}, {valid command}]` — commands 1 and 3 apply, command 2 reports a specific error, and the tool result lists all three.
- [ ] Manual: send a command missing a required field — the result names the field.

---

### Task 5: Measure schema cost, and the fallback if it is too big

**Files:** possibly `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** After Task 2, log `JSON.stringify(ALL_TOOLS).length` for Excel before and after. Record both numbers in the commit message.
- [ ] **Step 2:** If the growth is material relative to the per-request budget (rough rule: more than ~4k added tokens), switch Excel to the **grouped** variant: keep full `oneOf` detail for the high-traffic, high-ambiguity kinds (`format_range`, `add_conditional_format`, `add_chart`, `edit_chart`, `add_shape`, `set_data_validation`, `add_pivot`) and collapse the rest into a single permissive branch that keeps only `kind`'s `enum` plus the generated prose line. That retains the entire discoverability win where guessing actually fails, at a fraction of the size.
- [ ] **Step 3:** Whichever variant ships, note it at the top of the generated-schema code so a later reader understands why some kinds are detailed and others are not.

**Verification:** a normal Excel request still completes within the provider's request limits, with no truncation of the tool list.
