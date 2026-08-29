# PP-14: Excel Conditional Formatting — Documented Fields and Full Match Modes

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-14 (P1) — four related findings on `add_conditional_format`.

**Goal:** Every rule kind has schema-discoverable fields; the text-match and duplicate/unique modes cover what Excel natively supports; and no unrecognized value is silently swallowed.

**Architecture:** Four defects in `AddConditionalFormat` (`ExcelAiAddIn/ExcelTools.cs:400-476`) and its schema line (`ExcelAiAddIn/web-src/entry.ts:222`):

1. **Undocumented per-kind fields.** The schema says `rule: {kind:"number"|"text"|"blank"|"duplicate"|"top10"|"formula"|"colorScale"|"dataBar", ...kind-specific fields, format?:{...}}`. The literal string "…kind-specific fields" is the entire contract for eight different shapes — the worst single instance of PP-5's root cause. The handler's real requirements, read from source: `number` needs `operator` + `value` (`value2` optional); `text` needs `text`; `top10` takes `rank`/`percent`/`bottom`; `formula` needs `formula`; `colorScale` takes `minColor`/`midColor`/`maxColor`; `dataBar` takes `color`; `blank` and `duplicate` take nothing.
2. **`MapCfOperator` silent default** (`:388-398`): anything not `greaterThan`/`lessThan`/`equal`/`between` becomes `xlEqual`. So `"notEqual"` — a perfectly reasonable guess — silently inverts the user's intent.
3. **Text kind hardcoded to `xlContains`** (`:418-423`). Excel's `XlContainsOperator` natively offers `xlContains`, `xlDoesNotContain`, `xlBeginsWith`, `xlEndsWith`. Three of four are unreachable and undocumented.
4. **Duplicate kind hardcoded to `xlDuplicate`** (`:427-430`), despite the code already going through `AddUniqueValues()` + `DupeUnique` — the one-line flip to highlight uniques instead is right there, unexposed.

Additionally, `MapCfOperator` covers only 4 of Excel's 8 `XlFormatConditionOperator` cell-value comparisons — `notEqual`, `greaterEqual`, `lessEqual`, and `notBetween` are all missing, which is why finding 2 bites so easily.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Excel`; TypeScript for the schema.

**Relationship to PP-5:** if `2026-08-23-pp05-gateway-tool-schemas.md` has landed, Task 1 here is largely done via its `EXCEL_OPS` table — verify and extend rather than duplicating. Tasks 2-4 are independent of it.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **No silent fallbacks anywhere in this file's conditional-formatting path.** Every unrecognized enum value becomes a specific error naming the value and listing valid ones. This is the item's whole point; a fix that adds a new silent default has failed.
- Do not change the working parts: `colorScale`'s 3-criteria construction, `dataBar`, `top10`'s early-return format application and its documented PIA quirk (`:465`), and the shared `format` block at `:471-476` all work — leave them alone except where a task explicitly says otherwise.
- Reuse `HexToOleColor` (`ExcelTools.cs:775-782`) for every color.
- No automated tests for COM executor methods (project convention). Verification is build + Task 6's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: Per-kind schema

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Write the discriminated union**, one branch per kind, each pinning `kind` and declaring that kind's real fields with `required` lists derived from the handler's `GetProperty` (required) vs. `TryGetProperty` (optional) calls:

```ts
const CF_RULE_SCHEMA = {
  oneOf: [
    { type: 'object', properties: { kind: { const: 'number' },
        operator: { type: 'string', enum: ['greaterThan','lessThan','equal','notEqual','greaterEqual','lessEqual','between','notBetween'] },
        value: { type: 'number' }, value2: { type: 'number', description: 'Upper bound; required for between/notBetween.' },
        format: CF_FORMAT_SCHEMA },
      required: ['kind', 'operator', 'value'] },
    { type: 'object', properties: { kind: { const: 'text' },
        text: { type: 'string' },
        match: { type: 'string', enum: ['contains','notContains','beginsWith','endsWith'], description: 'Default contains.' },
        format: CF_FORMAT_SCHEMA },
      required: ['kind', 'text'] },
    { type: 'object', properties: { kind: { const: 'blank' }, format: CF_FORMAT_SCHEMA }, required: ['kind'] },
    { type: 'object', properties: { kind: { const: 'duplicate' },
        mode: { type: 'string', enum: ['duplicate','unique'], description: 'Default duplicate.' },
        format: CF_FORMAT_SCHEMA },
      required: ['kind'] },
    { type: 'object', properties: { kind: { const: 'top10' },
        rank: { type: 'number', description: 'Default 10.' },
        percent: { type: 'boolean' }, bottom: { type: 'boolean' }, format: CF_FORMAT_SCHEMA },
      required: ['kind'] },
    { type: 'object', properties: { kind: { const: 'formula' },
        formula: { type: 'string', description: 'Excel formula relative to the range\'s top-left cell, e.g. "=$C1>100".' },
        format: CF_FORMAT_SCHEMA },
      required: ['kind', 'formula'] },
    { type: 'object', properties: { kind: { const: 'colorScale' },
        minColor: { type: 'string' }, midColor: { type: 'string' }, maxColor: { type: 'string' } },
      required: ['kind'] },
    { type: 'object', properties: { kind: { const: 'dataBar' }, color: { type: 'string' } }, required: ['kind'] },
  ],
}

const CF_FORMAT_SCHEMA = {
  type: 'object',
  properties: { bold: { type: 'boolean' }, fontColor: { type: 'string' }, fillColor: { type: 'string' } },
  description: 'Applied when the rule matches. Not supported for colorScale/dataBar, which carry their own visual.',
}
```

- [ ] **Step 2:** Note in `colorScale`/`dataBar`'s descriptions that `format` is ignored for them — the handler returns early (`:462`, `:471`) before reaching the shared format block, and today nothing says so.
- [ ] **Step 3:** Note that `formula`'s expression is relative to the range's anchor cell — the single most common way a `formula` rule silently applies to the wrong cells.
- [ ] **Step 4:** Replace the single description line at `ExcelAiAddIn/web-src/entry.ts:222` with a short pointer to the schema, and wire `CF_RULE_SCHEMA` into `add_conditional_format`'s `rule` property.

**Verification:** bundle rebuilds; a natural-language "highlight cells over 100 in red" produces a well-formed `number` rule on the first attempt.

---

### Task 2: Complete the operator map, remove the silent default

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Extend `MapCfOperator` (`:388-398`) to all eight comparisons: `greaterThan`→`xlGreater`, `lessThan`→`xlLess`, `equal`→`xlEqual`, `notEqual`→`xlNotEqual`, `greaterEqual`→`xlGreaterEqual`, `lessEqual`→`xlLessEqual`, `between`→`xlBetween`, `notBetween`→`xlNotBetween`.
- [ ] **Step 2:** Replace `default: return xlEqual` with a throw listing the valid names.
- [ ] **Step 3:** Validate `value2`: required for `between`/`notBetween`, meaningless otherwise. Missing it today produces `null` for `Formula2` (`:411`) and a rule that silently matches nothing. Throw a specific error instead.
- [ ] **Step 4:** `value` is read with `GetDouble()` (`:410`) — a string value like `"100"` throws an opaque `InvalidOperationException`. Accept both number and numeric string, and throw a clear error otherwise.

**Verification:** build; `operator: 'notEqual'` produces a real ≠ rule; `operator: 'nonsense'` errors specifically; `between` without `value2` errors specifically.

---

### Task 3: Full text match modes

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Read the optional `match` field and map it: `contains`→`xlContains`, `notContains`→`xlDoesNotContain`, `beginsWith`→`xlBeginsWith`, `endsWith`→`xlEndsWith`. Default `contains`, preserving today's behavior for calls that omit it.
- [ ] **Step 2:** Unknown `match` → specific error.
- [ ] **Step 3:** Confirm at build time that this PIA exposes all four `XlContainsOperator` members. `ExcelTools.cs:22-26` sets the precedent for what to do if one is missing: drop it, remove it from the schema enum, and record why in a comment — never substitute a near-match.
- [ ] **Step 4:** `xlTextString` is the right `XlFormatConditionType` for all four (`:420`); verify `xlBeginsWith`/`xlEndsWith` work with it in this PIA rather than requiring `xlBeginsWith`-typed conditions, and adjust if the live test says otherwise.

**Verification:** build; each of the four modes produces the correct rule in Excel's own Manage Rules dialog.

---

### Task 4: Unique-values mode

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Read the optional `mode` field and set `((Excel.UniqueValues)fc).DupeUnique = mode == "unique" ? Excel.XlDupeUnique.xlUnique : Excel.XlDupeUnique.xlDuplicate;` — a one-line change to `:429`.
- [ ] **Step 2:** Default `duplicate`, preserving today's behavior.
- [ ] **Step 3:** Unknown `mode` → specific error.
- [ ] **Step 4:** Consider whether the kind name `duplicate` should become `duplicateOrUnique`. **Recommendation: keep `duplicate`.** Renaming breaks any saved conversation and buys little; the `mode` field plus the schema description carries the meaning.

**Verification:** build; `mode: 'unique'` highlights the values that appear once.

---

### Task 5: Rule-kind dispatch hardening

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** The `switch (kind)` (`:407-472`) has **no `default` branch** — an unrecognized kind falls through with `fc == null`, skips the format block (`:474`), and returns having done nothing, while `ProposeOperations` reports `add_conditional_format: ok` (`:552`). Another false success. Add `default: throw new ArgumentException(...)` listing valid kinds.
- [ ] **Step 2:** `rule.GetProperty("kind")` (`:406`) throws a bare `KeyNotFoundException` when `kind` is missing. Use `TryGetProperty` with a specific message.
- [ ] **Step 3:** `op.GetProperty("range")` (`:401`) likewise — and an invalid A1 address throws an opaque COM error. Wrap with a message naming the address.
- [ ] **Step 4:** Report what was created. `ProposeOperations`' line is a bare `ok`; make `AddConditionalFormat` return a short description (kind, range, and the effective operator/mode) that the batch loop appends, so the transcript — and PP-3's output view — shows what rule actually landed.

**Verification:** build; an unknown kind errors specifically instead of reporting `ok`.

---

### Task 6: Manual verification matrix

- [ ] Every one of the 8 kinds creates the expected rule, checked in Excel's Home > Conditional Formatting > Manage Rules dialog.
- [ ] All 8 `number` operators, including the four newly added.
- [ ] `between` and `notBetween` with `value2`; without it → specific error.
- [ ] All 4 text `match` modes, including "highlight cells NOT containing X" — the request the source item names as currently impossible.
- [ ] `duplicate` with `mode: 'unique'` — the second named-impossible request.
- [ ] Unknown values for `kind`, `operator`, `match`, `mode` each produce a specific error naming valid values, and create no rule.
- [ ] `format: {bold, fontColor, fillColor}` applies on `number`/`text`/`blank`/`duplicate`/`top10`/`formula`.
- [ ] `colorScale` and `dataBar` still work unchanged.
- [ ] `top10` with `percent`/`bottom` still works unchanged (the early-return path).
- [ ] Missing `kind` / bad range address → specific errors.
- [ ] Natural language end-to-end: "highlight cells in B2:B50 that don't contain 'approved' in red" succeeds on the first attempt with no field-name guessing.
