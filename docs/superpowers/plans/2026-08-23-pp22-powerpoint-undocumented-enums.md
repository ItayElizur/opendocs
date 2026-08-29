# PP-22: PowerPoint Undocumented Enums — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-22 (P2).

**Goal:** Four PowerPoint parameters that accept a closed set of strings get declared `enum` schemas and reject out-of-set values, following the `add_shape.shapeType` pattern already used in this same file.

**Architecture:** Four instances of one pattern — a handler with a `TryGetValue`-or-default lookup, and a schema field typed as bare `{ type: 'string' }`:

| Parameter | Handler | Current fallback |
|---|---|---|
| `add_chart.kind` | `PowerPointTools.cs:441` | unrecognized → 51 (documented as "bar", actually xlColumnClustered) |
| `add_smartart.layout` | `PowerPointTools.cs:564` (`ResolveSmartArtLayout`) | unrecognized → "Basic Block List" |
| `edit_table_structure.kind` | `PowerPointTools.cs` table section | 4 values, smaller set, lower severity |
| `edit_table_style.borderPreset` | `PowerPointTools.cs` table section | `"all"`/`"none"`, lowest severity |

`add_shape.shapeType` already declares `enum: ['rectangle','oval','roundRect']` (`PowerPointAiAddIn/web-src/entry.ts:220`) — the pattern is known and available in this codebase; these four just never got it.

Two of the four carry real hazards beyond the missing enum:

- **`add_chart.kind` shares `PptChartTypeMap` with `edit_chart.chartType`**, whose `"bar"` entry maps to xlColumnClustered rather than xlBarClustered. PP-21 Task 1 Step 2 fixes the map. **Land PP-21 first**, then derive this enum from the corrected map — otherwise this plan ships an enum that advertises a wrong mapping.
- **`add_smartart.layout`** maps seven keys to English display names (`PowerPointTools.cs:553-562`) which are then matched against `Application.SmartArtLayouts` by name (`:564-575`). The file's own comment (`:549-552`) records that the live cross-check against a real Office install was never performed — it needs interactive GUI access. So the enum here is only as good as those display names, and on a non-English Office install the lookup falls back to "Basic Block List" for *every* layout. Task 2 addresses that directly; it is the one genuinely uncertain item in this plan.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.PowerPoint`; TypeScript for the schema.

**Dependency:** PP-21 (`2026-08-23-pp21-powerpoint-edit-chart-silent-noop.md`) for the chart-type map correction.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Every enum in the schema must be generated from, or verified line-by-line against, the handler's map. An enum that lists a value the handler does not accept is worse than no enum.
- No silent fallbacks. Unrecognized value → specific error listing valid ones, built from the map's keys so it cannot drift.
- Existing valid values keep working.
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: `add_chart.kind`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Confirm PP-21 Task 1 Step 2 has landed and `PptChartTypeMap` maps `"bar"` to 57 and `"column"` to 51. If not, stop and do that first.
- [ ] **Step 2:** Replace the fallback at `:441`:

```csharp
string kindStr = input.GetProperty("kind").GetString();
int typeCode;
if (!PptChartTypeMap.TryGetValue(kindStr, out typeCode))
    throw new ArgumentException("add_chart: unknown kind '" + kindStr + "'. Valid: " +
                                string.Join(", ", PptChartTypeMap.Keys) + ".");
```

- [ ] **Step 3: Schema** — `kind: { type: 'string', enum: [...] }` from the corrected map, replacing `{ type: 'string' }` at `entry.ts:340`. It must be the *same* list `edit_chart.chartType` gets in PP-21 Task 1 Step 4 — one vocabulary, two fields.
- [ ] **Step 4:** While here, validate `categories`/`series` shape. `AddChartPpt` writes a grid from `categories` and each series' `values` (`:465-480`); a length mismatch produces a silently wrong chart. Throw if any series' `values` length differs from `categories` length.
- [ ] **Step 5:** Return the created chart's shape index in the result, so the model can `edit_chart` it without re-reading the slide. `AddChartPpt` currently returns a flat `"Chart added."` (`:502`).

**Verification:** `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug`; each kind produces the right chart; an unknown kind errors specifically; a ragged series errors specifically.

---

### Task 2: `add_smartart.layout` — enum plus the localization hazard

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Do the live cross-check the code comment asks for** (`PowerPointTools.cs:549-552`). On a machine with interactive Office access, enumerate `Application.SmartArtLayouts` and dump every `layout.Name` — a few lines in the VBA immediate window or a temporary tool call. Compare against `SmartArtLayoutNames`' seven display names.
- [ ] **Step 2:** Correct any name that does not match, and record the Office version and locale the check was performed on, replacing the "remains a manual follow-up" comment with the result.
- [ ] **Step 3: Fix the fallback.** `ResolveSmartArtLayout` (`:564-575`) falls back to `"Basic Block List"` twice: once when the key is unknown (`:566`), and once implicitly if the name match finds nothing. Make both throw:
  - unknown key → error listing the seven valid keys;
  - key valid but no matching layout found in `SmartArtLayouts` → a distinct error naming the display name it looked for and stating that this Office install may be non-English. That second message is what will actually diagnose a localized install, and it cannot be produced by a silent fallback.
- [ ] **Step 4: Consider index-based lookup as a locale-independent alternative.** `SmartArtLayouts` is index-addressable, and the built-in gallery order is stable across installs. **Recommendation: do not switch** — index stability across Office versions is an assumption no better founded than the name assumption, and names at least fail loudly with Step 3's message. Record the option and the reasoning in a comment.
- [ ] **Step 5: Schema** — `layout: { type: 'string', enum: ['list','process','cycle','hierarchy','pyramid','matrix','venn'] }`. The description already lists these values in prose (`entry.ts:363`); make them structural.
- [ ] **Step 6:** If Step 1's cross-check cannot be performed (no GUI access), ship Steps 3 and 5 anyway — the loud failure is a strict improvement — and leave the comment saying the names remain unverified, with the specific error message as the diagnostic path.

**Verification:** build; each layout key produces a visibly different diagram, or fails with the diagnostic message; an unknown key errors specifically.

---

### Task 3: Table parameters

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: `edit_table_structure.kind`** — `enum: ['insert-row','delete-row','insert-col','delete-col']` (schema at `entry.ts:317`, currently bare string with the values only in prose). Confirm the exact spellings against the handler's dispatch before writing them; a hyphen-vs-underscore mismatch would break every call.
- [ ] **Step 2:** Add a `default:` throw to that dispatch if it lacks one — check whether an unrecognized `kind` currently falls through to a no-op-plus-success.
- [ ] **Step 3: `edit_table_style.borderPreset`** — `enum: ['all','none']` (schema at `entry.ts:329`). Compare with Excel's border preset vocabulary after PP-13 Task 2 (`none`/`outline`/`all`/`thick-outline`); if PowerPoint's handler can support `outline` cheaply, add it and align the two. If not, leave the two-value enum and note the difference in `docs/ai-tool-surface.md` rather than pretending parity.
- [ ] **Step 4:** Check `edit_table_structure`'s `index` bounds. Inserting at an out-of-range index throws a raw COM error; validate against the table's current row/column count and report a specific message including the valid range.
- [ ] **Step 5:** Note the index-shift hazard in `edit_table_structure`'s description — deleting row 2 shifts every later row's index, the same trap as PP-19's `delete_slide`.

**Verification:** build; each `kind` value works; unknown values error; an out-of-range index errors specifically.

---

### Task 4: Sweep the rest of the PowerPoint surface

**Files:**
- Modify: `PowerPointAiAddIn/web-src/entry.ts` (and handlers as found)

- [ ] **Step 1:** Read all 23 tool schemas in `PowerPointAiAddIn/web-src/entry.ts` (`:146-400`) and list every `{ type: 'string' }` field whose handler compares against a closed set. Known candidates beyond the four above: `set_slide_background`'s parameters, `crop_image`'s, `set_element_stroke`'s dash style if it has one.
- [ ] **Step 2:** For each, add the enum and a throwing default, same as Tasks 1-3.
- [ ] **Step 3:** Skip fields owned by other plans: `edit_chart.chartType` and `legendPos` are PP-21; `add_shape.shapeType` breadth is PP-20.
- [ ] **Step 4:** Record the completed list in `docs/ai-tool-surface.md`'s PowerPoint section so the next audit starts from a verified baseline rather than re-deriving it.

**Verification:** no `{ type: 'string' }` field in the PowerPoint schema maps to a closed handler set without an enum, except ones explicitly noted as owned elsewhere.

---

### Task 5: Manual verification matrix

- [ ] `add_chart` with each valid `kind` → the right chart type (particularly `bar` vs `column` after PP-21's map fix).
- [ ] `add_chart {kind: 'nonsense'}` → specific error listing valid kinds; no chart added.
- [ ] `add_chart` with a series whose `values` length differs from `categories` → specific error.
- [ ] `add_smartart` with each of the seven layouts → seven visibly different diagrams (this is the check that validates or invalidates the display-name table).
- [ ] `add_smartart {layout: 'nonsense'}` → specific error listing valid layouts.
- [ ] If a layout key resolves to no gallery entry → the distinct "may be a non-English Office install" message, not a silent Basic Block List.
- [ ] `edit_table_structure` with each of the four kinds → correct insert/delete.
- [ ] `edit_table_structure {kind: 'nonsense'}` and an out-of-range `index` → specific errors.
- [ ] `edit_table_style {borderPreset: 'all'|'none'}` → correct; an unknown value → specific error.
- [ ] Natural language: "add an org chart", "add a cycle diagram", "insert a row after row 2" each work on the first attempt.
