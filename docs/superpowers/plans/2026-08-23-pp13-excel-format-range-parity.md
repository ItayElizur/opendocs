# PP-13: Excel `format_range` Property Parity — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-13 (P1).

**Goal:** `format_range` covers the property set genoffice's equivalent does — font family/size/color, underline, strikethrough, horizontal and vertical alignment, wrap, rotation, indent, and borders — so ordinary requests like "center this and add borders" or "wrap the text in column C" become achievable at all.

**Architecture:** `FormatRange` (`ExcelAiAddIn/ExcelTools.cs:597-611`) handles exactly four properties: `bold`, `italic`, `numberFormat`, `fillColor`. Every missing property maps directly onto a COM property that is already available on `Excel.Range` — there is no interop limitation here, only unimplemented code:

| Requested | COM target |
|---|---|
| fontName / fontSize / fontColor | `range.Font.Name` / `.Size` / `.Color` |
| underline / strikethrough | `range.Font.Underline` (`XlUnderlineStyle`) / `.Strikethrough` |
| horizontalAlignment / verticalAlignment | `range.HorizontalAlignment` / `.VerticalAlignment` (`XlHAlign`/`XlVAlign`) |
| wrapText | `range.WrapText` |
| textRotation | `range.Orientation` (-90..90, or `xlVertical`) |
| indent | `range.IndentLevel` (0..15) |
| borders | `range.Borders[XlBordersIndex]` — `.LineStyle`, `.Weight`, `.Color` |

The one genuinely non-trivial member is **borders**, because "add borders" has many meanings (outline only, all cells, one edge) and Excel models each edge plus the two interior grids separately. Task 2 gives it its own nested object rather than a boolean.

A second, easily-missed consequence: `read_formats` (`ExcelTools.cs:177-198`) reports only bold/italic/underline/numberFormat. After this change the model can *write* ten properties it still cannot *read*, so a read-modify-write cycle silently drops them. Task 3 closes that.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Excel`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Every new property is **optional**; an absent property leaves the cell's current value alone. `format_range` is additive, never a reset — existing calls must not start clearing properties they never mentioned.
- Every closed value set gets a real `enum` in the schema and a specific error on an unrecognized value. No silent fallbacks — that is the pattern PP-14/15/16 exist to remove, and this plan must not add new instances of it.
- Reuse the existing `HexToOleColor` helper (`ExcelTools.cs:775-782`) for every color; do not write a second hex parser. Note that `FormatRange`'s current `fillColor` branch inlines its own duplicate hex parsing (`:604-610`) — Task 1 Step 1 collapses it onto the shared helper.
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.
- If PP-5 (gateway schemas) has landed, add these properties to its `EXCEL_OPS` table rather than editing the description string.

---

### Task 1: Font, alignment, wrap, rotation, indent

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1: Collapse the duplicate hex parser**

Replace `FormatRange`'s inline `fillColor` parsing (`ExcelTools.cs:604-610`) with `range.Interior.Color = HexToOleColor(fc.GetString());`. Same behavior, one parser. Do this first — every color property added below depends on it.

- [ ] **Step 2: Font properties**

```csharp
if (op.TryGetProperty("fontName", out var fn) && fn.ValueKind == JsonValueKind.String) range.Font.Name = fn.GetString();
if (op.TryGetProperty("fontSize", out var fs) && fs.ValueKind == JsonValueKind.Number) range.Font.Size = fs.GetDouble();
if (op.TryGetProperty("fontColor", out var fcol) && fcol.ValueKind == JsonValueKind.String) range.Font.Color = HexToOleColor(fcol.GetString());
if (op.TryGetProperty("strikethrough", out var st)) range.Font.Strikethrough = st.ValueKind == JsonValueKind.True;
```

- [ ] **Step 3: Underline as an enum, not a boolean**

Excel supports `xlUnderlineStyleNone`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`. Accept `underline: 'none'|'single'|'double'|'singleAccounting'|'doubleAccounting'`, and **also** accept a plain boolean for ergonomics (`true` → single, `false` → none), since that is what a model will most often send. Reject any other string with a specific error listing valid values.

- [ ] **Step 4: Alignment**

```csharp
// horizontal: general|left|center|right|fill|justify|centerAcrossSelection|distributed
// vertical:   top|center|bottom|justify|distributed
```
Map each name to its `XlHAlign`/`XlVAlign` member via two small `Dictionary` statics, mirroring the file's existing `ShapeTypeMap`/`ExcelChartTypeMap` pattern (`ExcelTools.cs:27-66`). Unknown value → specific error.

- [ ] **Step 5: Wrap, rotation, indent**

```csharp
if (op.TryGetProperty("wrapText", out var wt)) range.WrapText = wt.ValueKind == JsonValueKind.True;
if (op.TryGetProperty("textRotation", out var tr) && tr.ValueKind == JsonValueKind.Number)
{
    int deg = tr.GetInt32();
    if (deg < -90 || deg > 90)
        throw new ArgumentOutOfRangeException("textRotation", "textRotation must be between -90 and 90 degrees.");
    range.Orientation = deg;
}
if (op.TryGetProperty("indent", out var ind) && ind.ValueKind == JsonValueKind.Number)
{
    int lvl = ind.GetInt32();
    if (lvl < 0 || lvl > 15)
        throw new ArgumentOutOfRangeException("indent", "indent must be between 0 and 15.");
    range.IndentLevel = lvl;
}
```

Explicit range validation matters: `Orientation = 120` throws a bare `COMException` whose message names neither the property nor the limit.

- [ ] **Step 6: Keep `bold`/`italic`/`numberFormat` exactly as they are** — they work, and a rewrite risks regressing the most-used path.

**Verification:** `MSBuild ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug` succeeds.

---

### Task 2: Borders

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Interfaces:**
- Produces: `private static void ApplyBorders(Excel.Range range, JsonElement borders)`.

- [ ] **Step 1: Shape**

```json
"borders": {
  "preset": "none" | "outline" | "all" | "thick-outline",
  "edges": ["top","bottom","left","right","insideHorizontal","insideVertical","diagonalDown","diagonalUp"],
  "style": "thin" | "medium" | "thick" | "double" | "dotted" | "dashed" | "none",
  "color": "#RRGGBB"
}
```

`preset` covers the overwhelmingly common requests in one field ("add borders" → `all`, "box this" → `outline`). `edges` is the escape hatch for precise control. When both are given, `preset` is applied first and `edges` refines it — state that ordering in the description rather than making them mutually exclusive.

- [ ] **Step 2: Implement** by mapping edge names to `XlBordersIndex` members and style names to `XlLineStyle` + `XlBorderWeight` pairs (`thin`→`xlContinuous`/`xlThin`, `medium`→`xlContinuous`/`xlMedium`, `thick`→`xlContinuous`/`xlThick`, `double`→`xlDouble`/`xlThick`, `dotted`→`xlDot`/`xlThin`, `dashed`→`xlDash`/`xlThin`, `none`→`xlLineStyleNone`). Two more small `Dictionary` statics, same pattern as Task 1.

- [ ] **Step 3:** `preset: "none"` clears all borders in the range — the only property in this whole plan that removes something, and it must be requested explicitly.

- [ ] **Step 4:** `insideHorizontal`/`insideVertical` on a single-cell range throw in Excel. Detect `range.Cells.Count == 1` and skip those two edges silently **with a note in the returned text** — this is the one place a silent skip is right, because "add all borders" to one cell has an obvious sane meaning, but the user should still be told.

- [ ] **Step 5:** Unknown edge or style name → specific error listing valid values.

**Verification:** build; `{preset:'all', style:'thin'}` over A1:C5 draws a full grid; `{preset:'outline', style:'thick', color:'#FF0000'}` draws only a thick red box.

---

### Task 3: Extend `read_formats` to match

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts`

**Problem:** without this, the model can set ten properties it cannot read back, so any read-modify-write cycle silently drops them.

- [ ] **Step 1:** Extend the per-cell report (`ExcelTools.cs:186-195`) with font name/size/color, strikethrough, both alignments, wrap, rotation, indent, and a compact border summary.
- [ ] **Step 2:** Preserve the "only explicitly-formatted cells" filter (`:193`) — but widen `hasDefaultFormat` to account for the new properties, or a cell that is *only* centered will still be filtered out as unformatted, which is precisely the bug this task exists to prevent.
- [ ] **Step 3: Watch the cell cap.** The cap is 200 cells (`:181`) and each cell now costs ~12 COM property reads instead of 4. Measure the wall-clock time for a 200-cell read after the change; if it exceeds ~2 seconds, either lower the cap or add a `properties?: string[]` parameter letting the caller ask for a subset. Record the measured number in a comment.
- [ ] **Step 4:** Keep the output one line per cell, `key=value` — it is compact, model-legible, and consistent with the current format. Do not switch to JSON.
- [ ] **Step 5:** Update the tool description (`ExcelAiAddIn/web-src/entry.ts:167`), which currently enumerates the four properties it reads.

**Verification:** format a cell with every new property, then `read_formats` it — every set value comes back.

---

### Task 4: Schema

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Update `format_range`'s entry in `propose_operations` — currently the one-liner at `ExcelAiAddIn/web-src/entry.ts:202`. If PP-5 has landed, edit `EXCEL_OPS` and let the generators produce both the schema and the prose.
- [ ] **Step 2:** Every closed set gets a real `enum`: `underline`, `horizontalAlignment`, `verticalAlignment`, `borders.preset`, `borders.style`, `borders.edges` items.
- [ ] **Step 3:** Document the numeric ranges (`textRotation` -90..90, `indent` 0..15) in the property descriptions, since JSON Schema `minimum`/`maximum` are not universally enforced by providers.
- [ ] **Step 4:** State in the description that omitted properties are left unchanged — the model should not need to resend the full property set to change one thing.

**Verification:** bundle rebuilds; a natural-language "center this and add thin borders" produces one correct `format_range` op.

---

### Task 5: Manual verification matrix

- [ ] Existing 4-property call (`bold`/`italic`/`numberFormat`/`fillColor`) behaves exactly as before.
- [ ] Font name, size, and color each apply independently.
- [ ] `underline: true` → single underline; `underline: 'double'` → double; `underline: 'nonsense'` → specific error, cell unchanged.
- [ ] `horizontalAlignment: 'center'` + `verticalAlignment: 'top'` both apply.
- [ ] `wrapText: true` on a long string wraps and grows the row.
- [ ] `textRotation: 45` rotates; `textRotation: 120` → specific error.
- [ ] `indent: 3` indents; `indent: 20` → specific error.
- [ ] `borders: {preset:'all', style:'thin'}` over a multi-cell range → full grid.
- [ ] `borders: {preset:'outline', style:'thick', color:'#FF0000'}` → thick red box only.
- [ ] `borders: {preset:'none'}` → all borders cleared.
- [ ] `borders: {preset:'all'}` on a single cell → outline drawn, result text notes the skipped interior edges.
- [ ] `read_formats` round-trips every property set above.
- [ ] Natural language: "make the header row bold, centered, white on dark blue, with a thick bottom border" produces one op that does all of it.
