# PP-20: PowerPoint `set_element_style` and `add_shape` Breadth — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-20 (P2).

**Goal:** `set_element_style` reaches the field coverage Word's `updateTextStyle` already has, and `add_shape` uses Excel's existing 26-type shape map instead of its own three-type ternary.

**Architecture:** Two capability gaps where this repo already contains the answer.

1. **`set_element_style`** (`PowerPointAiAddIn/PowerPointTools.cs:158-173`) handles `bold`, `italic`, `fontSize`, `color`. Word's `UpdateTextStyle` (`WordAiAddIn/WordTools.cs:379-416`) handles nine fields, and PP-12 adds `highlight` for ten. PowerPoint's `TextRange.Font` exposes direct equivalents for almost all of them — `Name`, `Underline`, `Shadow`, `Subscript`/`Superscript`, plus `TextRange.ParagraphFormat.Alignment` for alignment. Unlike Word's version, this is at least *honest*: the schema (`PowerPointAiAddIn/web-src/entry.ts:168-183`) advertises exactly the four fields it implements, so nothing silently no-ops. That makes this a pure capability gap, and it means the fix must not introduce the silent-no-op pattern while closing it.
2. **`add_shape`** (`PowerPointTools.cs:198-213`) maps three names via a ternary: `oval`, `roundRect`, and anything-else → rectangle. Its schema does declare `enum: ['rectangle','oval','roundRect']` (`entry.ts:217-233`) — the pattern PP-22 wants everywhere — so the enum and handler agree. The gap is breadth: `ExcelTools.cs:27-56` already holds a verified 26-entry `MsoAutoShapeType` map in this same repo, built against this same PIA, and it was not reused.

Note the naming discrepancy to resolve deliberately: PowerPoint uses `rectangle`/`oval`, Excel's map uses `rect`/`ellipse`. Task 2 Step 3 handles it.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.PowerPoint` + `Microsoft.Office.Core`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **Do not break the existing enum contract.** `rectangle`/`oval`/`roundRect` must keep working; any saved conversation and any prompt the model has learned uses them.
- Every added field is optional and additive — an absent field leaves the current value alone.
- No silent fallbacks: an unrecognized value errors with the valid list. This applies to the new `add_shape` names too, replacing the current anything-else-becomes-rectangle ternary.
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: Widen `set_element_style`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Add the font fields**

```csharp
if (input.TryGetProperty("fontName", out var fontName) && fontName.ValueKind == JsonValueKind.String)
    range.Font.Name = fontName.GetString();
if (input.TryGetProperty("underline", out var underline))
    range.Font.Underline = underline.ValueKind == JsonValueKind.True
        ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
if (input.TryGetProperty("shadow", out var shadow))
    range.Font.Shadow = shadow.ValueKind == JsonValueKind.True
        ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
```

Follow the existing `MsoTriState` idiom used for `bold`/`italic` at `:161-162` rather than introducing a second boolean convention.

- [ ] **Step 2: Strikethrough.** PowerPoint's `Font` (the older `TextRange.Font`) may not expose `Strikethrough` in this PIA; `TextRange2.Font.Strikethrough` (via `TextFrame2`) does. Try `TextFrame2.TextRange.Font.Strikethrough` and, if it does not compile or does not work against this PIA, **omit the field from both handler and schema** and record why in a comment — following the precedent at `ExcelTools.cs:22-26`. Do not ship a schema field the handler ignores; that is the PP-12 defect being imported.
- [ ] **Step 3: Alignment**

```csharp
// alignment: left|center|right|justify
if (input.TryGetProperty("alignment", out var align) && align.ValueKind == JsonValueKind.String) { ... }
```
Map to `PowerPoint.PpParagraphAlignment` (`ppAlignLeft`/`Center`/`Right`/`Justify`) on `range.ParagraphFormat.Alignment`. Unknown value → specific error listing valid ones. Use a small `Dictionary` static rather than a ternary chain, matching this file's `PptChartTypeMap`/`SmartArtLayoutNames` pattern.

- [ ] **Step 4: Baseline offset.** Word supports `SUPERSCRIPT`/`SUBSCRIPT`/`NONE` (`WordTools.cs:407-412`). PowerPoint's `Font` exposes `Superscript`/`Subscript` as `MsoTriState`. Implement the same three-value field with the same value names, so the two apps' vocabularies match.
- [ ] **Step 5: Reuse the hex parser.** The color branch inlines its own hex parsing (`:164-171`). Extract a `HexToOle` private static in this file (or reuse one if it already exists elsewhere in it) and use it for every color property, matching `ExcelTools.cs:775-782`.
- [ ] **Step 6: Scope the change.** `set_element_style` applies to the shape's entire `TextFrame.TextRange`. Consider optional `startChar`/`length` for a sub-range — but only if the manual matrix shows a real need; whole-shape styling is the common case and adding character addressing multiplies the failure modes.
- [ ] **Step 7: Schema** — add every implemented field with real `enum`s for `alignment` and `baselineOffset`, and update the description, which currently just says "Changes text formatting of one shape without changing its text."
- [ ] **Step 8: Report what applied.** The handler returns a flat `"Style updated."` (`:172`). Return the list of properties actually applied, so a typo'd-and-rejected field is distinguishable in the transcript from a successful one.

**Verification:** `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug`; each new field visibly changes the shape's text.

---

### Task 2: `add_shape` uses the full preset vocabulary

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Port the map.** Copy `ShapeTypeMap` from `ExcelTools.cs:27-56` into `PowerPointTools.cs` as a `private static readonly Dictionary<string, MsoAutoShapeType>` with `StringComparer.OrdinalIgnoreCase`. Copy the PIA-omission comment at `ExcelTools.cs:22-26` with it — the same two members are missing from the same PIA, and without the comment someone will re-add them and hit the same CS0117.
- [ ] **Step 2: Do not extract to `OfficeAi.Shared`.** It would need a `Microsoft.Office.Core` reference in a project that currently holds only app-agnostic plumbing (`ChatStore`, `ToolProtocol`, `WebViewBridgeHost`). Two copies with cross-referencing comments is the cheaper trade at this size; revisit if Word gains a third (PP-11 does not need one).
- [ ] **Step 3: Reconcile the names.** PowerPoint's existing enum uses `rectangle`/`oval`; Excel's map uses `rect`/`ellipse`. Register **both spellings as aliases** pointing at the same `MsoAutoShapeType`, keep both in the schema enum, and mark the Excel spellings as canonical in the description. This keeps every existing PowerPoint call working, gives the model one vocabulary across both apps, and costs two extra dictionary entries.
- [ ] **Step 4: Replace the ternary** at `:201-204` with a `TryGetValue` + throw listing the valid names, built from the map's keys so it cannot drift.
- [ ] **Step 5: Schema** — replace the three-value enum with the full list and update the description ("Creates a shape (rectangle/oval/roundRect) with optional text"), which will otherwise contradict the enum.
- [ ] **Step 6: `add_shape` has no fill/line parameters** — `set_element_fill` and `set_element_stroke` exist as separate tools (`entry.ts:252-272`). Leave that split alone, but say so in `add_shape`'s description so the model chains the calls rather than sending a `fillColor` that is silently dropped.

**Verification:** build; `shapeType: 'star5'` and `shapeType: 'rightArrow'` produce the right shapes; `'rectangle'` still works; an unknown name errors specifically.

---

### Task 3: Cross-app consistency check

**Files:** none modified (audit), then whichever needs correcting

- [ ] **Step 1:** After Tasks 1-2, compare the text-style field vocabulary across `WordTools.UpdateTextStyle`, `PowerPointTools.SetElementStyle`, and `ExcelTools.FormatRange` (as PP-13 leaves it). They should use the same names for the same concepts (`fontName` vs `font`, `fontSize` vs `sizeHalfPoints`, `color` vs `fontColor`).
- [ ] **Step 2:** Note every mismatch. **Recommendation: do not rename anything to fix them** — renaming breaks saved conversations across three apps for a cosmetic gain. Record the mapping in `docs/ai-tool-surface.md` instead, so the inconsistency is at least documented in one place.
- [ ] **Step 3:** One exception worth flagging for a decision: Word's `sizeHalfPoints` (halved at `WordTools.cs:400`) is a genoffice/OOXML-ism that no other app in this repo uses, and a model that sends `24` meaning 24pt gets 12pt silently. It is the one naming difference that produces a *wrong result* rather than a rejected call. File it as its own item rather than folding it in here.

---

### Task 4: Documentation

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1:** Update the PowerPoint section with the new `set_element_style` field list and the widened `add_shape` vocabulary.
- [ ] **Step 2:** Record the alias decision from Task 2 Step 3 and which spellings are canonical.
- [ ] **Step 3:** Record any field omitted for PIA reasons (Task 1 Step 2's strikethrough, if it lands that way) so it is a documented limitation rather than an apparent oversight.

---

### Task 5: Manual verification matrix

- [ ] Existing `set_element_style` call with `bold`/`italic`/`fontSize`/`color` → unchanged behavior.
- [ ] `fontName: 'Georgia'` → font changes.
- [ ] `underline: true` / `false` → toggles.
- [ ] `shadow: true` → text shadow appears.
- [ ] `alignment` with each of the four values → visible change; an unknown value → specific error.
- [ ] `baselineOffset: 'SUPERSCRIPT'` / `'SUBSCRIPT'` / `'NONE'` → correct rendering.
- [ ] Strikethrough — either works, or is absent from the schema entirely (never present-but-ignored).
- [ ] `add_shape` with `rectangle`, `oval`, `roundRect` → unchanged behavior.
- [ ] `add_shape` with `rect`, `ellipse` (Excel spellings) → same shapes.
- [ ] `add_shape` with `star5`, `rightArrow`, `heart`, `cloud` → correct shapes.
- [ ] `add_shape` with an unknown name → specific error listing valid names; no shape created.
- [ ] Natural language: "make this text underlined and centered in Georgia" → one correct call; "add a star" → a star, not a rectangle.
