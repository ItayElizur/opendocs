# PP-16: Document Excel's Shape Vocabulary — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-16 (P2).

**Goal:** The 26 shape presets `add_shape` already supports become discoverable as a real `enum`, and an unrecognized name errors instead of silently becoming a rectangle.

**Architecture:** `ShapeTypeMap` (`ExcelAiAddIn/ExcelTools.cs:27-56`) holds 26 entries — `rect`, `roundRect`, `ellipse`, `triangle`, `rtTriangle`, `parallelogram`, `trapezoid`, `diamond`, `pentagon`, `hexagon`, `octagon`, `pie`, `chord`, `donut`, `foldedCorner`, `heart`, `lightningBolt`, `sun`, `moon`, `cloud`, `arc`, `star5`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow` — plus `"textbox"` handled separately in `AddShapeExcel` (`:754-757`). The schema mentions none of them (`ExcelAiAddIn/web-src/entry.ts:213` reads simply `"add_shape" (sheet?, shapeType, anchorCell, fillColor?, text?)`), and an unrecognized name falls through `TryGetValue` to `msoShapeRectangle` (`:761`).

So a request for "a star" or "an arrow pointing right" has roughly a 1-in-27 chance of the model guessing the exact key, and otherwise silently produces a rectangle. Both halves — the missing enum and the silent fallback — are one-session fixes.

This is the smallest item in the Excel group and a good candidate to land early: it is a pure schema-plus-error change with no behavioral risk to working calls.

**Tech Stack:** C# 7.3 / .NET Framework 4.8; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Do not rename any existing key. The names are OOXML preset-geometry names, they match genoffice's vocabulary, and renaming would break saved conversations for no gain.
- Do not remove the documented PIA omission. `ExcelTools.cs:22-26` records that `msoShapePlus`/`msoShapeMathPlus` do not exist in this project's referenced PIA (confirmed by a CS0117 compile failure) and were deliberately dropped. Leave that comment intact and do not re-attempt them.
- No automated tests for COM executor methods (project convention). Verification is build + Task 4's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.
- If PP-5 has landed, add the enum through its `EXCEL_OPS` table.

---

### Task 1: Export the vocabulary to the schema

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Declare the enum as a named constant so it is greppable and reusable:

```ts
// Mirrors ShapeTypeMap in ExcelAiAddIn/ExcelTools.cs:27-56 exactly, plus the
// separately-handled 'textbox'. Edit both together.
const EXCEL_SHAPE_TYPES = [
  'textbox',
  'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'parallelogram', 'trapezoid',
  'diamond', 'pentagon', 'hexagon', 'octagon', 'pie', 'chord', 'donut', 'foldedCorner',
  'heart', 'lightningBolt', 'sun', 'moon', 'cloud', 'arc', 'star5',
  'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
] as const
```

- [ ] **Step 2:** Verify the list against the source map key by key — 26 entries plus `textbox` = 27. A typo here reintroduces the exact bug this plan fixes, since a schema-valid-but-map-missing name would still fall back to rectangle.
- [ ] **Step 3:** Wire it into `add_shape`'s `shapeType`. Since `propose_operations` currently has no per-op structure, either land PP-5 first, or (if going alone) at minimum enumerate the names inline in the description string so they reach the model somehow:

```
"add_shape" (sheet?, shapeType: one of textbox|rect|roundRect|ellipse|triangle|rtTriangle|parallelogram|trapezoid|diamond|pentagon|hexagon|octagon|pie|chord|donut|foldedCorner|heart|lightningBolt|sun|moon|cloud|arc|star5|rightArrow|leftArrow|upArrow|downArrow, anchorCell, fillColor?, text?)
```

- [ ] **Step 4:** Check whether `edit_shape` (`ExcelAiAddIn/web-src/entry.ts:214`) accepts a shape type too; if it does, give it the same enum.

**Verification:** bundle rebuilds; "add a star" produces `shapeType: 'star5'`.

---

### Task 2: Error instead of silent rectangle

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Replace the fallback at `:761`:

```csharp
Microsoft.Office.Core.MsoAutoShapeType msoType;
if (!ShapeTypeMap.TryGetValue(shapeType, out msoType))
    throw new ArgumentException("add_shape: unknown shapeType '" + shapeType + "'. Valid: textbox, " +
                                string.Join(", ", ShapeTypeMap.Keys) + ".");
shape = Sheet(op).Shapes.AddShape(msoType, left, top, width, height);
```

Building the message from `ShapeTypeMap.Keys` means it can never drift from the actual map.

- [ ] **Step 2:** Add a comment above `ShapeTypeMap` pointing at `EXCEL_SHAPE_TYPES` in `entry.ts` and stating the two must be edited together.
- [ ] **Step 3:** Case sensitivity. The map is a plain `Dictionary<string, ...>` — ordinal, case-sensitive — so `"Star5"` or `"RECT"` currently becomes a rectangle and will now error. Make it `StringComparer.OrdinalIgnoreCase`: it removes an entire class of near-miss failures at zero cost, and the enum still communicates the canonical spelling.

**Verification:** build; `shapeType: 'nonsense'` errors with the full valid list; `shapeType: 'STAR5'` works.

---

### Task 3: Return the shape's name

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

**Context:** `edit_shape` and `delete_visual` address shapes by `visualId` = the shape's name, but `AddShapeExcel` returns `void` and the batch reports a bare `add_shape: ok` (`:533`). The model must guess `"Rectangle 1"` to edit what it just created.

- [ ] **Step 1:** Return `shape.Name` from `AddShapeExcel` and surface it in the batch result line.
- [ ] **Step 2:** Optionally accept a `name` parameter and set `shape.Name`, so the model can pick a stable id up front.
- [ ] **Step 3:** This duplicates PP-15 Task 4 Step 4 — if that has landed, verify rather than redo.

**Verification:** add a shape, then edit it by the returned name in the next operation of the same batch.

---

### Task 4: Manual verification matrix

- [ ] Each of the 26 presets plus `textbox` produces the correct shape in real Excel. Work through the list; this is the only way to catch a mismapped `MsoAutoShapeType` member.
- [ ] `shapeType: 'nonsense'` → specific error listing every valid name; no shape created.
- [ ] `shapeType: 'RECT'` (wrong case) → works, produces a rectangle.
- [ ] `add_shape` returns a usable name; `edit_shape` with it succeeds.
- [ ] Natural language: "add a star", "add an arrow pointing right", "add a heart" each produce the right shape on the first attempt — the concrete failure the source item describes.
- [ ] `fillColor` and `text` still apply as before on every shape type.
