# PP-23: Word Table and SmartArt Support (+ PowerPoint read-visibility fix) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source:** user-reported gap, 2026-08-24 ("i also noticed there isn't a create/edit/read table and create/edit/read smartart"). Not from the original audit — confirmed absent by direct source read (`grep -ni "table\|smartart" WordAiAddIn/WordTools.cs WordAiAddIn/web-src/entry.ts` returns nothing relevant) after every other Word plan (PP-9 through PP-12) had already landed.

**Goal:** Word gets create/edit/read for native tables and native SmartArt diagrams — six new tools total (`add_table`, `edit_table`, `read_table`, `add_smartart`, `edit_smartart`, `read_smartart`), closing the last major content-type gap in Word's tool surface. A seventh, smaller fix (Task 8) closes a related PowerPoint gap surfaced while building this plan: PowerPoint's own `read_slide`/`get_deck_context` are blind to table and SmartArt shape content today.

## What's already proven, and what needs live verification

PowerPoint already has working table (`AddTable`/`EditTableCell`/`EditTableStructure`/`EditTableStyle`, `PowerPointAiAddIn/PowerPointTools.cs:524-664`) and SmartArt-create (`AddSmartArt`/`SmartArtLayoutNames`/`ResolveSmartArtLayout`, `:910-978`) code that **compiles against statically-typed `PowerPoint.Table`/`PowerPoint.Shape`** (tables) and is proven to run (SmartArt's node-writing pattern, `smartArt.Nodes.Add()` / `node.TextFrame2.TextRange.Text = ...`, is exercised by every `add_smartart` call). This plan ports both patterns to Word rather than inventing new ones — but **Word's object model is genuinely different from PowerPoint's in the table case**, not just a renamed copy:

- PowerPoint's table is a floating `Shape` with a nested `.Table`; a cell's text lives at `table.Cell(r,c).Shape.TextFrame.TextRange.Text` (a cell *contains* a shape).
- Word's table is a native, flow-integrated part of the document; `Document.Tables` is the collection, and a cell's text lives at `table.Cell(r,c).Range.Text` directly (a cell *is* a range — no nested shape).

This is the standard, well-documented Word Interop shape (`Document.Tables.Add(Range, rows, cols)`, `Table.Cell(row, col)`, `Table.Rows.Add`/`.Columns.Add`, `Table.Rows[n].Delete()`) — high confidence on the core API. **Lower confidence, needs live verification during implementation:** the exact optional-parameter shape of `Tables.Add` in this project's PIA (older Word Interop signatures carry trailing `ref object` optional params that may need `Type.Missing` rather than C#'s named-optional-argument sugar), and the exact property names for style toggles (`Table.ApplyStyleHeadingRows`/`Table.ApplyStyleBandedRows` are the standard names, analogous to PowerPoint's `FirstRow`/`HorizBanding`, but not cross-checked against this PIA).

SmartArt is the OFFICE-shared object model (`Microsoft.Office.Interop.SmartArt`), not PowerPoint-specific — `Application.SmartArtLayouts` and `Shape.SmartArt.Nodes` are exposed identically regardless of which host app the `Shapes` collection belongs to. `Word.Document.Shapes.AddSmartArt(layout, left, top, width, height)` is documented as available since Word 2010, mirroring PowerPoint's own call at `PowerPointTools.cs:967` almost exactly — **medium-high confidence**, not verified live.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Word` (statically typed for tables; `dynamic` for SmartArt, matching this file's existing chart-code convention and the reason stated for it at `WordTools.cs:483-486`).

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Six **standalone top-level tools**, not `apply_commands` kinds — matches how `edit_chart`/`read_chart` are standalone, and how PowerPoint's own table/SmartArt tools are standalone rather than routed through a batch gateway (Word has no such gateway for anything but formatting/structure edits; tables and SmartArt are a different content type, same tier as charts and images).
- Every new tool follows this file's existing addressing conventions: 0-based indices everywhere at the tool boundary, converted to Word's 1-based COM collections at the point of use, never left ambiguous.
- Read tools (`read_table`, `read_smartart`) go in `AlwaysAllowedTools` and `readOnlyTools`, matching `read_chart`'s precedent (`WordTools.cs:34-37`, `entry.ts:406`).
- No silent fallbacks: an out-of-range index, an unknown style name, or a malformed request throws a specific error naming the problem and (where the set is closed) listing valid values — the governing rule this whole project has converged on.
- No automated tests for COM executor methods (project convention). Verification is build + the manual matrix in Task 7, plus the same honest "dynamic/unverified" risk-flagging PP-9's chart work and PP-11's image work already established for this file.
- Rebuild the bundle and re-run MSBuild after every `entry.ts` change (4-alias esbuild command in `docs/superpowers/plans/STATUS.md`).
- Add every new tool's `toolDisplay` entry (English + Hebrew) in the same edit that adds its schema — FT-1's settings screen warns loudly on a missing entry; don't let one drift.
- Update `WordTools.cs`'s `PP-5` schema comment discipline: if `2026-08-23-pp05-gateway-tool-schemas.md`'s structural-schema pattern is ever extended to standalone tools (it currently isn't — `edit_chart`/`read_chart`/`add_image` are all plain `inputSchema` objects, not `oneOf` branches), these six tools should follow the same plain shape, not invent a new one.

---

### Task 1: Table addressing and read helper

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Produces: `private static Word.Table ResolveTable(JsonElement input)` — consumed by Tasks 2 and 3.

- [x] **Step 1: Table resolution by index**

```csharp
// 0-based at the tool boundary, matching every other index in this file;
// Document.Tables is 1-based in COM.
private static Word.Table ResolveTable(JsonElement input)
{
    int tableIndex = input.GetProperty("tableIndex").GetInt32();
    Word.Tables tables = ActiveDoc.Tables;
    if (tableIndex < 0 || tableIndex >= tables.Count)
        throw new ArgumentOutOfRangeException("tableIndex",
            "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
    return tables[tableIndex + 1];
}
```

`Document.Tables` addresses tables in the order they appear in the document (document order), which is the natural, predictable convention — unlike charts, there's no inline-vs-floating split to worry about, since Word tables are always flow content.

- [x] **Step 2: Verify `Tables.Add`'s real signature in this PIA** before Task 2 depends on it. Write a throwaway one-line test call in a scratch method (or just attempt the real implementation directly and let the compiler report the actual required argument shape) — the standard signature is `Tables.Add(Range Range, int NumRows, int NumColumns, ref object DefaultTableBehavior, ref object AutoFitBehavior)`, where the two trailing `ref object` parameters are typically satisfiable via `Type.Missing` (classic COM-interop optional-parameter idiom) or, if this PIA marks them `[Optional]`, via C#'s ordinary optional-argument omission. Record which one this project's PIA actually needs.

**Verification:** compiles; no runtime check possible until Task 2 has a caller.

---

### Task 2: `add_table`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [x] **Step 1: Handler**

```csharp
private static ToolResult AddTable(JsonElement input)
{
    int rows = input.GetProperty("rows").GetInt32();
    int cols = input.GetProperty("cols").GetInt32();
    if (rows < 1 || cols < 1)
        throw new ArgumentException("add_table: rows and cols must each be at least 1.");

    int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
        ? abEl.GetInt32() : (int?)null;
    Word.Range at = afterBlockIndex.HasValue ? RangeAfterBlock(afterBlockIndex.Value) : EndOfDocumentRange();

    Word.Table table = ActiveDoc.Tables.Add(at, rows, cols /*, verify trailing optional params per Task 1 Step 2 */);

    if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
    {
        int r = 0;
        foreach (JsonElement rowEl in cells.EnumerateArray())
        {
            if (r >= rows) break; // ignore extra rows beyond the declared size rather than throwing mid-write
            int c = 0;
            foreach (JsonElement cellEl in rowEl.EnumerateArray())
            {
                if (c >= cols) break;
                table.Cell(r + 1, c + 1).Range.Text = cellEl.GetString();
                c++;
            }
            r++;
        }
    }

    int newIndex = ActiveDoc.Tables.Count - 1; // Tables.Add appends; stable immediately after the call, matching PP-9/PP-22's "return the new index" precedent
    return new ToolResult
    {
        Output = "Table added at index " + newIndex + " (" + rows + " rows x " + cols + " cols).",
        Mutated = true,
        Summary = "add_table",
    };
}
```

`EndOfDocumentRange()` — a one-line helper (`ActiveDoc.Content` collapsed to the end) if one doesn't already exist under another name in this file; check before adding a duplicate.

- [x] **Step 2: `afterBlockIndex` semantics.** Reuse `RangeAfterBlock` from PP-10 exactly as PP-9's chart and PP-11's image insertion already do — `-1` = start of document, omitted = end of document. Do not invent a third convention.

- [x] **Step 3: Schema**

```ts
{
  name: 'add_table',
  description:
    'Adds a native Word table, optionally pre-filled with cell text (row-major array of arrays; extra cells beyond rows/cols are ignored). ' +
    'afterBlockIndex is the 0-based paragraph index to insert after (-1 = start of document; omit = end of document).',
  inputSchema: {
    type: 'object',
    properties: {
      rows: { type: 'number' },
      cols: { type: 'number' },
      cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
      afterBlockIndex: { type: 'number' },
    },
    required: ['rows', 'cols'],
  },
}
```

**Verification:** builds; a real Word test creates a 3x3 table with pre-filled text at a specific paragraph position.

---

### Task 3: `edit_table` and `read_table`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**One `edit_table` tool, not three separate ones** (unlike PowerPoint's `edit_table_cell`/`edit_table_structure`/`edit_table_style` split) — Word's table operations are few enough, and similar enough in shape, to keep as one tool with a `kind` field, closer to `apply_commands`'s per-kind dispatch than PowerPoint's three-tools split. Either shape works; this plan picks one tool to keep Word's already-large tool list (10, after PP-9's `read_chart`) from growing by three instead of one. If this later needs to be a gateway kind of its own, that is a straightforward follow-up, not a redesign.

- [x] **Step 1: `edit_table` dispatch**

```csharp
private static ToolResult EditTable(JsonElement input)
{
    Word.Table table = ResolveTable(input);
    string kind = input.GetProperty("kind").GetString();
    switch (kind)
    {
        case "set_cell":
        {
            int row = input.GetProperty("row").GetInt32();
            int col = input.GetProperty("col").GetInt32();
            if (row < 0 || row >= table.Rows.Count)
                throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (table.Rows.Count - 1) + ".");
            if (col < 0 || col >= table.Columns.Count)
                throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (table.Columns.Count - 1) + ".");
            table.Cell(row + 1, col + 1).Range.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Cell [" + row + "," + col + "] updated.", Mutated = true, Summary = "edit_table" };
        }
        case "insert_row":
        case "delete_row":
        case "insert_col":
        case "delete_col":
        {
            // Same index-always-existing, before/after-picks-side convention
            // as PowerPoint's edit_table_structure (PowerPointTools.cs:568-608) -
            // ported directly, including its out-of-range validation.
            int index = input.GetProperty("index").GetInt32();
            bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
            if (kind == "insert_row" || kind == "delete_row")
            {
                if (index < 0 || index >= table.Rows.Count)
                    throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for " + kind + ".");
                if (kind == "insert_row") table.Rows.Add(table.Rows[before ? index + 1 : Math.Min(index + 2, table.Rows.Count + 1)]);
                else table.Rows[index + 1].Delete();
            }
            else
            {
                if (index < 0 || index >= table.Columns.Count)
                    throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for " + kind + ".");
                if (kind == "insert_col") table.Columns.Add(table.Columns[before ? index + 1 : Math.Min(index + 2, table.Columns.Count + 1)]);
                else table.Columns[index + 1].Delete();
            }
            return new ToolResult
            {
                Output = kind + " applied at index " + index + ". Row/column indices after this point have shifted - re-read the table before another structural edit in the same run.",
                Mutated = true,
                Summary = "edit_table",
            };
        }
        case "set_style":
        {
            if (input.TryGetProperty("styleName", out var styleEl) && styleEl.ValueKind == JsonValueKind.String)
            {
                try { table.Style = styleEl.GetString(); }
                catch (Exception ex) { throw new ArgumentException("edit_table: '" + styleEl.GetString() + "' is not a valid table style name in this document/template. " + ex.Message); }
            }
            if (input.TryGetProperty("headerRow", out var hdr))
                table.ApplyStyleHeadingRows = hdr.ValueKind == JsonValueKind.True; // verify exact property name, Task 1 Step 2's discipline applies here too
            if (input.TryGetProperty("bandedRows", out var band))
                table.ApplyStyleBandedRows = band.ValueKind == JsonValueKind.True;
            return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table" };
        }
        default:
            throw new ArgumentException("edit_table: unknown kind '" + kind + "'. Valid: set_cell, insert_row, delete_row, insert_col, delete_col, set_style.");
    }
}
```

`Word.Table.Style` assigned by name is documented to throw if the name isn't a recognized/available style (built-in Word table styles like `"Table Grid"`, `"Light List Accent 1"`, etc., or a custom style already in the document) — the `try/catch` turns that into a specific, actionable error rather than a raw COM exception, following this file's established pattern (e.g. PP-9's `chartIndex` errors, PP-17's defined-name errors elsewhere in this repo).

Note the unknown-`kind` case is a `default: throw`, not a silent no-op-with-`ok` — the exact defect class PP-12/PP-14/PP-22 all fixed elsewhere in this codebase; don't reintroduce it here.

- [x] **Step 2: `read_table`**

```csharp
private static ToolResult ReadTable(JsonElement input)
{
    Word.Tables tables = ActiveDoc.Tables;
    if (tables.Count == 0)
        return new ToolResult { Output = "No tables in this document.", Summary = "read_table" };

    int tableIndex = input.TryGetProperty("tableIndex", out var ti) && ti.ValueKind == JsonValueKind.Number ? ti.GetInt32() : 0;
    if (tableIndex < 0 || tableIndex >= tables.Count)
        throw new ArgumentOutOfRangeException("tableIndex", "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
    Word.Table table = tables[tableIndex + 1];

    var sb = new System.Text.StringBuilder();
    sb.AppendLine("Table " + tableIndex + " of " + tables.Count + " (" + table.Rows.Count + " rows x " + table.Columns.Count + " cols):");
    for (int r = 0; r < table.Rows.Count; r++)
    {
        var cellsOut = new List<string>();
        for (int c = 0; c < table.Columns.Count; c++)
        {
            // A merged cell can make Cell(r,c) throw for the cells it no
            // longer owns - report a placeholder rather than failing the
            // whole read over one merged region.
            try { cellsOut.Add(table.Cell(r + 1, c + 1).Range.Text.TrimEnd('\r', '\a')); }
            catch { cellsOut.Add("(merged)"); }
        }
        sb.AppendLine("[" + r + "] " + string.Join(" | ", cellsOut));
    }
    return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_table" };
}
```

The `(merged)` fallback is a real, necessary guard: Word's `Table.Cell(r,c)` throws `WdCannotAccessIndividualCellsException`-shaped COM errors for a cell inside a merged region that isn't the merge's top-left anchor — a table with any merged cells would otherwise fail this tool entirely rather than degrading gracefully.

- [x] **Step 3: Register both in `Execute`'s switch and `read_table` in `AlwaysAllowedTools`.**

- [x] **Step 4: Schemas**

```ts
{
  name: 'edit_table',
  description:
    'Edits an existing table. kind: "set_cell" (row,col,text), "insert_row"/"delete_row"/"insert_col"/"delete_col" (index,before?), "set_style" (styleName?,headerRow?,bandedRows?). ' +
    'tableIndex addresses the table (0-based, document order); omit to target the first table. ' +
    'Structural edits shift later indices - re-read the table (read_table) before a second structural edit in the same run.',
  inputSchema: {
    type: 'object',
    properties: {
      tableIndex: { type: 'number' },
      kind: { type: 'string', enum: ['set_cell', 'insert_row', 'delete_row', 'insert_col', 'delete_col', 'set_style'] },
      row: { type: 'number' }, col: { type: 'number' }, text: { type: 'string' },
      index: { type: 'number' }, before: { type: 'boolean' },
      styleName: { type: 'string' }, headerRow: { type: 'boolean' }, bandedRows: { type: 'boolean' },
    },
    required: ['kind'],
  },
},
{
  name: 'read_table',
  description: 'Reads an existing table\'s cell contents, one row per line. tableIndex addresses the table (0-based, document order); omit to read the first table.',
  inputSchema: { type: 'object', properties: { tableIndex: { type: 'number' } }, required: [] },
},
```

**Verification:** builds; a real Word session round-trips `add_table` → `read_table` → `edit_table(set_cell)` → `read_table` confirming the change, plus one structural edit (`insert_row`) and one style change.

---

### Task 4: SmartArt layout map and `add_smartart`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [x] **Step 1: Port the layout map and resolver from PowerPoint verbatim**

Copy `SmartArtLayoutNames` and `ResolveSmartArtLayout` from `PowerPointAiAddIn/PowerPointTools.cs:914-954` into `WordTools.cs`, unchanged in content — same seven keys, same display names, same two-distinct-errors design (unknown key vs. valid-key-but-not-in-this-install's-gallery, the latter naming the possible non-English-install cause). `Globals.ThisAddIn.Application.SmartArtLayouts` resolves against `WordAiAddIn`'s own `Globals.ThisAddIn` this time — same property, different host app, same Office-shared object model.

Carry the existing comment about the live cross-check never having been performed (`PowerPointTools.cs:910-913`) — it applies here too, and doubly so since it has now never been verified against *either* host app.

- [x] **Step 2: `add_smartart`**

```csharp
private static ToolResult AddSmartArt(JsonElement input)
{
    string layoutKey = input.GetProperty("layout").GetString();
    dynamic layout = ResolveSmartArtLayout(layoutKey);

    int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
        ? abEl.GetInt32() : (int?)null;
    float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
    float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

    dynamic doc = ActiveDoc;
    dynamic shape;
    if (afterBlockIndex.HasValue)
    {
        // Mirrors PP-9's anchored-chart-creation path exactly, including its
        // caveat: whether Shapes.AddSmartArt truly accepts a named Anchor
        // parameter in this PIA is UNVERIFIED - flag this specific path as
        // elevated risk in the verification file, same as PP-9's did.
        Word.Range at = RangeAfterBlock(afterBlockIndex.Value);
        dynamic floatingAtAnchor = doc.Shapes.AddSmartArt(layout, 0, 0, width, height, Anchor: at);
        shape = floatingAtAnchor.ConvertToInlineShape();
    }
    else
    {
        float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
        float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
        shape = doc.Shapes.AddSmartArt(layout, left, top, width, height);
    }

    dynamic smartArt = shape.SmartArt;
    foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
    {
        dynamic node = smartArt.Nodes.Add();
        node.TextFrame2.TextRange.Text = item.GetString();
    }
    return new ToolResult { Output = "SmartArt added (" + input.GetProperty("items").GetArrayLength() + " node(s)).", Mutated = true, Summary = "add_smartart" };
}
```

Node-writing (`smartArt.Nodes.Add()` / `node.TextFrame2.TextRange.Text = ...`) is copied verbatim from PowerPoint's proven, already-running code (`PowerPointTools.cs:974-975`) — this part carries PowerPoint's confidence level, not PP-9's chart-anchor uncertainty. Only the `Anchor:`-parameter path is the new, unverified piece, and it is *only* reached when `afterBlockIndex` is given — omitting it (floating placement, matching the plan's non-anchored fallback pattern) is safe and already-proven-shaped.

- [x] **Step 3: Register in `Execute`'s switch.**

- [x] **Step 4: Schema**

```ts
{
  name: 'add_smartart',
  description:
    'Adds a shape-composed SmartArt diagram. layout: "list"|"process"|"cycle"|"hierarchy"|"pyramid"|"matrix"|"venn". items are flat node texts, one per top-level node. ' +
    'afterBlockIndex (0-based paragraph index, -1 = start, omit = end of document) inserts inline; omitting both x/y and afterBlockIndex places a floating shape at a default position.',
  inputSchema: {
    type: 'object',
    properties: {
      layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'] },
      items: { type: 'array', items: { type: 'string' } },
      x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
      afterBlockIndex: { type: 'number' },
    },
    required: ['layout', 'items'],
  },
}
```

**Verification:** builds; real-Word test creates each of the 7 layouts with a small flat item list, both floating and (separately) anchored via `afterBlockIndex` — the anchored path is the one to check first and most carefully, per the risk note above.

---

### Task 5: `edit_smartart` and `read_smartart`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**Neither of these exists in PowerPoint today** — this task is genuinely new work, not a port. Design conservatively: flat node list only (matching `add_smartart`'s own flat-list scope, and genoffice's own SmartArt tool, which — per the earlier genoffice comparison audit — only ever produces a flat item list too, no nested tree-building). Do not attempt nested/hierarchical node editing; that is a real SmartArt capability but a materially larger scope than this plan's six tools.

- [x] **Step 1: SmartArt resolution helper**

```csharp
// SmartArt shapes are not chart shapes and are not tables - a small,
// separate list-and-resolve helper, mirroring ListChartShapes'/ResolveTable's
// shape but for shape.HasSmartArt instead of shape.HasChart.
private static List<dynamic> ListSmartArtShapes(dynamic doc)
{
    var shapes = new List<dynamic>();
    foreach (dynamic shp in doc.InlineShapes)
    {
        try { if ((bool)shp.HasSmartArt) shapes.Add(shp); } catch { }
    }
    foreach (dynamic shp in doc.Shapes)
    {
        try { if ((bool)shp.HasSmartArt) shapes.Add(shp); } catch { }
    }
    return shapes;
}
```

**Verify `HasSmartArt`'s exact type during implementation** — PP-9's `HasChart` check used `(int)shp.HasChart == -1 /* msoTrue */` (an `MsoTriState`-shaped comparison, `WordTools.cs:492-497`), not a plain bool; `HasSmartArt` may follow the same `MsoTriState` shape rather than a real `bool`. Match whichever this PIA actually returns rather than assuming.

- [x] **Step 2: `read_smartart`**

```csharp
private static ToolResult ReadSmartArt(JsonElement input)
{
    dynamic doc = ActiveDoc;
    var shapes = ListSmartArtShapes(doc);
    if (shapes.Count == 0)
        return new ToolResult { Output = "No SmartArt diagrams in this document.", Summary = "read_smartart" };

    int index = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number ? si.GetInt32() : 0;
    if (index < 0 || index >= shapes.Count)
        throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");

    dynamic smartArt = shapes[index].SmartArt;
    dynamic nodes = smartArt.Nodes;
    int count = (int)nodes.Count;
    var sb = new System.Text.StringBuilder();
    sb.AppendLine("SmartArt " + index + " of " + shapes.Count + " (" + count + " node(s)):");
    for (int i = 1; i <= count; i++)
    {
        dynamic node = nodes.Item(i);
        string text = "";
        try { text = (string)node.TextFrame2.TextRange.Text; } catch { }
        sb.AppendLine("[" + (i - 1) + "] " + text);
    }
    return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_smartart" };
}
```

`nodes.Item(i)` vs `nodes[i]` — **verify which indexer form this dynamic COM collection actually accepts**; both are plausible for a `Microsoft.Office.Interop.SmartArt.IMsoDiagramNodes`-shaped collection, and getting it wrong throws a `RuntimeBinderException` at runtime rather than a compile error, same risk class as every other `dynamic` call in this file.

- [x] **Step 3: `edit_smartart`**

```csharp
private static ToolResult EditSmartArt(JsonElement input)
{
    dynamic doc = ActiveDoc;
    var shapes = ListSmartArtShapes(doc);
    if (shapes.Count == 0)
        throw new InvalidOperationException("edit_smartart: no SmartArt diagrams in this document.");

    int index = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number ? si.GetInt32() : 0;
    if (index < 0 || index >= shapes.Count)
        throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");

    dynamic smartArt = shapes[index].SmartArt;
    dynamic nodes = smartArt.Nodes;
    string kind = input.GetProperty("kind").GetString();
    switch (kind)
    {
        case "set_text":
        {
            int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
            int count = (int)nodes.Count;
            if (nodeIndex < 0 || nodeIndex >= count)
                throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
            nodes.Item(nodeIndex + 1).TextFrame2.TextRange.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Node " + nodeIndex + " updated.", Mutated = true, Summary = "edit_smartart" };
        }
        case "add_node":
        {
            dynamic newNode = nodes.Add();
            if (input.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
                newNode.TextFrame2.TextRange.Text = textEl.GetString();
            return new ToolResult { Output = "Node added at index " + ((int)nodes.Count - 1) + ".", Mutated = true, Summary = "edit_smartart" };
        }
        case "delete_node":
        {
            int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
            int count = (int)nodes.Count;
            if (nodeIndex < 0 || nodeIndex >= count)
                throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
            nodes.Item(nodeIndex + 1).Delete();
            return new ToolResult { Output = "Node " + nodeIndex + " deleted. Later node indices have shifted - re-read (read_smartart) before another node edit in the same run.", Mutated = true, Summary = "edit_smartart" };
        }
        default:
            throw new ArgumentException("edit_smartart: unknown kind '" + kind + "'. Valid: set_text, add_node, delete_node.");
    }
}
```

- [x] **Step 4: Register both in `Execute`'s switch; `read_smartart` also in `AlwaysAllowedTools`.**

- [x] **Step 5: Schemas**

```ts
{
  name: 'edit_smartart',
  description:
    'Edits an existing SmartArt diagram\'s flat node list. kind: "set_text" (nodeIndex,text), "add_node" (text?), "delete_node" (nodeIndex). ' +
    'smartArtIndex addresses the diagram (0-based, document order); omit to target the first one. ' +
    'delete_node shifts later node indices - re-read (read_smartart) before another node edit in the same run.',
  inputSchema: {
    type: 'object',
    properties: {
      smartArtIndex: { type: 'number' },
      kind: { type: 'string', enum: ['set_text', 'add_node', 'delete_node'] },
      nodeIndex: { type: 'number' }, text: { type: 'string' },
    },
    required: ['kind'],
  },
},
{
  name: 'read_smartart',
  description: 'Reads an existing SmartArt diagram\'s node texts, one per line. smartArtIndex addresses the diagram (0-based, document order); omit to read the first one.',
  inputSchema: { type: 'object', properties: { smartArtIndex: { type: 'number' } }, required: [] },
},
```

**Verification:** builds; real-Word test round-trips `add_smartart` → `read_smartart` → `edit_smartart(set_text)` → `read_smartart` confirming the change, then `add_node`/`delete_node`.

---

### Task 6: Schema, `toolDisplay`, and system-prompt integration

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`

- [x] **Step 1:** Add all six `toolDisplay` entries (English + Hebrew), matching the style of the existing `edit_chart`/`read_chart` entries — short label, one-sentence description, UI register not the model's schema register (per FT-1's own distinction, `2026-08-23-ft01-full-settings-screen.md` Task 5 Step 2).
- [x] **Step 2:** Add `read_table` and `read_smartart` to `readOnlyTools`, matching `read_chart`'s precedent.
- [x] **Step 3:** Update the Word skill's `systemPrompt` to mention all six tools and, specifically, the "edit_table/edit_smartart replace nothing implicitly, but insert_row/delete_row/insert_col/delete_col/delete_node shift later indices - re-read before a second structural edit in the same run" caveat, mirroring the caveat PP-9's `read_chart` addition already added for charts (`entry.ts`, the `edit_chart REPLACES the whole dataset` sentence added in the prior session turn).
- [x] **Step 4:** Word's tool count goes from 10 to 16. Note the new count in `docs/superpowers/plans/STATUS.md`'s eventual entry for this plan, so it doesn't silently drift the way the original PP-1 estimate did.

**Verification:** `npx tsc --noEmit` clean.

---

### Task 7: Manual verification matrix

Everything in this plan is either statically-typed-but-COM-fallible (tables) or `dynamic`-typed (SmartArt) — build success does not mean runtime success. Test in this order, cheapest/most-certain first:

- [ ] `add_table {rows:3, cols:3, cells:[["a","b","c"],["d","e","f"],["g","h","i"]]}` → a real 3x3 table with the given text, at the end of the document.
- [ ] `add_table` with `afterBlockIndex` set → table appears at that position, inline with the flow.
- [ ] `read_table` → returns the exact grid just created.
- [ ] `edit_table {kind:'set_cell', row:1, col:1, text:'X'}` → cell updates; `read_table` confirms.
- [ ] `edit_table {kind:'insert_row', index:0, before:false}` → new row appears after row 0; `read_table` shows the shift.
- [ ] `edit_table {kind:'delete_col', index:2}` → column removed; `read_table` confirms the new width.
- [ ] `edit_table {kind:'set_style', styleName:'Table Grid', headerRow:true, bandedRows:true}` → visible style change. **If `ApplyStyleHeadingRows`/`ApplyStyleBandedRows` don't compile or don't visibly apply, that is the first thing to fix per Task 1 Step 2's verification note.**
- [ ] A table with at least one merged cell → `read_table` reports `(merged)` for the cells it can't individually address, rather than failing entirely.
- [ ] `add_smartart {layout:'process', items:['Step 1','Step 2','Step 3']}` (no position given, floating) → diagram appears with 3 nodes. **Do this before the anchored variant.**
- [ ] `add_smartart` with `afterBlockIndex` set → **the specific unverified path (the `Anchor:` named parameter on `Shapes.AddSmartArt`) — test this in isolation and be ready for it to fail; if it does, the safe fallback is dropping just the anchored-creation sub-feature and keeping floating creation, per PP-9's own precedent for the identical situation with charts.**
- [ ] Each of the other 6 layout keys (`list`, `cycle`, `hierarchy`, `pyramid`, `matrix`, `venn`) → visibly distinct diagram types.
- [ ] `read_smartart` → returns the node texts just created, in order.
- [ ] `edit_smartart {kind:'set_text', nodeIndex:1, text:'Changed'}` → node updates; `read_smartart` confirms.
- [ ] `edit_smartart {kind:'add_node', text:'New'}` → node count increases; `read_smartart` shows it appended.
- [ ] `edit_smartart {kind:'delete_node', nodeIndex:0}` → node removed; `read_smartart` shows the shift.
- [ ] Unknown `kind` on both `edit_table` and `edit_smartart` → specific error listing valid kinds; nothing changed.
- [ ] Out-of-range `tableIndex`/`smartArtIndex`/`row`/`col`/`nodeIndex` on every tool → specific error naming the valid range; nothing changed.
- [ ] All of the above in Track Changes mode → edits appear as tracked revisions where Word's own table/SmartArt editing supports that (verify Word's own native behavior here first — not everything in Word tracks changes for structural table edits identically to text edits, independent of anything this plan does).
- [ ] Natural language end-to-end: "add a 2x2 table with headers Name and Score", "add a process diagram with 4 steps", "remove the last row from that table", "change the third step to say 'Review'" — each should resolve to a correct tool call on the first attempt.

---

### Task 8: Fix PowerPoint's `read_slide`/`get_deck_context` blindness to table/SmartArt content

**Not part of the "all 6" Word request** — a related, smaller bug surfaced while reading PowerPoint's table/SmartArt code as this plan's reference pattern, folded in here at the user's request rather than tracked separately.

**The bug:** `ShapeText(PowerPoint.Shape shape)` (`PowerPointAiAddIn/PowerPointTools.cs:116-122`) only checks `shape.HasTextFrame`/`shape.TextFrame.HasText`. A table shape or a SmartArt shape has neither — `HasTextFrame` is `msoFalse` for both — so `ShapeText` silently returns `""` for them. Both call sites inherit the blindness: `read_slide` (`:158`, `sb.AppendLine($"[{shapeIndex}] {shape.Name}: {ShapeText(shape)}")`) lists the shape by name with empty content, and `get_deck_context`'s per-slide preview (`:134`) drops it from the joined preview string entirely since empty strings are filtered by the `t.Length > 0` check at `:135`. **Practical effect:** the model can add a table or SmartArt diagram (it already has working `add_table`/`add_smartart` tools), then immediately lose all visibility into what it just created — `read_slide` reports the shape exists but shows no content, so a follow-up edit request ("update the second row") has nothing to ground itself on beyond guessing.

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`

- [x] **Step 1: Extend `ShapeText` to describe table and SmartArt content, not just text frames**

```csharp
private static string ShapeText(PowerPoint.Shape shape)
{
    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
    {
        return shape.TextFrame.TextRange.Text;
    }
    if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
    {
        PowerPoint.Table table = shape.Table;
        var rowsOut = new System.Collections.Generic.List<string>();
        for (int r = 1; r <= table.Rows.Count; r++)
        {
            var cellsOut = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= table.Columns.Count; c++)
            {
                cellsOut.Add(table.Cell(r, c).Shape.TextFrame.TextRange.Text.Replace("\r", " ").Trim());
            }
            rowsOut.Add(string.Join(" | ", cellsOut));
        }
        return "[table " + table.Rows.Count + "x" + table.Columns.Count + ": " + string.Join(" / ", rowsOut) + "]";
    }
    dynamic dshape = shape;
    bool hasSmartArt = false;
    try { hasSmartArt = (bool)(dshape.HasSmartArt == Microsoft.Office.Core.MsoTriState.msoTrue); } catch { }
    if (hasSmartArt)
    {
        dynamic nodes = dshape.SmartArt.Nodes;
        int count = (int)nodes.Count;
        var nodeTexts = new System.Collections.Generic.List<string>();
        for (int i = 1; i <= count; i++)
        {
            try { nodeTexts.Add(((string)nodes.Item(i).TextFrame2.TextRange.Text).Replace("\r", " ").Trim()); }
            catch { }
        }
        return "[SmartArt " + count + " node(s): " + string.Join(", ", nodeTexts) + "]";
    }
    return "";
}
```

`shape.HasTable`/`shape.Table`/`Table.Cell(r,c).Shape.TextFrame.TextRange.Text` reuse exactly the statically-typed pattern `AddTable`/`EditTableCell` already use elsewhere in this same file (`PowerPointTools.cs:524-624`) — no new risk there. The SmartArt branch is `dynamic` (matching `AddSmartArt`'s own existing pattern at `:967-978`) since `HasSmartArt`/`.SmartArt` aren't exposed on the statically-typed `PowerPoint.Shape` interface this project references; **verify `HasSmartArt`'s exact return type live** — same caveat as PP-23 Task 5 Step 1 raises for Word's identical check, and the two should be fixed consistently if one turns out to need a different comparison shape (e.g. a plain `bool` instead of `MsoTriState`).

- [x] **Step 2: No call-site changes needed.** `read_slide` and `get_deck_context` both already call `ShapeText` and already filter/format on its return value — extending what `ShapeText` returns is sufficient; neither call site's own logic needs to change.

**Verification:** `MSBuild` clean (mixes static and `dynamic` typing in one method, so this is a partial compile-time check only). Manual: a slide with one table and one SmartArt diagram → `read_slide` shows a `[table RxC: ...]` / `[SmartArt N node(s): ...]` line for each instead of blank content; `get_deck_context`'s preview line for that slide includes both summaries instead of omitting them.
