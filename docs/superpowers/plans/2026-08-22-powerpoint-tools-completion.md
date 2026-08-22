# PowerPoint Tools Completion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the gap between `PowerPointAiAddIn/PowerPointTools.cs` (currently 8 tools: `get_deck_context`, `read_slide`, `set_element_text`, `set_element_style`, `set_element_transform`, `add_text_box`, `add_shape`, `delete_element`) and the in-scope subset of genoffice's real `apps/slides` tool surface: `add_slide`, `add_chart`/`edit_chart`, `add_smartart`, `add_table`+its 3 edit tools, `set_slide_background`, `ungroup_element`, `set_element_fill`/`set_element_stroke`, `crop_image`, `set_picture_opacity`, `replace_image` (local-path variant). Explicitly excluded (unchanged from the original project scope): `execute_slide_script`, the entire deck-generation pipeline, and anything requiring internet access (`web_search`/`image_search`/`generate_image`/`analyze_media`/`insert_web_image`).

**Architecture:** Two tools carry genuine, flagged uncertainty that the rest of this codebase's tools don't: `add_smartart` needs an exact layout-name lookup against `Application.SmartArtLayouts` (PowerPoint's SmartArt catalog is name-addressed, not enum-addressed, and exact display-name strings vary slightly by Office version/locale), and `set_picture_opacity` needs live verification of which Interop property actually controls picture transparency on this machine's installed Office build. Both are handled as explicit spike/verification steps within their tasks (Task 6 Step 1, Task 7 Step 3) rather than written blind — consistent with how this project has always resolved genuine COM uncertainty (e.g. Word's original `edit_chart` spike). Every other tool maps to a direct, well-documented native PowerPoint COM call, reusing the existing `ResolveShape(JsonElement)` positional (`slideIndex`/`shapeIndex`) addressing already established in this file — no new addressing scheme is introduced.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (VSTO COM Interop against `Microsoft.Office.Interop.PowerPoint`), matching every other file in `PowerPointAiAddIn/`.

**Spec:** Parameter shapes and semantics below are read directly from genoffice's real source (`apps/slides/src/renderer/ai/slides-skill.ts`), adapted to this project's positional shape addressing instead of genoffice's opaque `sourceId`.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Do not modify the 8 existing tools (`get_deck_context`, `read_slide`, `set_element_text`, `set_element_style`, `set_element_transform`, `add_text_box`, `add_shape`, `delete_element`) — only add new top-level tools and new private helper methods, matching this file's discrete-tools-not-a-gateway pattern (unlike Excel, PowerPoint tools stay one-tool-per-capability, consistent with what's already there).
- Every new mutating tool respects the existing editing-mode gate in `PowerPointTools.Execute` — new read-only tools (none in this plan; all 13 new tools mutate) are NOT added to `AlwaysAllowedTools`.
- `dynamic` typing is used for chart/SmartArt COM objects, following this codebase's established convention (`WordTools.EditChart`, `ExcelTools.AddSparkline`) — a pragmatic choice to sidestep exact Interop type-name uncertainty for less-common COM interfaces.
- Any COM object explicitly created for a chart's embedded-workbook data write (Task 5) must be released deterministically before the tool returns, exactly as required for Excel's `edit_chart` in the sibling Excel plan — a leaked reference orphans a hidden Excel process.
- `add_image`/`replace_image`-equivalent tools reject `http://`/`https://` paths — this deployment is air-gapped; only local file paths are supported, deliberately narrower than genoffice's own URL-accepting schema.
- No automated tests for COM-executor methods — verification is build + manual interactive testing in real PowerPoint.
- Rebuild the esbuild bundle and re-run MSBuild after any `entry.ts` change (Task 8): `npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap`, then `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug`.

---

### Task 1: `add_slide`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`

**Interfaces:**
- Consumes: existing `ActivePresentation` property.
- Produces: nothing new for other tasks.

**Context:** currently there is no way for the AI to add a new slide at all — the most basic structural gap in the deck.

- [ ] **Step 1: Implement `add_slide`**

Add `case "add_slide": return AddSlide(input);` to `Execute`'s switch, and:
```csharp
private static ToolResult AddSlide(JsonElement input)
{
    int sourceIndex = input.GetProperty("sourceIndex").GetInt32();
    bool clearText = !input.TryGetProperty("clearText", out var ct) || ct.ValueKind != JsonValueKind.False;
    PowerPoint.Slides slides = ActivePresentation.Slides;
    if (sourceIndex < 0 || sourceIndex >= slides.Count)
    {
        return new ToolResult { Output = "Invalid sourceIndex.", IsError = true, Summary = "add_slide" };
    }
    dynamic source = slides[sourceIndex + 1];
    dynamic dupRange = source.Duplicate(); // returns a SlideRange containing exactly the new slide
    dynamic newSlide = dupRange[1];
    newSlide.MoveTo(sourceIndex + 2);
    if (clearText)
    {
        foreach (PowerPoint.Shape shape in newSlide.Shapes)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Text = "";
            }
        }
    }
    return new ToolResult { Output = "Slide added after index " + sourceIndex + ".", Mutated = true, Summary = "add_slide" };
}
```

- [ ] **Step 2: Add the tool schema to `entry.ts`**

In `PowerPointAiAddIn/web-src/entry.ts`'s tool array, add:
```typescript
    {
      name: 'add_slide',
      description: 'Clones an existing slide\'s layout as a new blank (or templated) slide inserted right after it.',
      inputSchema: {
        type: 'object',
        properties: { sourceIndex: { type: 'number' }, clearText: { type: 'boolean' } },
        required: ['sourceIndex'],
      },
    },
```

- [ ] **Step 3: Build and manually verify**

Run: `npx tsc --noEmit` then the esbuild command from Global Constraints, then MSBuild. Expected: 0 errors.

Manually verify: `add_slide` with `{"sourceIndex":0}` inserts a new blank slide (same layout/background as slide 0) as slide 2.

- [ ] **Step 4: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add add_slide - closes the no-way-to-add-a-slide gap"
```

---

### Task 2: `set_element_fill`, `set_element_stroke`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: existing `ResolveShape(JsonElement)`.
- Produces: `private static int HexToOle(string hex)` — a shared helper Tasks 3-7 also use (the existing `SetElementStyle` method keeps its own inline hex conversion untouched — this is a new, separate helper for new tools only, not a refactor of existing code).

- [ ] **Step 1: Add the shared `HexToOle` helper and both tools**

Add `case "set_element_fill": return SetElementFill(input);` and `case "set_element_stroke": return SetElementStroke(input);` to `Execute`'s switch, and:
```csharp
private static int HexToOle(string hex)
{
    hex = hex.TrimStart('#');
    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
    return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
}

private static ToolResult SetElementFill(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    string fill = input.GetProperty("fill").GetString();
    if (fill == "none")
    {
        shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
    }
    else
    {
        shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        shape.Fill.ForeColor.RGB = HexToOle(fill);
    }
    return new ToolResult { Output = "Fill updated.", Mutated = true, Summary = "set_element_fill" };
}

private static ToolResult SetElementStroke(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    bool remove = input.TryGetProperty("remove", out var r) && r.ValueKind == JsonValueKind.True;
    if (remove)
    {
        shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
    }
    else
    {
        shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (input.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
        {
            shape.Line.ForeColor.RGB = HexToOle(color.GetString());
        }
        shape.Line.Weight = input.TryGetProperty("widthPt", out var width) && width.ValueKind == JsonValueKind.Number ? (float)width.GetDouble() : 1f;
    }
    return new ToolResult { Output = "Stroke updated.", Mutated = true, Summary = "set_element_stroke" };
}
```

- [ ] **Step 2: Add both tool schemas to `entry.ts`**

```typescript
    {
      name: 'set_element_fill',
      description: 'Sets a shape\'s solid fill color, or "none" to remove its fill.',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, fill: { type: 'string' } },
        required: ['slideIndex', 'shapeIndex', 'fill'],
      },
    },
    {
      name: 'set_element_stroke',
      description: 'Sets a shape\'s outline/stroke color and width, or removes it.',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
          color: { type: 'string' }, widthPt: { type: 'number' }, remove: { type: 'boolean' },
        },
        required: ['slideIndex', 'shapeIndex'],
      },
    },
```

- [ ] **Step 3: Build and manually verify**

Run the same verification commands as Task 1 Step 3. Expected: 0 errors.

Manually verify: `set_element_fill` with `{"fill":"#4a9eff"}` fills a shape blue; `{"fill":"none"}` removes it; `set_element_stroke` with `{"color":"#000000","widthPt":3}` adds a thick black outline; `{"remove":true}` removes it.

- [ ] **Step 4: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add set_element_fill and set_element_stroke"
```

---

### Task 3: `set_slide_background`, `ungroup_element`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: `HexToOle` (Task 2), `ResolveShape` (existing).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement both tools**

Add `case "set_slide_background": return SetSlideBackground(input);` and `case "ungroup_element": return UngroupElement(input);`, and:
```csharp
private static ToolResult SetSlideBackground(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    int oleColor = HexToOle(input.GetProperty("color").GetString());
    PowerPoint.Slides slides = ActivePresentation.Slides;

    void Apply(PowerPoint.Slide s)
    {
        s.Background.Fill.ForeColor.RGB = oleColor;
        s.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
    }

    if (slideIndex == -1)
    {
        foreach (PowerPoint.Slide s in slides) Apply(s);
    }
    else
    {
        Apply(slides[slideIndex + 1]);
    }
    return new ToolResult { Output = "Background updated.", Mutated = true, Summary = "set_slide_background" };
}

private static ToolResult UngroupElement(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    shape.Ungroup();
    return new ToolResult { Output = "Shape ungrouped - re-read the slide (read_slide) to get updated shape indices before addressing the promoted children.", Mutated = true, Summary = "ungroup_element" };
}
```

- [ ] **Step 2: Add both tool schemas to `entry.ts`**

```typescript
    {
      name: 'set_slide_background',
      description: 'Sets a solid background color for one slide, or slideIndex=-1 for every slide in the deck.',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, color: { type: 'string' } },
        required: ['slideIndex', 'color'],
      },
    },
    {
      name: 'ungroup_element',
      description: 'Promotes a group shape\'s direct children to top-level shapes. Shape indices change after this call - re-read the slide before addressing the promoted shapes.',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } },
        required: ['slideIndex', 'shapeIndex'],
      },
    },
```

- [ ] **Step 3: Build and manually verify**

Run the same verification commands. Expected: 0 errors.

Manually verify: `set_slide_background` with `{"slideIndex":-1,"color":"#f0f0f0"}` changes every slide's background; `ungroup_element` on a manually-grouped set of shapes promotes them to top-level (confirm via `read_slide` showing more independently-addressable shapes than before).

- [ ] **Step 4: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add set_slide_background and ungroup_element"
```

---

### Task 4: Native tables — `add_table`, `edit_table_cell`, `edit_table_structure`, `edit_table_style`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: `ResolveShape`, `HexToOle`.
- Produces: `private static PowerPoint.Table ResolveTable(JsonElement)`, used only within this task's 3 edit tools.

**Scope decision:** `edit_table_style`'s named whole-table style presets (genoffice's `styleName`, e.g. `"lightGrid"`) require PowerPoint's built-in table-style GUID catalog, which classic Interop exposes only as opaque GUID strings via `Table.ApplyStyle(guid)` — not a named-constant enum. Pinning down the exact 8 GUIDs needs a one-time lookup against a live Office install (apply each style manually via the ribbon, read back `Table.TableStyle` — a GUID string — in the debugger). This plan deliberately scopes `edit_table_style` to the granular COM properties only (`firstRow`, `bandRow`, `shadingColor`, `borderColor`/`borderWidthPt`/`borderPreset`), which need no GUID lookup and already cover the common styling needs; named-preset support is a documented follow-up, not silently dropped.

- [ ] **Step 1: Implement `add_table`**

Add `case "add_table": return AddTable(input);` and:
```csharp
private static ToolResult AddTable(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    int rows = input.GetProperty("rows").GetInt32();
    int cols = input.GetProperty("cols").GetInt32();
    float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
    float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
    float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
    float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 200f;

    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    PowerPoint.Shape tableShape = slide.Shapes.AddTable(rows, cols, left, top, width, height);
    if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
    {
        int r = 0;
        foreach (JsonElement rowEl in cells.EnumerateArray())
        {
            int c = 0;
            foreach (JsonElement cellEl in rowEl.EnumerateArray())
            {
                tableShape.Table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = cellEl.GetString();
                c++;
            }
            r++;
        }
    }
    return new ToolResult { Output = "Table added.", Mutated = true, Summary = "add_table" };
}
```

- [ ] **Step 2: Implement `edit_table_cell`, `edit_table_structure`, `edit_table_style`**

Add the 3 `case` branches and:
```csharp
private static PowerPoint.Table ResolveTable(JsonElement input)
{
    return ResolveShape(input).Table;
}

private static ToolResult EditTableCell(JsonElement input)
{
    PowerPoint.Table table = ResolveTable(input);
    int row = input.GetProperty("row").GetInt32();
    int col = input.GetProperty("col").GetInt32();
    string text = input.GetProperty("paragraphs").GetString();
    table.Cell(row + 1, col + 1).Shape.TextFrame.TextRange.Text = text;
    return new ToolResult { Output = "Cell updated.", Mutated = true, Summary = "edit_table_cell" };
}

private static ToolResult EditTableStructure(JsonElement input)
{
    PowerPoint.Table table = ResolveTable(input);
    string kind = input.GetProperty("kind").GetString();
    int index = input.GetProperty("index").GetInt32();
    bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
    switch (kind)
    {
        case "insert-row": table.Rows.Add(before ? index + 1 : index + 2); break;
        case "delete-row": table.Rows[index + 1].Delete(); break;
        case "insert-col": table.Columns.Add(before ? index + 1 : index + 2); break;
        case "delete-col": table.Columns[index + 1].Delete(); break;
        default: return new ToolResult { Output = "Unknown structure kind: " + kind, IsError = true, Summary = "edit_table_structure" };
    }
    return new ToolResult { Output = "Table structure updated.", Mutated = true, Summary = "edit_table_structure" };
}

private static ToolResult EditTableStyle(JsonElement input)
{
    PowerPoint.Table table = ResolveTable(input);
    if (input.TryGetProperty("firstRow", out var firstRow))
    {
        table.FirstRow = firstRow.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }
    if (input.TryGetProperty("bandRow", out var bandRow))
    {
        table.HorizBanding = bandRow.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }
    if (input.TryGetProperty("shadingColor", out var shading) && shading.ValueKind == JsonValueKind.String)
    {
        int color = HexToOle(shading.GetString());
        foreach (PowerPoint.Row row in table.Rows)
        {
            foreach (PowerPoint.Cell cell in row.Cells)
            {
                cell.Shape.Fill.ForeColor.RGB = color;
            }
        }
    }
    if (input.TryGetProperty("borderColor", out _) || input.TryGetProperty("borderWidthPt", out _) || input.TryGetProperty("borderPreset", out _))
    {
        bool visible = !(input.TryGetProperty("borderPreset", out var bp) && bp.ValueKind == JsonValueKind.String && bp.GetString() == "none");
        float weight = input.TryGetProperty("borderWidthPt", out var bw) && bw.ValueKind == JsonValueKind.Number ? (float)bw.GetDouble() : 1f;
        int color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String ? HexToOle(bc.GetString()) : HexToOle("#000000");
        PowerPoint.PpBorderType[] sides = { PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom, PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight };
        foreach (PowerPoint.Row row in table.Rows)
        {
            foreach (PowerPoint.Cell cell in row.Cells)
            {
                foreach (PowerPoint.PpBorderType side in sides)
                {
                    PowerPoint.Border border = cell.Borders[side];
                    border.Visible = visible ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                    border.Weight = weight;
                    border.ForeColor.RGB = color;
                }
            }
        }
    }
    return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table_style" };
}
```

- [ ] **Step 3: Add all 4 tool schemas to `entry.ts`**

```typescript
    {
      name: 'add_table',
      description: 'Adds a native PowerPoint table, optionally pre-filled with cell text (row-major array of arrays).',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, rows: { type: 'number' }, cols: { type: 'number' },
          cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
          x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        },
        required: ['slideIndex', 'rows', 'cols'],
      },
    },
    {
      name: 'edit_table_cell',
      description: 'Replaces one table cell\'s text (0-based row/col).',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, row: { type: 'number' }, col: { type: 'number' }, paragraphs: { type: 'string' } },
        required: ['slideIndex', 'shapeIndex', 'row', 'col', 'paragraphs'],
      },
    },
    {
      name: 'edit_table_structure',
      description: 'Inserts or deletes a table row/column. kind: "insert-row"|"delete-row"|"insert-col"|"delete-col".',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, kind: { type: 'string' }, index: { type: 'number' }, before: { type: 'boolean' } },
        required: ['slideIndex', 'shapeIndex', 'kind', 'index'],
      },
    },
    {
      name: 'edit_table_style',
      description: 'Applies granular table styling: firstRow/bandRow (header row / banded rows), shadingColor (all cells), borderColor/borderWidthPt/borderPreset ("all"|"none").',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
          firstRow: { type: 'boolean' }, bandRow: { type: 'boolean' }, shadingColor: { type: 'string' },
          borderColor: { type: 'string' }, borderWidthPt: { type: 'number' }, borderPreset: { type: 'string' },
        },
        required: ['slideIndex', 'shapeIndex'],
      },
    },
```

- [ ] **Step 4: Build and manually verify**

Run the same verification commands. Expected: 0 errors.

Manually verify: `add_table` with `{"slideIndex":0,"rows":3,"cols":3,"cells":[["A","B","C"],["1","2","3"],["4","5","6"]]}` creates a filled 3x3 table; `edit_table_structure` inserts/deletes a row and column; `edit_table_style` with `{"firstRow":true,"bandRow":true}` visibly styles the header row and alternating row shading; `borderColor`/`borderWidthPt` visibly changes cell borders.

- [ ] **Step 5: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add native table tools (add_table, edit_table_cell/structure/style)"
```

---

### Task 5: `add_chart`, `edit_chart`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: `ResolveShape`.
- Produces: nothing new for other tasks.

**Context:** this is one of the two headline findings from the original feasibility report that justified choosing VSTO over Office.js for PowerPoint (no chart object model exists in PowerPoint's Office.js API at all) — never built until now.

- [ ] **Step 1: Implement `add_chart`**

Add `case "add_chart": return AddChartPpt(input);` and:
```csharp
private static readonly Dictionary<string, int> PptChartTypeMap = new Dictionary<string, int>
{
    ["bar"] = 51,          // xlColumnClustered
    ["barStacked"] = 52,   // xlColumnStacked
    ["line"] = 4,          // xlLine
    ["area"] = 1,          // xlArea
    ["pie"] = 5,           // xlPie
    ["doughnut"] = -4120,  // xlDoughnut
};

private static ToolResult AddChartPpt(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    string kindStr = input.GetProperty("kind").GetString();
    int typeCode = PptChartTypeMap.TryGetValue(kindStr, out var t) ? t : 51;
    var categories = new List<string>();
    foreach (JsonElement c in input.GetProperty("categories").EnumerateArray()) categories.Add(c.GetString());
    float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
    float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
    float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
    float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    dynamic chartShape = slide.Shapes.AddChart2(-1, typeCode, left, top, width, height);
    dynamic chart = chartShape.Chart;

    // Chart data lives in an embedded Excel workbook - open, write the grid,
    // close, and RELEASE explicitly so no hidden Excel host process leaks.
    dynamic dataWorkbook = chart.ChartData.Workbook;
    try
    {
        dynamic sheet = dataWorkbook.Worksheets[1];
        JsonElement seriesArray = input.GetProperty("series");
        int colIdx = 0;
        foreach (JsonElement s in seriesArray.EnumerateArray())
        {
            sheet.Cells[1, colIdx + 2].Value = s.GetProperty("name").GetString();
            colIdx++;
        }
        for (int r = 0; r < categories.Count; r++)
        {
            sheet.Cells[r + 2, 1].Value = categories[r];
        }
        colIdx = 0;
        foreach (JsonElement s in seriesArray.EnumerateArray())
        {
            int r = 0;
            foreach (JsonElement v in s.GetProperty("values").EnumerateArray())
            {
                sheet.Cells[r + 2, colIdx + 2].Value = v.GetDouble();
                r++;
            }
            colIdx++;
        }
        dynamic usedRange = sheet.UsedRange;
        chart.SetSourceData(usedRange);
    }
    finally
    {
        dataWorkbook.Close(SaveChanges: true);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook);
    }

    if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
    {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title.GetString();
    }
    return new ToolResult { Output = "Chart added.", Mutated = true, Summary = "add_chart" };
}
```

- [ ] **Step 2: Implement `edit_chart`**

Add `case "edit_chart": return EditChartPpt(input);` and:
```csharp
private static ToolResult EditChartPpt(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    dynamic chart = shape.Chart;

    if (input.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String && PptChartTypeMap.TryGetValue(ct.GetString(), out var typeCode))
    {
        chart.ChartType = typeCode;
    }
    if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
    {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title.GetString();
    }
    if (input.TryGetProperty("legendPos", out var legendPos) && legendPos.ValueKind == JsonValueKind.String)
    {
        string pos = legendPos.GetString();
        if (pos == "none")
        {
            chart.HasLegend = false;
        }
        else
        {
            chart.HasLegend = true;
            chart.Legend.Position = pos == "r" ? -4152 : pos == "t" ? -4160 : pos == "l" ? -4131 : -4107;
        }
    }
    if (input.TryGetProperty("dataLabels", out var dl))
    {
        bool show = dl.ValueKind == JsonValueKind.True;
        foreach (dynamic series in chart.SeriesCollection())
        {
            series.HasDataLabels = show;
        }
    }
    if (input.TryGetProperty("gridlines", out var gl))
    {
        chart.Axes(2 /* xlValue */).HasMajorGridlines = gl.ValueKind == JsonValueKind.True;
    }
    return new ToolResult { Output = "Chart updated.", Mutated = true, Summary = "edit_chart" };
}
```
(Series data repointing via `edit_chart` is deliberately deferred — it would reuse `add_chart`'s embedded-workbook-write pattern above; type/title/legend/labels/gridlines already cover the common "adjust an existing chart" requests.)

- [ ] **Step 3: Add both tool schemas to `entry.ts`**

```typescript
    {
      name: 'add_chart',
      description: 'Adds a native, editable PowerPoint chart. kind: "bar"|"barStacked"|"line"|"area"|"pie"|"doughnut".',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, kind: { type: 'string' }, title: { type: 'string' },
          categories: { type: 'array', items: { type: 'string' } },
          series: { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } } } },
          x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        },
        required: ['slideIndex', 'kind', 'categories', 'series'],
      },
    },
    {
      name: 'edit_chart',
      description: 'Modifies an existing chart\'s type/title/legend position/data labels/gridlines.',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
          chartType: { type: 'string' }, title: { type: 'string' },
          legendPos: { type: 'string' }, dataLabels: { type: 'boolean' }, gridlines: { type: 'boolean' },
        },
        required: ['slideIndex', 'shapeIndex'],
      },
    },
```

- [ ] **Step 4: Build and manually verify**

Run the same verification commands. Expected: 0 errors.

Manually verify: `add_chart` with `{"slideIndex":0,"kind":"bar","title":"Q1 Sales","categories":["Jan","Feb","Mar"],"series":[{"name":"Revenue","values":[10,20,15]}]}` produces a genuine, editable native chart (Chart Design ribbon tab appears when selected, and double-clicking opens the real embedded-workbook data editor). Watch Task Manager for an orphaned `EXCEL.EXE` process after the call. `edit_chart` with `{"chartType":"line","legendPos":"none"}` visibly changes it.

- [ ] **Step 5: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add add_chart and edit_chart - closes a headline VSTO-justification gap"
```

---

### Task 6: `add_smartart`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new for other tasks.

**Context:** the second headline VSTO-justification gap from the original feasibility report (no SmartArt creation/edit API exists in Office.js at all). PowerPoint's SmartArt COM API addresses layouts by display name, not a fixed enum — genoffice itself only ever produces flat item lists (no nested hierarchy), which maps cleanly to sequential top-level `SmartArt.Nodes.Add()` calls.

- [ ] **Step 1: Verify the 7 layout display-name strings against this machine's real Office install (required before Step 2)**

The lookup table below uses the standard English display names for PowerPoint's built-in SmartArt layout gallery. Before relying on it, confirm these names actually match `Application.SmartArtLayouts` on this machine's installed Office version — names can vary slightly by version/locale. In Visual Studio's Immediate window (or a temporary throwaway debug print) while a PowerPoint instance with this add-in loaded is running, evaluate something equivalent to:
```csharp
foreach (dynamic layout in Globals.ThisAddIn.Application.SmartArtLayouts) { System.Diagnostics.Debug.WriteLine(layout.Name); }
```
and cross-check that `"Basic Block List"`, `"Basic Process"`, `"Basic Cycle"`, `"Organization Chart"`, `"Basic Pyramid"`, `"Basic Matrix"`, `"Basic Venn"` are all present verbatim. If any differ, update `SmartArtLayoutNames` in Step 2 to match the real strings before proceeding — do not guess or skip this check, since a wrong name throws at runtime with no other fallback.

- [ ] **Step 2: Implement `add_smartart`**

Add `case "add_smartart": return AddSmartArt(input);` and:
```csharp
private static readonly Dictionary<string, string> SmartArtLayoutNames = new Dictionary<string, string>
{
    ["list"] = "Basic Block List",
    ["process"] = "Basic Process",
    ["cycle"] = "Basic Cycle",
    ["hierarchy"] = "Organization Chart",
    ["pyramid"] = "Basic Pyramid",
    ["matrix"] = "Basic Matrix",
    ["venn"] = "Basic Venn",
};

private static dynamic ResolveSmartArtLayout(string layoutKey)
{
    string targetName = SmartArtLayoutNames.TryGetValue(layoutKey, out var name) ? name : "Basic Block List";
    dynamic layouts = Globals.ThisAddIn.Application.SmartArtLayouts;
    foreach (dynamic layout in layouts)
    {
        if (string.Equals((string)layout.Name, targetName, StringComparison.OrdinalIgnoreCase))
        {
            return layout;
        }
    }
    throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName + "' found - see plan Task 6 Step 1.");
}

private static ToolResult AddSmartArt(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    string layoutKey = input.GetProperty("layout").GetString();
    float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
    float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
    float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
    float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

    dynamic layout = ResolveSmartArtLayout(layoutKey);
    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    dynamic shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
    dynamic smartArt = shape.SmartArt;

    // genoffice's own version only ever produces a flat item list - maps
    // 1:1 to sequential top-level nodes, no nested tree-building needed.
    foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
    {
        dynamic node = smartArt.Nodes.Add();
        node.TextFrame2.TextRange.Text = item.GetString();
    }
    return new ToolResult { Output = "SmartArt added.", Mutated = true, Summary = "add_smartart" };
}
```

- [ ] **Step 3: Add the tool schema to `entry.ts`**

```typescript
    {
      name: 'add_smartart',
      description: 'Adds a shape-composed SmartArt diagram. layout: "list"|"process"|"cycle"|"hierarchy"|"pyramid"|"matrix"|"venn". items are flat node texts, one per top-level node.',
      inputSchema: {
        type: 'object',
        properties: {
          slideIndex: { type: 'number' }, layout: { type: 'string' },
          items: { type: 'array', items: { type: 'string' } },
          x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        },
        required: ['slideIndex', 'layout', 'items'],
      },
    },
```

- [ ] **Step 4: Build and manually verify**

Run the same verification commands. Expected: 0 errors.

Manually verify: `add_smartart` with `{"slideIndex":0,"layout":"process","items":["Plan","Build","Ship"]}` inserts a real, editable SmartArt process diagram with those 3 node texts. Try at least 2 more layout keys to confirm the lookup table is correct.

- [ ] **Step 5: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add add_smartart - closes the second headline VSTO-justification gap"
```

---

### Task 7: `crop_image`, `set_picture_opacity`, `replace_image`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: `ResolveShape`.
- Produces: nothing new — this is the last tool-adding task in the plan.

**Context:** all three operate on an image already embedded in the deck — no internet/generation needed, so all three are in-scope despite being paired with AI generation in genoffice's own version.

- [ ] **Step 1: Implement `crop_image`**

Add `case "crop_image": return CropImage(input);` and:
```csharp
private static ToolResult CropImage(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    float l = (float)input.GetProperty("l").GetDouble();
    float t = (float)input.GetProperty("t").GetDouble();
    float r = (float)input.GetProperty("r").GetDouble();
    float b = (float)input.GetProperty("b").GetDouble();
    // Approximation, documented deliberately: fractions are applied against
    // the shape's CURRENT on-slide size, not the original uncropped source
    // image - classic Interop has no reliable "natural size" property once a
    // picture has already been resized/cropped on the slide. Correct for a
    // freshly-inserted, never-before-cropped picture; imprecise under
    // repeated crop calls on the same shape.
    shape.PictureFormat.CropLeft = l * shape.Width;
    shape.PictureFormat.CropTop = t * shape.Height;
    shape.PictureFormat.CropRight = r * shape.Width;
    shape.PictureFormat.CropBottom = b * shape.Height;
    return new ToolResult { Output = "Image cropped.", Mutated = true, Summary = "crop_image" };
}
```

- [ ] **Step 2: Implement `replace_image`**

Add `case "replace_image": return ReplaceImagePpt(input);` and:
```csharp
private static ToolResult ReplaceImagePpt(JsonElement input)
{
    string localPath = input.GetProperty("localPath").GetString();
    if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
    {
        return new ToolResult { Output = "replace_image: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "replace_image" };
    }
    PowerPoint.Shape oldShape = ResolveShape(input);
    bool keepCrop = input.TryGetProperty("keepCrop", out var kc) && kc.ValueKind == JsonValueKind.True;

    float left = oldShape.Left, top = oldShape.Top, width = oldShape.Width, height = oldShape.Height, rotation = oldShape.Rotation;
    int zPos = oldShape.ZOrderPosition;
    float cropLeft = 0, cropTop = 0, cropRight = 0, cropBottom = 0;
    if (keepCrop)
    {
        cropLeft = oldShape.PictureFormat.CropLeft;
        cropTop = oldShape.PictureFormat.CropTop;
        cropRight = oldShape.PictureFormat.CropRight;
        cropBottom = oldShape.PictureFormat.CropBottom;
    }
    PowerPoint.Slide slide = (PowerPoint.Slide)oldShape.Parent;
    oldShape.Delete();

    PowerPoint.Shape newShape = slide.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
    newShape.Rotation = rotation;
    if (keepCrop)
    {
        newShape.PictureFormat.CropLeft = cropLeft;
        newShape.PictureFormat.CropTop = cropTop;
        newShape.PictureFormat.CropRight = cropRight;
        newShape.PictureFormat.CropBottom = cropBottom;
    }
    // Restore approximate z-order: send to back, then bring forward to the
    // original stack position.
    newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
    for (int i = 1; i < zPos; i++)
    {
        newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
    }
    return new ToolResult { Output = "Image replaced.", Mutated = true, Summary = "replace_image" };
}
```

- [ ] **Step 3: Verify `set_picture_opacity`'s real Interop property, then implement it**

Research flagged genuine, unresolved uncertainty here: classic Interop historically had no dedicated opacity property for picture-type shapes (the same Office.js gap the original feasibility report flagged), but Office 2016+'s "Picture Transparency" UI control likely maps to `Shape.Fill.Transparency` even for a picture-type shape's underlying `Fill` object. Confirm against this machine's actual installed PIA before shipping:

In a live PowerPoint session with an inserted picture selected, in the Immediate window (or a throwaway debug print) try:
```csharp
dynamic shape = /* the selected picture Shape */;
shape.Fill.Transparency = 0.5f;
```
and visually confirm the picture becomes semi-transparent. If this works, implement:
```csharp
private static ToolResult SetPictureOpacity(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    float opacity = (float)input.GetProperty("opacity").GetDouble();
    dynamic dShape = shape;
    dShape.Fill.Transparency = 1f - opacity;
    return new ToolResult { Output = "Opacity updated.", Mutated = true, Summary = "set_picture_opacity" };
}
```
If the live check shows `Fill.Transparency` has no visible effect on a picture-type shape on this Office build, do not ship a no-op tool — instead return a clear, honest error from `SetPictureOpacity` (`IsError = true`, explaining picture opacity isn't supported on this Office build) rather than silently pretending to succeed, and note the finding in the commit message for Step 5.

Add `case "set_picture_opacity": return SetPictureOpacity(input);` to `Execute`'s switch regardless of which branch above applies.

- [ ] **Step 4: Add all 3 tool schemas to `entry.ts`**

```typescript
    {
      name: 'crop_image',
      description: 'Non-destructively crops a picture shape. l/t/r/b are 0..1 fractions of the current on-slide image size cut from each edge; all zero clears the crop.',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, l: { type: 'number' }, t: { type: 'number' }, r: { type: 'number' }, b: { type: 'number' } },
        required: ['slideIndex', 'shapeIndex', 'l', 't', 'r', 'b'],
      },
    },
    {
      name: 'set_picture_opacity',
      description: 'Sets a picture shape\'s overall opacity, 0 (invisible) to 1 (fully opaque).',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, opacity: { type: 'number' } },
        required: ['slideIndex', 'shapeIndex', 'opacity'],
      },
    },
    {
      name: 'replace_image',
      description: 'Swaps a picture shape\'s image content in place from a local file path, keeping position/size/rotation/approximate z-order.',
      inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, localPath: { type: 'string' }, keepCrop: { type: 'boolean' } },
        required: ['slideIndex', 'shapeIndex', 'localPath'],
      },
    },
```

- [ ] **Step 5: Build and manually verify**

Run the same verification commands. Expected: 0 errors.

Manually verify against a slide with an inserted picture: `crop_image` with `{"l":0.1,"t":0.1,"r":0.1,"b":0.1}` visibly crops 10% off each edge; `set_picture_opacity` with `{"opacity":0.5}` makes the picture semi-transparent (or returns the honest error from Step 3 if unsupported on this build); `replace_image` with a different local `.png` path swaps the visible image while keeping its position/size.

- [ ] **Step 6: Commit**

```bash
git add PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/web-src/entry.ts
git commit -m "feat(powerpoint): add crop_image, set_picture_opacity, replace_image (local-path only)"
```

---

### Task 8: Update the system prompt to mention the new tools

**Files:**
- Modify: `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing new.

- [ ] **Step 1: Update `powerPointSkill`'s system prompt**

Find the `systemPrompt` string in `PowerPointAiAddIn/web-src/entry.ts` and extend it to mention the 13 new tools added across Tasks 1-7 (add_slide, add_chart/edit_chart, add_smartart, add_table + its 3 edit tools, set_slide_background, ungroup_element, set_element_fill/set_element_stroke, crop_image/set_picture_opacity/replace_image), alongside whatever tools it already names, so the model knows they exist. Match the tone/format already established for Word's equivalent update.

- [ ] **Step 2: Typecheck and rebuild**

Run `npx tsc --noEmit` then the esbuild command from Global Constraints. Expected: 0 errors.

- [ ] **Step 3: Commit**

```bash
git add PowerPointAiAddIn/web-src/entry.ts
git commit -m "docs(powerpoint): mention the 13 new tools in the system prompt"
```
