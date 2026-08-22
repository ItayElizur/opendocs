# Word Tools Completion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the gap between `WordAiAddIn/WordTools.cs`'s `apply_commands` tool (currently 4 command kinds: `set_bold`, `set_italic`, `set_heading`, `find_replace`) and genoffice's real `apps/docs` `apply_commands` surface (10 command kinds total, per `C:\Dev\genoffice\docs\ai-tool-surface.md` and `apps/docs/src/renderer/ai/commands.ts`), by adding the 8 missing kinds: `updateTextStyle`, `updateParagraphStyle`, `deleteBlocks`, `moveBlocks`, `createParagraphBullets`, `deleteParagraphBullets`, `updateImageProperties`, `insertToc`.

**Architecture:** genoffice addresses commands via a `Target` object (`nodeType`/`headingLevel`/`containsText`/`blockIndexes`/`scope`) matched against ProseMirror's flat top-level block sequence — Word's COM object model has no equivalent unified block collection mixing paragraphs and images, so this plan introduces a `Target`-resolution helper that maps the same addressing semantics onto Word's 0-based `Paragraphs` collection instead, with image targeting deliberately excluded from the generic `Target` path (images get their own index-based addressing in `updateImageProperties`, since `nodeType:'image'` has no clean paragraph-index equivalent — confirmed via direct source research, not guessed). The 4 existing command kinds (`set_bold`/`set_italic`/`set_heading`/`find_replace`) are left completely untouched — this plan only adds new `case` branches to the same `switch` in `ApplyCommands`, so no existing behavior, already-shipped tool schema, or test can regress.

**Tech Stack:** C# 7.3 / .NET Framework 4.8 (VSTO COM Interop against `Microsoft.Office.Interop.Word`), matching every other file in `WordAiAddIn/`.

**Spec:** Command semantics are read directly from genoffice's real source (`apps/docs/src/renderer/ai/commands.ts`) — this plan's Global Constraints section restates the exact parameter shapes found there, adapted for paragraph-index addressing.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Do not modify the 4 existing `apply_commands` kinds (`set_bold`, `set_italic`, `set_heading`, `find_replace`) or any other existing tool (`get_document_context`, `insert_content`, `edit_chart`, `read_blocks`, `replace_blocks`, `add_comment`) — only add new `case` branches and new private helper methods to `WordTools.cs`.
- No automated tests for COM-executor methods (existing project convention) — verification is build + manual interactive testing in real Word, same pattern as every prior Word task.
- Rebuild the esbuild bundle and re-run MSBuild after any `entry.ts` change (Task 5): `npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --target=chrome100 --format=iife --sourcemap`, then `MSBuild WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug`.
- Every new command kind must respect the existing editing-mode gate in `WordTools.Execute` (already wraps the whole `switch` — no per-command gating needed, just don't bypass it).
- `Target`'s `scope` field, when `"selection"`, must use `Globals.ThisAddIn.Application.Selection.Range` — never assume a selection exists without checking `Selection.Type != WdSelectionType.wdNoSelection` first (an empty/collapsed selection should fall back to matching nothing rather than throwing).

### `Target` object — shared addressing contract for this plan's new commands

```json
{
  "nodeType": "heading" | "paragraph" | "listItem" | null,
  "headingLevel": 1-6 (optional, only meaningful with nodeType:'heading'),
  "containsText": "string (optional)",
  "matchCase": false,
  "blockIndexes": [0, 2, 5] (optional, 0-based paragraph indices),
  "scope": "document" | "selection" (default "document")
}
```
All given fields are AND-combined. At least one of `nodeType`/`containsText`/`blockIndexes` must be present (a bare `scope`-only target with nothing else is invalid — reject with an error, don't silently match everything). `nodeType:"image"` is NOT supported by this `Target` — `updateImageProperties` (Task 4) uses separate `InlineShapes`-index addressing instead.

---

### Task 1: `Target` resolution helper, `updateTextStyle`, `updateParagraphStyle`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Consumes: nothing new.
- Produces: `private static List<int> ResolveTargetParagraphs(JsonElement target)` — returns 0-based paragraph indices matching the `Target` contract above. Tasks 2 and 3 reuse this method for their own `Target`-addressed commands.

- [ ] **Step 1: Implement `ResolveTargetParagraphs`**

Add to `WordTools.cs`:
```csharp
private static List<int> ResolveTargetParagraphs(JsonElement target)
{
    string nodeType = target.TryGetProperty("nodeType", out var nt) && nt.ValueKind == JsonValueKind.String ? nt.GetString() : null;
    int? headingLevel = target.TryGetProperty("headingLevel", out var hl) && hl.ValueKind == JsonValueKind.Number ? hl.GetInt32() : (int?)null;
    string containsText = target.TryGetProperty("containsText", out var ct) && ct.ValueKind == JsonValueKind.String ? ct.GetString() : null;
    bool matchCase = target.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
    HashSet<int> blockIndexes = null;
    if (target.TryGetProperty("blockIndexes", out var bi) && bi.ValueKind == JsonValueKind.Array)
    {
        blockIndexes = new HashSet<int>();
        foreach (JsonElement e in bi.EnumerateArray()) blockIndexes.Add(e.GetInt32());
    }
    string scope = target.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String ? sc.GetString() : "document";

    if (nodeType == null && containsText == null && blockIndexes == null)
    {
        throw new ArgumentException("Target must specify at least one of nodeType, containsText, or blockIndexes.");
    }

    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    int selStart = -1, selEnd = -1;
    if (scope == "selection")
    {
        Word.Selection sel = Globals.ThisAddIn.Application.Selection;
        if (sel.Type != Word.WdSelectionType.wdNoSelection)
        {
            selStart = sel.Range.Start;
            selEnd = sel.Range.End;
        }
    }

    var result = new List<int>();
    for (int i = 0; i < paragraphs.Count; i++)
    {
        Word.Paragraph p = paragraphs[i + 1];

        if (scope == "selection")
        {
            if (selStart == -1) continue;
            if (p.Range.Start > selEnd || p.Range.End < selStart) continue;
        }

        if (blockIndexes != null && !blockIndexes.Contains(i)) continue;

        string styleName = p.Range.get_Style().NameLocal;
        bool isHeading = styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
        bool isListItem = p.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;

        if (nodeType == "heading" && !isHeading) continue;
        if (nodeType == "paragraph" && (isHeading || isListItem)) continue;
        if (nodeType == "listItem" && !isListItem) continue;

        if (nodeType == "heading" && headingLevel.HasValue)
        {
            string levelDigits = new string(styleName.Where(char.IsDigit).ToArray());
            if (!int.TryParse(levelDigits, out int actualLevel) || actualLevel != headingLevel.Value) continue;
        }

        if (containsText != null)
        {
            string text = p.Range.Text ?? "";
            bool found = matchCase
                ? text.Contains(containsText)
                : text.IndexOf(containsText, StringComparison.OrdinalIgnoreCase) >= 0;
            if (!found) continue;
        }

        result.Add(i);
    }
    return result;
}
```
(Requires `using System.Linq;` — already present in `WordTools.cs`.)

- [ ] **Step 2: Implement `updateTextStyle`**

Add a `case "updateTextStyle":` branch to `ApplyCommands`'s switch, calling a new method:
```csharp
private static void UpdateTextStyle(JsonElement cmd)
{
    List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
    JsonElement style = cmd.GetProperty("style");
    HashSet<string> fields = new HashSet<string>();
    foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    foreach (int i in indexes)
    {
        Word.Range range = paragraphs[i + 1].Range;
        if (fields.Contains("bold") && style.TryGetProperty("bold", out var bold))
            range.Font.Bold = bold.ValueKind == JsonValueKind.True ? 1 : 0;
        if (fields.Contains("italic") && style.TryGetProperty("italic", out var italic))
            range.Font.Italic = italic.ValueKind == JsonValueKind.True ? 1 : 0;
        if (fields.Contains("underline") && style.TryGetProperty("underline", out var underline))
            range.Font.Underline = underline.ValueKind == JsonValueKind.True ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
        if (fields.Contains("strike") && style.TryGetProperty("strike", out var strike))
            range.Font.StrikeThrough = strike.ValueKind == JsonValueKind.True ? 1 : 0;
        if (fields.Contains("sizeHalfPoints") && style.TryGetProperty("sizeHalfPoints", out var size) && size.ValueKind == JsonValueKind.Number)
            range.Font.Size = (float)(size.GetDouble() / 2.0);
        if (fields.Contains("font") && style.TryGetProperty("font", out var font) && font.ValueKind == JsonValueKind.String)
            range.Font.Name = font.GetString();
        if (fields.Contains("color") && style.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
            range.Font.Color = HexToWdColor(color.GetString());
        if (fields.Contains("baselineOffset") && style.TryGetProperty("baselineOffset", out var baseline) && baseline.ValueKind == JsonValueKind.String)
        {
            string b = baseline.GetString();
            range.Font.Superscript = b == "SUPERSCRIPT" ? 1 : 0;
            range.Font.Subscript = b == "SUBSCRIPT" ? 1 : 0;
        }
        if (fields.Contains("link") && style.TryGetProperty("link", out var link) && link.ValueKind == JsonValueKind.Object)
        {
            string url = link.GetProperty("url").GetString();
            ActiveDoc.Hyperlinks.Add(range, url);
        }
    }
}

private static Word.WdColor HexToWdColor(string hex)
{
    hex = hex.TrimStart('#');
    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
    return (Word.WdColor)System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
}
```
(`highlight` from genoffice's real `style` shape is intentionally NOT ported: Word's `Font.Highlight` only accepts a fixed enum of ~16 named colors (`WdColorIndex`), not arbitrary hex, unlike genoffice's free-color highlight — mark this a known, documented simplification rather than attempting a lossy hex→nearest-enum mapping.)

- [ ] **Step 3: Implement `updateParagraphStyle`**

Add a `case "updateParagraphStyle":` branch, calling:
```csharp
private static void UpdateParagraphStyle(JsonElement cmd)
{
    List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
    JsonElement style = cmd.GetProperty("style");
    HashSet<string> fields = new HashSet<string>();
    foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    foreach (int i in indexes)
    {
        Word.ParagraphFormat fmt = paragraphs[i + 1].Format;
        if (fields.Contains("align") && style.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
        {
            switch (align.GetString())
            {
                case "left": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                case "center": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                case "right": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                case "justify": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
            }
        }
        if (fields.Contains("lineSpacing") && style.TryGetProperty("lineSpacing", out var ls) && ls.ValueKind == JsonValueKind.Number)
            fmt.LineSpacing = (float)ls.GetDouble();
        if (fields.Contains("indentLeft") && style.TryGetProperty("indentLeft", out var il) && il.ValueKind == JsonValueKind.Number)
            fmt.LeftIndent = (float)il.GetDouble();
        if (fields.Contains("indentRight") && style.TryGetProperty("indentRight", out var ir) && ir.ValueKind == JsonValueKind.Number)
            fmt.RightIndent = (float)ir.GetDouble();
        if (fields.Contains("indentFirstLine") && style.TryGetProperty("indentFirstLine", out var ifl) && ifl.ValueKind == JsonValueKind.Number)
            fmt.FirstLineIndent = (float)ifl.GetDouble();
        if (fields.Contains("spaceBefore") && style.TryGetProperty("spaceBefore", out var sb) && sb.ValueKind == JsonValueKind.Number)
            fmt.SpaceBefore = (float)sb.GetDouble();
        if (fields.Contains("spaceAfter") && style.TryGetProperty("spaceAfter", out var sa) && sa.ValueKind == JsonValueKind.Number)
            fmt.SpaceAfter = (float)sa.GetDouble();
        if (fields.Contains("pageBreakBefore") && style.TryGetProperty("pageBreakBefore", out var pbb))
            fmt.PageBreakBefore = pbb.ValueKind == JsonValueKind.True ? 1 : 0;
        if (fields.Contains("shadingFill") && style.TryGetProperty("shadingFill", out var shading) && shading.ValueKind == JsonValueKind.String)
            paragraphs[i + 1].Shading.BackgroundPatternColor = HexToWdColor(shading.GetString());
        if (fields.Contains("borders") && style.TryGetProperty("borders", out var borders))
        {
            bool on = borders.ValueKind == JsonValueKind.True;
            foreach (Word.Border border in paragraphs[i + 1].Borders)
            {
                border.LineStyle = on ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleNone;
            }
        }
    }
}
```
(`borders` is simplified to an all-sides on/off toggle — genoffice's real per-side border shape is richer; document this as an intentional MVP simplification, not a bug, since the research pass didn't need to fully open that structure.)

- [ ] **Step 4: Build and manually verify**

Run from `WordAiAddIn/`: `"C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ../WordAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal` (adjust path if run from repo root). Expected: 0 errors.

Manually verify in real Word: send a message that triggers `apply_commands` with an `updateTextStyle` command targeting `{"containsText": "<some word already in the doc>"}` with `style:{bold:true}, fields:["bold"]` — confirm the matched text becomes bold. Repeat for `updateParagraphStyle` with `{"target":{"blockIndexes":[0]}, "style":{"align":"center"}, "fields":["align"]}` — confirm paragraph 0 centers.

- [ ] **Step 5: Commit**

```bash
git add WordAiAddIn/WordTools.cs
git commit -m "feat(word): add Target-resolution helper, updateTextStyle, updateParagraphStyle"
```

---

### Task 2: `deleteBlocks`, `moveBlocks`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Consumes: `ResolveTargetParagraphs` (Task 1).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `deleteBlocks`**

Add a `case "deleteBlocks":` branch, calling:
```csharp
private static void DeleteBlocksCmd(JsonElement cmd)
{
    List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    if (indexes.Count >= paragraphs.Count)
    {
        // Deleting every paragraph would leave zero - clear content instead,
        // leaving one empty paragraph (mirrors genoffice's own guard).
        ActiveDoc.Content.Text = "";
        return;
    }
    // Delete in descending order so earlier indices don't shift as later ones are removed.
    indexes.Sort();
    indexes.Reverse();
    foreach (int i in indexes)
    {
        paragraphs[i + 1].Range.Delete();
    }
}
```

- [ ] **Step 2: Implement `moveBlocks`**

Add a `case "moveBlocks":` branch, calling:
```csharp
private static void MoveBlocksCmd(JsonElement cmd)
{
    var blockIndexes = new List<int>();
    foreach (JsonElement e in cmd.GetProperty("blockIndexes").EnumerateArray()) blockIndexes.Add(e.GetInt32());
    int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();

    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    int count = paragraphs.Count;
    if (blockIndexes.Any(i => i < 0 || i >= count) || afterBlockIndex < -1 || afterBlockIndex >= count)
    {
        throw new ArgumentException("moveBlocks: index out of range.");
    }
    if (blockIndexes.Contains(afterBlockIndex))
    {
        throw new ArgumentException("moveBlocks: afterBlockIndex cannot be one of the moved blocks.");
    }

    blockIndexes.Sort();
    // Capture each moved paragraph's formatted content (preserves character
    // formatting) before any deletion shifts indices.
    var captured = blockIndexes.Select(i => paragraphs[i + 1].Range.FormattedText).ToList();

    // Delete moved paragraphs in descending order.
    var deleteOrder = new List<int>(blockIndexes);
    deleteOrder.Reverse();
    foreach (int i in deleteOrder)
    {
        paragraphs[i + 1].Range.Delete();
    }

    // Recompute the insertion point: afterBlockIndex shifts down by however
    // many moved blocks were originally BEFORE it.
    int shift = blockIndexes.Count(i => i < afterBlockIndex);
    int adjustedAfter = afterBlockIndex - shift;

    Word.Range insertionPoint = adjustedAfter == -1
        ? ActiveDoc.Range(0, 0)
        : ActiveDoc.Paragraphs[adjustedAfter + 1].Range;
    insertionPoint.Collapse(adjustedAfter == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);

    foreach (Word.Range block in captured)
    {
        insertionPoint.InsertParagraphAfter();
        insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        insertionPoint.FormattedText = block;
        insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
    }
}
```

- [ ] **Step 3: Build and manually verify**

Run the same MSBuild command as Task 1 Step 4. Expected: 0 errors.

Manually verify: `deleteBlocks` with `{"target":{"blockIndexes":[1]}}` on a 3+ paragraph document removes exactly paragraph 1, shifting the rest up. `moveBlocks` with `{"blockIndexes":[0], "afterBlockIndex":2}` on a 3+ paragraph document moves paragraph 0 to after the (new) paragraph 2, preserving its bold/italic formatting if any was present.

- [ ] **Step 4: Commit**

```bash
git add WordAiAddIn/WordTools.cs
git commit -m "feat(word): add apply_commands deleteBlocks and moveBlocks"
```

---

### Task 3: `createParagraphBullets`, `deleteParagraphBullets`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Consumes: `ResolveTargetParagraphs` (Task 1).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement both commands**

Add `case "createParagraphBullets":` and `case "deleteParagraphBullets":` branches, calling:
```csharp
private static void CreateParagraphBullets(JsonElement cmd)
{
    List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
    string preset = cmd.TryGetProperty("bulletPreset", out var bp) && bp.ValueKind == JsonValueKind.String ? bp.GetString() : "BULLET";
    bool numbered = preset.StartsWith("NUMBERED", StringComparison.OrdinalIgnoreCase);

    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    foreach (int i in indexes)
    {
        Word.Range range = paragraphs[i + 1].Range;
        string styleName = range.get_Style().NameLocal;
        if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) continue; // headings are matched but left unchanged, mirrors genoffice
        if (numbered) range.ListFormat.ApplyNumberDefault();
        else range.ListFormat.ApplyBulletDefault();
    }
}

private static void DeleteParagraphBullets(JsonElement cmd)
{
    List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    foreach (int i in indexes)
    {
        Word.Range range = paragraphs[i + 1].Range;
        if (range.ListFormat.ListType == Word.WdListType.wdListNoNumbering) continue; // non-list-item matches silently skipped, mirrors genoffice
        range.ListFormat.RemoveNumbers();
    }
}
```

- [ ] **Step 2: Build and manually verify**

Run the same MSBuild command. Expected: 0 errors.

Manually verify: `createParagraphBullets` with `{"target":{"blockIndexes":[0,1]}}` turns paragraphs 0-1 into a bulleted list; `deleteParagraphBullets` on the same target reverts them to plain paragraphs.

- [ ] **Step 3: Commit**

```bash
git add WordAiAddIn/WordTools.cs
git commit -m "feat(word): add apply_commands createParagraphBullets and deleteParagraphBullets"
```

---

### Task 4: `updateImageProperties`, `insertToc`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Interfaces:**
- Consumes: nothing from Tasks 1-3 (both commands use their own addressing, not `Target`).
- Produces: nothing new for other tasks.

- [ ] **Step 1: Implement `updateImageProperties`**

Add a `case "updateImageProperties":` branch, calling:
```csharp
private static void UpdateImageProperties(JsonElement cmd)
{
    int imageIndex = cmd.GetProperty("imageIndex").GetInt32(); // 0-based index into doc.InlineShapes
    Word.InlineShapes shapes = ActiveDoc.InlineShapes;
    if (imageIndex < 0 || imageIndex >= shapes.Count)
    {
        throw new ArgumentException("updateImageProperties: imageIndex out of range.");
    }
    Word.InlineShape shape = shapes[imageIndex + 1];
    JsonElement properties = cmd.GetProperty("properties");
    HashSet<string> fields = new HashSet<string>();
    foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

    const float pxToPoints = 0.75f; // 96dpi px -> points, matches genoffice's own pixel model
    float? newWidth = null, newHeight = null;
    if (fields.Contains("widthPx") && properties.TryGetProperty("widthPx", out var w) && w.ValueKind == JsonValueKind.Number)
        newWidth = (float)w.GetDouble() * pxToPoints;
    if (fields.Contains("heightPx") && properties.TryGetProperty("heightPx", out var h) && h.ValueKind == JsonValueKind.Number)
        newHeight = (float)h.GetDouble() * pxToPoints;

    if (newWidth.HasValue && !newHeight.HasValue)
    {
        newHeight = shape.Height * (newWidth.Value / shape.Width); // proportional scale from current size
    }
    else if (newHeight.HasValue && !newWidth.HasValue)
    {
        newWidth = shape.Width * (newHeight.Value / shape.Height);
    }
    if (newWidth.HasValue) shape.Width = newWidth.Value;
    if (newHeight.HasValue) shape.Height = newHeight.Value;

    if (fields.Contains("align") && properties.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
    {
        switch (align.GetString())
        {
            case "left": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
            case "center": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
            case "right": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
        }
    }
}
```
(Addresses `InlineShapes` only — floating/anchored shapes are out of scope for this command, matching how `insert_content`/other tools only ever produce inline images. If a future task needs floating-image AI edits, that's a separate addressing scheme, not an extension of this one.)

- [ ] **Step 2: Implement `insertToc`**

Add a `case "insertToc":` branch, calling:
```csharp
private static void InsertTocCmd(JsonElement cmd)
{
    int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();
    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    bool hasHeadings = false;
    foreach (Word.Paragraph p in paragraphs)
    {
        if (p.Range.get_Style().NameLocal.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) { hasHeadings = true; break; }
    }
    if (!hasHeadings)
    {
        throw new InvalidOperationException("insertToc: document has no heading-styled paragraphs to build a table of contents from.");
    }
    if (afterBlockIndex < -1 || afterBlockIndex >= paragraphs.Count)
    {
        throw new ArgumentException("insertToc: afterBlockIndex out of range.");
    }

    Word.Range insertionPoint = afterBlockIndex == -1
        ? ActiveDoc.Range(0, 0)
        : paragraphs[afterBlockIndex + 1].Range;
    insertionPoint.Collapse(afterBlockIndex == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);
    insertionPoint.InsertParagraphAfter();
    insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

    // Word's own native TOC field - auto-scans heading-styled paragraphs and
    // produces real, page-numbered entries directly. This is a more direct,
    // simpler native equivalent than genoffice's own hand-built TOC field-XML
    // workaround (real Word already paginates; genoffice's web renderer doesn't).
    ActiveDoc.TablesOfContents.Add(insertionPoint, UseHeadingStyles: true);
}
```

- [ ] **Step 3: Build and manually verify**

Run the same MSBuild command. Expected: 0 errors.

Manually verify: on a document with at least one `Heading 1`-styled paragraph, `updateImageProperties` with `{"imageIndex":0, "properties":{"widthPx":200}, "fields":["widthPx"]}` (document must have at least one inline image already) resizes it proportionally; `insertToc` with `{"afterBlockIndex":-1}` inserts a real, clickable, page-numbered table of contents at the top of the document, and reopening/scrolling the doc shows Word treats it as a native field (right-click shows "Update Field", "Edit Field", etc. — not plain text).

- [ ] **Step 4: Commit**

```bash
git add WordAiAddIn/WordTools.cs
git commit -m "feat(word): add apply_commands updateImageProperties and insertToc"
```

---

### Task 5: Update the `apply_commands` tool schema/system prompt

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Consumes: nothing new (documentation-only change, no code path affected).
- Produces: nothing new.

- [ ] **Step 1: Update `apply_commands`'s tool description**

In `WordAiAddIn/web-src/entry.ts`'s `ALL_WORD_TOOLS` array, find the `apply_commands` entry's `description` field (currently lists only `set_bold`/`set_italic`/`set_heading`/`find_replace`) and extend it to also document the 8 new kinds, e.g.:
```typescript
    {
      name: 'apply_commands',
      description:
        'Applies a batch of formatting/editing commands. Each command has a "kind": ' +
        '"set_bold"/"set_italic" (fields: startIndex, endIndex, value:boolean), ' +
        '"set_heading" (fields: index, level:0-9, 0=Normal style), ' +
        '"find_replace" (fields: find:string, replace:string, matchCase?:boolean), ' +
        '"updateTextStyle"/"updateParagraphStyle" (fields: target:Target, style:object, fields:string[] - only listed style keys apply), ' +
        '"deleteBlocks" (fields: target:Target), ' +
        '"moveBlocks" (fields: blockIndexes:number[], afterBlockIndex:number, -1=start), ' +
        '"createParagraphBullets"/"deleteParagraphBullets" (fields: target:Target, bulletPreset?:string), ' +
        '"updateImageProperties" (fields: imageIndex:number, properties:object, fields:string[]), ' +
        '"insertToc" (fields: afterBlockIndex:number, -1=start; requires at least one Heading-styled paragraph in the document). ' +
        'Target = {nodeType?:"heading"|"paragraph"|"listItem", headingLevel?:1-6, containsText?:string, matchCase?:boolean, blockIndexes?:number[], scope?:"document"|"selection"} - at least one of nodeType/containsText/blockIndexes required.',
      inputSchema: { type: 'object', properties: { commands: { type: 'array', items: { type: 'object' } } }, required: ['commands'] },
    },
```

- [ ] **Step 2: Typecheck and rebuild**

Run from `WordAiAddIn/`: `npx tsc --noEmit` then the esbuild command from Global Constraints. Expected: 0 errors, successful bundle.

- [ ] **Step 3: Commit**

```bash
git add WordAiAddIn/web-src/entry.ts
git commit -m "docs(word): document the 8 new apply_commands kinds in the tool schema"
```
