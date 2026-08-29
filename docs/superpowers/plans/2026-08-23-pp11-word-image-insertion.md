# PP-11: Word Image Insertion — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-11 (P1).

**Goal:** Give Word a local-file image-insertion tool, following the air-gapped pattern Excel and PowerPoint already use, so Word stops being the only app in this repo that cannot place an image at all.

**Architecture:** The pattern to copy is `AddImageExcel` (`ExcelAiAddIn/ExcelTools.cs:991-1002`): take a `path`, reject `http://`/`https://` with an explicit `NotSupportedException` naming the air-gapped constraint, then `Shapes.AddPicture(path, LinkToFile: msoFalse, SaveWithDocument: msoTrue, left, top, -1, -1)`. The `-1, -1` for width/height means "use the image's natural size", and `SaveWithDocument: msoTrue` embeds the bytes so the document is portable — both are exactly right for Word too.

Word-specific decisions:
- **Inline vs. floating.** Word has two picture collections. `InlineShapes.AddPicture(...)` places a picture in the text flow at a `Range`; `Shapes.AddPicture(...)` creates a floating shape. Inline is the correct default for a document — it flows with the text, moves when paragraphs above it are edited, and cannot land on top of the prose. Floating is offered as an option.
- **Positioning** reuses the same 0-based `afterBlockIndex` convention every other block-addressed Word tool uses, and specifically `RangeAfterBlock` from PP-10 Task 1 if that has landed.
- **Addressability.** `apply_commands`' `updateImageProperties` already addresses images by 0-based index into `doc.InlineShapes` (`WordAiAddIn/WordTools.cs:573-613`). An inline insert is therefore immediately editable by the existing command; a floating insert is not. That asymmetry has to be stated in the schema, or the model will insert floating and then fail to resize.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Word`.

**Soft dependency:** PP-10 (`2026-08-23-pp10-word-rich-content-and-positional-insert.md`) Task 1 produces `RangeAfterBlock`. If PP-10 has not landed, implement that helper here and note that PP-10 should consume it rather than duplicating it.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **Local file paths only.** Remote URLs are rejected with the same message shape Excel uses (`ExcelTools.cs:996`). This is a deployment constraint, not an oversight — do not add a download path, and do not accept `file://` URLs as a workaround (normalize or reject them explicitly).
- The tool is mutating, so it must sit behind the existing editing-mode gate. Confirm it is **not** added to `READ_ONLY_TOOL_NAMES` in `WordAiAddIn/web-src/entry.ts:165`.
- No automated tests for COM executor methods (project convention). Verification is build + Task 4's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: The `add_image` tool

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `case "add_image"` in `WordTools.Execute`'s switch (`WordAiAddIn/WordTools.cs:55-70`) and `private static ToolResult AddImage(JsonElement input)`.

- [ ] **Step 1: Path validation helper**

```csharp
private static string ValidateLocalImagePath(string path)
{
    if (string.IsNullOrWhiteSpace(path))
        throw new ArgumentException("add_image: path is required.");
    if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        throw new NotSupportedException(
            "add_image: remote URLs are not supported in this air-gapped deployment - use a local file path.");
    if (!System.IO.File.Exists(path))
        throw new System.IO.FileNotFoundException("add_image: no file at '" + path + "'.");
    return path;
}
```

The `File.Exists` check is the addition over Excel's version. `AddPicture` on a missing file throws a bare `COMException` with a useless message; a specific error lets the model correct the path on the next turn. Consider back-porting this check into `AddImageExcel` as a one-line follow-up, but do not change Excel's behavior as part of this plan.

- [ ] **Step 2: The handler**

```csharp
private static ToolResult AddImage(JsonElement input)
{
    string path = ValidateLocalImagePath(input.GetProperty("path").GetString());
    bool floating = input.TryGetProperty("floating", out var fl) && fl.ValueKind == JsonValueKind.True;
    int afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var ab) && ab.ValueKind == JsonValueKind.Number
        ? ab.GetInt32() : int.MinValue;   // sentinel: append at end

    Word.Document doc = ActiveDoc;
    Word.Range at;
    if (afterBlockIndex == int.MinValue)
    {
        at = doc.Content;
        at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
    }
    else
    {
        at = RangeAfterBlock(afterBlockIndex);   // PP-10 Task 1
    }

    // ... insert, then size, then report (Steps 3-5)
}
```

`int.MinValue` as the sentinel rather than `-1`, because `-1` already means "start of document" in this file's block-index convention.

- [ ] **Step 3: Insert**

Inline (default): `Word.InlineShape shape = doc.InlineShapes.AddPicture(path, LinkToFile: false, SaveWithDocument: true, Range: at);`

Floating: `Word.Shape shape = doc.Shapes.AddPicture(path, LinkToFile: false, SaveWithDocument: true, Left: ..., Top: ..., Width: -1, Height: -1, Anchor: at);`

`LinkToFile: false` + `SaveWithDocument: true` embeds the bytes — required, since a linked image breaks the moment the document leaves the machine, which is a silent-later-failure in an air-gapped deployment where the source file may be a temp path.

- [ ] **Step 4: Optional sizing**

Accept `widthPoints?` / `heightPoints?`. When exactly one is given, scale the other proportionally from the natural size (read the shape's `Width`/`Height` immediately after insertion, before overwriting). When neither is given, leave natural size. Do **not** silently distort by defaulting the missing dimension to a constant.

- [ ] **Step 5: Report**

Return the image's 0-based `InlineShapes` index (for inline inserts) alongside the final width/height, so the model can immediately drive `apply_commands`' `updateImageProperties` without a re-read. For floating inserts, say explicitly in the output that the image is not addressable by `updateImageProperties`.

- [ ] **Step 6: Register** the `case "add_image": return AddImage(input);` branch in `Execute` (`WordAiAddIn/WordTools.cs:55-70`).

- [ ] **Step 7: Schema**

```ts
{
  name: 'add_image',
  description:
    'Inserts an image from a LOCAL FILE PATH into the document (no URLs - this deployment is air-gapped). ' +
    'Inserts inline in the text flow by default, after the paragraph given by afterBlockIndex (0-based; -1 = start; omit = end of document). ' +
    'Inline images are addressable afterwards by apply_commands/updateImageProperties via their 0-based index, which this tool returns; floating images are not.',
  inputSchema: {
    type: 'object',
    properties: {
      path: { type: 'string' },
      afterBlockIndex: { type: 'number' },
      floating: { type: 'boolean' },
      widthPoints: { type: 'number' },
      heightPoints: { type: 'number' },
    },
    required: ['path'],
  },
}
```

- [ ] **Step 8:** Update the Word skill's `systemPrompt` (`WordAiAddIn/web-src/entry.ts:250-257`) to mention image insertion and the local-path-only rule — otherwise the model will keep telling users it cannot insert images.

**Verification:** `MSBuild WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug`; bundle rebuilds.

---

### Task 2: Alt text (accessibility)

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Accept `altText?: string` and set it after insertion — `shape.AlternativeText` for a floating `Shape`, `shape.AlternativeText` for an `InlineShape` (both expose it; confirm at build time).
- [ ] **Step 2:** Check whether `updateImageProperties` (`WordAiAddIn/WordTools.cs:573-613`) already handles an alt-text field. If it does, reuse the same property name for consistency; if not, note the gap for a follow-up rather than adding it there in this plan.
- [ ] **Step 3:** Recommend alt text in the schema description. Do not make it required — a required accessibility field that the model fills with junk is worse than an absent one.

**Verification:** inserted image shows the alt text in Word's own alt-text pane.

---

### Task 3: Reuse check against the other two apps

**Files:** none modified (assessment)

- [ ] **Step 1:** Compare `ValidateLocalImagePath` with the inline checks in `AddImageExcel` (`ExcelAiAddIn/ExcelTools.cs:993-997`) and PowerPoint's `replace_image`/`insert_web_image`. Confirm the rejection message is worded consistently across all three, so a user (and the model) sees one rule, not three.
- [ ] **Step 2:** If PowerPoint's `insert_web_image` accepts a remote URL despite the air-gapped constraint, that is a separate finding — file it, do not fix it here.
- [ ] **Step 3:** Do not extract a shared helper into `OfficeAi.Shared` for ~10 lines of validation; note the three copies in a comment in each file instead.

---

### Task 4: Manual verification matrix

- [ ] `add_image({path: 'C:\\...\\logo.png'})` → image appears at the end of the document at natural size, embedded (close the document, move the source file, reopen — the image is still there).
- [ ] `add_image({path, afterBlockIndex: 2})` → image sits between paragraphs 3 and 4 and moves when a paragraph above it is deleted (proving inline, not floating).
- [ ] `add_image({path: 'https://example.com/x.png'})` → specific air-gap error; document unchanged.
- [ ] `add_image({path: 'C:\\nope.png'})` → specific file-not-found error naming the path; document unchanged.
- [ ] `add_image({path, widthPoints: 200})` → proportional scaling, no distortion.
- [ ] Returned index feeds `apply_commands`/`updateImageProperties` successfully in the very next turn.
- [ ] `add_image({path, floating: true})` → floating image; the result says it is not addressable by `updateImageProperties`.
- [ ] In Read only mode → the tool is not offered and a direct call is refused by the mode gate.
- [ ] In Track Changes mode → the insertion appears as a tracked revision.
- [ ] A large (>5 MB) image inserts without hanging the pane, or fails with a clear message.
