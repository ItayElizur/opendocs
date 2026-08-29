# PP-10: Word Rich Content and Positional Insert — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-10 (P1 — "the single biggest capability gap between Word's document-editing tools and genoffice's docs app").

**Goal:** The model can insert and replace content at an arbitrary position in a Word document with basic rich formatting — bold/italic/underline, headings, and bulleted/numbered lists — preserved through a read → edit → write round trip, instead of today's plain-text, append-only-at-the-end behavior.

**Architecture:** Three current limitations, all in `WordTools.cs`:
- `InsertContent` (`:111-126`) collapses a range to the document end, adds a paragraph, and assigns `range.Text` — append-only, plain text, one paragraph.
- `ReadBlocks` (`:171-191`) emits `[i] <plain text>` per paragraph, discarding all formatting.
- `ReplaceBlocks` (`:193-208`) assigns `range.Text` over a paragraph span — which also **destroys** the formatting of the replaced range, silently.

genoffice's docs app round-trips restricted HTML. Word has a native, purpose-built equivalent that avoids writing an HTML parser: **`Range.InsertXML`** with WordProcessingML, and `Range.WordOpenXML` for reading. That is the highest-fidelity path but is verbose and easy to get subtly wrong, and a malformed fragment throws an opaque COM error.

This plan therefore takes the **restricted-HTML-via-clipboard-free-InsertXML** route in a specific, staged way:

- **Task 1-2** deliver the highest-value, lowest-risk half: *positional* insert and non-destructive replace, still plain text. This alone removes "the AI can only append at the end."
- **Task 3-5** add the rich layer: a small, closed HTML subset (`<b> <i> <u> <h1>-<h3> <ul>/<ol>/<li> <p> <br>`) converted to formatting by walking the fragment and applying Word formatting per run — **not** by a general HTML parser and not by `InsertXML`. Converting a closed tag set by hand is ~150 lines, has no dependency, fails loudly on unknown tags, and produces exactly the same primitives `apply_commands`' `updateTextStyle`/`createParagraphBullets` already drive.
- **Task 4** makes `read_blocks` emit that same subset, closing the round trip.

Staging matters: if scope is cut, Tasks 1-2 still ship a real improvement, and Tasks 3-5 remain a coherent follow-on.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Word`.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **Never** use the clipboard (`Selection.Paste`, `Range.PasteSpecial`) to inject content. It clobbers the user's clipboard and races with anything else in the session.
- Backward compatibility: `insert_content({text})` with no new fields must keep appending at the document end, and `replace_blocks({startIndex, endIndex, text})` must keep working with plain text. Existing prompts and any saved chat history depend on it.
- Reject unknown HTML tags with a specific error naming the tag and listing the supported set. Silently dropping tags is the failure mode this whole plan exists to remove; silently rendering them as literal text is worse.
- Escape nothing on output that was not escaped on input: `read_blocks`' HTML mode must HTML-escape the document's own text (`&`, `<`, `>`) or a document containing `<` will produce a fragment the model cannot round-trip.
- Respect the existing editing-mode gate — `Execute` already wraps the whole dispatch (`WordTools.cs:28-55`); do not bypass it.
- No automated tests for COM executor methods (project convention). Verification is build + Task 6's manual matrix.
- Rebuild bundle + MSBuild after `entry.ts` changes.

---

### Task 1: Positional insert

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `private static Word.Range RangeAfterBlock(int afterBlockIndex)` — resolves an insertion point; consumed by Tasks 3 and by PP-11's image tool.

- [ ] **Step 1: The insertion-point helper**

```csharp
// afterBlockIndex is 0-based over ActiveDoc.Paragraphs, matching every other
// block-addressed tool in this file. -1 means "before the first paragraph",
// matching insertToc/moveBlocks' existing convention.
private static Word.Range RangeAfterBlock(int afterBlockIndex)
{
    Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
    if (afterBlockIndex < -1 || afterBlockIndex > paragraphs.Count - 1)
        throw new ArgumentOutOfRangeException("afterBlockIndex",
            "afterBlockIndex must be between -1 and " + (paragraphs.Count - 1) + ".");
    if (afterBlockIndex == -1)
    {
        Word.Range start = paragraphs[1].Range;
        start.Collapse(Word.WdCollapseDirection.wdCollapseStart);
        return start;
    }
    Word.Range r = paragraphs[afterBlockIndex + 1].Range;
    r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
    return r;
}
```

Explicit bounds validation rather than clamping: silently inserting at the end when the model asked for paragraph 400 of a 30-paragraph document is the silent-wrong-outcome pattern this whole product plan is about.

- [ ] **Step 2: `InsertContent` takes an optional position**

Keep the existing end-of-document path when `afterBlockIndex` is absent (byte-for-byte, so the no-arg case cannot regress). When present, use `RangeAfterBlock` and `InsertParagraphAfter` + set text on the new paragraph.

- [ ] **Step 3: Multi-paragraph text**

Today a `text` containing `\n` produces one paragraph with literal line breaks. Split on `\n` and insert one paragraph per line — the model already writes multi-line content and currently gets a mangled result. Document this in the schema description.

- [ ] **Step 4: Schema**

```ts
inputSchema: {
  type: 'object',
  properties: {
    text: { type: 'string', description: 'Plain text; newlines create separate paragraphs.' },
    afterBlockIndex: { type: 'number', description: '0-based paragraph index to insert after; -1 = start of document; omit = end of document.' },
  },
  required: ['text'],
}
```

- [ ] **Step 5: Report the inserted range** — return the 0-based index range of the newly created paragraphs, so the model can address them in a follow-up command without re-reading the whole document.

**Verification:** build + bundle; inserting at index 2 in a 10-paragraph document places content between paragraphs 3 and 4 and leaves everything else untouched.

---

### Task 2: Non-destructive replace

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Problem:** `range.Text = text` (`WordTools.cs:206`) drops all character and paragraph formatting in the replaced span — a "fix the typo in paragraph 4" request silently strips paragraph 4's heading style, bolding, and list membership. Nothing tells the user.

- [ ] **Step 1:** Capture the first replaced paragraph's style before writing (`range.get_Style()`) and reapply it after, so a single-paragraph replacement keeps its heading/list identity. This is the common case and the cheap 80% fix.
- [ ] **Step 2:** For multi-paragraph replacements where the source paragraphs have differing styles, the correct behavior is genuinely ambiguous. Choose: apply the **first** replaced paragraph's style to all new paragraphs, and say so in the schema description. Do not attempt a per-paragraph style mapping — the counts differ in general.
- [ ] **Step 3:** Add `preserveFormatting?: boolean` (default `true`) so a caller can explicitly ask for the old strip-everything behavior.
- [ ] **Step 4:** Extend the result string to report what was preserved, so it is visible in the transcript (and via PP-3's output view).
- [ ] **Step 5:** Empty `text` still deletes the range, unchanged (documented behavior in the current schema, `WordAiAddIn/web-src/entry.ts:214`).

**Verification:** build; replacing the text of a Heading 2 paragraph leaves it a Heading 2; replacing a bulleted item keeps the bullet.

---

### Task 3: Restricted-HTML insertion

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `private static void InsertHtmlFragment(Word.Range at, string html)` — consumed by `insert_content` and `replace_blocks`.

- [ ] **Step 1: Fix the supported subset, and write it down**

Block: `<p>`, `<h1>`–`<h3>`, `<ul>`/`<ol>` with `<li>`. Inline: `<b>`/`<strong>`, `<i>`/`<em>`, `<u>`, `<br>`. Nothing else — no tables (use a future table tool), no images (PP-11), no styles/classes/attributes, no nesting of lists.

- [ ] **Step 2: Parse with `System.Xml.XmlReader`, not a regex or a hand-rolled scanner**

Wrap the fragment in a single synthetic root and parse it as XML. This gives well-formedness checking for free — a malformed fragment throws before anything is written to the document, rather than half-applying. Note the constraint this creates: the model must send well-formed XHTML (`<br/>`, closed tags). State that explicitly in the schema description, and normalize the two or three common void-element spellings (`<br>` → `<br/>`) before parsing so the most likely mistake does not fail the call.

- [ ] **Step 3: Walk and apply**

Maintain a formatting stack while walking elements. For each block element, insert a paragraph and set its style (`"Heading 1"`–`"Heading 3"`, or Normal); for `<li>`, insert a paragraph and call `ListFormat.ApplyBulletDefault()` / `ApplyNumberDefault()` per the enclosing list type; for inline elements, set `Font.Bold`/`Italic`/`Underline` on the run being written. Every primitive here already exists in this file (`UpdateTextStyle` at `:379-416`, `CreateParagraphBullets` at `:544-559`) — reuse the same COM calls so behavior is consistent with `apply_commands`.

- [ ] **Step 4: Unknown tag → specific error**, naming the tag and listing the supported set. Validate the whole fragment **before** writing anything, so an unsupported tag halfway through cannot leave a partial insert.

- [ ] **Step 5: Wire it in**

`insert_content` and `replace_blocks` each gain `html?: string` as an alternative to `text`. Exactly one of the two must be present; both present is an error, not a silent precedence rule.

- [ ] **Step 6: Schema** — document the exact supported tag list inline in the description. The model cannot discover it any other way, and an unlisted-but-attempted tag is a wasted turn.

**Verification:** build + bundle; inserting `<h2>Summary</h2><p>Some <b>bold</b> text.</p><ul><li>one</li><li>two</li></ul>` produces a real Heading 2, a paragraph with genuine bold, and a real two-item bulleted list.

---

### Task 4: Rich `read_blocks` — closing the round trip

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Add `format?: 'text' | 'html'` (default `'text'`, so existing behavior and existing chat history are unaffected).
- [ ] **Step 2:** In `'html'` mode, emit the same restricted subset Task 3 accepts, still prefixed `[i] ` per block so paragraph indices stay addressable. Derive the block tag from the paragraph's style name (`Heading N` → `<hN>` for N ≤ 3, list membership → `<li>` inside a `<ul>`/`<ol>`, otherwise `<p>`), and walk `Range.Words`/`Range.Characters` runs to emit `<b>`/`<i>`/`<u>` where the font properties change.
- [ ] **Step 3: Performance.** Per-character COM property reads are slow across a large range. Cap `'html'` mode to a smaller span than text mode (suggest 100 paragraphs) and return a specific error above that, directing the caller to page. Measure on a 50-page document before fixing the number.
- [ ] **Step 4: HTML-escape** the document's own text on the way out.
- [ ] **Step 5:** Verify the loop: `read_blocks(html)` → hand the fragment back to `replace_blocks(html)` unchanged → the document is byte-identical in the supported subset (headings, bold/italic/underline, list membership survive; anything outside the subset, e.g. font color, is documented as not surviving).

**Verification:** the round trip above, on a document containing a heading, a bold run, an italic run, and a numbered list.

---

### Task 5: System-prompt and doc updates

**Files:**
- Modify: `WordAiAddIn/web-src/entry.ts`, `docs/ai-tool-surface.md`

- [ ] **Step 1:** Update the Word skill's `systemPrompt` (`WordAiAddIn/web-src/entry.ts:250-257`) — it currently says the assistant can "insert content, read and replace paragraph ranges". Say that content can be inserted at any position and in a restricted HTML subset, and name the subset.
- [ ] **Step 2:** Update `docs/ai-tool-surface.md`'s Word section: the new parameters, the supported tag list, the documented lossiness of the round trip, and the `'html'`-mode paragraph cap.

**Verification:** the doc's Word section matches the shipped schemas.

---

### Task 6: Manual verification matrix

- [ ] `insert_content({text})` with no position → appends at the end, exactly as before.
- [ ] `insert_content({text, afterBlockIndex: 0})` → lands after the first paragraph.
- [ ] `insert_content({text, afterBlockIndex: -1})` → lands before the first paragraph.
- [ ] `insert_content({text, afterBlockIndex: 9999})` → specific out-of-range error; document unchanged.
- [ ] Multi-line `text` → one paragraph per line, not literal line breaks in one paragraph.
- [ ] `replace_blocks` over a Heading 2 → still Heading 2 afterwards.
- [ ] `replace_blocks` with `preserveFormatting: false` → old strip behavior.
- [ ] `insert_content({html})` with the full supported subset → correct native Word structures.
- [ ] `insert_content({html: '<table>...'})` → specific error naming `table`; document unchanged.
- [ ] Malformed fragment (`<b>unclosed`) → specific error; document unchanged (nothing partially written).
- [ ] `read_blocks(format:'html')` → fragment reflects headings/bold/lists; feeding it back through `replace_blocks` reproduces the same document.
- [ ] All of the above in Track Changes mode → edits appear as revisions, not direct writes.
