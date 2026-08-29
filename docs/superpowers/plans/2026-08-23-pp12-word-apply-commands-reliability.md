# PP-12: Word `apply_commands` Reliability — Highlight, Bullet Presets, Batch Isolation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-12 (P1) — three related findings.

**Goal:** Stop `apply_commands` from reporting success for work it did not do. Specifically: implement text highlighting, make `bulletPreset` mean something, and make one malformed command fail only itself instead of aborting the rest of the batch.

**Architecture:** Three independent defects sharing one root cause — an unenforced contract (PP-5).

1. **`highlight`** — `UpdateTextStyle` handles 9 style fields (`WordAiAddIn/WordTools.cs:379-416`): `bold`, `italic`, `underline`, `strike`, `sizeHalfPoints`, `font`, `color`, `baselineOffset`, `link`. `highlight` is absent. Because the command is `fields`-driven (`:383-384`), a request naming `highlight` matches no `if`, applies nothing, and falls out of the loop to `lines.AppendLine(kind + ": ok")` (`:237-238`) — a literal false success. Word's native `Range.HighlightColorIndex` (a `WdColorIndex`, a fixed 16-value palette — *not* an arbitrary RGB) is the correct target.
2. **`bulletPreset`** — `CreateParagraphBullets` (`:544-559`) reduces every preset to a boolean: `preset.StartsWith("NUMBERED")` picks `ApplyNumberDefault()`, everything else `ApplyBulletDefault()`. `DISC_CIRCLE_SQUARE`, `DECIMAL_ALPHA_ROMAN`, and every other well-formed genoffice preset name silently collapses to the same generic bullet.
3. **Batch abort** — `string kind = cmd.GetProperty("kind").GetString();` sits *outside* the per-command `try` (`:216-218`). A command missing `kind` throws `KeyNotFoundException` out of the whole `foreach`, aborting every remaining command, while commands already applied stay applied. The caller gets one generic error that understates what changed.

Finding 3 overlaps with PP-5 Task 4 Step 1. **If PP-5 has landed, Task 3 here is already done — verify and skip it.** If PP-5 has not landed, do Task 3 here; it is a five-line fix that should not wait on a schema project.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Word`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- **The governing rule for this plan: never report `ok` for work that did not happen.** Where a capability genuinely cannot be delivered, the tool must say so with a specific error the model can act on. That is the desired outcome the source item states, and it takes precedence over "make the call succeed".
- Do not change the 4 original command kinds (`set_bold`, `set_italic`, `set_heading`, `find_replace`) — a constraint carried from `2026-08-22-word-tools-completion.md`.
- Highlight is a **palette index**, not RGB. Do not accept a hex color and silently snap it to the nearest palette entry — that is the same silent-wrong-result pattern. Accept named colors, reject anything else.
- No automated tests for COM executor methods (project convention). Verification is build + Task 5's manual matrix.
- Rebuild bundle + MSBuild after any `entry.ts` change.

---

### Task 1: Implement `highlight`

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `private static readonly Dictionary<string, Word.WdColorIndex> HighlightColors`.

- [ ] **Step 1: The palette map**

```csharp
// Word highlighting is a fixed 16-entry palette (WdColorIndex), NOT arbitrary
// RGB - unlike Font.Color, which UpdateTextStyle's "color" field uses. Accept
// only these names; anything else is an error rather than a silent nearest-match.
private static readonly Dictionary<string, Word.WdColorIndex> HighlightColors =
    new Dictionary<string, Word.WdColorIndex>(StringComparer.OrdinalIgnoreCase)
{
    ["none"] = Word.WdColorIndex.wdNoHighlight,
    ["yellow"] = Word.WdColorIndex.wdYellow,
    ["brightGreen"] = Word.WdColorIndex.wdBrightGreen,
    ["turquoise"] = Word.WdColorIndex.wdTurquoise,
    ["pink"] = Word.WdColorIndex.wdPink,
    ["blue"] = Word.WdColorIndex.wdBlue,
    ["red"] = Word.WdColorIndex.wdRed,
    ["darkBlue"] = Word.WdColorIndex.wdDarkBlue,
    ["teal"] = Word.WdColorIndex.wdTeal,
    ["green"] = Word.WdColorIndex.wdGreen,
    ["violet"] = Word.WdColorIndex.wdViolet,
    ["darkRed"] = Word.WdColorIndex.wdDarkRed,
    ["darkYellow"] = Word.WdColorIndex.wdDarkYellow,
    ["gray50"] = Word.WdColorIndex.wdGray50,
    ["gray25"] = Word.WdColorIndex.wdGray25,
    ["black"] = Word.WdColorIndex.wdBlack,
    ["white"] = Word.WdColorIndex.wdWhite,
};
```

Confirm every enum member name against the referenced `Microsoft.Office.Interop.Word` PIA at build time. `ExcelTools.cs:22-26` documents a precedent where two `MsoAutoShapeType` members did not exist in this project's PIA and had to be dropped — apply the same discipline: if a member fails to compile, remove that entry and record why in a comment rather than substituting something else.

- [ ] **Step 2: Handle the field** in `UpdateTextStyle`, alongside the existing nine:

```csharp
if (fields.Contains("highlight") && style.TryGetProperty("highlight", out var highlight) && highlight.ValueKind == JsonValueKind.String)
{
    Word.WdColorIndex idx;
    if (!HighlightColors.TryGetValue(highlight.GetString(), out idx))
        throw new ArgumentException("updateTextStyle: unknown highlight color '" + highlight.GetString() +
                                    "'. Valid: " + string.Join(", ", HighlightColors.Keys) + ".");
    range.HighlightColorIndex = idx;
}
```

Throwing (rather than skipping) is deliberate: the per-command `catch` at `:264-267` converts it into a per-command `ERROR - ...` line, which is exactly the honest reporting this item asks for.

- [ ] **Step 3: Detect unknown style fields generally**

The same false-success hole exists for any misspelled field name. After the nine-plus-one `if` chain, compare `fields` against the set of keys actually handled and throw naming any unrecognized one. This converts a whole class of future silent no-ops into actionable errors, and is the piece that generalizes beyond `highlight`:

```csharp
private static readonly HashSet<string> KnownTextStyleFields = new HashSet<string>
{ "bold", "italic", "underline", "strike", "sizeHalfPoints", "font", "color", "baselineOffset", "link", "highlight" };
```

- [ ] **Step 4:** Do the same for `UpdateParagraphStyle`'s field set (`WordTools.cs:427-473`) — read which keys it actually handles and add the equivalent guard.

- [ ] **Step 5: Schema** — add `highlight` to the `style` properties and to the `fields` enum, with the color names as an `enum`. If PP-5 has landed, edit its `WORD_COMMAND_SCHEMAS` table; otherwise update the description string at `WordAiAddIn/web-src/entry.ts:225-238`.

**Verification:** build; "highlight every mention of 'draft' in yellow" produces real yellow highlighting; an invalid color name produces a specific error listing valid names.

---

### Task 2: Real bullet presets

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts`

- [ ] **Step 1: Decide the preset vocabulary**

genoffice's presets follow the Google Docs naming convention (`BULLET_DISC_CIRCLE_SQUARE`, `NUMBERED_DECIMAL_ALPHA_ROMAN`, …), which encodes the glyph sequence for nesting levels 1/2/3. Word's model is different: `ListFormat.ApplyListTemplateWithLevel(ListTemplate, ...)` against a `ListGallery` (`wdBulletGallery` / `wdNumberGallery` / `wdOutlineNumberGallery`), each holding 7 templates whose per-level glyphs are configurable via `ListLevels[n].NumberFormat` / `.NumberStyle`.

Choose the **explicit, verifiable** mapping: support a fixed set of preset names, each implemented by picking a gallery template and setting level 1-3 glyph styles explicitly, rather than trusting a gallery index to be stable across Office versions/locales. Support at minimum:
`BULLET_DISC_CIRCLE_SQUARE`, `BULLET_DIAMOND_X`, `BULLET_CHECKBOX`, `NUMBERED_DECIMAL`, `NUMBERED_DECIMAL_ALPHA_ROMAN`, `NUMBERED_UPPERALPHA`, `NUMBERED_UPPERROMAN`.

- [ ] **Step 2: Implement** a `private static void ApplyBulletPreset(Word.Range range, string preset)` that maps each name to a `WdListGalleryType` + template index + explicit `ListLevels[1..3].NumberStyle` assignments (`wdListNumberStyleBullet` with a specific `NumberFormat` character for bullets, `wdListNumberStyleArabic`/`LowercaseLetter`/`LowercaseRoman` etc. for numbering).

- [ ] **Step 3: Unknown preset → error**, listing supported names — replacing the current `StartsWith("NUMBERED")` collapse (`:547-548`).

- [ ] **Step 4: Keep the default.** An absent `bulletPreset` keeps today's `ApplyBulletDefault()` behavior, so the common "make this a bulleted list" request is unchanged and cannot regress.

- [ ] **Step 5: Keep the heading skip.** `CreateParagraphBullets` deliberately matches but skips Heading-styled paragraphs (`:553`, mirroring genoffice). Preserve that, and additionally report how many paragraphs were skipped in the result line — currently invisible.

- [ ] **Step 6: Schema** — `bulletPreset` gets an `enum` of the supported names.

- [ ] **Step 7 (scope valve):** If Step 2 proves unreliable across Office versions during manual testing, fall back to supporting only the subset that verifies cleanly and **remove the rest from the enum**. A short honest enum beats a long one that silently degrades — that is the whole point of this item.

**Verification:** each enum value produces a visibly distinct list style in real Word; an unknown preset errors specifically.

---

### Task 3: Per-command isolation in the batch loop

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

**Skip this task if PP-5 (`2026-08-23-pp05-gateway-tool-schemas.md`) Task 4 Step 1 has already landed** — verify by checking whether `kind` is read inside the `try` at `WordTools.cs:216`.

- [ ] **Step 1:** Move `kind` extraction inside the per-command `try` and use `TryGetProperty` with an explicit error:

```csharp
string kind = null;
try
{
    JsonElement kindEl;
    if (!cmd.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
        throw new ArgumentException("Command is missing a string \"kind\" field.");
    kind = kindEl.GetString();
    switch (kind) { /* unchanged */ }
}
catch (Exception ex)
{
    lines.AppendLine((kind ?? "(unknown kind)") + ": ERROR - " + ex.Message);
    anyError = true;
}
```

- [ ] **Step 2: Number the lines.** With partial batches now the norm, `lines.AppendLine` should prefix each with the command's 0-based position (`"[2] updateTextStyle: ok"`), so the model can tell *which* command in its batch failed. Today it only gets the kind, which is ambiguous when a batch contains three `updateTextStyle`s.

- [ ] **Step 3: Summarize.** Prepend a header line such as `"Applied 5 of 7 commands (2 failed)."` so the outcome is legible at a glance in the transcript.

- [ ] **Step 4: State the no-rollback decision explicitly** in a comment above the loop: Word COM offers no batch transaction, a hand-rolled undo would be less reliable than the honest per-command report, and the user retains Word's own Ctrl+Z. This is a decision, not an omission.

- [ ] **Step 5:** Apply the identical change to Excel's `ProposeOperations` loop (`ExcelAiAddIn/ExcelTools.cs:483-486`), which has the same defect — unless PP-5 already did.

**Verification:** a batch of `[valid, missing-kind, valid]` applies commands 1 and 3, reports command 2's specific error with its index, and the summary line reads "Applied 2 of 3".

---

### Task 4: Sweep for other false-success paths

**Files:**
- Modify: `WordAiAddIn/WordTools.cs`

- [ ] **Step 1:** Read every `case` in `ApplyCommands` (`:220-266`) and, for each, ask: can this reach `": ok"` having changed nothing? Known candidates beyond Tasks 1-2: `DeleteParagraphBullets` silently skips non-list paragraphs (`:566`); `MoveBlocksCmd` and `InsertTocCmd` may no-op under specific conditions; a `Target` that matches zero paragraphs makes every target-addressed command a no-op that still reports `ok`.
- [ ] **Step 2: Fix the zero-match case first** — it is the most likely in practice. When `ResolveTargetParagraphs` returns an empty list, report `"kind: no paragraphs matched target"` instead of `"ok"`, and do not set `anyMutated`. That single change tells the user "your instruction matched nothing" instead of "done!".
- [ ] **Step 3:** For each remaining no-op path, report a count (`"createParagraphBullets: 4 applied, 2 headings skipped"`) rather than a bare `ok`.
- [ ] **Step 4:** Do not set `Mutated = true` for a command that changed nothing — it drives the "✓ Applied" tag in the UI (`shared/chat-ui/chat-ui.ts:352-357`), which is currently shown for no-op batches.

**Verification:** a `deleteBlocks` with a `containsText` that matches nothing reports "no paragraphs matched", shows no "Applied" tag, and does not claim success.

---

### Task 5: Manual verification matrix

- [ ] `updateTextStyle` with `fields: ['highlight']`, `style: {highlight: 'yellow'}` → real yellow highlight.
- [ ] Same with `highlight: 'none'` → highlight removed.
- [ ] Same with `highlight: '#FFFF00'` → specific error (hex is not a palette name), no change.
- [ ] `fields: ['nonsenseField']` → specific error naming the field.
- [ ] Each supported `bulletPreset` → visibly distinct list formatting.
- [ ] Unknown `bulletPreset` → specific error listing valid presets.
- [ ] `createParagraphBullets` over a range including headings → result reports the skipped count.
- [ ] Batch `[valid, missing-kind, valid]` → both valid commands applied, indexed error for the middle one, accurate summary line.
- [ ] Any target-addressed command whose target matches nothing → "no paragraphs matched", no "Applied" tag.
- [ ] Everything above in Track Changes mode → each change appears as a tracked revision.
