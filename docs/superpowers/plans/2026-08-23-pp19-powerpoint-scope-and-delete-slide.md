# PP-19: PowerPoint Scripting/Generation Scope Decision + `delete_slide`

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. **Task 1 ships independently and does not wait on the decision in Task 2.**

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-19 (P2, large scope — "confirm appetite before planning implementation").

**Goal:** Ship the small, clearly-in-scope piece (`delete_slide` plus slide-ordering operations) now, and put the large piece (scripting DSL / deck-generation pipeline / automatic QC) in front of a product owner as an explicit scope decision rather than letting it sit as an open-ended "gap".

**Architecture:** The source item bundles two very different things.

- **`delete_slide` is a five-line omission.** `add_slide` exists (`PowerPointAiAddIn/PowerPointTools.cs`, schema at `PowerPointAiAddIn/web-src/entry.ts:243-251`), `delete_element` exists for shapes, and `Presentation.Slides[i].Delete()` is a single COM call. The only reason it is missing is that `docs/tool-surface-todo.md`'s exclusion list groups `delete_slide` in with the deck-generation pipeline — a grouping that made sense for genoffice, whose `delete_slide` was part of `regenerate_slide`'s machinery, but does not describe anything about officeoffice. PP-8 Task 1 Step 4 flags the same inconsistency.
- **The DSL + generation pipeline + QC is ~3,500 lines in genoffice** (a sandboxed AST-walked scripting language, `ask_clarification`/`plan_deck`/`generate_deck`/`regenerate_slide`/`save_style_template`/`list_style_templates`, and a geometric + vision-based QC pass after each generated page). It is explicitly out of scope per this project's existing planning docs. It is not an oversight and should not be planned as one.

Task 1 does the former. Tasks 2-3 handle the latter as a decision, with a middle option that captures most of the practical value at a fraction of the cost.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.PowerPoint`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Slide indices are 0-based at the tool boundary and 1-based in COM. Every existing PowerPoint handler does `Slides[slideIndex + 1]` (e.g. `PowerPointTools.cs:439`, `:451`); keep that convention exactly.
- Deletion is destructive and not undoable from the tool side. `delete_slide` must be gated by the editing mode (it already will be, via `Execute`'s mode gate) and must never accept a "delete all" shorthand.
- No automated tests for COM executor methods (project convention).
- Rebuild bundle + MSBuild after the `entry.ts` change.

---

### Task 1: `delete_slide` and slide ordering

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

**Interfaces:**
- Produces: `case "delete_slide"` / `case "move_slide"` / `case "duplicate_slide"` in `PowerPointTools.Execute`'s switch.

- [ ] **Step 1: `delete_slide`**

```csharp
private static ToolResult DeleteSlide(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    PowerPoint.Presentation pres = ActivePresentation;
    if (slideIndex < 0 || slideIndex >= pres.Slides.Count)
        throw new ArgumentOutOfRangeException("slideIndex",
            "slideIndex must be between 0 and " + (pres.Slides.Count - 1) + ".");
    if (pres.Slides.Count == 1)
        throw new InvalidOperationException(
            "delete_slide: cannot delete the only slide in the presentation.");
    pres.Slides[slideIndex + 1].Delete();
    return new ToolResult
    {
        Output = "Deleted slide " + slideIndex + ". " + pres.Slides.Count + " slide(s) remain; " +
                 "slides after it have shifted down by one index.",
        Mutated = true,
        Summary = "delete_slide",
    };
}
```

The index-shift sentence in the output is load-bearing: a model deleting slides 2, 3, 4 in one run will otherwise delete the wrong ones. Say it in the schema description too.

- [ ] **Step 2: Multi-delete safety.** Do **not** add a `slideIndexes: number[]` parameter. If it is added later it must delete in descending index order; the safer answer is one slide per call, with the shift warning above making the model re-read between deletes. State this decision in a comment.
- [ ] **Step 3: `move_slide`** — `pres.Slides[from + 1].MoveTo(toIndex + 1)`, with the same bounds validation. "Move the summary slide to the end" is a common request with no current tool.
- [ ] **Step 4: `duplicate_slide`** — `pres.Slides[i + 1].Duplicate()` inserts a copy right after. Distinct from `add_slide`, which clones the *layout* and optionally clears text (`entry.ts:243-251`); duplicate keeps the content. Both are useful; document the difference in both descriptions so the model can pick.
- [ ] **Step 5: Register** all three cases in `Execute`'s switch and confirm none is added to the read-only tool set in `PowerPointAiAddIn/web-src/entry.ts`.
- [ ] **Step 6: Schemas**

```ts
{ name: 'delete_slide',
  description: 'Deletes one slide (0-based slideIndex). Slides after it shift DOWN by one index - re-read the deck with get_deck_context before deleting another slide in the same run. Cannot delete the last remaining slide.',
  inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] } },
{ name: 'move_slide',
  description: 'Moves a slide to a new 0-based position; other slides shift accordingly.',
  inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' }, toIndex: { type: 'number' } }, required: ['slideIndex', 'toIndex'] } },
{ name: 'duplicate_slide',
  description: 'Inserts a copy of a slide (content included) directly after it. Use add_slide instead to create a new slide from a slide\'s layout without its content.',
  inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] } },
```

- [ ] **Step 7:** Update the PowerPoint skill's `systemPrompt` to mention slide deletion and reordering.
- [ ] **Step 8:** Update `docs/ai-tool-surface.md`'s PowerPoint section and — per PP-8 Task 1 Step 4 — remove `delete_slide` from the out-of-scope list, since it is now in scope and shipped.

**Verification:** `MSBuild PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug`; the manual matrix in Task 4.

---

### Task 2: Scope decision on the DSL / generation pipeline / QC

**Files:** none. Output is a recorded decision.

- [ ] **Step 1: Present the options**

  **(A) Full port.** genoffice's `execute_slide_script` (sandboxed AST interpreter), the six-tool generation pipeline, and the geometric + vision QC pass. Approximately 3,500 lines in genoffice, and the vision half of QC has no equivalent here — officeoffice has no image-analysis capability, so QC would be geometry-only or would need one added. **This is its own initiative with its own feasibility pass, not a task in a fix-up plan.**

  **(C) Decline, keep the current boundary.** The add-in stays a document-editing assistant, not a deck generator. Zero cost. The user can still build a deck slide by slide through existing tools.

  **(B) Middle option — batch + QC without the DSL and without generation.** Two pieces, each independently useful:
  - a `apply_slide_operations` gateway matching Word's `apply_commands` and Excel's `propose_operations` — a batch of the *existing* PowerPoint tool calls applied in one round trip. This is the actual pain the source item names ("any multi-property, multi-element edit must go tool-by-tool"), and it needs no DSL, no interpreter, and no sandbox: it is the same batching pattern already shipped twice in this repo.
  - a `check_slide_layout` read-only tool reporting geometric problems on one slide: shapes overflowing slide bounds, overlapping shapes, text overflowing its frame (`TextFrame2.TextRange.BoundHeight` vs. the frame's height), and off-grid alignment. Read-only, no vision, ~150 lines, and it gives the model the feedback loop that QC exists to provide.

- [ ] **Step 2: Recommend (B).** It addresses the concrete complaint at roughly 5% of (A)'s cost and does not commit the project to owning a scripting language. (A) should be reconsidered only if "generate me a deck about X" becomes a stated product goal, and then as its own initiative.
- [ ] **Step 3: Record** the decision, its date, and its rationale in `docs/ai-tool-surface.md`'s scope section, replacing the current bare "explicitly out of scope" line with a decision that has reasoning attached.
- [ ] **Step 4:** If (B) is chosen, write it up as two new plan documents (`...-pp19b-powerpoint-batch-operations.md` and `...-pp19c-slide-layout-check.md`) rather than expanding this one. If (A) is chosen, the first deliverable is a feasibility pass, not code.

---

### Task 3: If (B) — scoping notes for the follow-on plans

**Files:** none (input to the follow-on plans)

- [ ] **Step 1 (batch gateway):** Follow Word's `ApplyCommands` shape (`WordAiAddIn/WordTools.cs:210-270`) including PP-12 Task 3's per-command isolation and indexed reporting — do not copy the pre-fix version. The `kind` vocabulary is the existing PowerPoint tool names, so the handler is a `switch` delegating to methods that already exist.
- [ ] **Step 2 (batch gateway):** Note the index-invalidation hazard, which Word and Excel do not have as acutely: `delete_element` and `delete_slide` shift every later index within the same batch. Either resolve all targets up front before applying anything, or document that deletes must come last in descending index order. Resolving up front is better and is the harder half of this work.
- [ ] **Step 3 (layout check):** Report findings as text the model can act on (`"Slide 2: shape 3 ('Revenue') overflows the right edge by 40pt"`), not as a score. Read-only, so it is safe in every editing mode and belongs in `READ_ONLY_TOOL_NAMES`.
- [ ] **Step 4 (layout check):** Get the slide dimensions from `Presentation.PageSetup.SlideWidth`/`SlideHeight` — do not assume 720×540 or 960×540, since both 4:3 and 16:9 decks are common.

---

### Task 4: Manual verification matrix (Task 1)

- [ ] `delete_slide {slideIndex: 1}` on a 3-slide deck → slide 2 gone, 2 slides remain, output states the index shift.
- [ ] `delete_slide` on the only slide → specific error, slide untouched.
- [ ] `delete_slide {slideIndex: 99}` → specific out-of-range error naming the valid range.
- [ ] `move_slide {slideIndex: 0, toIndex: 2}` → the first slide becomes the third.
- [ ] `duplicate_slide {slideIndex: 0}` → a content-identical copy at position 2.
- [ ] `add_slide` still behaves as before (layout clone, optional text clear) — no regression.
- [ ] In Read only mode → none of the three is offered and a direct call is refused.
- [ ] Natural language: "delete the third slide" and "move the summary to the end" each work on the first attempt.
