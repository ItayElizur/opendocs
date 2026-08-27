# Phases 1 + 3 (unified) — Split the Giant Tool Files

> **For agentic workers:** Steps use checkbox (`- [ ]`) syntax. Each Task ends with its own build + test + commit and is independently revertable.

**Goal:** Turn the three oversized `*Tools.cs` files into navigable `partial class` file sets grouped by tool area, with zero logic change. Done when all three add-ins build clean (Debug **and** Release), `dotnet test` still passes 90/90, and no file in the set exceeds ~450 lines.

**Parent plan:** `docs/superpowers/plans/2026-08-27-refactor-proposal.md` (Phases 1 and 3).

**Prerequisite:** Phase 0 — **complete** (`2026-08-27-phase0-test-seam.md`). Its 90 tests are the regression net this phase leans on.

---

## Why these two phases merge into one

The refactor proposal listed them separately:

- **Phase 1** — add `#region`/section-header comments to the three files, grouped by tool area. Plus: archive stale plan docs.
- **Phase 3** — split each file into `partial class` files *along that same grouping*.

Doing them in sequence means **touching every method twice**: once to wrap it in a region, again to move it into a file. And the end state makes the first pass redundant — once `WordTools.Charts.cs` exists, a `#region Charts` inside it is noise. **The file boundary *is* the section marker.**

There is a real argument for doing Phase 1 first anyway: a comment-only diff is trivially reviewable, so you could validate the grouping cheaply before committing to moves. **That argument does not survive contact with this codebase.** The current method ordering is only *roughly* grouped — several areas are split across non-contiguous line ranges (verified below):

| File | Interleaving found |
|---|---|
| `WordTools.cs` | Charts at **186–410 *and* 804–1072** (split by HTML + content helpers); content/blocks at **560–803 *and* 1610–1737**; `apply_commands` at **1807–2438 *and* 2523–end** (`InsertTocCmd` sits after the image helpers) |
| `ExcelTools.cs` | Layout at **1275–1301 *and* 1554–1750**; charts/visuals at **1302–1553 *and* 1751–1797** — the two are interleaved with each other |
| `PowerPointTools.cs` | Essentially contiguous already — only trivial adjustments needed |

Comment-only regions over that ordering would produce `#region Charts (part 1)` … `#region Charts (part 2)`, which documents the mess rather than fixing it. Making the groups contiguous requires moving methods — and once a method is moving, moving it *into its destination file* is the same edit.

**So the unified shape is: one move per method, straight into its new file.** The only genuinely independent Phase 1 item is the docs archiving, which lands here as Task 5.

---

## The mechanical constraint that will bite you

**The three add-in projects use classic (non-SDK) `.csproj` with explicit `<Compile Include>` items.** They do **not** glob `*.cs`. Verified:

```xml
<Compile Include="WordTools.cs">
  <SubType>Code</SubType>
</Compile>
```

Every new partial file must be added to its `.csproj` by hand. A file that exists on disk but is missing from the csproj **is silently not compiled** — and the resulting error is "method does not exist" pointing at the *call site* in a different file, not at the file you forgot. Expect to lose time to this at least once if you don't front-load it.

Use `<DependentUpon>` so the parts nest under the parent in Solution Explorer, matching how `ThisAddIn.Designer.cs` already nests:

```xml
<Compile Include="WordTools.Charts.cs">
  <SubType>Code</SubType>
  <DependentUpon>WordTools.cs</DependentUpon>
</Compile>
```

All three classes are currently `public static class` and must become `public static partial class` — **in every part, including the original file.**

---

## Global Constraints

- **Structure only. Zero logic change.** No renames, no signature changes, no "while I'm here" fixes. If a bug is spotted mid-move, write it down and leave it — a structure commit that also changes behavior is the one you can't bisect later.
- **Do not de-duplicate anything in this phase**, even though the inventory below surfaces real duplication. See "Out of scope" — that is Phase 2's job and it changes more than structure.
- Keep every `// why` comment attached to the code it explains. These are the repo's most valuable asset and the easiest thing to drop in a copy-paste move.
- `LangVersion 7.3` — unchanged; no new syntax.
- After **each** file's split: build all three add-ins **and** run `dotnet test`. A green test run does not prove the add-ins compile; the tests only cover `OfficeAi.Shared`.
- One commit per source file split (three code commits), so a regression bisects to a single file.

---

## How Phase 0's outcome changes this phase

Phase 0 could not test the pure helpers *in place* (they are `private static` inside VSTO projects, unreachable from a test project), so it extracted them to `OfficeAi.Shared` and tested them there — 23 → 90 tests. That extraction **also delivered part of Phase 2's de-duplication** as a side effect: `ColorUtil` (hex color, was 3 copies) and `ShapeTypes` (shape maps, was 2 copies) are already done, along with `TextUtil`, `ToolArgs`, `GeometryUtil`, `JsonUtil`.

**So Phase 2 is materially smaller than the refactor proposal describes.** What is left of it, per the inventory below, is ~185 lines across two files — under 3% of the 6,422 lines this phase relocates.

**That does not change the split layout**, and it does not reorder the phases. Every remaining Phase 2 member has an obvious destination file here (chart maps → `.Charts.cs`, SmartArt layouts → `.SmartArt.cs`); when Phase 2 later removes them, those files simply get slightly smaller and nothing crosses a size threshold.

**Why Phase 2 still comes after, not before:** this phase's entire verification story is *"the member set is identical."* Phase 2 deliberately **changes** the member set — members leave the app assemblies and appear in `OfficeAi.Shared`. Run them together and the one check that makes a 6,400-line move-diff reviewable stops working: a member that vanished is no longer distinguishable from a bug. Doing the mechanical change first also makes Phase 2's own diffs easier — extracting `SmartArtLayoutNames` from a 280-line `WordTools.SmartArt.cs` is far more reviewable than from a 2,555-line file.

> **One Phase 0 rule that does NOT carry over.** Phase 0 established: *never expose an Office interop type as a generic type argument from `OfficeAi.Shared`* (`CS1769`). That constraint is about **crossing an assembly boundary** and does **not** apply here — partial class files are all in the same assembly, so a `Dictionary<string, MsoAutoShapeType>` inside `ExcelTools.Charts.cs` is perfectly legal. Do not "helpfully" convert app-internal enum-valued dictionaries to `int` during the split; that would be an unrequested behavior-adjacent change, and it is not needed.

---

## Out of scope — record, do not fix

The member inventory surfaced **cross-app duplication between Word and PowerPoint** that Phase 0 did not catch:

| Duplicated | Word | Excel | PowerPoint | Status |
|---|---|---|---|---|
| ~~Chart-type maps~~ | ~~186~~ | ~~48~~ | ~~1243~~ | **DONE 2026-08-27** — united into `OfficeAi.Shared.ChartTypes`; Word gained `barStacked` first. See below. |
| `TransientComHResults` + `RetryTransientCom` | 211 / 224 | — | 1263 / 1270 | **Near**-identical — Word's takes a `label` param, PowerPoint's does not |
| `SmartArtLayoutNames` + `ResolveSmartArtLayout` | 1333 / 1344 | — | 1543 / 1561 | Same shape, different layout key sets |

> **Chart types were pulled forward out of Phase 2 (user request, 2026-08-27)** and are already done — two commits, ahead of this phase:
> 1. `feat(word): add barStacked chart type` — Word's map had 7 entries to the others' 8, so Word could not draw a stacked bar chart. Ordering this **first** was deliberate: it made all three maps byte-identical, which turned step 2 into a pure zero-behavior extraction. It also fixed a `read_chart` gap (the reverse type-code lookup reported *"unrecognized chart type code 58"*).
> 2. `refactor(shared): unite the three chart-type maps into OfficeAi.Shared.ChartTypes` — tests assert the exact `xlChartType` codes, because a wrong one here is a *silent* wrong result, and that bug has shipped in this repo before (PowerPoint's `"bar"` mapped to `51`/`xlColumnClustered` instead of `57`, drawing a column chart and reporting success).
>
> **The `WordChartTypeMap` / `ExcelChartTypeMap` / `PptChartTypeMap` line items are therefore gone from all three files** — the layout tables below already reflect this, and the Task 0 baselines were captured *before* it, so re-capture them rather than reusing the numbers in this document if you have not started yet.

**Do not fix the two remaining rows here.** Moving them to `OfficeAi.Shared` changes assembly ownership and forces a signature decision for `RetryTransientCom` — not structure-only, and folding it in would blur a mechanical commit into a semantic one. Leave them for Phase 2.

Also unchanged: `ParagraphIndexResolver`, all COM-bound helpers, every tool schema, every `entry.ts`, every system prompt.

---

## Target file layout

Line counts are estimates from the current member inventory; they will land within ~10%.

### PowerPoint — `PowerPointTools.cs` (1,695 lines → 10 files)

| File | Contents | ~lines |
|---|---|---|
| `PowerPointTools.cs` | mode gate, `Execute` dispatch, `IsMutationAllowed`, `ModeLabel`, `ActivePresentation` | 115 |
| `.Read.cs` | `ShapeText`, `GetDeckContext`, `ReadSlide`, `FindTextPpt`, `ReplaceTextPpt` | 230 |
| `.Elements.cs` | `ResolveShape`, notes helpers, `ApplyAutoDirection`, `ApplyBulletSetting`, `SetElementText`, `SetSlideNotes`, `AlignmentMap`, `SetElementStyle`, `SetElementTransform`, `ZOrderMap`, `SetElementOrder`, `AddTextBox`, `AddShape`, `DeleteElement` | 270 |
| `.Slides.cs` | `AddSlide`, `DeleteSlide`, `MoveSlide`, `DuplicateSlide` | 95 |
| `.LayoutAnim.cs` | `SlideLayoutMap`, `ResolveCustomLayout`, `SetSlideLayout`, `TransitionEffectMap`, `SetSlideTransition`, animation maps, `AddAnimation`, `ReadAnimations`, `EditAnimation` | 280 |
| `.Styling.cs` | `SetElementFill`, `SetElementStroke`, `SetSlideBackground`, `UngroupElement` | 65 |
| `.Tables.cs` | `AddTable`, `ResolveTable`, `EditTableCell`, `EditTableStructure`, `EditTableStyle` | 165 |
| `.Charts.cs` | `PptChartTypeMap`, `TransientComHResults`, `RetryTransientCom`, `AddChartPpt`, `PptLegendPositions`, `EditChartPpt` | 300 |
| `.SmartArt.cs` | `SmartArtLayoutNames`, `ResolveSmartArtLayout`, `AddSmartArt` | 77 |
| `.Images.cs` | `CropImage`, `ReplaceImagePpt`, `SetPictureOpacity` | 76 |

### Excel — `ExcelTools.cs` (2,172 lines → 10 files)

| File | Contents | ~lines |
|---|---|---|
| `ExcelTools.cs` | mode gate, `ExcelErrorTexts`, `Execute` dispatch, `Sheet` | 110 |
| `.Read.cs` | `GetWorkbookContext`, `ReadRange`, `ReadCells`, `SelectRange`, `ReadFormats`, `HAlignName`, `VAlignName` | 145 |
| `.Search.cs` | `ResolveSheetsToSearch`, `NativeFindInSheet`, `FindCells`, `FindReplaceExcel`, `NativeFindReplaceInSheet` | 225 |
| `.SheetFeatures.cs` | `ReadSheetFeatures`, `FindDataTables`, `TracePrecedents`, `TraceDependents` | 180 |
| `.Operations.cs` | `RequiredFields`, `ProposeOperations`, `SetRangeValues` | 185 |
| `.Formatting.cs` | `HAlignMap`, `VAlignMap`, `BorderEdgeMap`, edge arrays, `FormatRange`, `ApplyBorders`, filters, conditional-format helpers + `AddConditionalFormat` | 410 |
| `.Layout.cs` | `InsertDeleteRows/Cols`, `SortRange`, row/col size + hidden, `SetFreeze`, `SetPageSetup`, sheet add/delete/duplicate/hide/move/protect | 225 |
| `.Charts.cs` | `ExcelChartTypeMap`, `ApplyChartCategoryRange`, `AddChart`, `EditChartExcel`, `AddSparkline`, `AddShapeExcel`, `ResolveShapeByName`, `EditShapeExcel`, `DeleteVisual`, `AddImageExcel` | 300 |
| `.Tables.cs` | table ops + `MapPivotAgg`, `AddPivot`, `RefreshPivot` | 190 |
| `.Data.cs` | `SetHyperlink`, `SetNote`, defined-name helpers, `SetDataValidation` | 190 |

### Word — `WordTools.cs` (2,555 lines → 9 files)

| File | Contents | ~lines |
|---|---|---|
| `WordTools.cs` | mode gate, `Execute` dispatch, `ActiveDoc`, `RangeAfterBlock`, `EndOfDocumentRange`, `TransientComHResults`, `RetryTransientCom`, `AddComment` | 230 |
| `.Content.cs` | `ParagraphIndexResolver`, `FindText`, `GetHeadings`, `GetDocumentContext`, `InsertContent`, read-cap consts, `ReadBlockAsHtml`, `ReadBlocks`, `ReplaceBlocks` | 380 |
| `.Html.cs` | `HtmlBlockTags`, `HtmlInlineTags`, `ParseHtmlFragment`, `ValidateHtmlTags`, `WriteInlineNodes`, `InsertHtmlFragment` | 150 |
| `.Commands.cs` | `RequiredFields`, `NonNullFields`, `ApplyCommands`, `ResolveTargetParagraphs`, `SetRunProperty`, `SetHeading`, `FindReplace` | 250 |
| `.Commands.Style.cs` | `HighlightColors`, `KnownTextStyleFields`, `KnownParagraphStyleFields`, `UpdateTextStyle`, `UpdateParagraphStyle`, `BulletPresets`, `ApplyBulletPreset`, `CreateParagraphBullets`, `DeleteParagraphBullets` | 330 |
| `.Commands.Blocks.cs` | `DeleteBlocksCmd`, `MoveBlocksCmd`, `UpdateImageProperties`, `InsertTocCmd` | 165 |
| `.Charts.cs` | `WordChartTypeMap`, `WriteChartData`, `ListChartShapes`, `ReadChart`, `EditChart` | 450 |
| `.Tables.cs` | `ResolveTable`, `AddTable`, `EditTable`, `ReadTable` | 260 |
| `.SmartArt.cs` | `SmartArtLayoutNames`, `ResolveSmartArtLayout`, `ResolveSmartArtGalleryItem`, `AddSmartArt`, `ListSmartArtShapes`, `ReadOneSmartArt`, `ReadSmartArt`, `EditSmartArt` | 280 |
| `.Images.cs` | `ValidateLocalImagePath`, `AddImage` | 85 |

> The three-way `apply_commands` split is the one judgment call here — `.Commands.Style.cs` / `.Commands.Blocks.cs` could reasonably be one 495-line file. Split as shown unless it fights you.

---

### Task 0: Build the verification harness first

**Files:** Create `tools/member-inventory.sh` (or run inline — it does not need to be committed).

A file split produces a diff git renders as hundreds of deletions plus hundreds of additions. Eyeballing that for "did a method get dropped?" does not work. Build the check *before* moving anything.

- [ ] **Step 1: Capture the pre-split member inventory for all three files**

```bash
inventory() {  # $1 = file
  grep -nE '^        (private|public|internal)' "$1" \
    | sed -E 's/^[0-9]+:[[:space:]]*//' | sort
}
inventory WordAiAddIn/WordTools.cs            > /tmp/word.before.txt
inventory ExcelAiAddIn/ExcelTools.cs          > /tmp/excel.before.txt
inventory PowerPointAiAddIn/PowerPointTools.cs > /tmp/ppt.before.txt
wc -l /tmp/*.before.txt
```

Line numbers are stripped and the list is sorted, so the inventory is **order-independent** — exactly what is needed when the whole point of the change is reordering.

- [ ] **Step 2: Confirm the after-check works**

After each file's split, re-run `inventory` across *all* parts of that class and diff:

```bash
cat WordAiAddIn/WordTools*.cs | grep -nE '^        (private|public|internal)' \
  | sed -E 's/^[0-9]+:[[:space:]]*//' | sort > /tmp/word.after.txt
diff /tmp/word.before.txt /tmp/word.after.txt && echo "MEMBER SET IDENTICAL"
```

**An empty diff is the pass condition for each of Tasks 1–3.** Anything else means a member was dropped, duplicated, or had its signature altered mid-move.

- [ ] **Step 3: Record baseline line counts** (`wc -l` on each of the three files) so Task 4 can check the totals reconcile.

**Baselines captured 2026-08-27** — the harness above was run and validated (including a negative control: deleting one member from the "after" list *is* caught by the diff). **Re-captured after the chart-type unification landed**, so these are current:

| Class | Members | Lines | `<Compile>` items in csproj |
|---|---|---|---|
| `WordTools` | 65 | 2,543 | 6 |
| `ExcelTools` | 82 | 2,157 | 6 |
| `PowerPointTools` | 62 | 1,683 | 6 |

(Was 66 / 83 / 63 members before chart types moved out — one map per file.) **Re-capture rather than trusting these** if further Phase 2 work lands before you start; the harness takes seconds and a stale baseline defeats its whole purpose.

After the split, `grep -c '<Compile' *.csproj` should read **14 (Word)**, **15 (Excel)**, **15 (PowerPoint)** — the original 6 plus 8 / 9 / 9 new parts.

---

### Task 1: Split `PowerPointTools.cs` — **do this one first**

PowerPoint's methods are already almost perfectly contiguous, so this is the split with the least reordering. Learn the mechanics — csproj edits, `partial` keyword, the member-set check — on the easy file, not on Word's interleaved one.

- [ ] **Step 1:** Mark the class `public static partial class PowerPointTools` in the original file.
- [ ] **Step 2:** Create the 9 new files per the layout table. Each starts with the same `using` block the original has (trim per file afterward — an unused `using` is a warning, not an error, so this is safe to do in one pass and tidy later), the same `namespace PowerPointAiAddIn`, and `public static partial class PowerPointTools`.
- [ ] **Step 3:** Give each new file a 2–4 line header comment stating what belongs in it. **Do not add `#region` markers** — the file name is the section marker now, and regions inside a 200-line file are noise. (The only files where an internal marker might earn its place are the two that stay >300 lines; add them only if they actually help.)
- [ ] **Step 4:** Move members. Cut from the original, paste into the destination, comments attached. Nothing else changes.
- [ ] **Step 5:** Add all 9 files to `PowerPointAiAddIn.csproj` with `<SubType>Code</SubType>` + `<DependentUpon>PowerPointTools.cs</DependentUpon>`.
- [ ] **Step 6:** Build all three add-ins; run `dotnet test`; run the Task 0 member-set diff. All must pass.
- [ ] **Step 7:** Commit.

```bash
git commit -m "refactor(powerpoint): split PowerPointTools into partial class files by tool area"
```

---

### Task 2: Split `ExcelTools.cs`

Same procedure. Two things specific to this file:

- **`ExcelChartTypeMap` currently sits at line 48**, up in the core region, but belongs with `.Charts.cs`. Moving it is correct and safe (partial class — same class, same accessibility).
- **Layout and charts are interleaved** (layout 1275–1301 *and* 1554–1750; charts 1302–1553 *and* 1751–1797). Take care not to drop the smaller fragment of either.

- [ ] **Step 1:** `public static partial class ExcelTools`.
- [ ] **Step 2–5:** Create the 9 new files, header comments, move members, add to `ExcelAiAddIn.csproj`.
- [ ] **Step 6:** Build all three; `dotnet test`; member-set diff clean.
- [ ] **Step 7:** Commit.

---

### Task 3: Split `WordTools.cs` — **most interleaved, do last**

Three areas are non-contiguous here (charts, content, `apply_commands`). Two specific hazards:

- **`ListChartShapes` and `ListSmartArtShapes` are `internal static`, not `private`** — and they are called from `WordAiAddIn/TaskPaneHost.cs`. The partial-class split keeps them reachable, but **they must stay `internal`**; silently "tidying" either to `private` breaks `TaskPaneHost.cs`. The member-set diff catches this (the accessibility keyword is part of the captured line).
- **`RetryTransientCom` goes in the core file, not `.Charts.cs`**, even though both current callers are chart code. It is general COM-retry infrastructure by name and design; burying it in a chart file means the next non-chart caller either reaches into the wrong file or duplicates it. Judgment call, made deliberately.

- [ ] **Step 1:** `public static partial class WordTools`.
- [ ] **Step 2–5:** Create the 8 new files, header comments, move members, add to `WordAiAddIn.csproj`.
- [ ] **Step 6:** Build all three; `dotnet test`; member-set diff clean.
- [ ] **Step 7:** Commit.

---

### Task 4: Reconcile and record

- [ ] **Step 1:** Check the line totals reconcile — for each class, the summed line count of all parts should be within ~5% of the Task 0 baseline (a little growth is expected: per-file `using` blocks, namespace wrappers, header comments).
- [ ] **Step 2:** Confirm no file in any set exceeds ~450 lines. If one does, split it further now rather than leaving the phase half-done.
- [ ] **Step 3:** Add the dated note to `docs/ai-tool-surface.md` (existing `> **Update YYYY-MM-DD (...)**` convention): what split into what, that no tool schema or behavior changed, and that the cross-app `RetryTransientCom` / `SmartArtLayoutNames` duplication was found and deliberately deferred to Phase 2.
- [ ] **Step 4:** Update `STATUS.md`'s build-commands block: note that each app's tool code is now a `partial class` across several files, and that **new files must be added to the classic `.csproj` by hand**.
- [ ] **Step 5:** Mark Phases 1 and 3 done in `2026-08-27-refactor-proposal.md`, noting they were merged and why.
- [ ] **Step 6:** Commit.

---

### Task 5: Archive stale plan docs (the independent half of Phase 1)

Unrelated to the file split — the one Phase 1 item not subsumed by Phase 3. Safe to do at any point, including first if you want a warm-up commit.

`docs/superpowers/plans/` holds 37 plan files plus 25 verification files. This has already gone wrong once: `tool-surface-todo.md` had to be retired for actively misleading about what was implemented.

- [ ] **Step 1:** Create `docs/superpowers/plans/archive/` and move every completed plan into it. **Archive, do not delete** — these are a genuinely useful audit trail that this session's own investigations leaned on repeatedly.
- [ ] **Step 2:** Leave only currently-relevant plans at the top level (this one, the refactor proposal, and anything not yet implemented).
- [ ] **Step 3:** Add one line at the top of `STATUS.md` naming `docs/ai-tool-surface.md` as the canonical current-state document, and `plans/archive/` as history rather than instructions.
- [ ] **Step 4:** Commit.

---

## Definition of done

- [ ] All three add-ins build clean in **Debug and Release** (Release matters — `deploy/package.ps1` builds Release and only Release signs manifests).
- [ ] `dotnet test` still passes **90/90** — unchanged, since this phase adds no testable logic.
- [ ] **Member-set diff empty for all three classes.** This is the load-bearing check: it is what makes an otherwise unreviewable move-diff trustworthy.
- [ ] No file over ~450 lines.
- [ ] Every new file present in its `.csproj` — cross-check the count: `grep -c "<Compile" *.csproj` should have risen by exactly the number of files added.
- [ ] `git log` shows one commit per split file, so a regression bisects to a single class.
- [ ] No `entry.ts`, tool schema, system prompt, or method body changed anywhere in this phase.

> **On skipping the smoke pass:** unlike Phase 0, this phase changes no method bodies and no COM call sites — a member-set diff plus a clean compile genuinely does cover it. A smoke pass is still worth doing opportunistically, but it is not the gate here that it was there.
