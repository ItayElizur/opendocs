# Docs Sync + Repo-Wide Comment Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Merge `main` into `docs/reorganize-docs` (PR #22), bring `docs/ai-tool-surface.md` fully current against everything that's shipped since this branch's last sync (~140 commits, mostly Outlook features), then sweep every hand-written `.cs`/`.ts` file in the repo to remove superfluous comments, relocate long-but-relevant ones into companion `.md` log files, and leave only short function-level docs plus inline comments where the code is genuinely hard to follow.

**Architecture:** Two independent phases. Phase 1 (Task 1) is a single docs-reconciliation task done inline by whoever has full context on the merge conflict. Phase 2 (Tasks 2-7) is a mechanical-but-judgment-heavy sweep split one task per project (`OfficeAi.Shared`, `WordAiAddIn`, `ExcelAiAddIn`, `PowerPointAiAddIn`, `OutlookAiAddIn`, `shared/*`), each fully independent of the others and safe to run in parallel via subagent-driven-development.

**Tech Stack:** C# (.NET Framework 4.8, VSTO), TypeScript (esbuild-bundled), MSBuild, `tsc`.

**Spec:** No separate spec doc — the spec is the user's own instructions, captured verbatim in Global Constraints below.

## Global Constraints

- **Workspace:** all work happens in the existing worktree `C:\Dev\opendocs-worktrees\docs-reorg`, branch `docs/reorganize-docs` (tracks PR #22). Do not create a new worktree.
- **Comment triage rule** (verbatim from the user, this is the spec for every comment-cleanup task):
  1. Superfluous comments (add no information beyond what the code already states) — **delete**.
  2. Relevant but overlong comments (real historical/rationale content, just too long to live inline) — **move verbatim** to a companion log file named `<source-file-name>.md` (e.g. `WordTools.cs` → `WordTools.cs.md`), sitting next to the source file. Do not summarize or paraphrase when moving — copy the comment text as-is into the `.md`, under a heading naming the symbol/line it came from, so no information is lost.
  3. Per-function documentation — keep, but trimmed to a short (1-3 line) summary of what the function does. If an existing header comment mixes "what it does" with "why/history," keep the "what" inline (short) and move the "why/history" to the `.md` per rule 2.
  4. Mid-function inline comments — keep **only** if the code they annotate is genuinely non-obvious (COM quirks, off-by-one tricks, a workaround for a specific library bug, a subtle invariant). If a mid-function comment just narrates what the next line obviously does, delete it (rule 1). If it's long rationale for a non-obvious choice, trim to a short note and move the full version to the `.md` (rule 2).
- **Excluded files** (generated/boilerplate, never hand-comment-reviewed): `*.Designer.cs`, `Properties/AssemblyInfo.cs`.
- **No behavior changes.** These are comment-only edits. Every file must still compile (C#) or typecheck (TypeScript) after its cleanup, with an identical AST/logic diff (only comment tokens added/removed/moved).
- **Docs must reflect verified current state**, not old prose carried forward. Where `docs/ai-tool-surface.md`'s two merge sides disagree, check the actual current source (`*/web-src/entry.ts` tool schemas, `*Tools*.cs` implementations) rather than assuming either side is right.

---

### Task 1: Resolve the `main` merge conflict in `docs/ai-tool-surface.md`, bring the whole doc current

**Files:**
- Modify: `docs/ai-tool-surface.md` (resolve 6 conflict regions, then a full current-state pass over the Outlook section)
- Check (no change expected, but verify): `CLAUDE.md`, `README.md`, `docs/ai-tool-surface-changelog.md`
- Verify against: `WordAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts`, `OutlookAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, and each app's `*Tools*.cs`

**Interfaces:** None — this is a standalone docs task, nothing downstream depends on its output.

**Context for whoever executes this:** `git merge origin/main` is already in progress in the worktree (started, not yet committed) and has left 6 conflict regions in `docs/ai-tool-surface.md` at (as of this writing) lines 661-676, 680-711, 824-972, 1027-1073, 1083-1151, 1168-1178 (re-run `grep -n "^<<<<<<<\|^=======\|^>>>>>>>" docs/ai-tool-surface.md` — line numbers will have shifted once earlier conflicts are resolved). Already determined by reading each region:
- The **Word** conflict (661-711) and **PowerPoint** conflict (824-972): `HEAD`'s side is the already-restructured, current-state-first version (this branch's own PR #22 work, last synced through PR #25) and is still accurate — nothing in the commits merged since then (`3d89f7c..origin/main`) added a new top-level Word or PowerPoint tool, only tweaked existing tool behavior (e.g. PR #33's `edit_table`/`edit_table_structure` row-0-isn't-a-header clarification, and the undo-granularity fix, neither of which changes the tool *list*). Resolution for these two regions: **keep `HEAD`, discard `origin/main`'s stale side** — but see Step 2 below for the one small content update this still needs.
- The **Outlook** conflicts (1027-1073, 1083-1151, 1168-1178, all within the `## Outlook` section spanning roughly lines 1001-1598) are the opposite: `origin/main`'s side has newer facts (e.g. `draft_edit_event`, `draft_cancel_event`, `draft_respond_meeting`, `tentative_meeting`, "seven auto-send tools" vs. `HEAD`'s "five") because the merged-in commit range contains roughly 15 Outlook feature PRs' worth of work (recurring events/series, shared calendars, meeting response tools, cancel/edit event, a custom undo/redo stack, `set_event_availability`, and more — see `git log 3d89f7c..origin/main --oneline` for the full list) that landed after this branch's last sync and were never folded into the restructured doc. `origin/main`'s prose is also still in the **old dated-block style** this PR eliminates everywhere else (e.g. "four tiers, added 2026-09-19"). Resolution: the Outlook section needs a genuine rewrite, not a pick-one-side resolution.

- [ ] **Step 1: Resolve the Word and PowerPoint conflicts (straightforward)**

For both the Word region (661-711) and PowerPoint region (824-972), delete the `<<<<<<< HEAD` / `=======` / `>>>>>>> origin/main` markers and the `origin/main` content in between, keeping only `HEAD`'s content. Then, in the Word table's `edit_table` row, add one clause reflecting PR #33's clarification (currently missing from `HEAD`'s table, which predates that PR):

Find the `edit_table` row in the restructured Word table and append to its Notes column: `Row/col index 0 is just the first physical row/column, including a header row - there's no separate header concept in the index space (PR review, 2026-10-02).`

Do the same check for PowerPoint's `edit_table_structure`/`edit_table_cell` rows (same clarification, PR #33).

- [ ] **Step 2: Rewrite the Outlook section against current source, not either merge side**

Read the full `## Outlook` section as it stands pre-conflict-resolution on **both** sides (`git show HEAD:docs/ai-tool-surface.md` and `git show origin/main:docs/ai-tool-surface.md`, both scoped to the Outlook section) to see everything either side already documented. Then cross-check against the actual current tool surface:

```bash
grep -n "name: '" OutlookAiAddIn/web-src/entry.ts
```

This lists every real current Outlook tool name — use it as the ground truth for what belongs in the rewritten section's tables, not either side's prose. For each tool found in `entry.ts` that isn't yet described in `HEAD`'s version of the section (expect: `cancel_event`, `edit_event`, `draft_edit_event`, `draft_cancel_event`, `draft_accept_meeting`, `draft_decline_meeting`, `draft_tentative_meeting`, `tentative_meeting`, `set_event_availability`, `undo_last_action`, `redo_last_action`, recurrence-related fields on `create_event`/`draft_event`, and the `mailbox`/shared-calendar parameter on `list_events`/`get_event`), add a row or note using the same by-area table structure `HEAD`'s Word/PowerPoint sections already use (Reading / Mail actions / Calendar / Tasks, etc. - follow whatever grouping the existing Outlook section head matter already started).

Resolve all three Outlook conflict markers by writing the merged section directly (not by picking a side) in this current-state-first style - no "added YYYY-MM-DD" prose inline; if a fact is purely historical (e.g. "editing modes used to only have two tiers"), move that sentence to `docs/ai-tool-surface-changelog.md` under a new dated entry instead of leaving it inline.

- [ ] **Step 3: Confirm no conflict markers remain**

```bash
grep -n "^<<<<<<<\|^=======\|^>>>>>>>" docs/ai-tool-surface.md
```

Expected: no output.

- [ ] **Step 4: Spot-check CLAUDE.md and README.md for staleness**

Read both files in full. They're short (per PR #22's description) and unlikely to need changes, but confirm: no tool count, file path, or specific claim they make has been invalidated by anything in `3d89f7c..origin/main`. If something is stale, fix it; otherwise leave untouched.

- [ ] **Step 5: Commit the merge**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
git add docs/ai-tool-surface.md docs/ai-tool-surface-changelog.md CLAUDE.md README.md
git commit -m "docs: merge main, bring ai-tool-surface.md current through PR #33"
```

(If Step 4 made no changes to `CLAUDE.md`/`README.md`, drop them from the `git add`.)

---

### Task 2: Comment cleanup — `OfficeAi.Shared`

**Files (29, apply the triage procedure to each):**
`OfficeAi.Shared/ActionHistory.cs`, `OfficeAi.Shared/AttachmentText/AttachmentTextExtractor.cs`, `OfficeAi.Shared/AttachmentText/OpenXmlReader.cs`, `OfficeAi.Shared/AttachmentText/PlainTextReader.cs`, `OfficeAi.Shared/ChartTypes.cs`, `OfficeAi.Shared/ChatStore.cs`, `OfficeAi.Shared/ColorUtil.cs`, `OfficeAi.Shared/ComRetry.cs`, `OfficeAi.Shared/ContactSearchFormat.cs`, `OfficeAi.Shared/DebugLog.cs`, `OfficeAi.Shared/DocSettingsStore.cs`, `OfficeAi.Shared/EwsAutodiscoverXml.cs`, `OfficeAi.Shared/GeometryUtil.cs`, `OfficeAi.Shared/JsonUtil.cs`, `OfficeAi.Shared/MeetingSlots.cs`, `OfficeAi.Shared/OfficeLanguage.cs`, `OfficeAi.Shared/OfficeTheme.cs`, `OfficeAi.Shared/OutlookDasl.cs`, `OfficeAi.Shared/PaneHostBase.cs`, `OfficeAi.Shared/RecurrenceSpec.cs`, `OfficeAi.Shared/RecurrenceValidator.cs`, `OfficeAi.Shared/RibbonBase.cs`, `OfficeAi.Shared/ShapeTypes.cs`, `OfficeAi.Shared/SharedCalendarEventFormat.cs`, `OfficeAi.Shared/SmartArtLayouts.cs`, `OfficeAi.Shared/TextUtil.cs`, `OfficeAi.Shared/ToolArgs.cs`, `OfficeAi.Shared/ToolProtocol.cs`, `OfficeAi.Shared/WebViewBridgeHost.cs`

**Interfaces:** None — purely comment edits, no public signatures change.

- [ ] **Step 1: Apply the triage procedure to every file in the list**

For each file: read it in full, then for every comment block (`//` or `/** */`):
- If it adds nothing beyond what the adjacent code already makes obvious -> delete it.
- If it's a function/class header mixing a short "what" with a long "why/history" -> keep a 1-3 line "what" summary in place, move the full original text verbatim into `<SameFileName>.md` (e.g. `ActionHistory.cs.md`) under a `## <method or type name>` heading, in the same relative directory as the source file (so `OfficeAi.Shared/AttachmentText/OpenXmlReader.cs`'s log file is `OfficeAi.Shared/AttachmentText/OpenXmlReader.cs.md`, not flattened into the project root).
- If it's a long comment with no short "what" to extract (pure rationale/history, e.g. explaining why an approach was chosen over an alternative) -> move the whole thing to the `.md` file, leave nothing inline.
- If it's a short mid-function comment explaining genuinely non-obvious code (a COM quirk, a workaround, a subtle invariant) -> keep as-is.
- If it's a short mid-function comment just narrating an obvious line -> delete it.

Only create a `<file>.md` if at least one comment from that file actually needs relocating - don't create empty log files.

- [ ] **Step 2: Build to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" OfficeAi.Shared/OfficeAi.Shared.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Expected: builds with no errors (this project has no behavior-affecting changes, only comments, so a clean build is the whole verification bar here).

- [ ] **Step 3: Spot-check for information loss**

For 3-4 of the files that had the longest comments removed, diff old vs. new and confirm every deleted/shortened comment's substantive content (if any) actually exists verbatim in its `.md` file:

```bash
git diff --stat OfficeAi.Shared/
```

- [ ] **Step 4: Commit**

```bash
git add OfficeAi.Shared/
git commit -m "docs(shared): move long rationale comments to companion .md files, trim the rest"
```

---

### Task 3: Comment cleanup — `WordAiAddIn`

**Files (14, excludes `WordAiAddIn/ThisAddIn.Designer.cs` and `WordAiAddIn/Properties/AssemblyInfo.cs` per Global Constraints):**
`WordAiAddIn/Ribbon.cs`, `WordAiAddIn/TaskPaneHost.cs`, `WordAiAddIn/ThisAddIn.cs`, `WordAiAddIn/WordTools.Charts.cs`, `WordAiAddIn/WordTools.Commands.Blocks.cs`, `WordAiAddIn/WordTools.Commands.Style.cs`, `WordAiAddIn/WordTools.Commands.cs`, `WordAiAddIn/WordTools.Content.cs`, `WordAiAddIn/WordTools.History.cs`, `WordAiAddIn/WordTools.Html.cs`, `WordAiAddIn/WordTools.Images.cs`, `WordAiAddIn/WordTools.SmartArt.cs`, `WordAiAddIn/WordTools.Tables.cs`, `WordAiAddIn/WordTools.cs`

**Interfaces:** None.

- [ ] **Step 1: Apply the comment triage rule from Global Constraints to every file in this list**

Same four rules (delete / move-with-short-summary / move-whole / keep-short-if-genuinely-hard). Concrete worked example already in this codebase, for calibration: `WordAiAddIn/WordTools.History.cs`'s `UndoLastAction`/`RedoLastAction` methods currently carry a long comment block explaining why `Application.Undo()`/`Redo()` don't exist and how reflection confirmed the real `Document.Undo()`/`Redo()` methods - that's exactly a "move whole thing to `WordTools.History.cs.md`, leave a 1-2 line summary inline" case (the summary: "Calls `Document.Undo()`/`Redo()` directly - `Application` has no such method, see WordTools.History.cs.md.").

- [ ] **Step 2: Build to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Expected: builds with no errors.

- [ ] **Step 3: Commit**

```bash
git add WordAiAddIn/
git commit -m "docs(word): move long rationale comments to companion .md files, trim the rest"
```

---

### Task 4: Comment cleanup — `ExcelAiAddIn`

**Files (13, excludes `ExcelAiAddIn/ThisAddIn.Designer.cs` and `ExcelAiAddIn/Properties/AssemblyInfo.cs`):**
`ExcelAiAddIn/ExcelTools.Charts.cs`, `ExcelAiAddIn/ExcelTools.Data.cs`, `ExcelAiAddIn/ExcelTools.Formatting.cs`, `ExcelAiAddIn/ExcelTools.Layout.cs`, `ExcelAiAddIn/ExcelTools.Operations.cs`, `ExcelAiAddIn/ExcelTools.Read.cs`, `ExcelAiAddIn/ExcelTools.Search.cs`, `ExcelAiAddIn/ExcelTools.SheetFeatures.cs`, `ExcelAiAddIn/ExcelTools.Tables.cs`, `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/Ribbon.cs`, `ExcelAiAddIn/TaskPaneHost.cs`, `ExcelAiAddIn/ThisAddIn.cs`

**Interfaces:** None.

- [ ] **Step 1: Apply the comment triage rule from Global Constraints to every file in this list**

- [ ] **Step 2: Build to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Expected: builds with no errors.

- [ ] **Step 3: Commit**

```bash
git add ExcelAiAddIn/
git commit -m "docs(excel): move long rationale comments to companion .md files, trim the rest"
```

---

### Task 5: Comment cleanup — `PowerPointAiAddIn`

**Files (17, excludes `PowerPointAiAddIn/ThisAddIn.Designer.cs` and `PowerPointAiAddIn/Properties/AssemblyInfo.cs`):**
`PowerPointAiAddIn/PowerPointTools.Charts.cs`, `PowerPointAiAddIn/PowerPointTools.CrossSlide.cs`, `PowerPointAiAddIn/PowerPointTools.Elements.cs`, `PowerPointAiAddIn/PowerPointTools.FormatPainter.cs`, `PowerPointAiAddIn/PowerPointTools.History.cs`, `PowerPointAiAddIn/PowerPointTools.Images.cs`, `PowerPointAiAddIn/PowerPointTools.LayoutAnim.cs`, `PowerPointAiAddIn/PowerPointTools.Master.cs`, `PowerPointAiAddIn/PowerPointTools.Read.cs`, `PowerPointAiAddIn/PowerPointTools.Slides.cs`, `PowerPointAiAddIn/PowerPointTools.SmartArt.cs`, `PowerPointAiAddIn/PowerPointTools.Styling.cs`, `PowerPointAiAddIn/PowerPointTools.Tables.cs`, `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/Ribbon.cs`, `PowerPointAiAddIn/TaskPaneHost.cs`, `PowerPointAiAddIn/ThisAddIn.cs`

**Interfaces:** None.

- [ ] **Step 1: Apply the comment triage rule from Global Constraints to every file in this list**

Concrete worked example for calibration: `PowerPointAiAddIn/PowerPointTools.History.cs`'s top-of-file comment (the ~35-line essay on `CommandBars.ExecuteMso` being the only undo/redo mechanism, referencing the `RemoveMasterElement` incident) is exactly a "move whole thing to `PowerPointTools.History.cs.md`" case - replace with a 2-3 line summary of what the two methods do and how (ExecuteMso dispatch, GetEnabledMso pre-check), pointing to the `.md` for the full rationale.

- [ ] **Step 2: Build to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Expected: builds with no errors.

- [ ] **Step 3: Commit**

```bash
git add PowerPointAiAddIn/
git commit -m "docs(powerpoint): move long rationale comments to companion .md files, trim the rest"
```

---

### Task 6: Comment cleanup — `OutlookAiAddIn`

**Files (15, excludes `OutlookAiAddIn/ThisAddIn.Designer.cs` and `OutlookAiAddIn/Properties/AssemblyInfo.cs`):**
`OutlookAiAddIn/OutlookEws.cs`, `OutlookAiAddIn/OutlookTools.Attachments.cs`, `OutlookAiAddIn/OutlookTools.Calendar.cs`, `OutlookAiAddIn/OutlookTools.Categories.cs`, `OutlookAiAddIn/OutlookTools.Compose.cs`, `OutlookAiAddIn/OutlookTools.Ews.cs`, `OutlookAiAddIn/OutlookTools.Folders.cs`, `OutlookAiAddIn/OutlookTools.Mail.cs`, `OutlookAiAddIn/OutlookTools.Search.cs`, `OutlookAiAddIn/OutlookTools.Tasks.cs`, `OutlookAiAddIn/OutlookTools.Undo.cs`, `OutlookAiAddIn/OutlookTools.cs`, `OutlookAiAddIn/Ribbon.cs`, `OutlookAiAddIn/TaskPaneHost.cs`, `OutlookAiAddIn/ThisAddIn.cs`

This is the largest and most heavily-commented app (per the file-system survey: ~1055 comment-only lines, the most of any project) - budget the most time here.

**Interfaces:** None.

- [ ] **Step 1: Apply the comment triage rule from Global Constraints to every file in this list**

- [ ] **Step 2: Build to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Expected: builds with no errors.

- [ ] **Step 3: Commit**

```bash
git add OutlookAiAddIn/
git commit -m "docs(outlook): move long rationale comments to companion .md files, trim the rest"
```

---

### Task 7: Comment cleanup — `shared/web-src` and `shared/chat-ui` (TypeScript)

**Files (19):**
`shared/chat-ui/chat-ui.test.ts`, `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/vitest.config.ts`, `shared/web-src/agent-core/id.ts`, `shared/web-src/agent-core/index.ts`, `shared/web-src/agent-core/loop.ts`, `shared/web-src/agent-core/skill.ts`, `shared/web-src/agent-core/types.ts`, `shared/web-src/ai-provider/fetch.ts`, `shared/web-src/ai-provider/http-error.ts`, `shared/web-src/ai-provider/index.ts`, `shared/web-src/ai-provider/providers.ts`, `shared/web-src/ai-provider/stream.ts`, `shared/web-src/ai-provider/types.ts`, `shared/web-src/ai-provider/watchdog.ts`, `shared/web-src/app-shell/bootstrap.ts`, `shared/web-src/app-shell/bridge.ts`, `shared/web-src/app-shell/index.ts`, `shared/web-src/app-shell/settings.ts`

**Interfaces:** None - comment-only edits, no exported signatures change.

- [ ] **Step 1: Apply the comment triage rule from Global Constraints to every file in this list**

Concrete worked example for calibration: `shared/chat-ui/chat-ui.ts`'s `trackMouseDownOutside` helper currently carries a long comment explaining the mousedown/mouseup/click browser-ancestor quirk and the PR-review history of why it's a shared helper - the quirk explanation is worth a short inline note (genuinely non-obvious browser behavior, keep per rule 4), but the "PR review, 2026-10-02" authorship/history framing belongs in `chat-ui.ts.md`, not inline.

- [ ] **Step 2: Typecheck to confirm no behavior change**

```bash
cd /c/Dev/opendocs
./shared/chat-ui/node_modules/.bin/tsc -p WordAiAddIn/tsconfig.json
./shared/chat-ui/node_modules/.bin/tsc -p ExcelAiAddIn/tsconfig.json
./shared/chat-ui/node_modules/.bin/tsc -p PowerPointAiAddIn/tsconfig.json
./shared/chat-ui/node_modules/.bin/tsc -p OutlookAiAddIn/tsconfig.json
```

(Run from the main repo checkout, not the worktree, since that's where `node_modules` lives - or symlink/copy if the worktree needs its own. Expected: no output from any of the four, meaning a clean typecheck.)

- [ ] **Step 3: Run the existing test suite to confirm no behavior change**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg/shared/chat-ui
npx vitest run
```

Expected: all 90 tests still pass (comment-only changes shouldn't affect test outcomes at all).

- [ ] **Step 4: Commit**

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
git add shared/
git commit -m "docs(shared-ts): move long rationale comments to companion .md files, trim the rest"
```

---

## Final step (after all 7 tasks): push

```bash
cd /c/Dev/opendocs-worktrees/docs-reorg
git push
```

This updates PR #22 with the merge + full comment-cleanup sweep. Do not merge the PR itself without the user's explicit go-ahead.
