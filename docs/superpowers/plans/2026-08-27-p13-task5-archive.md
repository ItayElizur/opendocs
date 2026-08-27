# P13 · Task 5 — Archive completed plan docs

**Part of:** `2026-08-27-phase1-3-file-split.md` (Phases 1+3).
**Prerequisites:** none. **This task is fully independent of Tasks 0–4** — it touches no source code and can be done before, during, or after the splits. It is the one piece of the original Phase 1 that the file split does not subsume.

**Time:** ~15 minutes. **Risk:** low — moving markdown files.

---

## Why this exists

`docs/superpowers/plans/` holds **45 markdown files**, plus 25 more in `docs/superpowers/verification/`. Almost all describe work that shipped months ago. A reader cannot tell which describe the current state and which are history.

This has already gone wrong once: `tool-surface-todo.md` had to be formally retired (see `2026-08-23-pp08-retire-stale-todo.md`) because it was actively misleading about what was implemented — it listed many things as unimplemented that had in fact shipped.

**Archive, do not delete.** These files are a genuinely useful audit trail; investigations during this refactor leaned on them repeatedly to answer "why is this like this?". The goal is to separate *history* from *instructions*, not to discard the history.

---

## Step 1 — Create the archive directory

```bash
cd /c/dev/officeoffice
mkdir -p docs/superpowers/plans/archive
```

---

## Step 2 — Move the completed plans

Everything dated **2026-08-22 through 2026-08-24** is shipped work. Move it:

```bash
cd /c/dev/officeoffice
git mv docs/superpowers/plans/2026-08-22-*.md docs/superpowers/plans/archive/
git mv docs/superpowers/plans/2026-08-23-*.md docs/superpowers/plans/archive/
git mv docs/superpowers/plans/2026-08-24-*.md docs/superpowers/plans/archive/
ls docs/superpowers/plans/*.md
```

Use `git mv`, not `mv` — it records the move so history follows the files.

**Expected to remain at the top level** (10 files):

```
2026-08-27-p13-task0-harness.md
2026-08-27-p13-task1-powerpoint.md
2026-08-27-p13-task2-excel.md
2026-08-27-p13-task3-word.md
2026-08-27-p13-task4-reconcile.md
2026-08-27-p13-task5-archive.md
2026-08-27-phase0-test-seam.md
2026-08-27-phase1-3-file-split.md
2026-08-27-refactor-proposal.md
STATUS.md
```

**Keep `2026-08-27-phase0-test-seam.md` at the top level even though it is done** — it is the immediate predecessor of the active work and is referenced by the current plans. Archive it once Phases 1+3 are also finished.

---

## Step 3 — Check nothing links to a moved file

Cross-references will now point at the wrong path:

```bash
cd /c/dev/officeoffice
grep -rn "plans/2026-08-2[234]-" docs/ --include=*.md | grep -v "/archive/" | head -20
```

**If this prints anything**, fix each hit to point at `plans/archive/…`. Bare filename mentions without a `plans/` path prefix are fine to leave.

Also check the source tree — code comments in this repo cite plan files by name:

```bash
grep -rn "docs/superpowers/plans/2026-08-2[234]" --include=*.cs --include=*.ts . | head -20
```

**Leave code comments alone even if they hit.** They cite a plan as provenance for a decision ("PP-15 Task 4 did X"), and rewriting dozens of comments to chase a docs reorganisation is churn with no benefit. Note it in Step 5's README instead.

---

## Step 4 — Do the same for `verification/` (optional)

`docs/superpowers/verification/` has 25 files, all records of manual test passes for shipped work. Same treatment if you want the symmetry:

```bash
cd /c/dev/officeoffice
mkdir -p docs/superpowers/verification/archive
git mv docs/superpowers/verification/*.md docs/superpowers/verification/archive/
```

**This is genuinely optional.** Unlike `plans/`, the verification folder has no mix of current and stale — *everything* in it is historical, so the folder name already carries that meaning. Skip it if you prefer.

---

## Step 5 — Add a one-line signpost to `STATUS.md`

At the very top of `docs/superpowers/plans/STATUS.md`, above `# Implementation status`:

```markdown
> **Where to look:** `docs/ai-tool-surface.md` is the canonical description of
> what the tools currently do. `plans/*.md` at this level are active or
> recently-completed work; `plans/archive/` is history — useful for *why* a
> decision was made, but **not** a description of current behaviour. Code
> comments cite archived plans by their original path; that is provenance, not
> a live link.
```

That last sentence is the one that prevents the next reader from "fixing" dozens of code comments.

---

## Step 6 — Verify and commit

```bash
cd /c/dev/officeoffice
ls docs/superpowers/plans/*.md | wc -l          # expect 10
ls docs/superpowers/plans/archive/*.md | wc -l  # expect 35
git status --short | head -5
```

`git status` should show the moves as renames (`R`), not delete+add. If it shows `D`/`??` pairs you used `mv` instead of `git mv` — recoverable, just `git add -A docs/superpowers/` and git will detect the renames.

```bash
git add -A docs/superpowers/
git commit -m "docs: archive completed plan docs, signpost the canonical current-state doc

Moves 35 shipped plans (2026-08-22..24) to plans/archive/ so the top level
holds only active work. Archived, not deleted - they are the audit trail for
why decisions were made, and this refactor leaned on them repeatedly.

Adds a STATUS.md signpost naming ai-tool-surface.md as canonical, and noting
that code comments cite archived plans as provenance rather than live links.

P13 Task 5 of docs/superpowers/plans/2026-08-27-phase1-3-file-split.md."
```

---

## Definition of done

- [ ] `docs/superpowers/plans/` has 10 `.md` files; `archive/` has 35.
- [ ] No markdown cross-reference points at a moved file's old path.
- [ ] `STATUS.md` carries the signpost.
- [ ] Git recorded renames, not delete+add.
- [ ] One commit, docs only.
