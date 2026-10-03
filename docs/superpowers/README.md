# `docs/superpowers/` — completed-work archive

This directory is a **dated historical record of already-completed implementation
work**, not live instructions, not current architecture, and not something a reader
(human or AI agent) should treat as a to-do list. Almost everything described in here
has already been built, merged, and folded into the current-state reference at
**`docs/ai-tool-surface.md`**.

If you want to know what the tool surface can do *today*, read `docs/ai-tool-surface.md`
(and, for the AI-agent-facing conventions this codebase follows, the root `CLAUDE.md`).
Come here only when you need the history behind a specific change — why a tool's
schema looks the way it does, what was tried and rejected, or what a specific bug fix
actually changed.

## What's in each subdirectory

- **`plans/`** — ~50 one-off planning documents (`2026-08-22-...md` through
  `2026-09-19-...md`), each written to drive a single implementation task (a "PP-n" or
  "FT-n" item, a phase of a larger port, or a standalone fix). Most describe work that
  is fully implemented and merged; a plan's existence here does not mean the work is
  pending. `STATUS.md` in this directory is the index — it tracks which plans landed,
  what deviated from the original plan, and what (if anything) is still only
  automated-checks-clean rather than confirmed against a live Office session.
- **`verification/`** — ~30 notes, generally one per plan, recording what was actually
  checked after a plan's implementation (automated builds/tests) versus what still
  needs a real Office session to confirm ("Manual verification matrix — none of this has
  been run" is a common, honest line in these files — this codebase's convention is to
  say so explicitly rather than claim untested code works; see the root `CLAUDE.md`).

## Why this exists as a separate archive

`docs/ai-tool-surface.md` used to interleave a growing sequence of dated update notes
with its current-state tool tables, which made it hard to tell "what is true now" from
"what changed and when." That document has since been restructured to be
current-state-first, with its own history moved to `docs/ai-tool-surface-changelog.md`.
This `plans/`/`verification/` archive is the more detailed, per-task version of that
same history — individual plan documents and their verification notes, rather than a
single chronological changelog. Neither this README nor the restructuring renamed,
moved, or edited any of the existing files in `plans/`/`verification/` — they are left
exactly as they were written, as a record of what was actually done and decided at the
time.
