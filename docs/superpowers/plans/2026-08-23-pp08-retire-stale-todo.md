# PP-8: Retire the Stale `tool-surface-todo.md` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-8 (P2, housekeeping).

**Goal:** Stop `docs/tool-surface-todo.md` (150 lines) from misleading planning work. It marks many command/operation kinds `[ ]` unimplemented that are fully implemented on current `main` — most egregiously "9 of 65 operations" for Excel, where the real figure is essentially all of them — while `docs/ai-tool-surface.md` already carries a source-verified account.

**Architecture:** The file is not entirely dead weight. Its **header** carries the only written statement of the project's *scope boundary* — what is deliberately excluded (`web_search`, `image_search`, `generate_image`, `analyze_media`, `read_attachment`, the PDF and Markdown apps, and PowerPoint's `execute_slide_script` DSL + deck-generation pipeline). `docs/ai-tool-surface.md:9` and the product plan's PP-19 both lean on that statement, and PP-19 explicitly notes the scope note is accurate even though the counts are not.

So: preserve the scope note by moving it into `docs/ai-tool-surface.md`, then replace the file's body with a pointer. Deleting outright would drop the scope statement; leaving it would keep the misleading counts.

**Tech Stack:** Markdown only. No code changes, no build.

## Global Constraints

- Do not delete the scope-boundary content — relocate it first, verify it reads correctly in its new home, then trim the source.
- Keep the file at its existing path with a stub rather than deleting it, so the three existing inbound references (`docs/ai-tool-surface.md:9` and `:105`, plus the product plan's PP-19 at `:503`) do not become dead links. Deletion is a follow-up once no document references it.
- Do not re-audit the tool surface as part of this task. `docs/ai-tool-surface.md` is already verified against current source (per its own top note); this plan trusts it. Any discrepancy noticed in passing gets filed as a new item, not fixed inline.

---

### Task 1: Move the scope boundary into `docs/ai-tool-surface.md`

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1:** Read `docs/tool-surface-todo.md:1-13` (the header and the "Explicitly out of scope everywhere" block) and `docs/ai-tool-surface.md`'s existing "explicitly out of scope" section.
- [ ] **Step 2:** Reconcile the two lists. If `ai-tool-surface.md`'s section already contains everything in the todo's block, note that and skip to Task 2. If anything is missing — particularly the full exclusion list `web_search`, `image_search`, `generate_image`, `analyze_media`, `read_attachment`, the PDF app, the Markdown app (folded into Word), and PowerPoint's `execute_slide_script` DSL plus `ask_clarification`/`plan_deck`/`generate_deck`/`regenerate_slide`/`delete_slide`/`save_style_template`/`list_style_templates` — add it verbatim.
- [ ] **Step 3:** Add one sentence of provenance: that this scope boundary originated in the feasibility report and the toolset-port plan's Global Constraints, and was previously stated in `tool-surface-todo.md`.
- [ ] **Step 4:** Flag one inconsistency for a later item, do not resolve it here: the exclusion list bundles `delete_slide` with the deck-generation pipeline, while PP-19 argues `delete_slide` is a small, clearly-in-scope fix (it has no dependency on the DSL, and `add_slide` already exists). Add a parenthetical noting the conflict and pointing at `2026-08-23-pp19-powerpoint-scope-and-delete-slide.md`.

**Verification:** `docs/ai-tool-surface.md` states the full scope boundary; nothing in the todo's header is now recorded only there.

---

### Task 2: Replace the todo's body with a pointer

**Files:**
- Modify: `docs/tool-surface-todo.md`

- [ ] **Step 1:** Replace the entire file with a short stub:

```markdown
# Tool surface checklist — RETIRED

This checklist was written against an early snapshot and its implementation
counts were badly out of date (it claimed "9 of 65" Excel operations when
essentially all of them were implemented). Planning off it produced work items
for things that already shipped.

**Use `docs/ai-tool-surface.md` instead** — it is verified against current
source and carries the tool-by-tool comparison, the full schema audit, and the
project's scope boundary (what is deliberately out of scope).

For prioritized, product-framed gaps see
`docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` and the per-item
plans beside it (`2026-08-23-pp*.md`).

Retired 2026-08-23 (see PP-8).
```

- [ ] **Step 2:** Do not preserve the old checkbox lists, even commented out or in git-history-only form — a `[ ]` in a retired file is exactly what caused the problem, and the content is recoverable from git.

**Verification:** the file is under ~20 lines and contains no checkbox syntax.

---

### Task 3: Fix inbound references

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1:** `docs/ai-tool-surface.md:9` currently explains that the todo "is stale". Update it to say the file is retired and points here, so the note matches reality.
- [ ] **Step 2:** `docs/ai-tool-surface.md:105` cites "9 of 65 implemented" as an example of the stale count. Keep the example (it is useful context) but make the tense past and reference the retirement.
- [ ] **Step 3:** Leave `2026-08-23-ai-addin-product-plan.md` unedited — it is a dated snapshot of an audit, and rewriting its evidence lines would falsify the record. PP-19's reference to the todo's scope note still resolves via the stub.
- [ ] **Step 4:** `grep -rn "tool-surface-todo" --include=*.md --include=*.ts --include=*.cs .` and confirm every remaining hit is either updated or deliberately left as a historical record.

**Verification:** no document tells a reader to plan work off the checklist.
