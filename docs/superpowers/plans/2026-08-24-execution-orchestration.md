# Execution Orchestration — Phases, Agents, and Parallelization

**Audience:** the model coordinating implementation of the 25 plans in this folder (`2026-08-23-pp*.md`, `2026-08-23-ft*.md`). Read this document **before** opening any plan. It tells you what order to work in, what may run at the same time, and what you are not allowed to do.

**Companion documents:** `2026-08-23-pp-index.md` (what each plan is), `2026-08-23-ai-addin-product-plan.md` (why each plan exists). Each individual plan is self-contained and task-by-task; this document never restates their content, only their sequencing.

---

## 0. How to use this document

1. Read Section 1 (ground rules). They apply to every task in every plan.
2. Do Section 2 (day-one actions) before scheduling any implementation work.
3. Work through Sections 4–9 (the phases) **in order**. Do not start a phase before its gate in Section 10 has passed.
4. Within a phase, Section 3's conflict map tells you what can run at the same time.
5. Keep `docs/superpowers/plans/STATUS.md` (Section 11) current after every task.

**If this document and an individual plan disagree, the individual plan wins on *what* to build; this document wins on *when* and *by whom*.**

---

## 1. Ground rules for every agent

These are the rules most likely to be broken. Put them in every subagent prompt verbatim (Section 8 has the template).

- **Stay inside your assigned files.** Every agent is given an explicit file list. Editing a file outside that list is the single most damaging thing you can do here, because parallel agents are working in other files and a stray edit is invisible until merge. If a task appears to require a file you were not assigned, **stop and report** — do not edit it.
- **C# 7.3 / .NET Framework 4.8 only.** No `using`-declarations, no target-typed `new()`, no switch expressions, no records, no nullable reference types. This is a hard compiler constraint, not a style preference.
- **Rebuild after every TypeScript change.** Run the esbuild bundle command for the affected app, then `MSBuild <App>/<App>.csproj -t:Build -p:Configuration=Debug`. A stale bundle silently ships old behavior — you will believe your change works when it is not in the build.
- **Never commit bundles.** `web/bundle.js`, `.map`, `bundle.css` are gitignored build artifacts.
- **Never report success for work you did not do.** Most of these plans exist because the *code* does this — a tool returns "ok" for a no-op. Do not do the same in your report. If a step is partially done, blocked, or unverified, say exactly that.
- **Do not fix things outside your plan's scope.** You will notice other defects; this codebase has many. Write them in your report. Do not fix them — an unplanned fix in a file another agent owns is a merge conflict, and an unplanned fix in your own file makes your change impossible to review.
- **Do not invent verification.** If a plan's verification step says "in real Word", you cannot do it. Say so and record it for the human checkpoint (Section 7). Never write "verified" for something you did not run.
- **When a plan says a step may already be done by an earlier plan, check before doing it.** Several plans overlap deliberately (Section 3.3). Re-doing a landed change creates conflicts.
- **Stop after 2 failed attempts at the same thing.** Report what you tried and what happened. Do not keep retrying variations.

---

## 2. Day-one actions (before any implementation)

- [ ] **Raise both decision gates with the product owner now**, because they block work that sits late in the schedule and the answers take time:
  - **PP-7** (`2026-08-23-pp07-data-provenance-guardrail.md` Task 0) — port the `dataSource` guardrail, prompt-only mitigation, or decline. Plan recommends prompt-only.
  - **PP-19 Task 2** — full DSL/deck-generation port, decline, or the batch-ops + layout-check middle option. Plan recommends the middle. **PP-19 Task 1 (`delete_slide`) ships regardless and is not blocked by this.**
- [ ] **Confirm the build invocation.** No `.csproj` contains an esbuild `Exec` target, so the bundle command is invoked manually or from an external script. Find it, write the exact command into `STATUS.md`, and give it to every agent. PP-0 Task 4 Step 3 also depends on knowing this.
- [ ] **Confirm who can run manual Office verification, and how often.** This is the schedule's real constraint (Section 7). If nobody can, say so now — most of these plans cannot be signed off without it.
- [ ] **Create `docs/superpowers/plans/STATUS.md`** from Section 11's template.
- [ ] **PP-8** (`2026-08-23-pp08-retire-stale-todo.md`) can be done immediately by anyone. It is documentation-only, touches no code, blocks nothing, and stops other people planning off a stale checklist. Do it first.

---

## 3. The dependency facts

Two kinds of dependency matter and they are **not** the same:

- **Logical** — plan B needs plan A's output to exist.
- **File-level** — plans A and B edit the same file, so they cannot run at the same time even if they are logically independent.

File-level conflicts are what actually break parallel work, and they are not visible from the index.

### 3.1 File ownership map

| File | Plans that modify it |
|---|---|
| `shared/chat-ui/chat-ui.ts` | PP-2, PP-3, PP-4, PP-6, FT-1, FT-2 |
| `shared/chat-ui/chat-ui.css` | PP-3, PP-6, FT-1 |
| `shared/chat-ui/chat-ui.test.ts` | PP-2, PP-3, PP-4, PP-6, FT-1, FT-2 |
| `shared/web-src/app-shell/*` (created by PP-0) | PP-4, PP-6, FT-1, FT-2 |
| `WordAiAddIn/web-src/entry.ts` | PP-0, PP-5, PP-9, PP-10, PP-11, PP-12, FT-1 |
| `ExcelAiAddIn/web-src/entry.ts` | PP-0, PP-5, PP-13…PP-18, FT-1 |
| `PowerPointAiAddIn/web-src/entry.ts` | PP-0, PP-19…PP-22, FT-1 |
| `WordAiAddIn/WordTools.cs` | PP-5, PP-9, PP-10, PP-11, PP-12 |
| `ExcelAiAddIn/ExcelTools.cs` | PP-5, PP-13…PP-18, **PP-21** |
| `PowerPointAiAddIn/PowerPointTools.cs` | PP-19, PP-20, PP-21, PP-22 |
| `*/ThisAddIn.cs`, `*/TaskPaneHost.cs` | PP-0, PP-1, FT-1, FT-2 |
| `OfficeAi.Shared/*` | PP-0, PP-1, FT-1 |

**Read this table as: any two plans on the same row must not run concurrently.**

### 3.2 The three bottleneck files

- `shared/chat-ui/chat-ui.ts` — six plans. This is the schedule's critical path and is almost entirely serial.
- Each app's `entry.ts` + `*Tools.cs` — the per-app lanes. Different apps are genuinely independent; plans **within** one app's lane are not.
- `*/ThisAddIn.cs` + `*/TaskPaneHost.cs` — PP-0's C# half and PP-1 both rewrite these, which is why PP-0 Task 9 hands the C# extraction to PP-1 rather than doing it separately.

### 3.3 Cross-lane hazards — read these before scheduling Phase 4

Three places where a plan reaches outside its own app. Each will cause a merge conflict or a wrong result if scheduled naively.

1. **PP-21 Task 2 Step 5 edits `ExcelTools.cs`.** The PowerPoint `legendPos` plan also fixes Excel's identical defect, deliberately, so it does not get forgotten. **Resolution: move that one step into the Excel lane** and have the PowerPoint agent skip it. Record the move in `STATUS.md` so it is not dropped.
2. **PP-20 Task 2 Step 1 copies `ShapeTypeMap` out of `ExcelTools.cs`, which PP-16 modifies** (adds `OrdinalIgnoreCase`). It only *reads* Excel's file, so there is no edit conflict — but copying the pre-PP-16 version ships the wrong thing. **Resolution: PP-16 must complete before PP-20 starts.** This is the one ordering dependency that crosses lanes.
3. **PP-4 and PP-6 both edit `shared/web-src/app-shell/settings.ts`** after PP-0. They must be serial. PP-4 first (it is one line; PP-6 is a rewrite of the same file).

### 3.4 Logical dependencies

```
PP-0 ──┬─> PP-2 ─> PP-3 ─> PP-4 ─> PP-6 ─> FT-1 ─> FT-2
       ├─> PP-1 ──────────────────────────> FT-1 Task 7b
       └─> PP-5 ──┬─> Word lane   (PP-9, PP-10 -> PP-11, PP-12)
                  ├─> Excel lane  (PP-13 … PP-18)
                  └─> PPT lane    (PP-19 T1, PP-16 -> PP-20, PP-21 -> PP-22)
PP-8  (independent, do immediately)
PP-7  (decision gate, then small)
```

---

## 4. Phase 1 — Foundation (serial, 1 agent)

**Goal:** the shared shell exists, so later phases edit one file instead of three.

| Order | Plan | Scope |
|---|---|---|
| 1 | PP-8 | Docs only. Can be done by anyone at any time. |
| 2 | PP-0 Tasks 1–6 | TypeScript app-shell extraction and the three `entry.ts` migrations. |

**Do not parallelize this.** PP-0 Task 5 Step 2 is an explicit correctness checkpoint: migrate Word first, verify by hand, and only then do Excel and PowerPoint. That sequencing is the whole safety mechanism of a pure refactor with no automated coverage.

**PP-0 Tasks 7–9 (the C# half) are deferred to Phase 2** and are owned by PP-1's agent. Do not do them here.

**Exit gate:** all three apps build; all three behave identically to before (PP-0 Task 5 Step 2's manual checklist, run in each app); `git diff --stat` shows a net deletion of roughly 400 lines.

---

## 5. Phase 2 — Two independent tracks (2 agents, parallel)

These two touch disjoint file sets — one is pure TypeScript, the other pure C#. They are the only clean large-scale parallelism in the schedule.

### Track A — shared UI chain (1 agent, strictly serial)

PP-2 → PP-3 → PP-4. All three edit `chat-ui.ts` in the same region; PP-2's lifecycle rework is the foundation the other two build on. One agent does all three in order. Do not split them across agents.

### Track B — host and windowing (1 agent)

PP-0 Tasks 7–9, then PP-1. Same agent, because PP-1 builds directly on `PaneHostBase`/`RibbonBase` and rewrites the same six files.

**Exit gate:** Track A — `npx vitest run` green in `shared/chat-ui`, plus PP-2/PP-3/PP-4's manual checks. Track B — PP-1 Task 7's full manual matrix in all three apps. Track B's gate is the expensive one; schedule the human checkpoint for it early.

---

## 6. Phase 3 — Schema foundation (serial, 1 agent)

**PP-5.** It touches Word's and Excel's `entry.ts` *and* both `*Tools.cs`, so it conflicts with both app lanes and must complete before Phase 4 starts.

Watch PP-5 Task 5: if Excel's generated schema turns out too large, the plan switches to a "grouped kinds" variant. That is a real decision the agent must make from a measurement, not a guess — require the before/after byte counts in its report.

PP-5 Task 4 Step 1 (per-command isolation in the batch loops) also satisfies PP-12 Task 3 and part of PP-14 Task 1. Record in `STATUS.md` that those are done, or the lane agents will redo them.

**Exit gate:** both projects build; PP-5 Task 4's malformed-batch manual check passes; Task 5's size measurement is recorded.

---

## 7. Phase 4 — Three app lanes (3 agents, parallel)

This is where parallelization pays off. Each lane owns one app's `entry.ts` and one `*Tools.cs` and touches nothing else.

| Lane | Files owned | Plans, in order |
|---|---|---|
| **Word** | `WordAiAddIn/WordTools.cs`, `WordAiAddIn/web-src/entry.ts` | PP-9, PP-10, PP-11 *(needs PP-10 Task 1)*, PP-12 |
| **Excel** | `ExcelAiAddIn/ExcelTools.cs`, `ExcelAiAddIn/web-src/entry.ts` | PP-13, PP-14, PP-16, PP-15, PP-17, PP-18, **+ PP-21 Task 2 Step 5** |
| **PowerPoint** | `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts` | PP-19 Task 1, PP-21 *(minus Task 2 Step 5)*, PP-22 *(needs PP-21)*, PP-20 *(needs Excel lane's PP-16)* |

**Within a lane the plans are strictly serial** — they all edit the same two files.

**Two scheduling notes:**
- Excel's lane runs **PP-16 third, before PP-15**, purely so the PowerPoint lane's PP-20 unblocks sooner. PP-16 is also the smallest plan in the set, so this costs nothing.
- PowerPoint's lane runs **PP-20 last**, because it is the one plan gated on another lane. If the Excel lane falls behind, PowerPoint's agent should report idle rather than starting PP-20 early against the pre-PP-16 map.

**Use one git worktree per lane** (`isolation: "worktree"` if spawning via the Agent tool). Three agents doing MSBuild in the same tree will collide on `obj/` and `bin/`.

**Exit gate per lane:** that lane's plans' manual matrices. These are long — PP-16 alone asks for 27 shape types to be eyeballed in real Excel. Budget for it; see Section 9.

---

## 8. Phase 5 — Provider and features (serial, 1 agent)

PP-6 → FT-1 → FT-2. All three converge on `chat-ui.ts` and `app-shell/`, so they are serial regardless of how much time is left.

- **PP-6** before **FT-1**, so FT-1's settings screen renders PP-6's provider selector rather than three text boxes that then have to be rebuilt (FT-1 Task 2 Step 4).
- **FT-1 Task 7b** additionally needs PP-1 Task 5 Step 1 (from Phase 2 Track B) — verify it landed before starting, because re-keying without the per-document unsaved id is destructive, not merely wrong.
- **FT-2** last. It depends on PP-0's shell (Phase 1) and PP-1's per-window routing (Phase 2), both long done by here.

**PP-7** slots in wherever its Task 0 answer arrives. If the answer is prompt-only, it is a 30-minute change to three system prompts and can be done by any agent between other work. If it is the full port, treat it as a fourth item in this phase.

---

## 9. Verification — the real constraint

**Most of these plans cannot be signed off by an agent.** There are no automated tests for COM executor methods (an explicit, long-standing project convention), so nearly every plan ends in a manual matrix to be run in real Word/Excel/PowerPoint. An agent in a worktree cannot open Office.

What agents *can* verify, and must:
- `MSBuild` succeeds for every affected project.
- `npx vitest run` in `shared/chat-ui` is green.
- `OfficeAi.Shared.Tests` passes (relevant to PP-0, PP-1, FT-1 Task 7/7b).
- `npx tsc --noEmit` per app.
- The bundle rebuilds.

What must go to a human:
- Every plan's "Manual verification matrix" task.

**Protocol:** each agent, on finishing a plan, writes `docs/superpowers/verification/<plan-id>.md` containing that plan's manual matrix as an unchecked checklist, plus anything it could not verify and any deviation it made. A human runs it and checks the boxes. **A plan is not done until its verification file is fully checked.**

Batch these into **checkpoints at phase boundaries** rather than after every plan — an Office testing session has fixed setup cost, and the matrices for one phase share a session. Four checkpoints total, matching the five phases (Phase 4's three lanes converge into one).

Be honest in scheduling: Phase 4's three lanes will produce a large combined matrix. If manual testing capacity is the bottleneck, run **fewer lanes at once** rather than accumulating unverified merged code — unverified silent breakage is precisely the failure class this whole plan set exists to remove.

---

## 10. Phase gates

Do not start a phase until the previous gate passes.

| Gate | Passes when |
|---|---|
| **G0 → Phase 1** | Build command confirmed and recorded; `STATUS.md` created; PP-7 and PP-19 decisions requested. |
| **G1 → Phase 2** | PP-0 Tasks 1–6 complete; all three apps build and behave as before; Phase 1 verification file fully checked. |
| **G2 → Phase 3** | Track A and Track B both complete and verified. In particular PP-1's matrix, since Phase 5's FT-1 Task 7b depends on it. |
| **G3 → Phase 4** | PP-5 complete and verified; overlap notes recorded in `STATUS.md` (which of PP-12 Task 3 / PP-14 Task 1 it already did). |
| **G4 → Phase 5** | All three lanes complete and verified; PP-21 Task 2 Step 5's relocation into the Excel lane confirmed done. |

---

## 11. `STATUS.md` template

Create this at `docs/superpowers/plans/STATUS.md` and update it after **every** task, not every plan.

```markdown
# Implementation status

**Build command:** <the exact esbuild invocation, filled in on day one>
**Manual verification owner:** <name> — availability: <...>

## Decisions
| Gate | Question | Answer | Date |
|---|---|---|---|
| PP-7 Task 0 | dataSource guardrail | PENDING | |
| PP-19 Task 2 | DSL / generation scope | PENDING | |

## Plans
| Plan | Phase | Agent | Code | Verified | Notes |
|---|---|---|---|---|---|
| PP-8 | 1 | — | done | n/a | docs only |
| PP-0 | 1 | | | | |
| … one row per plan … |

## Cross-plan overlaps already handled
<!-- e.g. "PP-5 Task 4 Step 1 landed per-command isolation; PP-12 Task 3 is a no-op — verify only." -->

## Deviations from plan
<!-- Anything an agent did differently, and why. Reviewers read this first. -->

## Deferred findings
<!-- Defects noticed but deliberately not fixed, per Ground Rule 6. -->
```

---

## 12. Subagent prompt template

Give each agent exactly this, filled in. Do not paraphrase the ground rules.

```
You are implementing ONE plan in the officeoffice repository.

PLAN: docs/superpowers/plans/<file>.md
Read it completely before editing anything. It is task-by-task with checkbox steps.
Work the tasks in order. Do not skip ahead.

FILES YOU MAY EDIT — this list is exhaustive:
  <exact paths>
If a task appears to need a file not on this list, STOP and report. Other agents
are working in other files right now; a stray edit is invisible until merge.

CONTEXT YOU MAY READ (do not edit):
  docs/superpowers/plans/2026-08-23-pp-index.md
  docs/superpowers/plans/STATUS.md          <- check "overlaps already handled" first
  any source file you need to understand

RULES:
- C# 7.3 / .NET Framework 4.8 only. No using-declarations, no target-typed new(),
  no switch expressions. This is a compiler constraint.
- After ANY TypeScript change: rebuild the bundle with
    <exact esbuild command>
  then: MSBuild <App>/<App>.csproj -t:Build -p:Configuration=Debug
  A stale bundle silently ships old behavior.
- Never commit web/bundle.js or its siblings — gitignored build artifacts.
- Do not fix defects outside this plan's scope. Record them in your report instead.
- If the plan says a step may already be done by an earlier plan, verify before
  doing it. Check STATUS.md.
- Stop after 2 failed attempts at the same problem and report.

VERIFICATION:
- Run every build/test command the plan names. Report actual output.
- You CANNOT verify anything requiring real Word/Excel/PowerPoint. Do not claim to.
  Write docs/superpowers/verification/<plan-id>.md containing the plan's manual
  matrix as an unchecked checklist for a human to run.

REPORT WHEN DONE:
1. Which tasks/steps are complete, and which are not.
2. Build and test output, verbatim.
3. Every deviation from the plan, and why.
4. Defects noticed but not fixed.
5. Anything you could not verify.

Never report success for work you did not do. Partial is fine; say so precisely.
```

---

## 13. Summary schedule

| Phase | Agents | Contents | Blocked by |
|---|---|---|---|
| 0 | — | Decisions requested, build command confirmed, `STATUS.md`, PP-8 | — |
| 1 | 1 | PP-0 Tasks 1–6 | G0 |
| 2 | 2 (parallel) | A: PP-2→PP-3→PP-4 · B: PP-0 T7–9→PP-1 | G1 |
| 3 | 1 | PP-5 | G2 |
| 4 | 3 (parallel) | Word / Excel / PowerPoint lanes | G3 |
| 5 | 1 | PP-6→FT-1→FT-2 (+PP-7 when answered) | G4 |

Peak parallelism is three agents, in Phase 4. Everything else is one or two — not from caution, but because `chat-ui.ts` and the shared host files are genuinely serial resources, and pretending otherwise produces merge conflicts rather than speed.
