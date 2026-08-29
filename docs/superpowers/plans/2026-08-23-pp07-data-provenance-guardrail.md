# PP-7: Data Provenance Guardrail — Decision Brief and Conditional Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. **Task 0 is a product-owner decision gate — do not start Task 1 until it is answered.**

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-7 (P2, "recommend product-owner call before scoping").

**Goal:** Decide whether to port genoffice's `dataSource` provenance gate to officeoffice, and — if yes — implement the reduced, air-gap-appropriate version scoped in Tasks 1-4.

---

### Task 0: Decision gate

**Files:** none. Output is a recorded decision.

**The question:** genoffice gates chart/data-bearing content behind a `dataSource` enum (`user` / `document` / `search` / `sample`) and rejects `'search'` unless a `web_search` actually happened in-conversation. officeoffice has no web search at all, so the specific failure it guards — content presented as researched when nothing was researched — largely cannot occur here.

**What remains applicable:** "the model invented numbers instead of reading them from the sheet." A user asks for a chart of Q3 revenue; the model doesn't call `read_range`, produces plausible figures, and the chart lands in the workbook looking authoritative. Nothing today distinguishes that from a chart built on real cell values.

**Arguments for porting a reduced version:**
- Excel and PowerPoint chart/data tools accept arbitrary literal values with no claim about origin (`ExcelTools.cs:643-663` `AddChart`, `PowerPointTools.cs:437+` `AddChartPpt`).
- Fabricated numbers inside a real .xlsx are materially worse than fabricated text in a chat window — the artifact outlives the conversation and loses the context that a model produced it.
- The reduced form is cheap: one required enum field on ~4 tools plus a transcript label. Days, not weeks.

**Arguments against:**
- It is a *declaration*, not a verification. A model that invents numbers can equally invent `dataSource: 'document'`. Only the `'search'` case in genoffice was actually checkable — and that is the one case that does not apply here.
- Softer partial mitigations exist and cost less: strengthen the system prompt to require reading before charting, and — once PP-3 ships — a user can see whether a read tool ran at all.
- It adds a required field to the highest-traffic authoring tools, i.e. a new way for a request to fail.

- [ ] **Step 1:** Present the above to the product owner and record one of:
  - **(A) Port the reduced version** → proceed to Tasks 1-4.
  - **(B) Prompt-only mitigation** → do only Task 1, skip 2-4.
  - **(C) Decline** → close the item; do Task 4 Step 1 only (record the decision in `docs/ai-tool-surface.md` so the next audit doesn't re-raise it).
- [ ] **Step 2:** Write the decision, its date, and its rationale into `docs/ai-tool-surface.md`'s scope section regardless of which branch is chosen.

**Recommendation if no owner is available:** take **(B)**. It captures most of the practical benefit at a fraction of the cost and does not add a failure mode to the authoring tools. Revisit (A) only if officeoffice ever gains a data-retrieval tool, at which point the checkable case returns.

---

## Tasks 1-4 — only if (A) or (B) was chosen

### Task 1: System-prompt provenance rule (both (A) and (B))

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Extend each app's `AgentSkill.systemPrompt` (Word's is at `WordAiAddIn/web-src/entry.ts:250-257`; the Excel and PowerPoint equivalents sit in the same position in their files) with a rule of the form:

> Never invent numeric data. Before creating or editing a chart, table, or any data-bearing content, read the actual values from the workbook/deck with a read tool. If the user supplied the numbers in the conversation, use exactly those. If you have neither, say so and ask — do not produce illustrative or placeholder figures presented as real.

- [ ] **Step 2:** Word's chart tool is single-series and unlabeled (see PP-9), so it carries the same risk; add the rule there too.
- [ ] **Step 3:** Verify by asking, in a workbook with no revenue data, "chart our Q3 revenue" — the model should ask or refuse rather than fabricate. Record the observed behavior; it is the baseline against which (A)'s added value is judged.

**Verification:** manual, as above, in all three apps.

---

### Task 2: `dataSource` field on data-bearing tools ((A) only)

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts`, `WordAiAddIn/web-src/entry.ts`
- Modify: `ExcelAiAddIn/ExcelTools.cs`, `PowerPointAiAddIn/PowerPointTools.cs`, `WordAiAddIn/WordTools.cs`

- [ ] **Step 1: Scope the field to exactly these tools/operations**
  - Excel `propose_operations`: `add_chart`, `add_sparkline`, `set_range`, `add_pivot`.
  - PowerPoint: `add_chart`, `add_table`.
  - Word: `edit_chart`.

  Deliberately **not** `set_cell`/`set_formula`/`edit_table_cell` — single-value edits are usually direct user instructions, and requiring provenance on every one would be noise that trains the model to fill it in reflexively.

- [ ] **Step 2: Schema** — add `dataSource: { type: 'string', enum: ['user', 'document', 'sample'] }` to each, required, with a description defining each value: `user` = values the user gave in this conversation; `document` = values read from this document/workbook/deck with a read tool in this conversation; `sample` = illustrative placeholder values, explicitly labeled as such. `'search'` is **excluded** — officeoffice has no web search, so offering it would invite exactly the unverifiable claim genoffice's gate exists to prevent. If PP-5 has landed, add these through its schema tables rather than by hand.

- [ ] **Step 3: Handler enforcement** — reject a missing or out-of-enum value with a specific error naming the allowed values. This is the only mechanical guarantee available; there is no way to verify a `document` claim from inside the handler.

- [ ] **Step 4: `sample` must be visible in the artifact, not just the transcript.** When `dataSource == "sample"`, append " (sample data)" to the chart title (or set a title if none was given). A placeholder chart that looks identical to a real one is the actual harm; a label in the chat that the user scrolls past is not a mitigation.

**Verification:** both projects build; a chart request without `dataSource` returns a specific error; a `sample` chart carries the label in the document.

---

### Task 3: Transcript surfacing ((A) only)

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

- [ ] **Step 1:** When a completed step's input carries `dataSource`, render a small badge on the step row next to the existing `✓ Applied` tag (`shared/chat-ui/chat-ui.ts:352-357`): neutral for `user`/`document`, warning-colored for `sample`.
- [ ] **Step 2:** Localize the three labels through `STRINGS`.
- [ ] **Step 3:** Test in `shared/chat-ui/chat-ui.test.ts` that the badge appears and reflects the value.

**Note:** if PP-3 has landed, the full tool input is already inspectable, and this task is a convenience rather than the only way to see the claim. Consider dropping it if scope needs trimming.

**Verification:** `cd shared/chat-ui && npx vitest run`; visual check in a real add-in.

---

### Task 4: Document the outcome (all branches)

**Files:**
- Modify: `docs/ai-tool-surface.md`

- [ ] **Step 1:** Record the Task 0 decision, its date, and its rationale in the scope section, including the explicit note that `'search'` is not offered because officeoffice has no web-search tool — so a future reader doesn't "restore parity" with genoffice by adding it.
- [ ] **Step 2 ((A) only):** Document the enum, which tools carry it, and the `sample` labeling rule.
