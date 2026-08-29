# PP-3: Inspectable Tool Output in the Transcript — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-3 (P1).

**Goal:** Let a user expand any tool call in the transcript and see the actual text that tool returned — the same string the model received — instead of only a ✓/✗ icon. This is the diagnostic surface that would have resolved the "`get_document_context` returns empty" report without a source dive.

**Architecture:** The data already flows all the way to the UI and is thrown away at the last step. `ToolStepHandle.complete()` accepts `{ output, isError, mutated }` (`shared/chat-ui/chat-ui.ts:48`) and its implementation (`:346-359`) reads only `isError` and `mutated`; `output` is dropped. All three `entry.ts` files already pass the real output through (`WordAiAddIn/web-src/entry.ts:335-340` and equivalents), sourced from `execution.output`, which is the C# `ToolResult.Output` string round-tripped through `OfficeAi.Shared/ToolProtocol.cs:SerializeToolResult`.

So the work is entirely presentational: render the output into a collapsible region under each step row, inside the existing `.ai-work-group` disclosure. Two design constraints shape it — tool outputs can be large (a `read_range` of 2000 cells), so the panel must cap what it renders inline with an explicit "show more"; and outputs are untrusted model/document text, so they must be inserted as `textContent`, never `innerHTML`.

**Tech Stack:** TypeScript + CSS in `shared/chat-ui/`, vitest/jsdom tests.

**Dependency:** Land PP-2 (`2026-08-23-pp02-tool-steps-chronological-order.md`) first. It reworks the same `beginToolGroup` block; doing PP-3 first guarantees a merge conflict there.

## Global Constraints

- `complete(result)`'s signature does not change — `output` is already in it. No `entry.ts` change is required for the basic feature (Task 4 is optional polish only).
- Never use `innerHTML` for tool output. Use `textContent` (or the existing `escapeHtml` helper at `shared/chat-ui/chat-ui.ts:70-74` if a template string is unavoidable) — a document's own text reaches this surface verbatim.
- Colors and spacing come from existing CSS custom properties in `shared/chat-ui/chat-ui.css`; no new raw hex outside a token definition.
- Default state is **collapsed**. A tool group already collapses by default; expanding output by default would bury the answer under a wall of text and undo PP-2's readability win.
- Rebuild each add-in's bundle and re-run MSBuild after any `chat-ui.ts`/`chat-ui.css` change (command in PP-2's Global Constraints); bundles are gitignored.

---

### Task 1: Render tool output into an expandable region

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`

**Interfaces:**
- Produces: a `.ai-step-output` element per completed step, and a `.ai-step-row.has-output` class the CSS in Task 2 hooks.

- [ ] **Step 1: Constants**

Near the existing `truncateForDisplay` helper (`shared/chat-ui/chat-ui.ts:76-78`):

```ts
/** Inline cap for a tool's rendered output; longer outputs get a "show all" toggle. */
const TOOL_OUTPUT_PREVIEW_CHARS = 2_000
```

- [ ] **Step 2: Make the step row a two-part element**

In `addStep` (`shared/chat-ui/chat-ui.ts:338-346`), the row currently renders icon + title only. Keep that markup and append an empty, hidden output container after it:

```ts
const outputEl = document.createElement('pre')
outputEl.className = 'ai-step-output'
outputEl.dir = 'auto'
rowEl.appendChild(outputEl)
```

`<pre>` preserves the line structure of multi-line outputs (`apply_commands` returns one line per command; `read_blocks` one line per paragraph) without a formatter. `dir="auto"` matches the per-message bidi rule used for chat bubbles.

- [ ] **Step 3: Populate it in `complete`**

Inside `complete(result)` (`shared/chat-ui/chat-ui.ts:347-358`), after the existing icon/`mutated` handling:

```ts
const text = result.output ?? ''
if (text.length > 0) {
  rowEl.classList.add('has-output')
  const truncated = text.length > TOOL_OUTPUT_PREVIEW_CHARS
  outputEl.textContent = truncated ? text.slice(0, TOOL_OUTPUT_PREVIEW_CHARS) : text
  if (truncated) {
    const more = document.createElement('button')
    more.className = 'ai-step-output-more'
    more.type = 'button'
    more.textContent = `Show all (${text.length} chars)`
    more.addEventListener('click', (e) => {
      e.stopPropagation()
      outputEl.textContent = text
      more.remove()
    })
    rowEl.appendChild(more)
  }
}
```

`stopPropagation` matters: the row lives inside the group summary's click-to-toggle region (`shared/chat-ui/chat-ui.ts:344`), and without it "show all" would also collapse the group.

- [ ] **Step 4: Per-row disclosure**

Make the step title itself the toggle, so a group of 5 tools doesn't dump 5 outputs at once:

```ts
const titleEl = rowEl.querySelector<HTMLElement>('.ai-step-title')!
titleEl.addEventListener('click', (e) => {
  if (!rowEl.classList.contains('has-output')) return
  e.stopPropagation()
  rowEl.classList.toggle('output-open')
})
```

- [ ] **Step 5: Error outputs**

When `result.isError` is true, still render the output (it carries the error message from `WordTools.Execute`'s catch path) and add `outputEl.classList.add('error')` so Task 2 can color it. Error rows start **expanded** — an error the user cannot see is the exact failure mode this item exists to fix:

```ts
if (result.isError && text.length > 0) rowEl.classList.add('output-open')
```

**Verification:** `cd shared/chat-ui && npx vitest run` still passes (new assertions come in Task 3).

---

### Task 2: Styling

**Files:**
- Modify: `shared/chat-ui/chat-ui.css`

- [ ] **Step 1:** Add rules next to the existing `.ai-step-row` / `.ai-step-icon` / `.ai-step-title` block:

```css
.ai-step-output { display: none; }
.ai-step-row.output-open .ai-step-output { display: block; }
```

Plus, on `.ai-step-output`: a monospace stack, `white-space: pre-wrap`, `word-break: break-word`, `max-height` with `overflow: auto`, a muted background and border from existing tokens, and comfortable small type. `.ai-step-output.error` uses the same error color token the `.ai-step-icon.error` rule already uses.

- [ ] **Step 2:** `.ai-step-row.has-output .ai-step-title` gets `cursor: pointer` and a small disclosure caret (reuse the `.caret` glyph pattern from `.ai-work-group-summary`), rotated via `.ai-step-row.output-open`.

- [ ] **Step 3:** `.ai-step-output-more` styled as a quiet inline text button, not a primary button.

- [ ] **Step 4:** Verify both light and dark rendering if the stylesheet defines a dark variant; if it only defines one theme, match it and don't introduce a second.

**Verification:** visual check in the running add-in (Task 5).

---

### Task 3: Tests

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`

- [ ] **Step 1:** `complete({ output: 'Paragraphs: 3, Words: 40' })` puts that exact string in the row's `.ai-step-output` `textContent`.
- [ ] **Step 2:** Output is hidden by default (`.output-open` absent) and appears after clicking `.ai-step-title`.
- [ ] **Step 3:** `complete({ output: '<img src=x onerror=alert(1)>' })` — the rendered output element has no `img` child; `textContent` equals the raw string. This is the XSS regression guard.
- [ ] **Step 4:** An output longer than `TOOL_OUTPUT_PREVIEW_CHARS` renders truncated with a `.ai-step-output-more` button; clicking it reveals the full string and removes the button.
- [ ] **Step 5:** `complete({ output: 'boom', isError: true })` renders expanded (`.output-open` present) with the error class.
- [ ] **Step 6:** `complete({ output: '' })` adds neither `.has-output` nor a visible output region — an empty result stays as quiet as it is today.

**Verification:** `cd shared/chat-ui && npx vitest run` — all green.

---

### Task 4 (optional polish): Input display parity

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`

The step title currently inlines `JSON.stringify(input)` truncated to 150 chars (`shared/chat-ui/chat-ui.ts:342`), which for a large `set_range` is unreadable and unrecoverable.

- [ ] **Step 1:** Render the full input into the same disclosure region, above the output, as a second `<pre class="ai-step-input">` with a small "Input" / "Output" label on each. Keep the truncated inline form in the title as the collapsed summary.
- [ ] **Step 2:** Pretty-print with `JSON.stringify(input, null, 2)`.
- [ ] **Step 3:** Add a test that a large input is fully present in the expanded region even though the title is truncated.

Skip this task if scope needs trimming; it does not block PP-3's goal.

---

### Task 5: Wire-through check and manual verification

**Files:** none modified (verification only)

- [ ] **Step 1:** Confirm `execution.output` is non-empty for read tools end-to-end: `WordTools.GetDocumentContext` builds `"Paragraphs: N, Words: M\nPreview: ..."` (`WordAiAddIn/WordTools.cs:100-109`), `ToolProtocol.SerializeToolResult` puts it on the wire as `output`, and `entry.ts`'s `onToolExecuted` passes `event.execution.output` straight into `complete` (`WordAiAddIn/web-src/entry.ts:335-340`). If the agent-core `ToolExecution` type drops it anywhere in between, fix that instead of working around it in the UI.
- [ ] **Step 2:** Rebuild all three bundles and MSBuild all three projects.
- [ ] **Step 3:** Manual in real Word: ask a question that triggers `get_document_context`; expand the step; confirm the paragraph/word counts and preview text are visible and match the document. This directly closes out the "returns empty" report — either the output is there (tool was fine, visibility was the bug) or it is genuinely empty (a real `WordTools` bug to file separately).
- [ ] **Step 4:** Manual in Excel: `read_range` over ~200 cells — confirm truncation + "show all" works and the panel stays scrollable.
- [ ] **Step 5:** Manual: force a tool error (e.g. `read_blocks` with an out-of-range index) and confirm the error text is visible without any clicking.
