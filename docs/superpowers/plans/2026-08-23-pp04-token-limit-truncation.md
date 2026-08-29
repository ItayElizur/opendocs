# PP-4: No More Silent "(no text)" Replies — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-4 (P0).

> **Updated 2026-08-24, post-PP-0:** `2026-08-23-pp00-shared-app-shell.md` has landed. `MAX_TOKENS`, `makeTransport`, and the `onDone` handler this plan targets now live once in `shared/web-src/app-shell/settings.ts` and `shared/web-src/app-shell/bootstrap.ts` respectively, not in three `entry.ts` files. Tasks 1-3 are rewritten as single-file edits below; this is exactly the win PP-0 was for.

**Goal:** A turn cut off by the output-token limit is never shown as an unexplained empty reply. Give the model enough budget that ordinary requests finish, and when a turn is genuinely truncated, say so in the transcript with a way to continue — matching how the loop already handles the analogous max-*turns* case with `TURN_LIMIT_NOTE`.

**Architecture:** Three layers, each independently wrong today.

1. **Budget.** `MAX_TOKENS = 1024` is hardcoded per turn, now in one place (`shared/web-src/app-shell/settings.ts`, previously duplicated across all three apps' `entry.ts`). A modern model that reasons before answering can spend the whole budget without emitting visible text — which is exactly why "summarize the *key points*" fails where "summarize the content" succeeds. 1024 is far below what any supported provider charges for or limits to.
2. **Detection.** It is already detected and already discarded. `stream.ts` maps `finish_reason: length` to `stopReason: 'max_tokens'` (`shared/web-src/ai-provider/stream.ts:819-826`), deliberately treating it as a *normal* stop reason rather than the abnormal-finish error path (correct — a truncated turn with partial text is still usable). `loop.ts` turns that into `result.truncated: true` on `onDone` (`shared/web-src/agent-core/loop.ts:523-529`). And the shell's `onDone` handler (`shared/web-src/app-shell/bootstrap.ts`, formerly duplicated in each `entry.ts`) ignores the flag entirely: `const finalText = result.text || '(no text)'`.
3. **Recovery.** There is none. `turnLimit` at least injects `TURN_LIMIT_NOTE` and gets a partial answer (`shared/web-src/agent-core/loop.ts:611-615`); `max_tokens` just ends the run.

The fix lands one change per layer: raise and centralize the budget; surface truncation honestly in the UI with a localized string instead of the hardcoded English `'(no text)'`; and offer a one-click continue.

**Tech Stack:** TypeScript across `shared/web-src/agent-core`, `shared/web-src/ai-provider`, `shared/chat-ui`, and `shared/web-src/app-shell`. Vitest for `chat-ui`.

## Global Constraints

- Do not change `stream.ts`'s classification of `finish_reason: length` as a non-error stop reason. A truncated turn that *did* emit text must keep that text; routing it to the error path would throw the partial answer away and regress a working case.
- Do not change `AgentRunResult`'s shape beyond what already exists — `truncated?: boolean` is already declared (`shared/web-src/agent-core/loop.ts:22-30`) and already set. Consumers only need to read it.
- Every user-visible string added to the UI goes through `chat-ui.ts`'s `STRINGS` table (`shared/chat-ui/chat-ui.ts:7-31`) with both `en` and `he` entries. No new hardcoded English in the panel — that is the specific defect `'(no text)'` represents.
- Rebuild each add-in's bundle and re-run MSBuild after any change (command in `2026-08-23-pp02-tool-steps-chronological-order.md`'s Global Constraints); bundles are gitignored.
- Task 3's "continue" affordance must not silently re-run the user's original prompt — that would re-execute mutating tools against an already-edited document. It sends an explicit continuation instruction instead.

---

### Task 1: Raise and centralize the per-turn token budget

**Files:**
- Modify: `shared/web-src/app-shell/settings.ts`

*(If PP-0 has not landed: modify `MAX_TOKENS` identically in `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts`.)*

**Interfaces:**
- Produces: a documented `MAX_TOKENS` value used by all three apps through the shared transport.

- [ ] **Step 1:** Raise `MAX_TOKENS` from `1024` to `8192` in `settings.ts` — one edit, not three — with a comment recording why:

```ts
// Per-turn output budget. 1024 was too small for models that spend budget on
// reasoning before emitting visible text: the provider returned
// finish_reason=length with zero content and the run ended in an unexplained
// empty reply (PP-4). 8192 is comfortably above a long tool-using turn's real
// output while staying well under every supported provider's per-request cap.
const MAX_TOKENS = 8192
```

- [ ] **Step 2:** Sanity-check the value against each provider's limits before committing to it. `shared/web-src/ai-provider/providers.ts` lists the supported providers/models; confirm 8192 is a legal `max_tokens` for the smallest-limit model in that list (in particular any older OpenAI-compatible endpoint), and lower it if not. Record the checked value in the comment.

- [ ] **Step 3:** Decide whether the budget should be user-configurable from Settings. **Recommendation: no** — it is a footgun with no good default the user could pick better than the code can, and Settings is already growing in PP-6. Note the decision in the comment so it is not revisited blindly.

**Verification:** all three bundles rebuild (one source edit still means three separate esbuild invocations — the shell is bundled into each app); a prompt that previously produced "(no text)" (e.g. "summarize the key points of this document" on a multi-page document) now produces a real answer.

---

### Task 2: Surface truncation honestly instead of "(no text)"

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts`

*(If PP-0 has not landed: modify `onDone` identically in all three `entry.ts` files.)*

**Interfaces:**
- Produces: `ChatUIHandle.showNotice(kind: 'truncated' | 'turnLimit', onContinue?: () => void): void` — a non-error, in-transcript informational row. Task 3 passes `onContinue`.

- [ ] **Step 1: Strings**

Add to `STRINGS` (`shared/chat-ui/chat-ui.ts:7-31`):

```ts
noticeTruncated: { en: 'The reply was cut off by the length limit.', he: 'התשובה נקטעה עקב מגבלת האורך.' },
noticeTurnLimit: { en: 'The tool-step limit for this request was reached.', he: 'הגעת למגבלת שלבי הכלים לבקשה זו.' },
noticeContinue:  { en: 'Continue',                                     he: 'המשך' },
emptyReply:      { en: '(no reply)',                                   he: '(אין תשובה)' },
```

- [ ] **Step 2: `showNotice`**

Implement on the returned handle, modeled on the existing `showError` (`shared/chat-ui/chat-ui.ts:370-376`) but with a distinct `.ai-msg-notice` class (informational, not red), a `data-t` attribute so the existing `applyStrings`/`setLang` machinery relocalizes it on a language switch, and — when `onContinue` is passed — a `.ai-notice-action` button wired to it that removes itself on click so it cannot be double-fired.

- [ ] **Step 3: CSS** — add `.ai-msg-notice` and `.ai-notice-action` to `shared/chat-ui/chat-ui.css` using existing tokens; muted/neutral, visually distinct from `.ai-msg-error`.

- [ ] **Step 4: Rewrite `onDone` in the shared shell**

Replacing the `onDone` handler in `startAddIn`'s `AgentLoop` construction (`shared/web-src/app-shell/bootstrap.ts`):

```ts
onDone: (result) => {
  const hasText = result.text.length > 0
  const finalText = hasText ? result.text : t('emptyReply')
  ui.endAssistantMessage(finalText)
  if (result.truncated) ui.showNotice('truncated', () => continueRun())
  else if (result.turnLimit) ui.showNotice('turnLimit')
  ui.setBusy(false)
  persistMessage('assistant', finalText)
},
```

Note the two behavior changes beyond the notice: the fallback is localized, and — importantly — a truncated turn that *did* stream partial text keeps that text and gets the notice appended below it, rather than being treated as empty.

- [ ] **Step 5:** `t()` here must read the panel's current language. If `chat-ui.ts` does not already export a lookup usable from `bootstrap.ts`, export a small `translate(key: string): string` bound to the mounted panel's current lang — one export, consumed once by the shell, rather than duplicated per app.

- [ ] **Step 6:** `persistMessage` writes the assistant text to `ChatStore` via the `append-message` bridge kind. Persist only the model's actual text (or the localized empty marker) — do **not** persist the notice, which is UI state about one run, not conversation content.

**Verification:**
- [ ] `cd shared/chat-ui && npx vitest run` — plus a new test that `showNotice('truncated', fn)` renders a `.ai-msg-notice` with a button that calls `fn` once.
- [ ] Manual: temporarily set `MAX_TOKENS = 32` in one app, rebuild, send any prompt → the transcript shows partial text (if any) plus the truncation notice, never a bare "(no text)". Restore the value afterwards.

---

### Task 3: One-click continue after truncation

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`

*(If PP-0 has not landed: modify all three `entry.ts` files identically.)*

**Interfaces:**
- Consumes: `ui.showNotice`'s `onContinue` from Task 2.
- Produces: `function continueRun(): void` inside `startAddIn`.

- [ ] **Step 1:**

```ts
// Sends a continuation instruction rather than re-running the user's original
// prompt: the run may already have applied mutating tools to the document, and
// replaying the prompt would apply them a second time.
const CONTINUE_INSTRUCTION =
  'Your previous reply was cut off by the length limit. Continue exactly where it stopped. ' +
  'Do not repeat what you already wrote and do not re-apply any edits you already made.'

function continueRun(): void {
  if (loop.busy) return
  ui.beginAssistantMessage()
  ui.setBusy(true)
  loop.run(CONTINUE_INSTRUCTION)
}
```

- [ ] **Step 2:** Confirm `AgentLoop.run()` appends to the existing history rather than resetting it (read `shared/web-src/agent-core/loop.ts` around `run`/`reset`) — continuation depends on the cut-off assistant turn still being in history. `finishTurn` pushes `this.turnText || COMPLETED_VIA_TOOLS_TEXT` into history before `onDone` (`shared/web-src/agent-core/loop.ts:521`), so the partial text is preserved; verify the empty-truncation case ends up with `COMPLETED_VIA_TOOLS_TEXT` rather than an empty assistant message, since an empty one is exactly what the comment there says poisons follow-ups.
- [ ] **Step 3:** Do not persist `CONTINUE_INSTRUCTION` as a user message in `ChatStore` — it is machine-generated plumbing, and showing it in restored history would be confusing. Add a comment saying so.
- [ ] **Step 4:** Disable/hide the continue button while `loop.busy` so a double-click cannot start two runs.

**Verification:** with `MAX_TOKENS` temporarily lowered, click Continue and confirm the reply resumes rather than restarting, and that no tool re-applies an edit it already made.

---

### Task 4: Guard against the regression returning

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`
- Modify: `shared/web-src/agent-core/loop.ts` (comments only)

- [ ] **Step 1:** Add a comment at `loop.ts:523-529`, where `truncated` is set, noting that the shared shell's `onDone` handler (`shared/web-src/app-shell/bootstrap.ts`, consumed by all three apps) depends on this flag and must surface it — so a future refactor doesn't drop it as "unused".
- [ ] **Step 2:** If the repo has (or can cheaply get) an agent-core test harness with a fake transport, add a test that a transport emitting `onStopReason('max_tokens')` with no deltas produces `onDone({ text: '', truncated: true })`. If no such harness exists, do not build one for this — note it as untested and rely on the `chat-ui` test plus manual verification.
- [ ] **Step 3:** Grep for any other `'(no text)'` literal across the repo and remove each one the same way.

**Verification:** `grep -rn "(no text)" --include=*.ts .` returns nothing outside build artifacts.
