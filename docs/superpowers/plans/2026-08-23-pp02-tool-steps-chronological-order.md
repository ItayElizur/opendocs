# PP-2: Tool Steps Render in Chronological Order — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-2 (P0).

> **Updated 2026-08-24, post-PP-0:** `2026-08-23-pp00-shared-app-shell.md` has landed. The `onSend`/`onToolStart`/`onTurnEnd` wiring this plan originally cited in three `entry.ts` files now lives once in `shared/web-src/app-shell/bootstrap.ts` (line numbers below are stale; the logic is unchanged). Task 3 is rewritten accordingly — it is now a single-file edit, not three.

**Goal:** Make the chat transcript read forwards — every tool-call group appears above the assistant text that depended on its results, for every turn of a multi-turn run — instead of the current inverted order where the final answer sits above the work that produced it.

**Architecture:** The inversion is purely a DOM-append-order artifact. `onSend` calls `ui.beginAssistantMessage()` *at send time* (originally `WordAiAddIn/web-src/entry.ts:289` et al., now `shared/web-src/app-shell/bootstrap.ts`'s `onSend`), which appends an empty assistant bubble immediately; tool groups are appended lazily on the first `onToolStart` (same file). The bubble is created first, so it is positioned first, and `updateAssistantMessage` later overwrites its text in place with the *final* answer.

The fix is to stop pre-creating the bubble and instead create it on first text delta, so DOM order matches causal order naturally. That requires `chat-ui.ts` to own bubble lifecycle rather than have it driven from outside: `beginAssistantMessage()` becomes a no-op-safe "arm" (it shows the busy/typing affordance without appending a bubble), `updateAssistantMessage(text)` lazily appends the bubble on the first call, and — critically for multi-turn runs — `beginToolGroup()` **closes the current bubble** so the next turn's text starts a fresh bubble below the group.

Multi-turn runs are the case that makes a naive "just move the `beginAssistantMessage()` call" fix wrong: a run can be text → tools → text → tools → text, and today `assistantBubble` is a single element reused across all of it. This plan makes each contiguous text segment its own bubble.

**Tech Stack:** TypeScript (esbuild-bundled), shared `shared/chat-ui/chat-ui.ts` component with vitest/jsdom coverage in `shared/chat-ui/chat-ui.test.ts`.

## Global Constraints

- The `ChatUIHandle` public method names (`beginAssistantMessage`/`updateAssistantMessage`/`endAssistantMessage`/`beginToolGroup`) must not change — `shared/web-src/app-shell/bootstrap.ts` calls them (once, for all three apps, post-PP-0) and this plan touches call sites minimally.
- UI chrome colors come from the existing CSS custom properties in `shared/chat-ui/chat-ui.css` — no new raw hex values outside a token definition line.
- Keep the `dir="auto"` per-message bidi behavior (`shared/chat-ui/chat-ui.ts:290-301`) on every bubble this plan creates — it is load-bearing for Hebrew.
- Rebuild each add-in's bundle and re-run MSBuild after any `chat-ui.ts` or `entry.ts`/`bootstrap.ts` change, in each of `WordAiAddIn/`, `ExcelAiAddIn/`, `PowerPointAiAddIn/`:
  `npx esbuild web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap`
  then `MSBuild <App>/<App>.csproj -t:Build -p:Configuration=Debug`. A stale bundle silently ships the old behavior; bundles are gitignored, do not commit them. (This is the canonical copy of the command — several other plans point back at this line rather than repeating it.)
- This plan changes ordering only. It must not change what text is displayed (that is PP-4) or add tool-output display (that is PP-3) — but Task 1's lifecycle rework is the foundation both of those build on, so land this first.

---

### Task 1: Lazy assistant-bubble creation in `chat-ui.ts`

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`

**Interfaces:**
- Consumes: nothing new.
- Produces: unchanged `ChatUIHandle` surface with new semantics — `beginAssistantMessage()` arms without appending; `updateAssistantMessage()` appends on first call; `beginToolGroup()` seals the open bubble. PP-3 and PP-4 both build on this.

- [ ] **Step 1: Add a "seal the current bubble" helper**

Alongside the existing `let assistantBubble` state near the returned handle (`shared/chat-ui/chat-ui.ts:305-336`):

```ts
function sealAssistantBubble(): void {
  if (assistantBubble) {
    assistantBubble.classList.remove('streaming')
    // An empty bubble means the turn produced tool calls but no prose — drop
    // the element entirely rather than leaving an empty box in the transcript.
    if (!assistantBubble.textContent) assistantBubble.remove()
  }
  assistantBubble = null
}
```

- [ ] **Step 2: Make `beginAssistantMessage` arm rather than append**

```ts
beginAssistantMessage() {
  // Deliberately does NOT append a bubble: the bubble is created lazily on the
  // first text delta (updateAssistantMessage), so a turn's tool-call group —
  // appended when its first tool starts — lands ABOVE the text that depended
  // on it. Pre-appending here is what made the transcript read backwards.
  sealAssistantBubble()
  scrollToBottom()
},
```

- [ ] **Step 3: Make `updateAssistantMessage` create on demand**

```ts
updateAssistantMessage(cumulativeText) {
  if (!assistantBubble) {
    assistantBubble = renderMessage('assistant', '')
    assistantBubble.classList.add('streaming')
  }
  assistantBubble.textContent = cumulativeText
  scrollToBottom()
},
```

`renderMessage` already sets `dir = 'auto'` and clears the empty state, so nothing else is needed.

- [ ] **Step 4: `endAssistantMessage` finalizes, creating a bubble if the run produced text but never streamed a delta**

```ts
endAssistantMessage(finalText) {
  if (!assistantBubble && finalText) {
    assistantBubble = renderMessage('assistant', '')
  }
  if (assistantBubble) assistantBubble.textContent = finalText
  sealAssistantBubble()
  scrollToBottom()
},
```

The `!assistantBubble && finalText` branch covers a non-streaming transport (or a turn where the whole text arrived after tools) — without it, such a run would show no answer at all.

- [ ] **Step 5: `beginToolGroup` seals first**

At the top of `beginToolGroup()` (`shared/chat-ui/chat-ui.ts:327`), before creating `groupEl`:

```ts
// Close out any text streamed earlier in this run so the group appends BELOW
// it and the next turn's text starts a fresh bubble BELOW the group —
// preserving true chronological order across multi-turn runs.
sealAssistantBubble()
```

**Verification:**
- [ ] `cd shared/chat-ui && npx vitest run` — existing tests pass (some will need the updates in Task 2).

---

### Task 2: Ordering tests

**Files:**
- Modify: `shared/chat-ui/chat-ui.test.ts`

**Interfaces:** consumes the Task 1 handle; produces no new exports.

- [ ] **Step 1: Single-turn ordering test**

Following the existing `setup()` helper's style (`shared/chat-ui/chat-ui.test.ts:5-21`):

```ts
it('renders a tool group above the assistant text that follows it', () => {
  const { root, handle } = setup()
  handle.addUserMessage('do the thing')
  handle.beginAssistantMessage()
  const group = handle.beginToolGroup()
  group.addStep('read_blocks', { startIndex: 0, endIndex: 5 }).complete({ output: 'ok' })
  group.end()
  handle.updateAssistantMessage('Here is the summary')
  handle.endAssistantMessage('Here is the summary')

  const nodes = Array.from(root.querySelectorAll('.ai-work-group, .ai-msg-assistant'))
  expect(nodes.map((n) => n.className.split(' ')[0])).toEqual(['ai-work-group', 'ai-msg-assistant'])
})
```

- [ ] **Step 2: Multi-turn ordering test** — text, tools, text, tools, text produces `assistant, group, assistant, group, assistant` in that DOM order.

- [ ] **Step 3: No-prose turn test** — `beginAssistantMessage()` then a tool group then `endAssistantMessage('')` leaves no empty `.ai-msg-assistant` element in the transcript.

- [ ] **Step 4: Non-streaming test** — `beginAssistantMessage()` then `endAssistantMessage('final')` with no `updateAssistantMessage` call still renders one bubble containing `final`.

- [ ] **Step 5:** Update any existing test that asserted a bubble exists immediately after `beginAssistantMessage()`.

**Verification:** `cd shared/chat-ui && npx vitest run` — all green.

---

### Task 3: Align the call site in the shared shell

**Files:**
- Modify: `shared/web-src/app-shell/bootstrap.ts`

**Updated post-PP-0:** this was originally three edits (one per `entry.ts`); it is now one, in the shared shell that all three apps call into. If PP-0 has *not* landed when this plan is picked up, fall back to editing `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts` identically, matching the shapes below.

- [ ] **Step 1:** Leave `ui.beginAssistantMessage()` in `onSend` where it is (in `startAddIn`'s `mountChatUI` options) — with Task 1's semantics it now correctly means "arm for a new assistant turn" and appends nothing. Add a short comment at the call site pointing at that, so nobody "fixes" it back by moving it.

- [ ] **Step 2:** Verify the `onTurnEnd` handler still only ends the group and nulls `currentToolGroup` — with Task 1 in place, the *next* turn's `beginToolGroup()` re-seals, so no extra call is needed here.

- [ ] **Step 3:** Rebuild all three bundles and MSBuild all three projects per the Global Constraints — one source edit, three rebuilds (the shell is bundled separately into each app).

**Verification:**
- [ ] All three add-ins build.
- [ ] Manual, in real Word: ask "summarize this document" and watch the transcript. Tool steps appear first, answer appears below them, and the answer does not jump above the steps when it finishes streaming.
- [ ] Manual, a request forcing several tool rounds (e.g. "read the first 10 paragraphs, then bold every heading, then tell me what you changed"): the transcript alternates text/group/text/group in the order the events actually happened.
- [ ] Manual, Excel and PowerPoint: same spot check.
