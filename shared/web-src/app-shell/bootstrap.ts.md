# bootstrap.ts — relocated comment history / rationale

## `SelectionContext` (type)

```
FT-2 Task 4: what the user has selected, classified from the raw bridge
payload into the vocabulary each app's tools actually take - Word's
paragraph-index tools get content, Excel's A1-addressed tools get an
address, PowerPoint's slideIndex/shapeIndex tools get indices. The `range`
variant carries more than the plan's minimal sketch (entireColumns/
entireRows/effectiveAddress/effectiveCellCount) because describeSelection's
whole-column/whole-row wording needs them - there is nowhere else for that
data to live.
```

## `defaultDescribeSelection`

```
FT-2 Task 4 Step 2: the per-turn sentence injected via buildContext() -
must state the addressing vocabulary explicitly (an A1 address or a
slideIndex/shapeIndex pair) so the selection is actionable, not merely
informative. Kind-keyed rather than per-app-configurable because each app
only ever produces its own subset of kinds (Word: text/none; Excel:
range/none; PowerPoint: slides/shapes/shapeText/none) - AddInConfig.
describeSelection lets an app override this if it ever needs to.
```

## `defaultDescribeSelection` 'text' case — object-selection check ordering

```
Post-hoc addition (2026-08-24, user-reported: selecting a table/
chart/SmartArt "doesn't appear under selection") - these objects'
own selection.Text is empty or a placeholder character, so this
must be checked before the ctx.fullText emptiness check below, or
an object selection would always fall through to "no selection".
```

## `defaultDescribeSelection` 'text' case — paragraph-range addressability

```
Post-hoc fix (2026-08-24, user-reported): previously gave only the
text with no addressability, so a request to transform the
selection in place (e.g. "translate this paragraph") had no way to
target replace_blocks at exactly the selected paragraphs and fell
back to insert_content, appending a new paragraph instead of
replacing the original. Now states the 0-based paragraph range
explicitly, matching FT-2's addressable wording for Excel/PowerPoint.
```

## `toSelectionScopeUpdate`

```
FT-2 Task 5: classifies the raw payload into the UI's SelectionExtent -
chat-ui.ts owns rendering/localizing the words, this only picks which case
applies and extracts the numbers, per Task 5 Step 3 ("a label the shell
computes"). Returns null for "no selection" (reverts the pill to its
per-app whole-scope label).
```

## `toSelectionScopeUpdate` — Word object-kind extent pill

```
Word (post-hoc addition, 2026-08-24): a table/chart/SmartArt selection
renders as a proper extent pill (e.g. "Table 2") instead of the
quoted-text form, which would otherwise show empty/placeholder text for
these object kinds - the exact "no pointer" gap the user reported.
```

## `AddInConfig.trackChangesExtraTools`

```
Additionally available in Track changes mode, ON TOP OF Comment only's
set (not a replacement) - `null`/absent (every app but Outlook) keeps
Track changes' original meaning: every tool, since Word/Excel/
PowerPoint's real edit tools are legitimately usable under native
track-changes recording. Outlook sets this to unlock
respond_meeting on top of its own
"Draft only" (commentOnly) tier - see availableForMode() below.
```

## `AddInConfig.availableModes`

```
Restricts the editing-mode menu to this subset, in this order. Defaults to
all four modes (Word/Excel/PowerPoint). Outlook passes all four too, with
commentOnly/trackChanges repurposed as "Draft only"/"Automate approvals"
(see modeOverrides) rather than their Word-ish original meaning.
```

## `todayContextLine`

```
Fix for: relative-date tool args (draft_event, find_meeting_slots, etc.)
were resolved by the model with no ground truth for "today" anywhere in
the system prompt - it had to guess both the date and the weekday from
training data, which is exactly how "next Tuesday" turned into
Wednesday. Called fresh from systemSuffix() below on every turn, not
frozen at conversation start: the full system prompt is already resent
on every turn regardless (see loop.ts's startTurn()), so recomputing this
~20-token line costs nothing extra, and it keeps a conversation that
spans midnight (or a laptop that sleeps overnight) correct instead of
stuck on its start-of-chat date.
```

## `startAddIn`

```
Boots one add-in's chat panel: WebView2 bridge, settings, transport,
chat-UI mount, and AgentLoop event plumbing. Everything here was
previously duplicated near-verbatim across WordAiAddIn/ExcelAiAddIn/
PowerPointAiAddIn's entry.ts (PP-0) - each app now supplies only what is
genuinely app-specific through `config`.
```

## Editing-mode control (`editingMode` initialization)

```
Task 11 (Word): editing-mode control. Client-side filtering only (first
line of defense - smaller prompts, fewer wasted turns); the real
enforcement is server-side in each app's *Tools.Execute, which gates
mutating tool calls even if the model somehow requests one that wasn't
offered here.
A fresh session no longer starts in Full Autonomy by default - see
defaultModeFor() in chat-ui.ts. This must resolve identically to
mountChatUI's own `defaultMode` computation below (same helper, same
inputs) or the UI's initial selection and this filtering state disagree
about what's actually available before the user ever touches the mode
menu.
```

### The starting mode must be pushed to C# too

`postMode(editingMode)` runs once right after `mountChatUI`, in addition to
`onModeChange`. Before that line existed, C# only learned about a mode *change*:
a pane that was never touched still showed "Track Changes" (the default) while every
app's `*Tools.ModeFor(docKey)` found no entry and returned its Full Autonomy fallback.
For Word that meant `TrackRevisions` was forced off, so edits were applied untracked
until the user flipped the mode once; the server-side gate was also looser than the UI
claimed. The C# fallback is still FullAutonomy: this change makes the client always say what it
shows, and leaves the fail-open default alone to stay minimal. Nothing in the repo's
tests depends on it (the COM tool classes have no unit coverage), so failing closed would
be a reasonable separate, behaviour-changing follow-up.
Caveat: Word keys the mode by `GetChatId()`, which re-keys once when an unsaved document
is first saved, so a mode set before that first save is not carried over by the key
change (not addressed here).

## Document guidelines message (`savedDocMessage`/`activeDocMessage`)

```
FT-1 Task 8: the document guidelines message. `savedDocMessage` is
whatever is currently persisted on disk (kept in sync by the bridge's
doc-settings-loaded response and by a successful Save); `activeDocMessage`
is what actually gets injected into the system prompt this conversation -
frozen by beginConversation() at conversation-start boundaries only
(initial load, New chat), never read live per-turn, so editing the
guidelines mid-conversation cannot retroactively change a run in progress.
The date context line is the opposite: recomputed live on every turn by
systemSuffix() below (todayContextLine()'s own comment explains why),
not frozen here.
```

## Theme reconciliation (`currentThemePref`/`lastKnownOfficeTheme`)

```
Theme reconciliation. `currentThemePref` is the persisted 3-way choice;
`lastKnownOfficeTheme` is Office's real theme, read exactly once from the
registry (via requestOfficeTheme() below) and cached for this pane's
whole lifetime - not re-checked while the pane stays open, by design.
C# never sees `currentThemePref`; it only ever reports "what does Office
look like right now", and this file alone decides whether that answer
gets applied (only when the preference is 'default').
```

## Language reconciliation (`currentLangPref`/`lastKnownOfficeLanguage`)

```
Language reconciliation - exact same shape as theme above.
`currentLangPref` is the persisted 3-way choice; `lastKnownOfficeLanguage`
is Office's real UI display language, read exactly once (via
requestOfficeLanguage() below, see OfficeAi.Shared/OfficeLanguage.cs) and
cached for this pane's whole lifetime. C# never sees `currentLangPref`;
it only ever reports "what language is Office's own UI in right now",
and this file alone decides whether that answer gets applied (only when
the preference is 'default').
```

## `onStop` wiring

```
Post-hoc addition (2026-08-24, user-requested): AgentLoop.cancel()
already existed (its own comment even anticipated "when the user
clicks stop") but was never reachable from any UI control until now -
wired here, same forward-reference-via-closure pattern onSend already
uses for `loop` below (declared further down this file).
```

## `onSend` — queueing while busy

```
Post-hoc change (2026-08-24, user-requested): previously a no-op
while busy (the textarea used to be disabled too, so this was
unreachable anyway). Now the textarea stays enabled during a run,
so a send while busy queues the message instead of dropping it -
dispatched automatically once the current run finishes (onDone
below), whether it finished normally or was stopped. The user's
message is shown and persisted immediately (chronologically
accurate - they sent it now), only the actual model run is
deferred; `pendingQueuedText` is declared further down this file
(same forward-reference-via-closure pattern already used for `loop`).
```

## `onSettingsSave` — lang persistence gap

```
lang used to be a pre-existing gap here - threaded into the payload
but never persisted or applied. Now folded into setSettings() above
and applied immediately, same as theme, so Save's effect (including
picking 'default') is visible without waiting for anything async.
```

## Theme flash avoidance (after `mountChatUI`)

```
Same idea for theme: an explicit Light/Dark preference is known
synchronously from localStorage, so apply it immediately to avoid a
flash of the wrong theme. 'default' has no synchronous answer (only C#
knows Office's real theme) - left as-is until the async reply below
resolves it; a one-frame flash there is accepted, not fixable without
delaying first paint.
```

## `currentToolGroup` grouping across back-to-back tool-only turns

```
Post-hoc fix (2026-08-24, user-reported): loop.ts's "turn" is one model
response - for a model that chains several tool calls back-to-back with
no text between them, that is often exactly one tool call per turn, so
closing/nulling currentToolGroup on every onTurnEnd split a single
logical batch into a separate "Ran 1 tool" box per call instead of one
group incrementing to "Ran N tools". The group should only actually
close when text has genuinely streamed since it opened (so a LATER
block of tools still gets its own group below that text, preserving
chronological order) - tracked here since only bootstrap.ts sees both
onText and onTurnEnd.
```

## `dispatchQueuedIfAny`

```
Post-hoc addition (2026-08-24, user-requested): dispatches a message
queued via onSend while the previous run was busy - the user bubble and
ChatStore persistence already happened at queue time, so this only
needs to actually start the model run. Called from onDone/onError below
regardless of how the prior run ended (finished normally, truncated, or
stopped via the stop button) - a queued message should still go out.
```

## `onDirectCall` (the "!" palette)

Runs a tool the user typed by hand through `callDotNetTool` directly (the same bridge call the model's tools use, but not `executeConfirmed`, because the user wrote the call themselves), shows it with the
ordinary user bubble and tool-step row, and then records the exchange into the loop's history with
`AgentLoop.recordToolExchange` (user text, assistant tool call, tool result, a closing assistant marker) so the next model turn
sees it. Refused while a run or another direct call is in flight (the text stays in the box); a normal send during a direct
call is queued like a send during a run; Continue is refused with the same busy message. The typed line and a one-line assistant
reply (`name: first 200 chars of the output`) are persisted to the ChatStore: `AgentLoop.restore` drops a user line that has no
reply, so without the reply a reloaded conversation would lose the call. New chat bumps `directCallGeneration` and clears
`directCallRunning`, so a call that finishes afterwards is ignored (as `AgentLoop` does for a reset model run) and a call that
never replies cannot wedge the panel. Tool arguments and results are never redacted; the loop's `redact` option (see
`redactModelKeys`) removes only the user's configured model API key from what they type. The palette's pill needs each tool's lowest mode, computed by `minModeOf` from the same
sets `availableForMode()` uses.

## `executeConfirmed`, `needsApproval` and `confirmActions` (Approve before a call)

`skill.executeTool` is `executeConfirmed`. The administrator's `policy.json` decides when the user is asked
(`requireApproval`, injected by the host as `window.__opendocsApprovals`; see `OfficeAi.Shared/ModePolicy.cs.md`).
`needsApproval` checks the tool's own entry, or `"*"` for any tool that is not read-only, against the current mode. If
the policy names a tool this app does not have, every non-read-only tool asks in every mode, and `showPolicyNotices`
says why. There is no user setting for this.

How the card looks comes from `config.confirmActions`. The first `ConfirmRule` that matches wins, and `when` narrows a
rule by the call's arguments. A rule with `preview` first calls the app's `preview_action` read tool. A `blocked`
answer is returned to the model as the result, and no card is shown. Otherwise the answer's type, subject, sender and
date become the card's rows; a failed preview is ignored, so the card shows only the question. The rule's `fields` add
rows from the call's own arguments. A tool with no rule gets a plain card: "Allow this action?", the tool's display
name, and up to four simple arguments, each cut to 120 characters. A rule's card never shows the body or the tool call
itself.

New chat always calls `ui.setBusy(false)`, because `loop.reset()` ends a run without an `onDone` and the stop button
would otherwise stay up. The `AbortSignal` the loop passes in is handed to the card: `AgentLoop.cancel()` and `reset()`
both abort it, so Stop and New chat deny a pending card and nothing runs. A denied call returns an error result,
`DECLINED_TEXT`, which tells the model not to retry and to ask the user. The mode is read per call, so a mode change
applies at once. The `!` palette bypasses this on purpose (see `onDirectCall`).
