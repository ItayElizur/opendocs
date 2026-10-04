# chat-ui.ts — relocated comment history / rationale

Comments moved here verbatim per the repo-wide comment cleanup sweep. Each heading names the symbol/line the comment originally sat above or next to; the current file keeps a short trimmed summary in place.

## `SelectionExtent` (type)

```
FT-2 Task 5: the structured facts of an Excel/PowerPoint selection - the
shell (bootstrap.ts) classifies the C# payload into one of these; this
component renders and localizes the actual words, the same way every other
piece of UI text here does, so the pill relocalizes on a language switch
for free (via refreshScopeHint(), already called from setLang()).
```

## `ChatUIOptions.onStop`

```
Post-hoc addition (2026-08-24, user-requested): stops the in-flight run. The send button becomes a stop button while busy; omit to leave it disabled while busy instead (previous behavior).
```

## `ChatUIOptions.modeOverrides`

```
Per-app override of a mode's label/description shown in the mode menu
and settings scope control - falls back to the shared STRINGS entry
(modeReadOnly/modeCommentOnly/modeTrackChanges/modeFullAutonomy and
their *Desc counterparts) when a mode has no override. Outlook uses
this to relabel commentOnly/trackChanges as "Draft only"/"Automate
approvals" (mail has no real "comment" or "track changes" concept) and
to note that Full autonomy sends mail/invites - Word/Excel/PowerPoint
leave this unset and keep the shared generic copy.
```

## `ChatUIHandle.setSelectionScope`

```
`preview` is Word's existing quoted-text-excerpt form (wrapped in the
localized `scopeSelectionPrefix` + `..."`). `extent` is Excel/PowerPoint's
form (FT-2 Task 5) - a structured, already-classified selection that this
component renders and localizes itself, shown as-is with no quoting/
ellipsis (an address is not a text excerpt). At most one of the two is
set at a time.
```

## Caret visual-line measurement section banner (above `MIRROR_CSS_PROPS`)

```
---- Caret visual-line measurement, for the ArrowUp/ArrowDown history-recall
gate (caretCollapsedAtFirstLine/caretCollapsedAtLastLine in mountChatUI
below). A message with no literal '\n' can still word-wrap into several
*visual* lines inside the textarea (.ai-textarea has no white-space:
nowrap) - checking only for '\n' (the old implementation) can't tell those
visual lines apart, so it always reported "on the first/last line", and
the very first ArrowUp/ArrowDown always recalled history instead of first
moving the caret up/down within the wrapped draft. This mirrors the
standard "textarea-caret-position" technique: clone the textarea's box
model into a hidden same-width div, insert the value up to a given
position plus a marker span, and read where that marker actually rendered.
```

## `measureCaretLineTop`

```
Returns the pixel offsetTop of the visual line that position `pos` in
`textarea.value` renders on. Two positions on the same visual line always
return the same number, two positions on different visual lines never do -
regardless of whether the line break between them is a hard '\n' or a
soft word-wrap - so callers compare this against the offset of position 0
(or value.length) rather than needing to separately count lines. Copies
direction/textAlign from the textarea's *computed* style, which already
reflects dir="auto" resolution (see updateTextareaDir) - so RTL messages
measure correctly too.
```

## `caretLineMeasurement` (exported test seam)

```
Exposed purely so tests can stub the measurement: jsdom (used by
chat-ui.test.ts) does not perform real text layout, so offsetTop on a
jsdom-rendered element is always 0 - measureCaretLineTop's *output* can't
be asserted against real pixel positions in that environment. Production
code always goes through this object's `measure`, calling the real
mirror-div implementation above; tests instead reassign `.measure` to a
fake per-position line map so the surrounding gating logic
(caretCollapsedAtFirstLine/caretCollapsedAtLastLine) - the actual bug fix -
can still be exercised deterministically.
```

## `chipDockHtml`

```
The reopened-conversation chip dock (see showHistoric()) - lives OUTSIDE
the scrolling .ai-chat flex column entirely (a sibling in the panel
skeleton), so it can never be crushed toward zero height the way
.ai-chat-empty was when appended inside .ai-chat after a divider (that
element's `flex: 1` + `overflow: hidden` gives it a zero automatic
minimum size once the transcript above it already fills the pane - see
the fix's PR description for the full flexbox explanation). Reuses the
same options.starters data as emptyStateHtml, not a second copy of it,
plus one extra "New conversation" chip.
```

## `endBufferEl` one-way latch (in `updateEndBufferActive`'s surrounding block)

```
One-way latch (user-requested behavior, 2026-09-30): once a conversation
is long enough to need scrolling, permanently reserve endBufferEl's fixed
height so the reply that starts filling it doesn't visibly shift/"jump"
the transcript - and it never turns back off, so the buffer's size stays
constant for the rest of the session regardless of what's sent next.
Checked here (rather than a ResizeObserver) because scrollToBottom()
already runs after every content change that could newly overflow the
pane. Measures BEFORE the potential activation below, while the buffer
is still height:0, so its own box never counts toward "is this
overflowing" - only real conversation content does.
```

## Stop button wiring (near `stopBtn.addEventListener('click', ...)`)

```
Post-hoc addition (2026-08-24, user-requested): a separate stop button
next to send (not send doubling as stop, per user feedback on the
first version of this) - shown only while busy. AgentLoop.cancel()
already existed and was fully wired end-to-end (bootstrap.ts's
onStop), just never reachable from any UI control before this.
```

## Send button wiring (near `sendBtn.addEventListener('click', doSend)`)

```
Post-hoc change (2026-08-24, user-requested): send is now always
clickable, including while busy - the textarea also stays enabled
(see setBusy below), so the user can type and queue their next
message instead of being locked out until the current run finishes.
Queueing itself is bootstrap.ts's job (it owns run/busy state); this
layer just always relays "user hit send with this text".
```

## Settings button click handler (single-handler note)

```
Note: the settings button's click handler lives further down (FT-1,
"one handler, not two rebound listeners") - it opens/closes this dropdown
in chat view and doubles as "back to conversation" in the settings view.
A second listener was briefly wired here too; removed - two listeners on
the same button each toggling .open cancelled each other out on every
click (confirmed repro: the button appeared completely unresponsive).
```

## `trackMouseDownOutside`

```
Shared by every dismiss-on-outside-click popup (PR review, 2026-10-02:
settings dropdown and mode menu had copy-pasted this identical
tracking block, which is exactly how the mode menu went without the
fix for a while in the first place) - a future popup gets this by
calling the helper, not by copy-pasting another block.
```

## Settings dropdown outside-click handler

```
Post-hoc addition (2026-08-24, user-requested): closes the quick
settings dropdown on an outside click - only applies to the dropdown
('open' class); the full inline settings VIEW has its own back/close
affordance (settingsBtn above) and unsaved-changes guard, so it is
deliberately untouched here.
```

## `setBusy` handler (textarea/send-button enablement)

```
Post-hoc change (2026-08-24, user-requested): neither the textarea
nor the send button are disabled while busy any more - the user can
keep typing (and queue) their next message during a run instead of
being locked out. The stop button (separate from send) is the only
thing that toggles with busy state.
```

## `showHistoric` (trailing-empty-state omission)

```
No trailing emptyStateHtml() append here (that used to be the bug):
.ai-chat-empty sets `flex: 1` + `overflow: hidden`, which per the
flexbox spec gives it a zero automatic minimum size - once the
replayed transcript above already fills the pane, the shrink
algorithm crushes this element toward zero height instead of the
message bubbles around it, hiding the welcome icon/title/starters.
The chip dock (a sibling outside .ai-chat entirely) replaces it as
this reopened conversation's way back to both actions, and stays
visible for the rest of the session (only resetToEmpty hides it).
```
