# chat-ui.test.ts — relocated comment history / rationale

Comments moved here verbatim per the repo-wide comment cleanup sweep. Each heading names the test/section the comment originally sat above; the current file keeps a short trimmed summary in place.

## 'More settings' hides itself once inside the settings view

```
The "More settings" button makes no sense once already inside the
settings view it opens - this only checks the `hidden` IDL attribute
the code sets, which jsdom does not render; it does NOT catch a CSS
rule silently overriding [hidden]'s effect (confirmed repro: a
`display: block` rule on this exact button did exactly that - see the
fix in chat-ui.css). Real visual verification needs a real browser.
```

## 'the gear button toggles the quick settings dropdown open and closed in chat view'

```
Regression test: a duplicate click listener on this button (one from
before FT-1, one from FT-1 itself) each toggled .open independently,
so every click cancelled itself out and the button appeared totally
unresponsive - confirmed by real-world testing, not caught by any
existing test since the ones above interact with the panel's fields
directly without ever asserting the dropdown's own open/closed state.
```

## Up/Down-arrow recall section banner

```
---- Up/Down-arrow recall of previously sent messages ----

caretCollapsedAtFirstLine/caretCollapsedAtLastLine (chat-ui.ts) decide
whether an arrow press should recall history or just move the caret,
by comparing caretLineMeasurement.measure(pos) - the pixel offsetTop of
the visual line `pos` renders on, via a hidden mirror div - against the
measurement at position 0 (top) / value.length (bottom). jsdom (used
here) performs no real text layout, so a *real* mirror-div measurement
always reports offsetTop 0 for every position in this environment -
there is no way to assert a genuine pixel-line answer from jsdom alone.
So every test in this section stubs caretLineMeasurement.measure with a
small fake that reproduces a specific, known line layout, and asserts
only the surrounding gating/recall logic against it. The default stub
below (hard '\n' counting) reproduces exactly the *old* behavior, which
is still supposed to work today - i.e. it's what real browsers do for
text with no soft-wrapping - so the pre-existing hard-newline tests
keep meaning what they say. The dedicated soft-wrap test further down
installs its own stub simulating word-wrap, to exercise the actual bug
fix (a long, single-line, no-'\n' draft that wraps across several
visual lines).
```

## 'a long single-line (soft-wrapped) draft does not recall on the first ArrowUp...'

```
No '\n' anywhere - under the old lastIndexOf('\n', ...)-only check this
was indistinguishable from a single-line draft, so caretCollapsedAtFirstLine()
was always true here and the very first ArrowUp always recalled
history. Simulate a real browser word-wrapping this into 4 visual
lines of 20 characters each (line index = floor(pos / 20)) - the exact
shape the mirror-div technique measures in production, stubbed here
because jsdom can't lay text out for real (see the section comment above).
```
