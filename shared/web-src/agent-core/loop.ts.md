# loop.ts — relocated comment history / rationale

## `finishTurn` — empty assistant text normalization

```
Models often end a tool-using run with an empty text turn ("I'm done").
Leaving assistant text empty in history then poisons the next user
prompt: Anthropic rejects empty content arrays, Gemini rejects empty
parts, and OpenAI-compatible routes send content:null with no tool_calls —
all of which make follow-up turns fail or return empty again (first prompt
works, second shows "no summary").
Same normalization as restore(), applied unconditionally: cancelled and
read-only empty turns poison follow-ups just the same. onDone still
reports the raw turn text so app UIs keep their localized fallbacks
instead of surfacing this English placeholder.
```

## `redact` option

```
Redacts the user's own configured model API key from an outgoing user message,
so a key pasted into the chat does not reach the model API verbatim. The app
shell supplies the function (it knows the configured keys). It used to be a
regex that rewrote anything key-shaped (sk-..., URL credentials, password
assignments); that was narrowed on purpose. Tool inputs and tool results are
never redacted: a direct "!" tool call stores its arguments as typed.
```
