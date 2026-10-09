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

## `sanitizeAgentPayload`

```
Redact secret-looking tokens from an outgoing user message so accidentally
pasted API keys, URL credentials, and password assignments don't reach
remote model APIs verbatim.

Imported from public PR #32 (BuiltByHarshil), with the credential pattern
narrowed to URL userinfo (scheme://user:pass@host) so ordinary "a:b@c"
prose is never rewritten.
```
