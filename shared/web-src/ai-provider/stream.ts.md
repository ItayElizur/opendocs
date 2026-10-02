# stream.ts — relocated comment history / rationale

## Anthropic stream: empty-stream detection

```
A stream with no content AND no message framing (no stop_reason ever seen)
is a gateway soft-failure, not a model turn — surface it instead of letting
it dissolve into an empty "successful" turn with no diagnostics. A genuine
empty closing turn (common after tool-heavy runs) still carries end_turn.
The "(empty stream)" suffix is a contract: app renderers match it to
classify the failure as empty output (fail fast, no billed retries).
```

## `openAiReasoningFields`

```
`reasoning_effort` is OpenAI's own field name, tried against any
OpenAI-compatible surface. `chat_template_kwargs` is vLLM-specific (it
forwards into the model's Jinja chat template, e.g. to toggle a
Qwen3-style `enable_thinking` switch) - only worth sending to a
self-hosted `custom` endpoint, never to the real OpenAI/DeepSeek APIs,
which don't use it.
```

## `openAiReasoningFields` — 'xhigh' clamp on the plain OpenAI API

```
'xhigh' is a real tier for vLLM-served models like Qwen3 (sent as-is
when includeThinkingKwargs is set), but the plain OpenAI API's
reasoning_effort enum only goes up to 'high' - sending 'xhigh' there
verbatim is a hard 400, not a silent no-op, unlike every other
unsupported-tier case this function handles (PR review, 2026-10-02).
```
