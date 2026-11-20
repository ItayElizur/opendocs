# stream.ts — relocated comment history / rationale

## Empty-stream detection

A stream with no content AND no finish reason from the provider is a gateway
soft-failure, not a model turn. It surfaces as "<label> returned no content
(empty stream)" instead of dissolving into an empty "successful" turn with no
diagnostics. A genuine empty closing turn (common after tool-heavy runs) still
carries a finish reason. The SDK already raises its own errors for this on the
Anthropic and OpenAI-compatible routes; `errorText` rewrites them to the same
message, and `streamForProvider` checks it directly for Gemini, where the SDK
ends quietly.

## JSON reply to a streaming request

Gateways can answer a `stream: true` request with a complete JSON body, observed
when credits are exhausted (HTTP 200, `Content-Type: application/json`, the
notice inside a regular message). The SDK would read no SSE events from it and
report only "no output". `watchedFetch` reads such a body before the SDK does:
a credits notice becomes the error text, anything else becomes "<label>
returned a JSON reply instead of a stream: <body>". Before the SDK swap, a
non-credits JSON reply was rendered as a normal turn.

## Watchdog instead of the SDK's timeouts

The SDK's `timeout` options count content parts, not bytes, so a stream that
only sends keepalive pings during a long thinking pause would be killed. A
prototype confirmed it (pings every 100 ms, killed at the first-chunk timeout).
`watchedFetch` touches the byte-level watchdog on every network chunk instead.

## `maxRetries: 0`

The SDK retries twice by default. A failing gateway would then be billed for
three attempts, and the user would wait three times as long for the error.
The transport has always made one attempt.

## Cancel ends in onError

`cancel()` aborts the request, and the turn ends with `onError("This operation
was aborted")`, as it did before the SDK swap. `AgentStreamHandle.cancel`'s doc
and `AgentLoop.cancel` say the transport emits `onDone` after a cancel; that
mismatch predates the swap and is pinned by the contract test until it is fixed
on its own.

## Anthropic thinking budget

The SDK adds Anthropic's thinking `budgetTokens` on top of `maxOutputTokens`,
so the reply always keeps the full `MAX_TOKENS`. Before the swap the budget
counted inside `max_tokens`, which left `xhigh` 1,024 tokens for the reply.

## `openAiReasoningFields`

`reasoning_effort` is OpenAI's own field name, tried against any
OpenAI-compatible surface. `chat_template_kwargs` is vLLM-specific (it
forwards into the model's Jinja chat template, e.g. to toggle a
Qwen3-style `enable_thinking` switch) - only worth sending to a
self-hosted `custom` endpoint, never to the real OpenAI/DeepSeek APIs,
which don't use it. The fields go through the openai-compatible provider's
`providerOptions`, which it copies into the request body.

## `openAiReasoningFields` — 'xhigh' clamp on the plain OpenAI API

'xhigh' is a real tier for vLLM-served models like Qwen3 (sent as-is
when includeThinkingKwargs is set), but the plain OpenAI API's
reasoning_effort enum only goes up to 'high' - sending 'xhigh' there
verbatim is a hard 400, not a silent no-op, unlike every other
unsupported-tier case this function handles (PR review, 2026-10-02).
