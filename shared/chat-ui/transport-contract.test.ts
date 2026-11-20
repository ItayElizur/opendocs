import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentStreamHandle, AgentStreamRequest } from '@officeai/agent-core'
import type { AiProviderId, ReasoningEffort } from '@officeai/ai-provider'
import { getSettings, makeTransport, setSettings } from '../web-src/app-shell/settings'

// Pins what users of makeTransport() observe (callback sequence, request fields
// each provider reads) so the provider layer can be swapped and proven equivalent.

type Protocol = 'anthropic' | 'gemini' | 'openai'
const PROVIDERS: AiProviderId[] = ['anthropic', 'gemini', 'openai', 'deepseek', 'custom']
const protocolOf = (p: AiProviderId): Protocol => (p === 'anthropic' || p === 'gemini' ? p : 'openai')
const MODELS: Record<AiProviderId, string> = {
  anthropic: 'claude-x',
  gemini: 'gemini-x',
  openai: 'gpt-x',
  deepseek: 'deepseek-chat',
  custom: 'qwen3',
}
const CUSTOM_BASE_URL = 'http://vllm.local/v1/'

function configure(provider: AiProviderId, reasoningEffort: ReasoningEffort = 'default', baseUrl = CUSTOM_BASE_URL): void {
  const s = structuredClone(getSettings())
  s.ai.provider = provider
  s.ai.providers[provider] = { apiKey: 'test-key', model: MODELS[provider], baseUrl: provider === 'custom' ? baseUrl : undefined }
  setSettings({ ...s, reasoningEffort })
}

// ---- fake network ----

interface Reply {
  status?: number
  contentType?: string
  body?: string
  /** keep the body open after `body` until the request is aborted */
  hang?: boolean
}
interface Sent {
  url: string
  headers: Record<string, string>
  body: any
}

const enc = new TextEncoder()

function serve(reply: Reply) {
  const sent: Sent[] = []
  let push = (_chunk: string) => {}
  vi.stubGlobal('fetch', async (url: string | URL, init: RequestInit) => {
    sent.push({
      url: String(url),
      headers: Object.fromEntries(new Headers(init.headers).entries()),
      body: JSON.parse(String(init.body)),
    })
    const signal = init.signal
    if (signal?.aborted) throw signal.reason
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        // what a real fetch does to an in-flight body when its signal aborts
        signal?.addEventListener('abort', () => c.error(signal.reason ?? new DOMException('This operation was aborted', 'AbortError')), { once: true })
        push = (chunk) => c.enqueue(enc.encode(chunk))
        // small chunks so SSE lines and JSON straddle reads, as on a real wire
        const body = reply.body ?? ''
        for (let i = 0; i < body.length; i += 40) push(body.slice(i, i + 40))
        if (!reply.hang) c.close()
      },
    })
    return new Response(stream, {
      status: reply.status ?? 200,
      headers: { 'content-type': reply.contentType ?? 'text/event-stream' },
    })
  })
  return { sent, push: (chunk: string) => push(chunk) }
}

// ---- recorded callback sequence ----

interface RecordedCall {
  id: unknown
  name: string
  input: Record<string, unknown>
  inputError: boolean
  truncated: boolean
}
type Ev = { delta: string } | { toolCall: RecordedCall } | { stop: string } | 'done' | { error: unknown }

const TOOL = {
  name: 'insert_text',
  description: 'Insert text',
  inputSchema: { type: 'object', properties: { text: { type: 'string' }, position: { type: 'string' } }, required: ['text'] },
}
const BASIC: AgentStreamRequest = { system: 'sys', messages: [{ role: 'user', text: 'go' }], tools: [TOOL] }

function record(request: AgentStreamRequest = BASIC, { cancelOnDelta = false } = {}) {
  const events: Ev[] = []
  let handle: AgentStreamHandle | undefined
  const finished = new Promise<Ev[]>((resolve) => {
    handle = makeTransport().stream(request, {
      onDelta: (text) => {
        const last = events.at(-1)
        // chunking is incidental; consecutive deltas are one observable run of text
        if (last && typeof last === 'object' && 'delta' in last) last.delta += text
        else events.push({ delta: text })
        if (cancelOnDelta) handle?.cancel()
      },
      onToolCall: (c) =>
        events.push({ toolCall: { id: c.id, name: c.name, input: c.input, inputError: !!c.inputError, truncated: !!c.truncated } }),
      onStopReason: (reason) => events.push({ stop: reason }),
      onDone: () => {
        events.push('done')
        resolve(events)
      },
      onError: (error) => {
        events.push({ error })
        resolve(events)
      },
    })
  })
  return { events, finished }
}

/** Waits for the terminal callback plus one macrotask, so a late duplicate callback would show up in the list. */
async function settle(finished: Promise<Ev[]>): Promise<Ev[]> {
  const events = await finished
  await new Promise((r) => setTimeout(r, 0))
  return events
}

// ---- canned wire formats ----

const sse = (events: unknown[]) => events.map((e) => `data: ${typeof e === 'string' ? e : JSON.stringify(e)}\n\n`).join('')

const GOOD = { text: 'hello', position: 'end' }
const GOOD_JSON = JSON.stringify(GOOD)
const BAD_JSON = '{"text":"say "hi" now"}'
const QUOTED = { text: 'say "hi" now' }
const TRUNC_JSON = '{"text":"a long outl'
const NOTICE = 'Your credits have been exhausted.'

const anth = {
  start: { type: 'message_start', message: { id: 'msg_1', type: 'message', role: 'assistant', model: 'claude-x', content: [], stop_reason: null, usage: { input_tokens: 1, output_tokens: 0 } } },
  textOpen: (index: number, ...parts: string[]) => [
    { type: 'content_block_start', index, content_block: { type: 'text', text: '' } },
    ...parts.map((text) => ({ type: 'content_block_delta', index, delta: { type: 'text_delta', text } })),
  ],
  text: (index: number, ...parts: string[]) => [...anth.textOpen(index, ...parts), { type: 'content_block_stop', index }],
  tool: (index: number, json: string) => [
    { type: 'content_block_start', index, content_block: { type: 'tool_use', id: 'toolu_1', name: 'insert_text', input: {} } },
    { type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: json.slice(0, 5) } },
    { type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: json.slice(5) } },
    { type: 'content_block_stop', index },
  ],
  end: (stop_reason: string) => [
    { type: 'message_delta', delta: { stop_reason, stop_sequence: null }, usage: { output_tokens: 5 } },
    { type: 'message_stop' },
  ],
}
const oai = {
  chunk: (delta: unknown, finish_reason: string | null = null) => ({ id: 'c1', object: 'chat.completion.chunk', created: 1, model: 'm', choices: [{ index: 0, delta, finish_reason }] }),
  tool: (json: string) => [
    oai.chunk({ tool_calls: [{ index: 0, id: 'call_1', type: 'function', function: { name: 'insert_text', arguments: '' } }] }),
    oai.chunk({ tool_calls: [{ index: 0, function: { arguments: json.slice(0, 5) } }] }),
    oai.chunk({ tool_calls: [{ index: 0, function: { arguments: json.slice(5) } }] }),
  ],
  end: (finish: string) => [oai.chunk({}, finish), '[DONE]'],
}
const gem = (parts: unknown[], finishReason?: string) => ({
  candidates: [{ content: { role: 'model', parts }, ...(finishReason ? { finishReason } : {}) }],
  usageMetadata: { promptTokenCount: 1, candidatesTokenCount: 1, totalTokenCount: 2 },
})
const gemCall = (args: Record<string, unknown>) => ({ functionCall: { name: 'insert_text', args } })

const call = (id: unknown, input: Record<string, unknown>, extra: Partial<RecordedCall> = {}): Ev => ({
  toolCall: { id, name: 'insert_text', input, inputError: false, truncated: false, ...extra },
})
// Gemini sends no call ids; the transport mints one, whose value is incidental
const GENERATED_ID = expect.stringMatching(/\S/)
const ABORTED = expect.stringMatching(/abort/i)
const json = (body: unknown): Reply => ({ contentType: 'application/json', body: JSON.stringify(body) })

interface Case {
  reply: Reply
  expected: Ev[]
}
type CaseTable = Record<string, { cancelOnDelta?: boolean } & Record<Protocol, Case>>

const ANTH_400 = JSON.stringify({ type: 'error', error: { type: 'invalid_request_error', message: 'max_tokens: 99999 > 8192, which is the maximum allowed' } })
const GEM_400 = JSON.stringify({ error: { code: 400, message: 'API key not valid. Please pass a valid API key.', status: 'INVALID_ARGUMENT' } })
const OAI_400 = JSON.stringify({ error: { message: 'Invalid model: gpt-nope', type: 'invalid_request_error', param: null, code: 'model_not_found' } })

const CASES: CaseTable = {
  'text then one tool call': {
    anthropic: {
      reply: { body: sse([anth.start, ...anth.text(0, 'Hi ', 'there'), ...anth.tool(1, GOOD_JSON), ...anth.end('tool_use')]) },
      expected: [{ delta: 'Hi there' }, call('toolu_1', GOOD), { stop: 'tool_use' }, 'done'],
    },
    gemini: {
      reply: { body: sse([gem([{ text: 'Hi ' }]), gem([{ text: 'there' }]), gem([gemCall(GOOD)], 'STOP')]) },
      expected: [{ delta: 'Hi there' }, call(GENERATED_ID, GOOD), 'done'],
    },
    openai: {
      reply: { body: sse([oai.chunk({ role: 'assistant', content: 'Hi ' }), oai.chunk({ content: 'there' }), ...oai.tool(GOOD_JSON), ...oai.end('tool_calls')]) },
      expected: [{ delta: 'Hi there' }, call('call_1', GOOD), 'done'],
    },
  },
  'tool arguments with unescaped quotes are repaired': {
    anthropic: {
      reply: { body: sse([anth.start, ...anth.tool(0, BAD_JSON), ...anth.end('tool_use')]) },
      expected: [call('toolu_1', QUOTED), { stop: 'tool_use' }, 'done'],
    },
    // Gemini sends args as a parsed object, so there is no raw JSON to repair
    gemini: {
      reply: { body: sse([gem([gemCall(QUOTED)], 'STOP')]) },
      expected: [call(GENERATED_ID, QUOTED), 'done'],
    },
    openai: {
      reply: { body: sse([...oai.tool(BAD_JSON), ...oai.end('tool_calls')]) },
      expected: [call('call_1', QUOTED), 'done'],
    },
  },
  'output cut off by max_tokens': {
    anthropic: {
      reply: { body: sse([anth.start, ...anth.tool(0, TRUNC_JSON), ...anth.end('max_tokens')]) },
      expected: [call('toolu_1', {}, { inputError: true, truncated: true }), { stop: 'max_tokens' }, 'done'],
    },
    // Gemini never streams partial arguments; the cut-off shows only as the stop reason
    gemini: {
      reply: { body: sse([gem([{ text: 'partial' }], 'MAX_TOKENS')]) },
      expected: [{ delta: 'partial' }, { stop: 'max_tokens' }, 'done'],
    },
    openai: {
      reply: { body: sse([...oai.tool(TRUNC_JSON), ...oai.end('length')]) },
      expected: [call('call_1', {}, { inputError: true, truncated: true }), { stop: 'max_tokens' }, 'done'],
    },
  },
  'cancel() mid-stream': {
    cancelOnDelta: true,
    anthropic: {
      reply: { hang: true, body: sse([anth.start, ...anth.textOpen(0, 'Hi there')]) },
      expected: [{ delta: 'Hi there' }, { error: ABORTED }],
    },
    gemini: {
      reply: { hang: true, body: sse([gem([{ text: 'Hi there' }])]) },
      expected: [{ delta: 'Hi there' }, { error: ABORTED }],
    },
    openai: {
      reply: { hang: true, body: sse([oai.chunk({ role: 'assistant', content: 'Hi there' })]) },
      expected: [{ delta: 'Hi there' }, { error: ABORTED }],
    },
  },
  'HTTP 400 with a JSON error body': {
    anthropic: { reply: { status: 400, contentType: 'application/json', body: ANTH_400 }, expected: [{ error: `Claude HTTP 400: ${ANTH_400}` }] },
    gemini: { reply: { status: 400, contentType: 'application/json', body: GEM_400 }, expected: [{ error: `Gemini HTTP 400: ${GEM_400}` }] },
    openai: { reply: { status: 400, contentType: 'application/json', body: OAI_400 }, expected: [{ error: `HTTP 400: ${OAI_400}` }] },
  },
  '200 with a JSON credits notice instead of SSE': {
    anthropic: {
      reply: json({ id: 'm', type: 'message', role: 'assistant', content: [{ type: 'text', text: NOTICE }], stop_reason: 'end_turn' }),
      expected: [{ error: NOTICE }],
    },
    gemini: {
      reply: json([{ candidates: [{ content: { role: 'model', parts: [{ text: NOTICE }] }, finishReason: 'STOP' }] }]),
      expected: [{ error: NOTICE }],
    },
    openai: {
      reply: json({ id: 'c', object: 'chat.completion', choices: [{ index: 0, message: { role: 'assistant', content: NOTICE }, finish_reason: 'stop' }] }),
      expected: [{ error: NOTICE }],
    },
  },
  // Since the SDK swap a non-credits JSON reply to a streaming request is an error, not a rendered turn.
  '200 with a complete JSON message instead of SSE': {
    anthropic: {
      reply: json({
        id: 'm',
        type: 'message',
        role: 'assistant',
        content: [{ type: 'text', text: 'Hi there' }, { type: 'tool_use', id: 'toolu_1', name: 'insert_text', input: GOOD }],
        stop_reason: 'tool_use',
      }),
      expected: [{ error: expect.stringMatching(/^Claude returned a JSON reply instead of a stream: [[{]/) }],
    },
    gemini: {
      reply: json([{ candidates: [{ content: { role: 'model', parts: [{ text: 'Hi there' }, gemCall(GOOD)] }, finishReason: 'STOP' }] }]),
      expected: [{ error: expect.stringMatching(/^Gemini returned a JSON reply instead of a stream: [[{]/) }],
    },
    openai: {
      reply: json({
        id: 'c',
        object: 'chat.completion',
        choices: [
          {
            index: 0,
            message: { role: 'assistant', content: 'Hi there', tool_calls: [{ id: 'call_1', type: 'function', function: { name: 'insert_text', arguments: GOOD_JSON } }] },
            finish_reason: 'tool_calls',
          },
        ],
      }),
      expected: [{ error: expect.stringMatching(/^The model returned a JSON reply instead of a stream: [[{]/) }],
    },
  },
  '200 with no events': {
    anthropic: { reply: { body: '' }, expected: [{ error: 'Claude returned no content (empty stream)' }] },
    gemini: { reply: { body: '' }, expected: [{ error: 'Gemini returned no content (empty stream)' }] },
    openai: { reply: { body: '' }, expected: [{ error: 'The model returned no content (empty stream)' }] },
  },
  '{"error":...} event inside a 200 stream': {
    anthropic: {
      reply: { body: sse([anth.start, ...anth.textOpen(0, 'Hi'), { error: { message: 'quota exceeded upstream' } }]) },
      expected: [{ delta: 'Hi' }, { error: 'quota exceeded upstream' }],
    },
    gemini: {
      reply: { body: sse([gem([{ text: 'Hi' }]), { error: { message: 'quota exceeded upstream' } }]) },
      expected: [{ delta: 'Hi' }, { error: 'quota exceeded upstream' }],
    },
    openai: {
      reply: { body: sse([oai.chunk({ role: 'assistant', content: 'Hi' }), { error: { message: 'quota exceeded upstream' } }]) },
      expected: [{ delta: 'Hi' }, { error: 'quota exceeded upstream' }],
    },
  },
}

const RESPONSE_ROWS = Object.entries(CASES).flatMap(([name, c]) =>
  PROVIDERS.map((provider) => ({ name, provider, cancelOnDelta: !!c.cancelOnDelta, ...c[protocolOf(provider)] })),
)

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('transport contract: callbacks', () => {
  it.each(RESPONSE_ROWS)('$name ($provider)', async ({ provider, reply, expected, cancelOnDelta }) => {
    configure(provider)
    serve(reply)
    const { finished } = record(BASIC, { cancelOnDelta })
    expect(await settle(finished)).toEqual(expected)
  })
})

describe('transport contract: idle timeout', () => {
  beforeEach(() => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
  })
  afterEach(() => {
    vi.useRealTimers()
  })

  it.each(PROVIDERS)('keepalive pings hold the stream open, then 180s of silence fails it (%s)', async (provider) => {
    configure(provider)
    const server = serve({ hang: true })
    const { events, finished } = record()
    await vi.advanceTimersByTimeAsync(0)
    for (let i = 0; i < 3; i++) {
      await vi.advanceTimersByTimeAsync(170_000)
      server.push(': ping\n\n')
      await vi.advanceTimersByTimeAsync(0)
    }
    await vi.advanceTimersByTimeAsync(179_000)
    expect(events).toEqual([])
    await vi.advanceTimersByTimeAsync(1_000)
    expect(await finished).toEqual([{ error: 'AI request timed out: no data received from the network for 180s' }])
  })
})

// ---- request side ----

const OK_REPLY: Record<Protocol, Reply> = {
  anthropic: CASES['text then one tool call']!.anthropic.reply,
  gemini: CASES['text then one tool call']!.gemini.reply,
  openai: CASES['text then one tool call']!.openai.reply,
}

async function sentRequest(provider: AiProviderId, request: AgentStreamRequest = BASIC, effort: ReasoningEffort = 'default'): Promise<Sent> {
  configure(provider, effort)
  const { sent } = serve(OK_REPLY[protocolOf(provider)])
  const events = await settle(record(request).finished)
  expect(events.at(-1)).toBe('done')
  expect(sent).toHaveLength(1)
  return sent[0]!
}

/** A string and a single text block are the same content to these APIs. */
const blocks = (content: unknown) => (typeof content === 'string' ? [{ type: 'text', text: content }] : content)
/** Gemini's JSON API accepts snake_case and camelCase field names alike. */
const camel = (value: unknown): unknown =>
  Array.isArray(value)
    ? value.map(camel)
    : value && typeof value === 'object'
      ? Object.fromEntries(Object.entries(value).map(([k, v]) => [k.replace(/_([a-z])/g, (_, c: string) => c.toUpperCase()), camel(v)]))
      : value

/** Gemini reads a tool's schema from `parameters` or `parametersJsonSchema`. */
const geminiTools = (tools: any[] | undefined) =>
  tools?.map((t) => ({
    ...t,
    functionDeclarations: t.functionDeclarations?.map(({ parametersJsonSchema, ...d }: any) => ({ ...d, parameters: d.parameters ?? parametersJsonSchema })),
  }))
/** A Gemini functionResponse.response is a free-form object; a success may carry its text as `result` or `content`. */
const geminiResponses = (contents: any[]) =>
  contents.map((c) => ({
    ...c,
    parts: c.parts.map((part: any) => {
      const r = part.functionResponse?.response
      if (!r || 'error' in r) return part
      return { functionResponse: { ...part.functionResponse, response: { result: r.result ?? r.content } } }
    }),
  }))

const ENDPOINT: Record<AiProviderId, { url: string; headers: Record<string, string> }> = {
  anthropic: {
    url: 'https://api.anthropic.com/v1/messages',
    headers: { 'x-api-key': 'test-key', 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
  },
  gemini: {
    url: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-x:streamGenerateContent?alt=sse',
    headers: { 'x-goog-api-key': 'test-key' },
  },
  openai: { url: 'https://api.openai.com/v1/chat/completions', headers: { authorization: 'Bearer test-key' } },
  deepseek: { url: 'https://api.deepseek.com/v1/chat/completions', headers: { authorization: 'Bearer test-key' } },
  custom: { url: 'http://vllm.local/v1/chat/completions', headers: { authorization: 'Bearer test-key' } },
}

/** The request fields each protocol needs, in the shape the provider reads them. */
const ESSENTIALS: Record<Protocol, (body: any) => unknown> = {
  anthropic: (b) => ({
    model: b.model,
    stream: b.stream,
    system: blocks(b.system),
    tools: b.tools,
    messages: b.messages.map((m: any) => ({ ...m, content: blocks(m.content) })),
  }),
  gemini: (b) => camel({ systemInstruction: b.systemInstruction, tools: geminiTools(b.tools), contents: b.contents }),
  openai: (b) => ({ model: b.model, stream: b.stream, tools: b.tools, messages: b.messages }),
}

const BASIC_ESSENTIALS: Record<AiProviderId, unknown> = {
  anthropic: {
    model: 'claude-x',
    stream: true,
    system: [{ type: 'text', text: 'sys' }],
    tools: [{ name: 'insert_text', description: 'Insert text', input_schema: TOOL.inputSchema }],
    messages: [{ role: 'user', content: [{ type: 'text', text: 'go' }] }],
  },
  gemini: {
    systemInstruction: { parts: [{ text: 'sys' }] },
    tools: [{ functionDeclarations: [{ name: 'insert_text', description: 'Insert text', parameters: TOOL.inputSchema }] }],
    contents: [{ role: 'user', parts: [{ text: 'go' }] }],
  },
  ...Object.fromEntries(
    (['openai', 'deepseek', 'custom'] as const).map((p) => [
      p,
      {
        model: MODELS[p],
        stream: true,
        tools: [{ type: 'function', function: { name: 'insert_text', description: 'Insert text', parameters: TOOL.inputSchema } }],
        messages: [{ role: 'system', content: 'sys' }, { role: 'user', content: 'go' }],
      },
    ]),
  ),
} as Record<AiProviderId, unknown>

describe('transport contract: request', () => {
  it.each(PROVIDERS)('URL, auth headers, model, system, tools and messages (%s)', async (provider) => {
    const { url, headers, body } = await sentRequest(provider)
    expect(url).toBe(ENDPOINT[provider].url)
    expect(headers).toMatchObject(ENDPOINT[provider].headers)
    expect(ESSENTIALS[protocolOf(provider)](body)).toMatchObject(BASIC_ESSENTIALS[provider] as object)
  })

  it('custom with no Base URL fails without sending a request', async () => {
    configure('custom', 'default', '')
    const { sent } = serve(OK_REPLY.openai)
    expect(await settle(record().finished)).toEqual([{ error: 'A custom provider requires a Base URL' }])
    expect(sent).toHaveLength(0)
  })

  // [provider, effort, output cap, reasoning fields]; Anthropic's thinking budget comes on top of the 8192 reply cap
  const anthropicThinking = (budget_tokens: number) => ({ thinking: { type: 'enabled', budget_tokens } })
  const geminiThinking = (thinkingBudget: number) => ({ thinkingConfig: { thinkingBudget } })
  const kwargs = (on: boolean) => ({ chat_template_kwargs: { enable_thinking: on, preserve_thinking: on } })
  const EFFORT_ROWS: Array<[AiProviderId, ReasoningEffort, number, Record<string, unknown>]> = [
    ['anthropic', 'default', 8192, {}],
    ['anthropic', 'off', 8192, {}],
    ['anthropic', 'low', 8192 + 1024, anthropicThinking(1024)],
    ['anthropic', 'medium', 8192 + 3072, anthropicThinking(3072)],
    ['anthropic', 'high', 8192 + 5120, anthropicThinking(5120)],
    ['anthropic', 'xhigh', 8192 + 7168, anthropicThinking(7168)],
    ['gemini', 'default', 8192, {}],
    ['gemini', 'off', 8192, geminiThinking(0)],
    ['gemini', 'low', 8192, geminiThinking(1024)],
    ['gemini', 'medium', 8192, geminiThinking(4096)],
    ['gemini', 'high', 8192, geminiThinking(8192)],
    ['gemini', 'xhigh', 8192, geminiThinking(16384)],
    ...(['openai', 'deepseek'] as const).flatMap((p): Array<[AiProviderId, ReasoningEffort, number, Record<string, unknown>]> => [
      [p, 'default', 8192, {}],
      [p, 'off', 8192, { reasoning_effort: 'minimal' }],
      [p, 'low', 8192, { reasoning_effort: 'low' }],
      [p, 'medium', 8192, { reasoning_effort: 'medium' }],
      [p, 'high', 8192, { reasoning_effort: 'high' }],
      [p, 'xhigh', 8192, { reasoning_effort: 'high' }],
    ]),
    ['custom', 'default', 8192, {}],
    ['custom', 'off', 8192, kwargs(false)],
    ['custom', 'low', 8192, { reasoning_effort: 'low', ...kwargs(true) }],
    ['custom', 'medium', 8192, { reasoning_effort: 'medium', ...kwargs(true) }],
    ['custom', 'high', 8192, { reasoning_effort: 'high', ...kwargs(true) }],
    ['custom', 'xhigh', 8192, { reasoning_effort: 'xhigh', ...kwargs(true) }],
  ]
  const REASONING: Record<Protocol, (body: any) => { cap: number; reasoning: Record<string, unknown> }> = {
    anthropic: (b) => ({ cap: b.max_tokens, reasoning: { thinking: b.thinking } }),
    gemini: (b) => ({ cap: b.generationConfig.maxOutputTokens, reasoning: { thinkingConfig: b.generationConfig.thinkingConfig } }),
    openai: (b) => ({ cap: b.max_tokens, reasoning: { reasoning_effort: b.reasoning_effort, chat_template_kwargs: b.chat_template_kwargs } }),
  }

  it.each(EFFORT_ROWS)('%s with reasoning effort %s sends output cap %d and %o', async (provider, effort, cap, reasoning) => {
    const { body } = await sentRequest(provider, BASIC, effort)
    // toEqual treats an undefined field as absent, so {} asserts no reasoning field was sent
    expect(REASONING[protocolOf(provider)](body)).toEqual({ cap, reasoning })
  })

  const HISTORY: AgentMessage[] = [
    { role: 'user', text: 'Add a title', images: [{ base64: 'iVBORw0KGgo=', mime: 'image/png' }] },
    {
      role: 'assistant',
      text: 'Sure.',
      toolCalls: [
        { id: 'toolu_1', name: 'insert_text', input: { text: 'Title' } },
        { id: 'toolu_2', name: 'read_doc', input: {} },
      ],
    },
    {
      role: 'tool',
      results: [
        { id: 'toolu_1', name: 'insert_text', output: 'ok' },
        { id: 'toolu_2', name: 'read_doc', output: 'boom', isError: true },
      ],
    },
    // a prior turn with no prose and no calls; providers reject an empty assistant turn
    { role: 'assistant', text: '' },
    { role: 'user', text: 'thanks' },
  ]

  const HISTORY_WIRE: Record<Protocol, { pick: (body: any) => unknown; expected: unknown[] }> = {
    anthropic: {
      pick: (b) =>
        b.messages.map((m: any) => ({
          role: m.role,
          content: blocks(m.content).map((c: any) => (c.type === 'tool_result' ? { ...c, content: blocks(c.content) } : c)),
        })),
      expected: [
        {
          role: 'user',
          content: [
            { type: 'text', text: 'Add a title' },
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: 'iVBORw0KGgo=' } },
          ],
        },
        {
          role: 'assistant',
          content: [
            { type: 'text', text: 'Sure.' },
            { type: 'tool_use', id: 'toolu_1', name: 'insert_text', input: { text: 'Title' } },
            { type: 'tool_use', id: 'toolu_2', name: 'read_doc', input: {} },
          ],
        },
        {
          role: 'user',
          content: [
            { type: 'tool_result', tool_use_id: 'toolu_1', content: [{ type: 'text', text: 'ok' }] },
            { type: 'tool_result', tool_use_id: 'toolu_2', content: [{ type: 'text', text: 'boom' }], is_error: true },
          ],
        },
        { role: 'assistant', content: [{ type: 'text', text: '(no content)' }] },
        { role: 'user', content: [{ type: 'text', text: 'thanks' }] },
      ],
    },
    gemini: {
      pick: (b) => geminiResponses(camel(b.contents) as any[]),
      expected: [
        { role: 'user', parts: [{ text: 'Add a title' }, { inlineData: { mimeType: 'image/png', data: 'iVBORw0KGgo=' } }] },
        {
          role: 'model',
          parts: [{ text: 'Sure.' }, { functionCall: { name: 'insert_text', args: { text: 'Title' } } }, { functionCall: { name: 'read_doc', args: {} } }],
        },
        {
          role: 'user',
          parts: [
            { functionResponse: { name: 'insert_text', response: { result: 'ok' } } },
            { functionResponse: { name: 'read_doc', response: { error: 'boom' } } },
          ],
        },
        { role: 'model', parts: [{ text: '(no content)' }] },
        { role: 'user', parts: [{ text: 'thanks' }] },
      ],
    },
    openai: {
      pick: (b) =>
        b.messages.map((m: any) => ({
          ...m,
          ...(m.role === 'user' || m.role === 'tool' ? { content: blocks(m.content) } : {}),
          ...(m.tool_calls
            ? { tool_calls: m.tool_calls.map((tc: any) => ({ ...tc, function: { ...tc.function, arguments: JSON.parse(tc.function.arguments) } })) }
            : {}),
        })),
      expected: [
        { role: 'system', content: 'sys' },
        {
          role: 'user',
          content: [
            { type: 'text', text: 'Add a title' },
            { type: 'image_url', image_url: { url: 'data:image/png;base64,iVBORw0KGgo=' } },
          ],
        },
        {
          role: 'assistant',
          content: 'Sure.',
          tool_calls: [
            { id: 'toolu_1', type: 'function', function: { name: 'insert_text', arguments: { text: 'Title' } } },
            { id: 'toolu_2', type: 'function', function: { name: 'read_doc', arguments: {} } },
          ],
        },
        // chat completions has no error flag on a tool message, so isError has nowhere to go
        { role: 'tool', tool_call_id: 'toolu_1', content: [{ type: 'text', text: 'ok' }] },
        { role: 'tool', tool_call_id: 'toolu_2', content: [{ type: 'text', text: 'boom' }] },
        { role: 'assistant', content: '(no content)' },
        { role: 'user', content: [{ type: 'text', text: 'thanks' }] },
      ],
    },
  }

  it.each(PROVIDERS)('history with an image, tool calls, an error result and an empty assistant turn (%s)', async (provider) => {
    const { body } = await sentRequest(provider, { system: 'sys', messages: HISTORY, tools: [TOOL] })
    const wire = HISTORY_WIRE[protocolOf(provider)]
    expect(wire.pick(body)).toMatchObject(wire.expected)
  })
})
