import type { AgentMessage, AgentToolCall, AgentToolDef } from '@officeai/agent-core'
import { createAnthropic } from '@ai-sdk/anthropic'
import { createGoogleGenerativeAI } from '@ai-sdk/google'
import { createOpenAICompatible } from '@ai-sdk/openai-compatible'
import { jsonSchema, streamText, tool, type JSONValue, type LanguageModel, type ModelMessage, type ToolSet } from 'ai'
import { httpBodyDetail } from './http-error'
import type { AiProviderConfig, AiProviderId, ReasoningEffort } from './types'
import { createStreamWatchdog, type StreamWatchdog } from './watchdog'

/** Anthropic's thinking budget; the SDK adds it on top of maxOutputTokens. */
const ANTHROPIC_THINKING_BUDGETS: Record<'low' | 'medium' | 'high' | 'xhigh', number> = {
  low: 1024,
  medium: 3072,
  high: 5120,
  xhigh: 7168,
}

/** Gemini's thinkingBudget is a separate allowance from maxOutputTokens, so it can run higher than Anthropic's. */
const GEMINI_THINKING_BUDGETS: Record<'low' | 'medium' | 'high' | 'xhigh', number> = {
  low: 1024,
  medium: 4096,
  high: 8192,
  xhigh: 16384,
}

export interface StreamCallbacks {
  onDelta: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
  /** normalized stop reason ('max_tokens' when the output was cut off by the token limit) */
  onStopReason?: (reason: string) => void
  signal: AbortSignal
}

/**
 * Models occasionally emit unescaped " inside string values (e.g. English quotes in Chinese copy).
 * Single-pass scan: a " inside a string whose next non-whitespace char is not structural gets escaped.
 */
function repairUnescapedQuotes(json: string): string {
  let out = ''
  let inStr = false
  for (let i = 0; i < json.length; i++) {
    const c = json[i]!
    if (!inStr) {
      if (c === '"') inStr = true
      out += c
      continue
    }
    if (c === '\\') {
      out += c + (json[++i] ?? '')
      continue
    }
    if (c === '"') {
      let j = i + 1
      while (j < json.length && ' \n\r\t'.includes(json[j]!)) j++
      const next = json[j]
      if (next === undefined || ',}]:'.includes(next)) {
        inStr = false
        out += c
      } else {
        out += '\\"'
      }
      continue
    }
    out += c
  }
  return out
}

/**
 * A gateway's credits-exhausted notice (e.g. "Your credits have been exhausted…"),
 * found anywhere in a JSON body.
 */
function creditsNoticeText(value: unknown): string | null {
  if (typeof value === 'string') {
    const t = value.toLowerCase()
    return t.includes('credit') && (t.includes('exhausted') || t.includes('insufficient')) ? value : null
  }
  if (value && typeof value === 'object') {
    for (const v of Object.values(value)) {
      const hit = creditsNoticeText(v)
      if (hit) return hit
    }
  }
  return null
}

function openAiReasoningFields(effort: ReasoningEffort | undefined, includeThinkingKwargs: boolean): Record<string, JSONValue> {
  if (!effort || effort === 'default') return {}
  if (effort === 'off') {
    return includeThinkingKwargs
      ? { chat_template_kwargs: { enable_thinking: false, preserve_thinking: false } }
      // No universal "off" in the plain OpenAI API - 'minimal' is the
      // closest a reasoning model can get to disabled.
      : { reasoning_effort: 'minimal' }
  }
  return {
    // Clamp 'xhigh' to 'high' for the plain OpenAI API: sending 'xhigh' there is a hard 400 (see `stream.ts.md`).
    reasoning_effort: !includeThinkingKwargs && effort === 'xhigh' ? 'high' : effort,
    ...(includeThinkingKwargs ? { chat_template_kwargs: { enable_thinking: true, preserve_thinking: true } } : {}),
  }
}

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
}

/** What the user reads in front of an error: today's per-provider wording. */
const ERROR_LABEL: Record<AiProviderId, string> = {
  anthropic: 'Claude',
  gemini: 'Gemini',
  deepseek: 'The model',
  openai: 'The model',
  custom: 'The model',
}

interface ProviderSetup {
  model: LanguageModel
  providerOptions: Record<string, Record<string, JSONValue>>
  temperature?: number
}

function providerSetup(provider: AiProviderId, config: AiProviderConfig, fetch: typeof globalThis.fetch): ProviderSetup {
  const effort = config.reasoningEffort
  switch (provider) {
    case 'anthropic': {
      const budget = effort && effort in ANTHROPIC_THINKING_BUDGETS ? ANTHROPIC_THINKING_BUDGETS[effort as keyof typeof ANTHROPIC_THINKING_BUDGETS] : 0
      return {
        model: createAnthropic({
          apiKey: config.apiKey,
          fetch,
          // Anthropic rejects browser-origin requests with 403 unless they opt in.
          headers: { 'anthropic-dangerous-direct-browser-access': 'true' },
        })(config.model),
        providerOptions: budget ? { anthropic: { thinking: { type: 'enabled', budgetTokens: budget } } } : {},
      }
    }
    case 'gemini': {
      const thinkingBudget =
        effort === 'off' ? 0 : effort && effort in GEMINI_THINKING_BUDGETS ? GEMINI_THINKING_BUDGETS[effort as keyof typeof GEMINI_THINKING_BUDGETS] : undefined
      return {
        model: createGoogleGenerativeAI({ apiKey: config.apiKey, fetch })(config.model),
        providerOptions: thinkingBudget === undefined ? {} : { google: { thinkingConfig: { thinkingBudget } } },
        temperature: 0.3,
      }
    }
    case 'deepseek':
    case 'openai':
    case 'custom': {
      const baseURL = provider === 'custom' ? config.baseUrl : OPENAI_COMPATIBLE_BASE_URLS[provider]
      if (!baseURL) throw new Error('A custom provider requires a Base URL')
      return {
        model: createOpenAICompatible({ name: 'compat', baseURL: baseURL.replace(/\/$/, ''), apiKey: config.apiKey, fetch }).chatModel(config.model),
        providerOptions: { compat: openAiReasoningFields(effort, provider === 'custom') },
        temperature: 0.3,
      }
    }
    default:
      throw new Error(`Unknown provider: ${provider as string}`)
  }
}

export function toModelMessages(messages: AgentMessage[]): ModelMessage[] {
  return messages.map((m): ModelMessage => {
    if (m.role === 'user') {
      if (!m.images?.length) return { role: 'user', content: m.text }
      return {
        role: 'user',
        content: [
          ...(m.text ? [{ type: 'text' as const, text: m.text }] : []),
          ...m.images.map((img) => ({ type: 'image' as const, image: img.base64, mediaType: img.mime })),
        ],
      }
    }
    if (m.role === 'assistant') {
      const calls = (m.toolCalls ?? []).map((c) => ({ type: 'tool-call' as const, toolCallId: c.id, toolName: c.name, input: c.input }))
      // Providers reject or drop an empty assistant turn, which poisons every follow-up.
      const text = m.text || (calls.length ? '' : '(no content)')
      return { role: 'assistant', content: [...(text ? [{ type: 'text' as const, text }] : []), ...calls] }
    }
    return {
      role: 'tool',
      content: m.results.map((r) => ({
        type: 'tool-result' as const,
        toolCallId: r.id,
        toolName: r.name,
        output: r.isError ? { type: 'error-text' as const, value: r.output } : { type: 'text' as const, value: r.output },
      })),
    }
  })
}

/** Raised from inside fetch so the reply body survives the SDK's own error wrapping. */
class ReplyError extends Error {}

/**
 * Feeds every received network chunk to the watchdog (the SDK's own timeouts
 * ignore keepalive pings), and turns a JSON reply to a streaming request -
 * gateways send one when credits run out - into a readable error.
 */
function watchedFetch(wd: StreamWatchdog, label: string): typeof globalThis.fetch {
  return async (input, init) => {
    let response: Response
    try {
      response = await fetch(input, init)
    } catch (e) {
      throw new ReplyError(`${label} fetch failed: ${e instanceof Error ? e.message : String(e)}`, { cause: e })
    }
    wd.touch()
    if (!response.ok) throw new ReplyError(`${label === 'The model' ? '' : `${label} `}HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
    if ((response.headers.get('content-type') ?? '').includes('application/json')) {
      const body = await response.text()
      let parsed: unknown
      try {
        parsed = JSON.parse(body)
      } catch {
        parsed = null
      }
      throw new ReplyError(creditsNoticeText(parsed) ?? `${label} returned a JSON reply instead of a stream: ${httpBodyDetail(body)}`)
    }
    if (!response.body) return response
    const body = response.body.pipeThrough(
      new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) {
          wd.touch()
          controller.enqueue(chunk)
        },
      }),
    )
    return new Response(body, { status: response.status, statusText: response.statusText, headers: response.headers })
  }
}

function errorText(error: unknown, label: string): string {
  for (let e: unknown = error; e; e = (e as { cause?: unknown }).cause) {
    if (e instanceof ReplyError) return e.message
  }
  const name = (error as { name?: string } | null)?.name ?? ''
  if (name === 'AI_NoOutputGeneratedError' || name === 'AI_InvalidResponseDataError') {
    return `${label} returned no content (empty stream)`
  }
  // A gateway's {"error": {...}} event on a 200 stream fails the SDK's schema check; show its message.
  const value = (error as { value?: { error?: { message?: unknown } | string } } | null)?.value
  const inner = typeof value?.error === 'string' ? value.error : value?.error?.message
  if (typeof inner === 'string' && inner) return inner
  return error instanceof Error ? error.message : String(error)
}

/** Streams one model turn (no tool execution; AgentLoop runs the tools). */
export async function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const label = ERROR_LABEL[provider]
  const wd = createStreamWatchdog(cb.signal)
  const setup = providerSetup(provider, config, watchedFetch(wd, label))
  const toolSet: ToolSet = Object.fromEntries(
    tools.map((t) => [t.name, tool({ description: t.description, inputSchema: jsonSchema(t.inputSchema) })]),
  )
  return wd.guard(async () => {
    const result = streamText({
      model: setup.model,
      system,
      messages: toModelMessages(messages),
      ...(tools.length ? { tools: toolSet } : {}),
      maxOutputTokens: maxTokens,
      ...(setup.temperature === undefined ? {} : { temperature: setup.temperature }),
      providerOptions: setup.providerOptions,
      abortSignal: wd.signal,
      // One attempt, never the SDK's two billed retries (see `stream.ts.md`).
      maxRetries: 0,
      repairToolCall: async ({ toolCall }) => {
        const repaired = repairUnescapedQuotes(toolCall.input)
        try {
          JSON.parse(repaired)
        } catch {
          return null
        }
        return { ...toolCall, input: repaired }
      },
      // Errors arrive as 'error' parts below; without this the SDK also logs each one to the console.
      onError: () => {},
    })
    const calls: AgentToolCall[] = []
    let finishReason: string | undefined
    let rawFinishReason: string | undefined
    let sawText = false
    for await (const part of result.fullStream) {
      switch (part.type) {
        case 'text-delta':
          if (part.text) {
            sawText = true
            cb.onDelta(part.text)
          }
          break
        case 'tool-call':
          calls.push(
            part.invalid
              ? {
                  id: part.toolCallId,
                  name: part.toolName,
                  input: {},
                  inputError: `${errorText(part.error, label)}; raw: ${String(part.input).slice(0, 500)}`,
                }
              : { id: part.toolCallId, name: part.toolName, input: part.input as Record<string, unknown> },
          )
          break
        case 'finish-step':
          finishReason = part.finishReason
          rawFinishReason = part.rawFinishReason
          break
        case 'error':
          throw new Error(errorText(part.error, label), { cause: part.error })
        case 'abort':
          // A cancel or watchdog timeout ends the turn as an error, as before the SDK; the guard renames a timeout (see `stream.ts.md`).
          throw new Error('This operation was aborted')
      }
    }
    // No content and no finish reason from the provider is a gateway soft-failure, not a turn (see `stream.ts.md`).
    if (!sawText && calls.length === 0 && rawFinishReason === undefined) throw new Error(`${label} returned no content (empty stream)`)
    const cutOff = finishReason === 'length'
    const last = calls.at(-1)
    if (cutOff && last) last.truncated = true
    for (const call of calls) cb.onToolCall(call)
    // Only 'max_tokens' is read downstream; Anthropic's raw reasons pass through as before.
    const stop = cutOff ? 'max_tokens' : provider === 'anthropic' ? rawFinishReason : undefined
    if (stop) cb.onStopReason?.(stop)
  })
}
