import { AgentLoop, type AgentSkill, type AgentStreamHandle, type AgentTransport, type ToolExecution } from '@genoffice/agent-core'
import { streamOpenAiCompatible, type AiProviderConfig } from '@genoffice/ai-provider'

// Spike 2: prove packages/agent-core's AgentLoop and packages/ai-provider's
// streamOpenAiCompatible run unmodified inside a WebView2 page hosted in a
// VSTO CustomTaskPane, talking to a real (local, for this spike) OpenAI-
// compatible HTTP+SSE endpoint - no Electron IPC hop anywhere in this path.
//
// Spike 3: real tools, executed via the WebView2 <-> .NET WebMessage bridge
// (chrome.webview.postMessage <-> CoreWebView2.PostWebMessageAsJson) instead
// of Electron IPC. Tool execution itself is real COM automation against the
// live Word document, handled in TaskPaneHost.cs / WordTools.cs.

declare const chrome: {
  webview: {
    postMessage(message: unknown): void
    addEventListener(type: 'message', listener: (ev: { data: unknown }) => void): void
  }
}

interface ToolCallMessage {
  kind: 'tool-call'
  requestId: string
  toolName: string
  input: Record<string, unknown>
}

interface ToolResultMessage {
  kind: 'tool-result'
  requestId: string
  output: string
  isError?: boolean
  mutated?: boolean
  summary: string
}

const pendingToolCalls = new Map<string, (result: ToolExecution) => void>()

chrome.webview.addEventListener('message', (ev) => {
  const data = ev.data as ToolResultMessage
  if (!data || data.kind !== 'tool-result') return
  const resolve = pendingToolCalls.get(data.requestId)
  if (!resolve) return
  pendingToolCalls.delete(data.requestId)
  resolve({
    output: data.output,
    isError: data.isError,
    mutated: data.mutated,
    summary: data.summary,
  })
})

function callDotNetTool(toolName: string, input: Record<string, unknown>): Promise<ToolExecution> {
  const requestId = crypto.randomUUID()
  return new Promise((resolve) => {
    pendingToolCalls.set(requestId, resolve)
    const msg: ToolCallMessage = { kind: 'tool-call', requestId, toolName, input }
    chrome.webview.postMessage(msg)
  })
}

const PROVIDER_CONFIG: AiProviderConfig = {
  apiKey: 'test',
  model: 'test-model',
}
const BASE_URL = 'http://127.0.0.1:9000/v1'
const MAX_TOKENS = 1024

function makeTransport(): AgentTransport {
  return {
    stream(request, callbacks): AgentStreamHandle {
      const controller = new AbortController()
      const t0 = performance.now()
      let chunkIndex = 0
      streamOpenAiCompatible(
        BASE_URL,
        PROVIDER_CONFIG,
        request.system,
        request.messages,
        request.tools,
        MAX_TOKENS,
        {
          onDelta: (text) => {
            chunkIndex++
            appendLine(`  [chunk ${chunkIndex} @ +${Math.round(performance.now() - t0)}ms] ${JSON.stringify(text)}`)
            callbacks.onDelta(text)
          },
          onToolCall: callbacks.onToolCall,
          onStopReason: callbacks.onStopReason,
          signal: controller.signal,
        },
      )
        .then(() => callbacks.onDone())
        .catch((e: unknown) => callbacks.onError(e instanceof Error ? e.message : String(e)))
      return { cancel: () => controller.abort() }
    },
  }
}

const wordSkill: AgentSkill = {
  id: 'spike3-word-tools',
  systemPrompt:
    'You are a test assistant running inside a VSTO Word add-in spike (spike 3: real COM tool execution). ' +
    'You can read the document, insert text, and create/edit a native Word chart. Use the tools when asked to.',
  tools: [
    {
      name: 'get_document_context',
      description: "Reads the active Word document's paragraph/word count and a text preview.",
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'insert_content',
      description: 'Inserts a paragraph of text at the end of the active Word document.',
      inputSchema: {
        type: 'object',
        properties: { text: { type: 'string' } },
        required: ['text'],
      },
    },
    {
      name: 'edit_chart',
      description:
        'Creates (if none exists) or edits a native Word chart: sets its title and its first series values.',
      inputSchema: {
        type: 'object',
        properties: {
          title: { type: 'string' },
          values: { type: 'array', items: { type: 'number' } },
        },
        required: ['title', 'values'],
      },
    },
  ],
  executeTool: (call) => callDotNetTool(call.name, call.input),
}

const loop = new AgentLoop({
  transport: makeTransport(),
  skill: wordSkill,
  events: {
    onText: (text) => setAssistantBubble(text),
    onToolStart: (call) => {
      appendLine(`  [tool call] ${call.name}(${JSON.stringify(call.input)})`)
    },
    onToolExecuted: (event) => {
      appendLine(
        `  [tool result] ${event.call.name} -> ${event.execution.isError ? 'ERROR: ' : ''}${event.execution.output}`,
      )
    },
    onTurnEnd: () => appendLine('  [turn end - back to model]'),
    onDone: (result) => {
      setAssistantBubble(result.text || '(no text)')
      setBusy(false)
    },
    onError: (error) => {
      appendLine(`[error] ${error}`)
      setBusy(false)
    },
  },
})

// ---- minimal chat UI ----

const transcript = document.getElementById('transcript') as HTMLDivElement
const input = document.getElementById('input') as HTMLInputElement
const sendBtn = document.getElementById('sendBtn') as HTMLButtonElement

let assistantBubble: HTMLDivElement | null = null

function appendLine(text: string): void {
  const div = document.createElement('div')
  div.className = 'line'
  div.textContent = text
  transcript.appendChild(div)
  transcript.scrollTop = transcript.scrollHeight
}

function setAssistantBubble(text: string): void {
  if (!assistantBubble) {
    assistantBubble = document.createElement('div')
    assistantBubble.className = 'line assistant'
    transcript.appendChild(assistantBubble)
  }
  assistantBubble.textContent = 'assistant: ' + text
  transcript.scrollTop = transcript.scrollHeight
}

function setBusy(busy: boolean): void {
  sendBtn.disabled = busy
  input.disabled = busy
}

function send(): void {
  const text = input.value.trim()
  if (!text || loop.busy) return
  appendLine('user: ' + text)
  input.value = ''
  assistantBubble = null
  setBusy(true)
  loop.run(text)
}

sendBtn.addEventListener('click', send)
input.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') send()
})

appendLine('[spike 2 ready] talking to ' + BASE_URL)
