import { AgentLoop, type AgentSkill, type AgentStreamHandle, type AgentTransport, type ToolExecution } from '@genoffice/agent-core'
import { streamOpenAiCompatible, type AiProviderConfig } from '@genoffice/ai-provider'
import { mountChatUI, type EditingMode } from '@officeai/chat-ui'

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

interface OtherMessage {
  kind: string
  [key: string]: unknown
}

const pendingToolCalls = new Map<string, (result: ToolExecution) => void>()

chrome.webview.addEventListener('message', (ev) => {
  const data = ev.data as OtherMessage & ToolResultMessage
  if (!data) return
  if (data.kind === 'tool-result') {
    const resolve = pendingToolCalls.get(data.requestId)
    if (!resolve) return
    pendingToolCalls.delete(data.requestId)
    resolve({
      output: data.output,
      isError: data.isError,
      mutated: data.mutated,
      summary: data.summary,
    })
    return
  }
  if (data.kind === 'history-loaded') {
    const messages = data.messages as Array<{ role: 'user' | 'assistant'; text: string }>
    if (messages.length > 0) {
      ui.showHistoric(messages)
      loop.restore(messages.map((m) => ({ role: m.role, text: m.text })))
    }
  }
})

function requestHistory(): void {
  chrome.webview.postMessage({ kind: 'load-history' })
}

function persistMessage(role: 'user' | 'assistant', text: string): void {
  chrome.webview.postMessage({ kind: 'append-message', role, text })
}

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
      streamOpenAiCompatible(
        BASE_URL,
        PROVIDER_CONFIG,
        request.system,
        request.messages,
        request.tools,
        MAX_TOKENS,
        {
          onDelta: callbacks.onDelta,
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
    {
      name: 'read_blocks',
      description: 'Reads paragraphs [startIndex, endIndex] (0-based, inclusive) of the active document, one per line prefixed with its index.',
      inputSchema: {
        type: 'object',
        properties: { startIndex: { type: 'number' }, endIndex: { type: 'number' } },
        required: ['startIndex', 'endIndex'],
      },
    },
    {
      name: 'replace_blocks',
      description: 'Replaces paragraphs [startIndex, endIndex] (0-based, inclusive) with new text (empty text deletes the range).',
      inputSchema: {
        type: 'object',
        properties: { startIndex: { type: 'number' }, endIndex: { type: 'number' }, text: { type: 'string' } },
        required: ['startIndex', 'endIndex', 'text'],
      },
    },
    {
      name: 'apply_commands',
      description:
        'Applies a batch of formatting/editing commands. Each command has a "kind": ' +
        '"set_bold"/"set_italic" (fields: startIndex, endIndex, value:boolean), ' +
        '"set_heading" (fields: index, level:0-9, 0=Normal style), ' +
        '"find_replace" (fields: find:string, replace:string, matchCase?:boolean).',
      inputSchema: { type: 'object', properties: { commands: { type: 'array', items: { type: 'object' } } }, required: ['commands'] },
    },
  ],
  executeTool: (call) => callDotNetTool(call.name, call.input),
}

const root = document.getElementById('root')!
const ui = mountChatUI(root, {
  title: 'Airchat Office',
  onSend: (text) => {
    if (loop.busy) return
    ui.addUserMessage(text)
    ui.beginAssistantMessage()
    ui.setBusy(true)
    persistMessage('user', text)
    loop.run(text)
  },
  onNewChat: () => {
    chrome.webview.postMessage({ kind: 'new-chat-divider' })
    loop.reset()
    ui.resetToEmpty()
  },
  onModeChange: (mode: EditingMode) => {
    // Task 11 wires the actual bridge call + tool-list filtering here.
  },
  onSettingsSave: (settings) => {
    // Not yet wired to the transport/provider config - deferred (Phase 5).
  },
})

let currentToolGroup: ReturnType<typeof ui.beginToolGroup> | null = null
const activeSteps = new Map<string, ReturnType<ReturnType<typeof ui.beginToolGroup>['addStep']>>()

const loop = new AgentLoop({
  transport: makeTransport(),
  skill: wordSkill,
  events: {
    onText: (text) => ui.updateAssistantMessage(text),
    onToolStart: (call) => {
      if (!currentToolGroup) currentToolGroup = ui.beginToolGroup()
      activeSteps.set(call.id, currentToolGroup.addStep(call.name, call.input))
    },
    onToolExecuted: (event) => {
      activeSteps.get(event.call.id)?.complete({
        output: event.execution.output,
        isError: event.execution.isError,
        mutated: event.execution.mutated,
      })
      activeSteps.delete(event.call.id)
    },
    onTurnEnd: () => {
      currentToolGroup?.end()
      currentToolGroup = null
    },
    onDone: (result) => {
      const finalText = result.text || '(no text)'
      ui.endAssistantMessage(finalText)
      ui.setBusy(false)
      persistMessage('assistant', finalText)
    },
    onError: (error) => {
      ui.showError(error)
      ui.setBusy(false)
    },
  },
})

requestHistory()
