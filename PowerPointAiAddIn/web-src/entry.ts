import { AgentLoop, type AgentSkill, type AgentStreamHandle, type AgentTransport, type ToolExecution } from '@genoffice/agent-core'
import { streamOpenAiCompatible, type AiProviderConfig } from '@genoffice/ai-provider'
import { mountChatUI, type EditingMode } from '@officeai/chat-ui'

// Spike 2: prove packages/agent-core's AgentLoop and packages/ai-provider's
// streamOpenAiCompatible run unmodified inside a WebView2 page hosted in a
// VSTO CustomTaskPane, talking to a real (local, for this spike) OpenAI-
// compatible HTTP+SSE endpoint - no Electron IPC hop anywhere in this path.
//
// PowerPoint scaffold (Task 17): mirrors WordAiAddIn's entry.ts wiring
// (WebView2 <-> .NET WebMessage bridge, chat persistence, mode/settings
// stubs) but with an empty tool list - PowerPointTools.cs is a stub until
// Task 18 (readers) / Task 19 (mutation tools + editing-mode gating).

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

const READER_TOOLS = [
  {
    name: 'get_deck_context',
    description: 'Reads a one-line-per-slide outline: slide index and a text preview of its shapes.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_slide',
    description: 'Reads full text of every shape on one slide (0-based index).',
    inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
  },
]

const MUTATION_TOOLS = [
  {
    name: 'set_element_text',
    description: 'Replaces the text content of one shape (0-based slideIndex, 0-based shapeIndex within that slide).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, text: { type: 'string' } },
      required: ['slideIndex', 'shapeIndex', 'text'],
    },
  },
  {
    name: 'set_element_style',
    description: 'Changes text formatting of one shape without changing its text.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: 'number' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        fontSize: { type: 'number' },
        color: { type: 'string' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_element_transform',
    description: 'Moves/resizes/rotates one shape (values in points; rotation in degrees).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: 'number' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_text_box',
    description: 'Creates a new text box on the given slide.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        text: { type: 'string' },
      },
      required: ['slideIndex', 'left', 'top', 'width', 'height', 'text'],
    },
  },
  {
    name: 'add_shape',
    description: 'Creates a shape (rectangle/oval/roundRect) with optional text.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeType: { type: 'string', enum: ['rectangle', 'oval', 'roundRect'] },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        text: { type: 'string' },
      },
      required: ['slideIndex', 'shapeType', 'left', 'top', 'width', 'height'],
    },
  },
  {
    name: 'delete_element',
    description: 'Deletes one shape from a slide.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_slide',
    description: 'Clones an existing slide\'s layout as a new blank (or templated) slide inserted right after it.',
    inputSchema: {
      type: 'object',
      properties: { sourceIndex: { type: 'number' }, clearText: { type: 'boolean' } },
      required: ['sourceIndex'],
    },
  },
]

const ALL_TOOLS = [...READER_TOOLS, ...MUTATION_TOOLS]

const powerPointSkill: AgentSkill = {
  id: 'powerpoint-tools',
  systemPrompt:
    'You are an AI assistant running inside a VSTO PowerPoint add-in. ' +
    'You can read the deck outline (get_deck_context) and the full text of any slide (read_slide). ' +
    'You can also edit the deck: set_element_text, set_element_style, set_element_transform, ' +
    'add_text_box, add_shape, and delete_element.',
  tools: ALL_TOOLS,
  executeTool: (call) => callDotNetTool(call.name, call.input),
}

const root = document.getElementById('root')!
const ui = mountChatUI(root, {
  starters: [
    { en: "Improve this slide's title and copy", he: 'שפר את הכותרת והטקסט של השקופית' },
    { en: "Make this slide's bullets more concise", he: 'קצר את התבליטים בשקופית' },
    { en: 'Check the whole deck for typos and fix them', he: 'בדוק שגיאות כתיב בכל המצגת ותקן אותן' },
  ],
  onCollapseChange: (collapsed) => {
    chrome.webview.postMessage({ kind: collapsed ? 'collapse-pane' : 'expand-pane' })
  },
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
    chrome.webview.postMessage({ kind: 'set-mode', mode })
    powerPointSkill.tools = mode === 'readOnly' || mode === 'commentOnly' ? READER_TOOLS : ALL_TOOLS
  },
  onSettingsSave: (settings) => {
    // Not yet wired to the transport/provider config - deferred (Phase 5).
  },
})

let currentToolGroup: ReturnType<typeof ui.beginToolGroup> | null = null
const activeSteps = new Map<string, ReturnType<ReturnType<typeof ui.beginToolGroup>['addStep']>>()

const loop = new AgentLoop({
  transport: makeTransport(),
  skill: powerPointSkill,
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
      const placeholder = `[Error: ${error}]`
      ui.endAssistantMessage(placeholder)
      persistMessage('assistant', placeholder)
      ui.showError(error)
      ui.setBusy(false)
    },
  },
})

requestHistory()
