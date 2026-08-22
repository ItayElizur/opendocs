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

const powerPointSkill: AgentSkill = {
  id: 'powerpoint-tools',
  systemPrompt:
    'You are an AI assistant running inside a VSTO PowerPoint add-in. ' +
    'You can read the deck outline (get_deck_context) and the full text of any slide (read_slide).',
  tools: [
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
    // Task 19 wires the actual bridge call + tool-list filtering here.
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
      ui.showError(error)
      ui.setBusy(false)
    },
  },
})

requestHistory()
