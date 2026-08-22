import { AgentLoop, type AgentSkill, type AgentStreamHandle, type AgentTransport, type ToolExecution } from '@genoffice/agent-core'
import { streamOpenAiCompatible, type AiProviderConfig } from '@genoffice/ai-provider'
import { mountChatUI, type EditingMode } from '@officeai/chat-ui'

// Excel add-in scaffold (Task 13): proves AgentLoop + streamOpenAiCompatible +
// the chat UI shell run inside a WebView2 page hosted in a VSTO Excel
// CustomTaskPane, wired to real chat persistence via the WebView2 <-> .NET
// WebMessage bridge (chrome.webview.postMessage <-> CoreWebView2.PostWebMessageAsJson).
// No Excel tools exist yet (excelSkill.tools is empty) - those land in
// Tasks 14-16, calling into ExcelTools.cs (currently a stub) via the same
// bridge, mirroring WordAiAddIn's pattern.

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

const ALL_TOOLS: AgentSkill['tools'] = [
  {
    name: 'get_workbook_context',
    description: "Reads the active sheet's name, used range, and current selection address.",
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_range',
    description: 'Reads cell values in a rectangular range (e.g. "A1:C10"), max 2000 cells. Optional sheet name defaults to the active sheet.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_cells',
    description: 'Reads specific scattered cell addresses (e.g. ["A1","C5"]).',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, addresses: { type: 'array', items: { type: 'string' } } }, required: ['addresses'] },
  },
  {
    name: 'propose_operations',
    description:
      'Applies a batch of spreadsheet operations. Each has a "kind": ' +
      '"set_cell" (sheet?, address, value), "set_formula" (sheet?, address, formula), ' +
      '"set_range" (sheet?, address, values: value[][]), ' +
      '"format_range" (sheet?, address, bold?, italic?, numberFormat?, fillColor? - hex like "#FFFF00"), ' +
      '"insert_rows"/"delete_rows" (sheet?, startRow:number 1-based, count:number), ' +
      '"insert_cols"/"delete_cols" (sheet?, startCol:number 1-based, count:number), ' +
      '"add_chart" (sheet?, dataRange:string, chartType?:"column"|"line"|"pie", title?:string).',
    inputSchema: { type: 'object', properties: { operations: { type: 'array', items: { type: 'object' } } }, required: ['operations'] },
  },
]

const READ_ONLY_TOOL_NAMES = new Set(['get_workbook_context', 'read_range', 'read_cells'])
const READ_ONLY_TOOLS = ALL_TOOLS.filter((tool) => READ_ONLY_TOOL_NAMES.has(tool.name))

let editingMode: EditingMode = 'fullAutonomy'

function toolsForMode(): AgentSkill['tools'] {
  // Excel has no add_comment-equivalent tool yet, so Comment Only mode
  // allows the same read-only set as Read Only mode (documented gap - see
  // Task 16 brief). Track Changes and Full Autonomy both get the full list;
  // ExcelTools.Execute enforces the same policy server-side regardless of
  // what the client sends.
  if (editingMode === 'readOnly' || editingMode === 'commentOnly') return READ_ONLY_TOOLS
  return ALL_TOOLS
}

const excelSkill: AgentSkill = {
  id: 'excel-tools',
  systemPrompt:
    'You are an assistant running inside a VSTO Excel add-in. You can help the user with their active workbook. ' +
    'You have tools to read the workbook (context, ranges, individual cells) and to propose batches of write ' +
    'operations (set cell values/formulas, set ranges, format cells, insert/delete rows and columns, add charts). ' +
    'Use the tools when asked to inspect or modify the spreadsheet.',
  get tools() {
    return toolsForMode()
  },
  executeTool: (call) => callDotNetTool(call.name, call.input),
}

const root = document.getElementById('root')!
const ui = mountChatUI(root, {
  starters: [
    { en: 'Summarize this sheet', he: 'סכם את הגיליון הזה' },
    { en: 'Add a totals row', he: 'הוסף שורת סיכום' },
    { en: 'Check the formulas', he: 'בדוק את הנוסחאות' },
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
    editingMode = mode
    chrome.webview.postMessage({ kind: 'set-mode', mode })
  },
  onSettingsSave: (settings) => {
    // Not yet wired to the transport/provider config - deferred (Phase 5).
  },
})

let currentToolGroup: ReturnType<typeof ui.beginToolGroup> | null = null
const activeSteps = new Map<string, ReturnType<ReturnType<typeof ui.beginToolGroup>['addStep']>>()

const loop = new AgentLoop({
  transport: makeTransport(),
  skill: excelSkill,
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
