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
  {
    name: 'set_element_fill',
    description: 'Sets a shape\'s solid fill color, or "none" to remove its fill.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, fill: { type: 'string' } },
      required: ['slideIndex', 'shapeIndex', 'fill'],
    },
  },
  {
    name: 'set_element_stroke',
    description: 'Sets a shape\'s outline/stroke color and width, or removes it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        color: { type: 'string' }, widthPt: { type: 'number' }, remove: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_slide_background',
    description: 'Sets a solid background color for one slide, or slideIndex=-1 for every slide in the deck.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, color: { type: 'string' } },
      required: ['slideIndex', 'color'],
    },
  },
  {
    name: 'ungroup_element',
    description: 'Promotes a group shape\'s direct children to top-level shapes. Shape indices change after this call - re-read the slide before addressing the promoted shapes.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_table',
    description: 'Adds a native PowerPoint table, optionally pre-filled with cell text (row-major array of arrays).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, rows: { type: 'number' }, cols: { type: 'number' },
        cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
      },
      required: ['slideIndex', 'rows', 'cols'],
    },
  },
  {
    name: 'edit_table_cell',
    description: 'Replaces one table cell\'s text (0-based row/col).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, row: { type: 'number' }, col: { type: 'number' }, paragraphs: { type: 'string' } },
      required: ['slideIndex', 'shapeIndex', 'row', 'col', 'paragraphs'],
    },
  },
  {
    name: 'edit_table_structure',
    description: 'Inserts or deletes a table row/column. kind: "insert-row"|"delete-row"|"insert-col"|"delete-col".',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, kind: { type: 'string' }, index: { type: 'number' }, before: { type: 'boolean' } },
      required: ['slideIndex', 'shapeIndex', 'kind', 'index'],
    },
  },
  {
    name: 'edit_table_style',
    description: 'Applies granular table styling: firstRow/bandRow (header row / banded rows), shadingColor (all cells), borderColor/borderWidthPt/borderPreset ("all"|"none").',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        firstRow: { type: 'boolean' }, bandRow: { type: 'boolean' }, shadingColor: { type: 'string' },
        borderColor: { type: 'string' }, borderWidthPt: { type: 'number' }, borderPreset: { type: 'string' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_chart',
    description: 'Adds a native, editable PowerPoint chart. kind: "bar"|"barStacked"|"line"|"area"|"pie"|"doughnut".',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, kind: { type: 'string' }, title: { type: 'string' },
        categories: { type: 'array', items: { type: 'string' } },
        series: { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } } } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
      },
      required: ['slideIndex', 'kind', 'categories', 'series'],
    },
  },
  {
    name: 'edit_chart',
    description: 'Modifies an existing chart\'s type/title/legend position/data labels/gridlines.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        chartType: { type: 'string' }, title: { type: 'string' },
        legendPos: { type: 'string' }, dataLabels: { type: 'boolean' }, gridlines: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_smartart',
    description: 'Adds a shape-composed SmartArt diagram. layout: "list"|"process"|"cycle"|"hierarchy"|"pyramid"|"matrix"|"venn". items are flat node texts, one per top-level node.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, layout: { type: 'string' },
        items: { type: 'array', items: { type: 'string' } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
      },
      required: ['slideIndex', 'layout', 'items'],
    },
  },
  {
    name: 'crop_image',
    description: 'Non-destructively crops a picture shape. l/t/r/b are 0..1 fractions of the current on-slide image size cut from each edge; all zero clears the crop.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, l: { type: 'number' }, t: { type: 'number' }, r: { type: 'number' }, b: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'l', 't', 'r', 'b'],
    },
  },
  {
    name: 'set_picture_opacity',
    description: 'Sets a picture shape\'s overall opacity, 0 (invisible) to 1 (fully opaque).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, opacity: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'opacity'],
    },
  },
  {
    name: 'replace_image',
    description: 'Swaps a picture shape\'s image content in place from a local file path, keeping position/size/rotation/approximate z-order.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, localPath: { type: 'string' }, keepCrop: { type: 'boolean' } },
      required: ['slideIndex', 'shapeIndex', 'localPath'],
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
