import { AgentLoop, type AgentSkill, type AgentStreamHandle, type AgentTransport, type ToolExecution } from '@genoffice/agent-core'
import { streamOpenAiCompatible, type AiProviderConfig } from '@genoffice/ai-provider'
import { mountChatUI, type EditingMode } from '@officeai/chat-ui'

// Excel add-in: AgentLoop + streamOpenAiCompatible + the chat UI shell run
// inside a WebView2 page hosted in a VSTO Excel CustomTaskPane, wired to
// real chat persistence via the WebView2 <-> .NET WebMessage bridge
// (chrome.webview.postMessage <-> CoreWebView2.PostWebMessageAsJson).
// Tools call into ExcelTools.cs via the same bridge, mirroring WordAiAddIn's
// pattern: 9 standalone tools plus propose_operations (a 50-operation-kind
// gateway covering writing, formatting, layout, sheet structure, charts/
// visuals, native tables, pivot tables, and misc data operations).

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
    name: 'select_range',
    description: 'Activates a sheet and selects/navigates to a range - UI navigation only, no data change.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_formats',
    description: 'Reads only explicitly-formatted cells (bold/italic/underline/number format) in a range, max 200 cells. Cells with default formatting are omitted.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_sheet_features',
    description: 'Reports a sheet\'s AutoFilter range, freeze panes, conditional-format rule count, defined names (sheet- and workbook-scoped), hidden/protected state, and shape/image count.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' } } },
  },
  {
    name: 'find_cells',
    description: 'Searches cell values/formulas across the workbook (or one sheet via sheetId) for a substring or regex, or scans for error-valued cells (#REF!, #DIV/0!, etc.) via errors_only. look_in: "values"|"formulas"|"both" (default "both"). Requires query or errors_only. Capped at max_results.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string' }, regex: { type: 'boolean' }, look_in: { type: 'string' },
        sheetId: { type: 'string' }, errors_only: { type: 'boolean' }, max_results: { type: 'number' },
      },
      required: ['max_results'],
    },
  },
  {
    name: 'trace_precedents',
    description: 'Lists the cells a formula directly reads from (same-sheet only), flagging any that hold an error value.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'trace_dependents',
    description: 'Lists every formula on the same sheet that reads a given cell.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'propose_operations',
    description:
      'Applies a batch of spreadsheet operations. Each has a "kind":\n' +
      'Writing: "set_cell" (sheet?, address, value), "set_formula" (sheet?, address, formula), "clear_cell" (sheet?, address), ' +
      '"set_range" (sheet?, address, values: value[][]), "clear_range" (sheet?, range).\n' +
      'Formatting: "format_range" (sheet?, address, bold?, italic?, numberFormat?, fillColor? hex).\n' +
      'Layout: "sort_range" (sheet?, range, byColumn, order:"asc"|"desc", hasHeader?), "merge_cells"/"unmerge_cells" (sheet?, range), ' +
      '"set_row_height" (sheet?, row, count?, heightPoints), "set_col_width" (sheet?, column, count?, widthPx), ' +
      '"set_rows_hidden"/"set_cols_hidden" (sheet?, row|column, count?, hidden), "set_freeze" (sheet?, rows, columns), ' +
      '"set_page_setup" (sheet?, orientation?:"portrait"|"landscape", scale?, fitToWidth?, fitToHeight?, printGridlines?, printHeadings?, printArea?, margins?:"normal"|"wide"|"narrow").\n' +
      'Structure: "insert_rows"/"delete_rows" (sheet?, startRow:number 1-based, count:number), "insert_cols"/"delete_cols" (sheet?, startCol:number 1-based, count:number), ' +
      '"add_sheet" (name), "delete_sheet" (sheet?), "duplicate_sheet" (sheet?, name?), "set_sheet_hidden" (sheet?, hidden), ' +
      '"move_sheet" (sheet?, position:number 1-based), "protect_sheet" (sheet?, protected:boolean), "rename_sheet" (sheet?, name).\n' +
      'Charts/visuals: "add_chart" (sheet?, dataRange:string, chartType?:"column"|"line"|"pie"|"bar"|"doughnut", title?:string), ' +
      '"edit_chart" (sheet?, chartPath:string - the chart\'s name, chartType?, title?, legend?:"none"|"right"|"top"|"left"|"bottom", dataLabels?:"none"|"value"|"percent", seriesColors?: {seriesIndex0based: hex}, seriesData?: [{name}]), ' +
      '"delete_visual" (sheet?, visualId:string - shape/chart name), "add_sparkline" (sheet?, dataRange, targetCell?, type?:"line"|"column"|"stacked", color?), ' +
      '"add_shape" (sheet?, shapeType, anchorCell, fillColor?, text?), "edit_shape" (sheet?, visualId, text?, fillColor?, anchorCell?), ' +
      '"add_image" (sheet?, path - LOCAL FILE PATH ONLY, no URLs, anchorCell).\n' +
      'Tables: "add_table" (sheet?, range, name?, style?, bandedRows?), "add_table_row" (sheet?, tableName, row?, count?), ' +
      '"add_table_column" (sheet?, tableName, column?, columnName, count?), "delete_table_row" (sheet?, tableName, row, count?), ' +
      '"delete_table_column" (sheet?, tableName, column, count?), "delete_table" (sheet?, tableName).\n' +
      'Pivot: "add_pivot" (sourceRange, targetCell, targetSheetId?, name?, rowFields?: string|string[], columnField?, pageFields?: string[], values: [{field, agg?:"sum"|"count"|"average"|"max"|"min", formula?, numFmt?}]), "refresh_pivot" (sheet?).\n' +
      'Data: "set_hyperlink" (sheet?, address, target: string|null to remove), "set_note" (sheet?, address, text: string|null to remove), ' +
      '"add_defined_name" (name, ref), "delete_defined_name" (name), "set_filter" (sheet?, range), "clear_filter" (sheet?), ' +
      '"set_filter_criteria" (sheet?, column:number 0-based, values: string[]|null to clear), ' +
      '"add_conditional_format" (sheet?, range, rule: {kind:"number"|"text"|"blank"|"duplicate"|"top10"|"formula"|"colorScale"|"dataBar", ...kind-specific fields, format?:{bold?,fontColor?,fillColor?}}), ' +
      '"clear_conditional_formats" (sheet?), ' +
      '"set_data_validation" (sheet?, range, validation: {kind:"list"|"listRef"|"numberBetween"|"dateBetween"|"formula", ...kind-specific fields}|null to clear - note "checkbox" is NOT supported, will error).',
    inputSchema: { type: 'object', properties: { operations: { type: 'array', items: { type: 'object' } } }, required: ['operations'] },
  },
]

const READ_ONLY_TOOL_NAMES = new Set([
  'get_workbook_context', 'read_range', 'read_cells',
  'select_range', 'read_formats', 'read_sheet_features', 'find_cells', 'trace_precedents', 'trace_dependents',
])
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
    'You have read tools (workbook context, ranges, individual cells, explicit cell formats, sheet features like ' +
    'filters/freeze-panes/defined-names, workbook-wide search including a native error-cell scan, and formula ' +
    'precedent/dependent tracing) and a single propose_operations gateway for all writes - covering cell/range edits, ' +
    'formatting, sorting, row/column sizing and visibility, freeze panes, page setup, sheet management (add/delete/' +
    'duplicate/hide/move/protect/rename), charts, sparklines, shapes and local images, native Excel Tables, native ' +
    'pivot tables, hyperlinks, cell notes, defined names, AutoFilter, conditional formatting, and data validation. ' +
    'Your available tools depend on the current editing mode (Read Only/Comment Only allow only the read tools; ' +
    'Track Changes/Full Autonomy allow everything) - only call tools currently offered to you. ' +
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
