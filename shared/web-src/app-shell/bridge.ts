import type { ToolExecution } from '@genoffice/agent-core'

// WebView2 <-> .NET WebMessage bridge (chrome.webview.postMessage <->
// CoreWebView2.PostWebMessageAsJson). This is the ONLY file in the app-shell
// package that touches `chrome` directly - every message kind is wrapped in a
// named export here, so no caller elsewhere constructs a message shape by hand.
declare const chrome: {
  webview: {
    postMessage(message: unknown): void
    addEventListener(type: 'message', listener: (ev: { data: unknown }) => void): void
  }
}

export interface ToolCallMessage {
  kind: 'tool-call'
  requestId: string
  toolName: string
  input: Record<string, unknown>
}

export interface ToolResultMessage {
  kind: 'tool-result'
  requestId: string
  output: string
  isError?: boolean
  mutated?: boolean
  summary: string
}

export interface OtherMessage {
  kind: string
  [key: string]: unknown
}

/** Word's original shape - kept as a distinct type since Word's own code (entry.ts) still names it directly. */
export interface SelectionState {
  hasSelection: boolean
  preview: string
  fullText: string
}

/**
 * FT-2: the raw 'selection-changed' WebMessage, before any app-specific
 * interpretation. Every field below is app-specific and optional - `app`
 * distinguishes Excel/PowerPoint's payloads from Word's (which carries no
 * `app` field at all, only `hasSelection`/`preview`/`fullText`). This bridge
 * module stays app-agnostic (per its own file-header rule) - only
 * bootstrap.ts's per-app describeSelection/classification logic interprets
 * these fields into a SelectionContext.
 */
export interface RawSelectionPayload {
  hasSelection: boolean
  // Word
  preview?: string
  fullText?: string
  // Word (post-hoc fix, 2026-08-24): 0-based paragraph range the selection
  // spans, so the model can address it with replace_blocks.
  startBlockIndex?: number
  endBlockIndex?: number
  // Word (post-hoc addition, 2026-08-24): set when the selection is a
  // table/chart/SmartArt object rather than plain text - the 0-based index
  // read_table/read_chart/read_smartart would use to address it.
  objectKind?: 'table' | 'chart' | 'smartart' | null
  objectIndex?: number
  // Excel ('app: "excel"'), PowerPoint ('app: "powerpoint"'), Outlook ('app: "outlook"')
  app?: 'excel' | 'powerpoint' | 'outlook'
  // Excel (Task 2)
  sheet?: string
  address?: string
  cellCount?: number
  rows?: number
  cols?: number
  firstRow?: number
  firstCol?: string
  entireColumns?: boolean
  entireRows?: boolean
  multi?: boolean
  areaCount?: number
  effectiveAddress?: string | null
  effectiveCellCount?: number
  effectiveRows?: number
  effectiveCols?: number
  // PowerPoint (Task 3)
  selKind?: 'slides' | 'shapes' | 'shapeText'
  slideIndexes?: number[]
  slideIndex?: number
  shapeIndexes?: number[]
  names?: string[]
  textPreview?: string[]
  shapeIndex?: number
  text?: string
  // Outlook ('app: "outlook"') - the Explorer's currently-selected mail
  // item(s) / conversation. subject is the first item's subject.
  count?: number
  entryIds?: string[]
  subject?: string
  senderName?: string
  folderName?: string
  conversationTopic?: string | null
}

export interface BridgeHandlers {
  onHistoryLoaded(messages: Array<{ role: 'user' | 'assistant'; text: string }>): void
  /**
   * Selection-change notifications, sent by all three apps' C# side (Word:
   * WordAiAddIn/TaskPaneHost.cs; Excel/PowerPoint: FT-2 Tasks 2/3). Wiring
   * this unconditionally costs nothing where an app's payload shape is inert
   * to another app's handler - each app's bootstrap.ts config only reads the
   * fields relevant to it.
   */
  onSelectionChanged(selection: RawSelectionPayload): void
  /** FT-1 Task 7/8: the per-document system message, sent once in response to postLoadDocSettings(). */
  onDocSettingsLoaded(systemMessage: string): void
}

const pendingToolCalls = new Map<string, (result: ToolExecution) => void>()

export function initBridge(handlers: BridgeHandlers): void {
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
      if (messages.length > 0) handlers.onHistoryLoaded(messages)
      return
    }
    if (data.kind === 'selection-changed') {
      handlers.onSelectionChanged(data as unknown as RawSelectionPayload)
      return
    }
    if (data.kind === 'doc-settings-loaded') {
      handlers.onDocSettingsLoaded((data as unknown as { systemMessage: string }).systemMessage ?? '')
    }
  })
}

export function requestHistory(): void {
  chrome.webview.postMessage({ kind: 'load-history' })
}

export function requestDocSettings(): void {
  chrome.webview.postMessage({ kind: 'load-doc-settings' })
}

export function saveDocSettings(systemMessage: string): void {
  chrome.webview.postMessage({ kind: 'save-doc-settings', systemMessage })
}

export function persistMessage(role: 'user' | 'assistant', text: string): void {
  chrome.webview.postMessage({ kind: 'append-message', role, text })
}

export function callDotNetTool(toolName: string, input: Record<string, unknown>): Promise<ToolExecution> {
  const requestId = crypto.randomUUID()
  return new Promise((resolve) => {
    pendingToolCalls.set(requestId, resolve)
    const msg: ToolCallMessage = { kind: 'tool-call', requestId, toolName, input }
    chrome.webview.postMessage(msg)
  })
}

export function postNewChatDivider(): void {
  chrome.webview.postMessage({ kind: 'new-chat-divider' })
}

export function postMode(mode: string): void {
  chrome.webview.postMessage({ kind: 'set-mode', mode })
}

export function postCollapse(collapsed: boolean): void {
  chrome.webview.postMessage({ kind: collapsed ? 'collapse-pane' : 'expand-pane' })
}

export function postTlsBypass(enabled: boolean): void {
  chrome.webview.postMessage({ kind: 'set-tls-bypass', enabled })
}
