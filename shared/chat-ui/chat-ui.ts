import './chat-ui.css'
import { AI_PROVIDERS, type AiProviderId } from '@genoffice/ai-provider'

export type EditingMode = 'readOnly' | 'commentOnly' | 'trackChanges' | 'fullAutonomy'

export type Lang = 'en' | 'he'

const STRINGS: Record<string, Record<Lang, string>> = {
  panelTitle:           { en: 'Airchat Office', he: "איירצ'אט אופיס" },
  inputPlaceholder:     { en: 'Ask Airchat Office to edit this document...', he: 'בקש מ-Airchat Office לערוך את המסמך...' },
  send:                 { en: 'Send', he: 'שלח' },
  stop:                 { en: 'Stop', he: 'עצור' },
  newChat:              { en: 'New chat', he: 'שיחה חדשה' },
  settings:             { en: 'Settings', he: 'הגדרות' },
  settingsTitle:        { en: 'Airchat Office Settings', he: 'הגדרות Airchat Office' },
  settingsBaseUrl:      { en: 'API Base URL', he: 'כתובת בסיס API' },
  settingsApiKey:       { en: 'API Key', he: 'מפתח API' },
  settingsModel:        { en: 'Model name', he: 'שם המודל' },
  settingsSkipTls:      { en: 'Skip TLS certificate verification (insecure - testing only)', he: 'דלג על אימות אישור TLS (לא מאובטח - לבדיקות בלבד)' },
  settingsLanguage:     { en: 'Language', he: 'שפה' },
  save:                 { en: 'Save', he: 'שמור' },
  collapse:             { en: 'Collapse panel', he: 'כווץ חלונית' },
  historySep:           { en: 'Earlier conversation', he: 'שיחה קודמת' },
  scopeWholeDoc:        { en: 'Whole document', he: 'כל המסמך' },
  emptyTitle:           { en: 'What can I help with?', he: 'איך אפשר לעזור?' },
  modeReadOnly:         { en: 'Read only', he: 'קריאה בלבד' },
  modeReadOnlyDesc:     { en: 'AI can only read, never edit', he: 'הבינה יכולה רק לקרוא, לא לערוך' },
  modeCommentOnly:      { en: 'Comment only', he: 'הערות בלבד' },
  modeCommentOnlyDesc:  { en: 'Adds comments, no content edits', he: 'מוסיפה הערות בלבד, ללא עריכת תוכן' },
  modeTrackChanges:     { en: 'Track changes', he: 'מעקב אחר שינויים' },
  modeTrackChangesDesc: { en: 'Edits as reviewable revisions', he: 'עריכות כתיקונים לאישור' },
  modeFullAutonomy:     { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
  modeFullAutonomyDesc: { en: 'Edits applied directly', he: 'עריכות מוחלות ישירות' },
  noticeTruncated:      { en: 'The reply was cut off by the length limit.', he: 'התשובה נקטעה עקב מגבלת האורך.' },
  noticeTurnLimit:      { en: 'The tool-step limit for this request was reached.', he: 'הגעת למגבלת שלבי הכלים לבקשה זו.' },
  noticeContinue:       { en: 'Continue', he: 'המשך' },
  emptyReply:           { en: '(no reply)', he: '(אין תשובה)' },
  thinking:             { en: 'Thinking', he: 'חושב' },
  settingsProvider:     { en: 'Provider', he: 'ספק' },
  testConnection:       { en: 'Test connection', he: 'בדוק חיבור' },
  testTesting:          { en: 'Testing...', he: 'בודק...' },
  testSuccess:          { en: 'Connection OK.', he: 'החיבור תקין.' },
  moreSettings:         { en: 'More settings', he: 'הגדרות נוספות' },
  backToChat:           { en: 'Back to conversation', he: 'חזרה לשיחה' },
  settingsScreenTitle:  { en: 'Settings', he: 'הגדרות' },
  sectionConnection:    { en: 'Connection', he: 'חיבור' },
  sectionLanguage:      { en: 'Language', he: 'שפה' },
  sectionScope:         { en: 'Edit scope', he: 'היקף עריכה' },
  sectionTools:         { en: 'Tools', he: 'כלים' },
  sectionDocMessage:    { en: 'Document guidelines', he: 'הנחיות למסמך' },
  scopeNote:            { en: 'Applies immediately and resets tool registration below.', he: 'חל מיידית ומאפס את רישום הכלים למטה.' },
  docMessagePlaceholder:{ en: 'Background and guidelines about this document - included at the start of every new conversation.', he: 'רקע והנחיות לגבי המסמך הזה - ייכלל בתחילת כל שיחה חדשה.' },
  docMessageSaveNote:   { en: 'Applies at the start of the next New chat, not the current conversation.', he: 'חל בתחילת השיחה החדשה הבאה, לא בשיחה הנוכחית.' },
  savedNote:            { en: 'Saved', he: 'נשמר' },
  toolOutOfScopeHint:   { en: 'Not available in the current edit scope - change the scope above to enable.', he: 'לא זמין בהיקף העריכה הנוכחי - שנה את ההיקף למעלה כדי לאפשר.' },
  toolsAvailable:       { en: 'Available', he: 'זמינים' },
  toolsOutOfScope:      { en: 'Out of scope', he: 'מחוץ להיקף' },
  discardChangesConfirm:{ en: 'Discard unsaved changes?', he: 'לבטל שינויים שלא נשמרו?' },
  lastToolRefused:      { en: 'At least one tool must stay registered.', he: 'לפחות כלי אחד חייב להישאר רשום.' },
  // FT-2 Task 5: per-app "no selection" wording and the selection pill's
  // small localized vocabulary (Excel's extent words, PowerPoint's counts).
  scopeWholeSheet:      { en: 'Whole sheet', he: 'כל הגיליון' },
  scopeWholeDeck:       { en: 'Whole deck', he: 'כל המצגת' },
  scopeSelectionPrefix: { en: 'Selection: "', he: 'בחירה: "' },
  scopeColumn:          { en: 'column', he: 'עמודה' },
  scopeColumns:         { en: 'columns', he: 'עמודות' },
  scopeRow:             { en: 'row', he: 'שורה' },
  scopeRowsWithData:    { en: 'rows with data', he: 'שורות עם נתונים' },
  scopeWithData:        { en: 'with data', he: 'עם נתונים' },
  scopeEmpty:           { en: 'empty', he: 'ריק' },
  scopeAreas:           { en: 'areas', he: 'אזורים' },
  scopeCells:           { en: 'cells', he: 'תאים' },
  scopeSlide:           { en: 'slide', he: 'שקופית' },
  scopeSlides:          { en: 'slides', he: 'שקופיות' },
  scopeShape:           { en: 'shape', he: 'צורה' },
  scopeShapes:          { en: 'shapes', he: 'צורות' },
  scopeTextIn:          { en: 'text in', he: 'טקסט ב' },
  // Post-hoc addition (2026-08-24, user-reported): Word table/chart/SmartArt
  // selection nouns, so the pill shows a real pointer instead of nothing.
  scopeTable:           { en: 'Table', he: 'טבלה' },
  scopeChart:           { en: 'Chart', he: 'תרשים' },
  scopeSmartArt:        { en: 'SmartArt', he: 'SmartArt' },
  // Outlook: mailbox scope + selected-mail pill.
  scopeWholeMailbox:    { en: 'Whole mailbox', he: 'כל תיבת הדואר' },
  scopeEmail:           { en: 'Selected email', he: 'הודעה נבחרת' },
  scopeEmails:          { en: 'emails', he: 'הודעות' },
}

/**
 * FT-2 Task 5: the structured facts of an Excel/PowerPoint selection - the
 * shell (bootstrap.ts) classifies the C# payload into one of these; this
 * component renders and localizes the actual words, the same way every other
 * piece of UI text here does, so the pill relocalizes on a language switch
 * for free (via refreshScopeHint(), already called from setLang()).
 */
export type SelectionExtent =
  | { kind: 'cell'; address: string }
  | { kind: 'range'; address: string; rows: number; cols: number }
  | { kind: 'wholeColumn'; col: string; dataRows: number | null }
  | { kind: 'wholeColumns'; firstCol: string; lastCol: string; cols: number; dataRows: number | null }
  | { kind: 'wholeRow'; row: number }
  | { kind: 'multiArea'; areaCount: number; cellCount: number }
  | { kind: 'slides'; count: number }
  | { kind: 'shapes'; count: number; primaryName: string | null }
  | { kind: 'shapeText'; shapeName: string | null }
  | { kind: 'wordObject'; objectKind: 'table' | 'chart' | 'smartart'; objectIndex: number }
  | { kind: 'mailSelection'; count: number; subject: string }

export interface SettingsSavePayload {
  provider: AiProviderId
  apiKey: string
  model: string
  baseUrl: string
  skipTlsVerify: boolean
  lang: Lang
  /** only present when saved from the full settings view (FT-1), not the quick dropdown */
  docSystemMessage?: string
  /** only present when saved from the full settings view (FT-1) - registration itself already took effect live via onToolRegistrationChange; this is an echo for symmetry with the other fields */
  registeredTools?: string[]
}

/** One tool's UI-only display info (FT-1 Task 5) - distinct from the tool's
 * model-facing JSON-schema `description`, which stays English and is tuned
 * for the model, not the user. */
export interface ToolDisplayEntry {
  name: string
  label: { en: string; he: string }
  description: { en: string; he: string }
}

export interface InitialSettings {
  provider?: AiProviderId
  apiKey?: string
  model?: string
  baseUrl?: string
  skipTlsVerify?: boolean
  /** per-provider saved values, so switching providers restores that provider's own key/model */
  providers?: Record<string, { apiKey: string; model: string; baseUrl?: string }>
}

export interface ChatUIOptions {
  onSend: (text: string) => void
  /** Post-hoc addition (2026-08-24, user-requested): stops the in-flight run. The send button becomes a stop button while busy; omit to leave it disabled while busy instead (previous behavior). */
  onStop?: () => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: SettingsSavePayload) => void
  /** Tries connecting with the form's current (unsaved) values; resolves with a short success message, rejects with the error text. */
  onSettingsTest?: (settings: SettingsSavePayload) => Promise<string>
  starters: Array<{ en: string; he: string }>
  onCollapseChange: (collapsed: boolean) => void
  /** Every tool this app implements, for the full settings view's tool list (FT-1). Omit to leave the section empty. */
  tools?: ToolDisplayEntry[]
  /** Fires immediately when a tool is toggled on/off in the settings view - registration takes effect right away, not gated behind Save. */
  onToolRegistrationChange?: (registered: string[]) => void
  // Pre-fills the settings form on mount (e.g. from the host app's own
  // persisted storage) so the user doesn't see blank fields every time they
  // reopen the document, even though a value was saved previously.
  initialSettings?: InitialSettings
  /** FT-2 Task 5: which "no selection" wording the scope-hint pill shows - defaults to 'doc' (Word). Excel passes 'sheet', PowerPoint 'deck', Outlook 'mailbox'. */
  scopeUnit?: 'doc' | 'sheet' | 'deck' | 'mailbox'
  /** Restricts the editing-mode menu to this subset, in this order. Defaults to all four modes. Outlook passes ['readOnly', 'fullAutonomy']. */
  modes?: EditingMode[]
}

export interface ToolStepHandle {
  complete(result: { output: string; isError?: boolean; mutated?: boolean }): void
}

export interface ToolGroupHandle {
  addStep(toolName: string, input: Record<string, unknown>): ToolStepHandle
  end(): void
}

export interface ChatUIHandle {
  addUserMessage(text: string): void
  beginAssistantMessage(): void
  updateAssistantMessage(cumulativeText: string): void
  endAssistantMessage(finalText: string): void
  /** One model turn (that called tools) finished; the run continues. Seals any
   *  midway reply so it can't be overwritten, and shows the thinking indicator
   *  again for the reasoning gap before the next turn. */
  endTurn(): void
  beginToolGroup(): ToolGroupHandle
  setBusy(busy: boolean): void
  showError(message: string): void
  resetToEmpty(): void
  showHistoric(messages: Array<{ role: 'user' | 'assistant'; text: string }>): void
  /**
   * `preview` is Word's existing quoted-text-excerpt form (wrapped in the
   * localized `scopeSelectionPrefix` + `..."`). `extent` is Excel/PowerPoint's
   * form (FT-2 Task 5) - a structured, already-classified selection that this
   * component renders and localizes itself, shown as-is with no quoting/
   * ellipsis (an address is not a text excerpt). At most one of the two is
   * set at a time.
   */
  setSelectionScope(selection: { hasSelection: boolean; preview?: string; extent?: SelectionExtent } | null): void
  /**
   * Non-error, in-transcript informational row (PP-4) - distinct from
   * showError, which is red/alarming. `onContinue`, when passed, renders a
   * button that fires once and then removes itself.
   */
  showNotice(kind: 'truncated' | 'turnLimit', onContinue?: () => void): void
  /** Looks up a STRINGS entry in the panel's current language - exposed so a
   * host app's shell (bootstrap.ts) can localize its own strings (e.g. the
   * "(no reply)" fallback) without duplicating this table. */
  translate(key: string): string
  /** FT-1 Task 4: pushes down which tools are in-scope for the current editing mode and which are currently registered. Called once after mount and again on every mode change. */
  setToolScope(available: string[], registered: string[]): void
  /** FT-1 Task 7/8: populates the document-guidelines textarea once the async load-doc-settings round trip resolves. */
  setDocSystemMessage(message: string): void
  /** Opens the full settings view programmatically (FT-1 Task 1) - the "More settings" button already does this; exposed for completeness. */
  openSettings(): void
}

const MODES: EditingMode[] = ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']

// The mode menu (composer) and the scope control (settings) both render from
// this list. `options.modes` narrows it per app - Outlook drops Comment only /
// Track changes, which have no meaning for mail. Order follows the passed list.
function resolveModes(requested?: EditingMode[]): EditingMode[] {
  if (!requested || requested.length === 0) return MODES
  return requested.filter((m) => MODES.indexOf(m) !== -1)
}

function modeStringKey(mode: EditingMode): string {
  return 'mode' + mode.charAt(0).toUpperCase() + mode.slice(1)
}

function modeMenuItemsHtml(modes: EditingMode[], selected: EditingMode): string {
  return modes
    .map((mode) => {
      const key = modeStringKey(mode)
      const sel = mode === selected ? ' selected' : ''
      return (
        `<div class="ai-mode-menu-item${sel}" data-mode="${mode}">` +
        `<span data-t="${key}">${STRINGS[key]?.en ?? mode}</span>` +
        `<span class="desc" data-t="${key}Desc">${STRINGS[key + 'Desc']?.en ?? ''}</span></div>`
      )
    })
    .join('')
}

const GEAR_ICON = '&#9881;'
// Shown in place of the gear while the full settings view is open, since at
// that point clicking the button navigates back to the chat, not to settings.
const CHAT_ICON = '<svg width="14" height="14" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="1.5" y="2.5" width="13" height="8" rx="1.8" fill="currentColor"/><polygon points="4,10.5 4,13.3 7.2,10.5" fill="currentColor"/></svg>'

function escapeHtml(s: string): string {
  const div = document.createElement('div')
  div.textContent = s
  return div.innerHTML
}

// Renders a practical subset of markdown (headings, bold/italic/code spans,
// links, unordered/ordered lists, tables, blockquotes, fenced code blocks,
// horizontal rules) for assistant messages. Every character of the model's
// raw text passes through escapeHtml before it lands in the output - the
// only real tags in the result are ones this function builds around markdown
// syntax it recognizes, so a literal "<script>" (or any other HTML) in the
// model's response renders as visible text instead of being parsed/executed.
function renderMarkdown(src: string): string {
  const lines = src.replace(/\r\n?/g, '\n').split('\n')

  function inline(text: string): string {
    let s = escapeHtml(text)
    s = s.replace(/`([^`]+)`/g, (_m, code: string) => `<code>${code}</code>`)
    s = s.replace(/\*\*([^\s*][^*]*?)\*\*/g, '<strong>$1</strong>')
    s = s.replace(/\*([^\s*][^*]*?)\*/g, '<em>$1</em>')
    s = s.replace(/(?<![A-Za-z0-9_])_([^\s_][^_]*?)_(?![A-Za-z0-9_])/g, '<em>$1</em>')
    s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>')
    return s
  }

  const isTableSepLine = (line: string): boolean =>
    /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(line) && line.includes('-')
  const isFence = (line: string): boolean => /^\s*(`{3,}|~{3,})/.test(line)
  const isHeading = (line: string): boolean => /^#{1,6}\s+/.test(line)
  const isRule = (line: string): boolean => /^(-{3,}|\*{3,}|_{3,})\s*$/.test(line.trim())
  const isBlockquote = (line: string): boolean => /^>\s?/.test(line)
  const isUl = (line: string): boolean => /^\s*[-*+]\s+/.test(line)
  const isOl = (line: string): boolean => /^\s*\d+[.)]\s+/.test(line)
  const isTableStart = (line: string, next: string | undefined): boolean =>
    line.includes('|') && next !== undefined && isTableSepLine(next)
  const isBlockStart = (line: string, next: string | undefined): boolean =>
    !line.trim() || isFence(line) || isHeading(line) || isRule(line) || isBlockquote(line) ||
    isUl(line) || isOl(line) || isTableStart(line, next)

  const splitRow = (line: string): string[] => {
    let row = line.trim()
    if (row.startsWith('|')) row = row.slice(1)
    if (row.endsWith('|')) row = row.slice(0, -1)
    return row.split('|').map((c) => c.trim())
  }

  const out: string[] = []
  let i = 0
  while (i < lines.length) {
    const line = lines[i]!

    if (!line.trim()) {
      i++
      continue
    }

    if (isFence(line)) {
      const marker = line.trim()[0]!
      const codeLines: string[] = []
      i++
      while (i < lines.length && !new RegExp(`^\\s*${marker}{3,}\\s*$`).test(lines[i]!)) {
        codeLines.push(lines[i]!)
        i++
      }
      i++ // skip closing fence, or run off the end if the stream hasn't closed it yet
      out.push(`<pre><code>${escapeHtml(codeLines.join('\n'))}</code></pre>`)
      continue
    }

    const heading = line.match(/^(#{1,6})\s+(.*)$/)
    if (heading) {
      const level = heading[1]!.length
      out.push(`<h${level}>${inline(heading[2]!.trim())}</h${level}>`)
      i++
      continue
    }

    if (isRule(line)) {
      out.push('<hr>')
      i++
      continue
    }

    if (isTableStart(line, lines[i + 1])) {
      const headCells = splitRow(line)
      i += 2
      const bodyRows: string[][] = []
      while (i < lines.length && lines[i]!.trim() && lines[i]!.includes('|')) {
        bodyRows.push(splitRow(lines[i]!))
        i++
      }
      const thead = `<tr>${headCells.map((c) => `<th>${inline(c)}</th>`).join('')}</tr>`
      const tbody = bodyRows.map((r) => `<tr>${r.map((c) => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')
      out.push(`<table><thead>${thead}</thead><tbody>${tbody}</tbody></table>`)
      continue
    }

    if (isBlockquote(line)) {
      const quoteLines: string[] = []
      while (i < lines.length && isBlockquote(lines[i]!)) {
        quoteLines.push(lines[i]!.replace(/^>\s?/, ''))
        i++
      }
      out.push(`<blockquote>${renderMarkdown(quoteLines.join('\n'))}</blockquote>`)
      continue
    }

    if (isUl(line)) {
      const items: string[] = []
      while (i < lines.length && isUl(lines[i]!)) {
        items.push(`<li>${inline(lines[i]!.replace(/^\s*[-*+]\s+/, ''))}</li>`)
        i++
      }
      out.push(`<ul>${items.join('')}</ul>`)
      continue
    }

    if (isOl(line)) {
      const items: string[] = []
      while (i < lines.length && isOl(lines[i]!)) {
        items.push(`<li>${inline(lines[i]!.replace(/^\s*\d+[.)]\s+/, ''))}</li>`)
        i++
      }
      out.push(`<ol>${items.join('')}</ol>`)
      continue
    }

    const paraLines: string[] = []
    while (i < lines.length && !isBlockStart(lines[i]!, lines[i + 1])) {
      paraLines.push(lines[i]!)
      i++
    }
    out.push(`<p>${paraLines.map(inline).join('<br>')}</p>`)
  }

  return out.join('')
}

function truncateForDisplay(s: string, max: number): string {
  return s.length > max ? s.slice(0, max) + '…' : s
}

/** Inline cap for a tool's rendered output; longer outputs get a "show all" toggle. */
const TOOL_OUTPUT_PREVIEW_CHARS = 2_000

function emptyStateHtml(options: ChatUIOptions, currentLang: Lang): string {
  const pills = options.starters
    .map((s) => `<div class="ai-starter">${escapeHtml(s[currentLang])}</div>`)
    .join('')
  const title = escapeHtml(STRINGS.emptyTitle[currentLang])
  return `<div class="ai-chat-empty"><img class="ai-chat-empty-bg" src="chat-empty-bg.svg" alt="" /><div class="ai-chat-empty-title" data-t="emptyTitle">${title}</div><div class="ai-starters">${pills}</div></div>`
}

export function mountChatUI(root: HTMLElement, options: ChatUIOptions): ChatUIHandle {
  const menuModes = resolveModes(options.modes)
  const defaultMode: EditingMode = menuModes.indexOf('fullAutonomy') !== -1 ? 'fullAutonomy' : menuModes[menuModes.length - 1]
  root.innerHTML = `
    <div class="ai-dock">
      <div class="ai-rail" data-t="panelTitle"></div>
      <div class="ai-panel">
      <div class="ai-panel-header">
        <div class="ai-panel-title"><img class="ai-logo" src="logo.png" alt="" /><span data-t="panelTitle">Airchat Office</span></div>
        <div class="ai-header-actions">
          <button class="ai-header-btn" data-t-title="newChat">+</button>
          <button class="ai-header-btn" data-t-title="settings">&#9881;</button>
          <button class="ai-header-btn" data-t-title="collapse">&#x276E;</button>
        </div>
        <div class="ai-settings-panel" id="settingsPanel">
          <h4 data-t="settingsTitle">Airchat Office Settings</h4>
          <div class="ai-field"><label data-t="settingsProvider">Provider</label><select class="ai-settings-provider"></select></div>
          <div class="ai-field ai-field-baseurl"><label data-t="settingsBaseUrl">API Base URL</label><input data-field="baseUrl" type="text" /></div>
          <div class="ai-field"><label data-t="settingsApiKey">API Key</label><input data-field="apiKey" type="password" /></div>
          <div class="ai-field">
            <label data-t="settingsModel">Model name</label>
            <select class="ai-settings-model-select" hidden></select>
            <input class="ai-settings-model-input" data-field="model" type="text" />
          </div>
          <div class="ai-field ai-field-checkbox">
            <label><input type="checkbox" data-field="skipTlsVerify" /> <span data-t="settingsSkipTls">Skip TLS certificate verification (insecure - testing only)</span></label>
          </div>
          <div class="ai-field">
            <label data-t="settingsLanguage">Language</label>
            <div class="ai-lang-toggle">
              <button data-lang="en" class="active">English</button>
              <button data-lang="he">עברית</button>
            </div>
          </div>
          <div class="ai-settings-test-result" id="settingsTestResult"></div>
          <div class="ai-settings-actions">
            <button class="ai-btn-secondary" data-t="testConnection" type="button">Test connection</button>
            <button class="ai-btn-primary" data-t="save">Save</button>
          </div>
          <button class="ai-more-settings-btn" id="moreSettingsBtn" type="button" data-t="moreSettings">More settings</button>
        </div>
      </div>
      <div class="ai-chat"></div>
      <div class="ai-settings-view" id="settingsView">
        <div class="ai-settings-section">
          <h4 data-t="sectionScope">Edit scope</h4>
          <p class="ai-settings-section-note" data-t="scopeNote">Applies immediately and resets tool registration below.</p>
          <div class="ai-scope-control" id="scopeControl"></div>
        </div>
        <div class="ai-settings-section" id="connectionSlot"></div>
        <div class="ai-settings-section">
          <h4><span data-t="sectionTools">Tools</span> <span class="ai-tools-count" id="toolsCount"></span></h4>
          <div class="ai-tools-list" id="toolsList"></div>
        </div>
        <div class="ai-settings-section">
          <h4 data-t="sectionDocMessage">Document guidelines</h4>
          <textarea class="ai-doc-message" id="docMessageInput" data-t-placeholder="docMessagePlaceholder" maxlength="8192"></textarea>
          <p class="ai-settings-section-note" data-t="docMessageSaveNote">Applies at the start of the next New chat, not the current conversation.</p>
        </div>
        <div class="ai-settings-view-actions">
          <div class="ai-settings-saved-note" id="settingsSavedNote" data-t="savedNote">Saved</div>
          <button class="ai-btn-primary" id="settingsViewSave" data-t="save">Save</button>
        </div>
      </div>
      <div class="ai-composer">
        <div class="ai-input-box">
          <span class="ai-scope-hint"><span class="dot"></span><span class="label" id="scopeHintLabel">Whole document</span></span>
          <textarea class="ai-textarea" rows="1" dir="auto" placeholder="Ask Airchat Office to edit this document..." data-t-placeholder="inputPlaceholder"></textarea>
          <div class="ai-input-footer">
            <div style="position: relative;">
              <button class="ai-mode-btn"><span class="dot"></span><span id="modeBtnLabel">Full autonomy</span></button>
              <div class="ai-mode-menu" id="modeMenu">${modeMenuItemsHtml(menuModes, defaultMode)}</div>
            </div>
            <div class="ai-send-group">
              <button class="ai-stop-btn" data-t-title="stop" hidden>&#9632;</button>
              <button class="ai-send-btn" data-t-title="send">&#10148;</button>
            </div>
          </div>
        </div>
      </div>
      </div>
    </div>
  `

  const chatEl = root.querySelector<HTMLDivElement>('.ai-chat')!
  const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
  const sendBtn = root.querySelector<HTMLButtonElement>('.ai-send-btn')!
  const stopBtn = root.querySelector<HTMLButtonElement>('.ai-stop-btn')!
  const newChatBtn = root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!
  const settingsBtn = root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!
  const settingsPanel = root.querySelector<HTMLDivElement>('#settingsPanel')!
  const modeBtn = root.querySelector<HTMLButtonElement>('.ai-mode-btn')!
  const modeMenu = root.querySelector<HTMLDivElement>('#modeMenu')!
  const modeBtnLabel = root.querySelector<HTMLSpanElement>('#modeBtnLabel')!
  const scopeHintLabel = root.querySelector<HTMLSpanElement>('#scopeHintLabel')!

  const dockEl = root.querySelector<HTMLDivElement>('.ai-dock')!
  const railEl = root.querySelector<HTMLDivElement>('.ai-rail')!
  const collapseBtn = root.querySelector<HTMLButtonElement>('[data-t-title="collapse"]')!

  // PP-6: provider + model controls. Two model elements share one slot in the
  // layout - only one carries data-field="model" at a time (toggled in
  // applyProviderUI), so a plain `[data-field="model"]` read/write (used by
  // Save and by the initial-settings prefill) always finds exactly the
  // active one, regardless of whether the current provider needs a select
  // (has models) or a free-text input (Custom).
  const providerSelect = root.querySelector<HTMLSelectElement>('.ai-settings-provider')!
  const baseUrlField = root.querySelector<HTMLDivElement>('.ai-field-baseurl')!
  const modelSelect = root.querySelector<HTMLSelectElement>('.ai-settings-model-select')!
  const modelInput = root.querySelector<HTMLInputElement>('.ai-settings-model-input')!
  const apiKeyInput = root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!

  providerSelect.innerHTML = AI_PROVIDERS.map((p) => `<option value="${p.id}">${escapeHtml(p.label)}</option>`).join('')

  // Remembered per-provider values, so switching providers and back restores
  // that provider's own key/model rather than showing blank fields - seeded
  // from initialSettings.providers, falling back to each provider's default
  // model and an empty key for one never configured before.
  const rememberedProviders = new Map<string, { apiKey: string; model: string; baseUrl?: string }>()
  for (const p of AI_PROVIDERS) {
    const saved = options.initialSettings?.providers?.[p.id]
    rememberedProviders.set(p.id, saved ?? { apiKey: '', model: p.defaultModel, baseUrl: p.needsBaseUrl ? '' : undefined })
  }

  function applyProviderUI(providerId: string): void {
    const meta = AI_PROVIDERS.find((p) => p.id === providerId) ?? AI_PROVIDERS[0]!
    const slot = rememberedProviders.get(providerId) ?? { apiKey: '', model: meta.defaultModel, baseUrl: '' }

    providerSelect.value = meta.id
    baseUrlField.hidden = !meta.needsBaseUrl
    // Always set (not just when visible) so a hidden field never leaks a
    // stale value from a previously-selected provider into the Save payload.
    root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value = meta.needsBaseUrl ? (slot.baseUrl ?? '') : ''

    apiKeyInput.placeholder = meta.keyPlaceholder
    apiKeyInput.value = slot.apiKey

    if (meta.models.length > 0) {
      modelSelect.innerHTML = meta.models.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('')
      modelSelect.value = slot.model || meta.defaultModel
      modelSelect.hidden = false
      modelSelect.setAttribute('data-field', 'model')
      modelInput.hidden = true
      modelInput.removeAttribute('data-field')
    } else {
      modelInput.value = slot.model
      modelInput.hidden = false
      modelInput.setAttribute('data-field', 'model')
      modelSelect.hidden = true
      modelSelect.removeAttribute('data-field')
    }
  }

  providerSelect.addEventListener('change', () => applyProviderUI(providerSelect.value))

  // 'custom' matches this repo's actual default provider (settings.ts), and
  // is also the only sane fallback for the pre-PP-6 flat-settings shape a
  // caller might still pass with no `provider` field at all - that old shape
  // was always a single OpenAI-compatible endpoint, i.e. today's Custom slot.
  const DEFAULT_PROVIDER = 'custom'

  if (options.initialSettings) {
    const init = options.initialSettings
    const providerId = init.provider ?? DEFAULT_PROVIDER
    if (init.apiKey || init.model || init.baseUrl) {
      rememberedProviders.set(providerId, { apiKey: init.apiKey ?? '', model: init.model ?? '', baseUrl: init.baseUrl })
    }
    applyProviderUI(providerId)
    if (init.skipTlsVerify) root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked = true
  } else {
    applyProviderUI(DEFAULT_PROVIDER)
  }

  function setCollapsed(collapsed: boolean): void {
    dockEl.classList.toggle('collapsed', collapsed)
    options.onCollapseChange(collapsed)
  }

  collapseBtn.addEventListener('click', () => setCollapsed(true))
  railEl.addEventListener('click', () => setCollapsed(false))

  let currentLang: Lang = 'en'
  let lastSelection: { hasSelection: boolean; preview?: string; extent?: SelectionExtent } | null = null
  const WHOLE_SCOPE_KEYS: Record<string, string> = { doc: 'scopeWholeDoc', sheet: 'scopeWholeSheet', deck: 'scopeWholeDeck', mailbox: 'scopeWholeMailbox' }
  const wholeScopeKey = WHOLE_SCOPE_KEYS[options.scopeUnit ?? 'doc']!

  chatEl.innerHTML = emptyStateHtml(options, currentLang)

  function t(key: string): string {
    return STRINGS[key]?.[currentLang] ?? key
  }

  function applyStrings(): void {
    root.querySelectorAll<HTMLElement>('[data-t]').forEach((el) => {
      el.textContent = t(el.dataset.t!)
    })
    root.querySelectorAll<HTMLElement>('[data-t-title]').forEach((el) => {
      const s = t(el.dataset.tTitle!)
      el.title = s
      el.setAttribute('aria-label', s)
    })
    root.querySelectorAll<HTMLTextAreaElement>('[data-t-placeholder]').forEach((el) => {
      el.placeholder = t(el.dataset.tPlaceholder!)
    })
    updateTextareaDir()
  }

  // dir="auto" only inspects the textarea's own value, never its placeholder
  // (spec behavior) - so an empty Hebrew-UI box still showed the placeholder
  // LTR. Pin dir to the current language while empty, and hand back to
  // "auto" (real bidi detection) as soon as there's typed content to judge.
  function updateTextareaDir(): void {
    textarea.dir = textarea.value ? 'auto' : currentLang === 'he' ? 'rtl' : 'ltr'
  }

  // FT-2 Task 5: renders an Excel/PowerPoint extent - the localized words
  // live here (not in the shell), so a language switch relocalizes the pill
  // for free via refreshScopeHint() below, same as everything else.
  function buildExtentLabel(extent: SelectionExtent): string {
    switch (extent.kind) {
      case 'cell':
        return extent.address
      case 'range':
        return `${extent.address} · ${extent.rows}×${extent.cols}`
      case 'wholeColumn':
        return extent.dataRows === null
          ? `${t('scopeColumn')} ${extent.col} · ${t('scopeEmpty')}`
          : `${t('scopeColumn')} ${extent.col} · ${extent.dataRows} ${t('scopeRowsWithData')}`
      case 'wholeColumns':
        return extent.dataRows === null
          ? `${t('scopeColumns')} ${extent.firstCol}–${extent.lastCol} · ${t('scopeEmpty')}`
          : `${t('scopeColumns')} ${extent.firstCol}–${extent.lastCol} · ${extent.dataRows}×${extent.cols} ${t('scopeWithData')}`
      case 'wholeRow':
        return `${t('scopeRow')} ${extent.row}`
      case 'multiArea':
        return `${extent.areaCount} ${t('scopeAreas')} · ${extent.cellCount} ${t('scopeCells')}`
      case 'slides':
        return `${extent.count} ${t(extent.count === 1 ? 'scopeSlide' : 'scopeSlides')}`
      case 'shapes':
        return extent.count === 1 && extent.primaryName ? extent.primaryName : `${extent.count} ${t('scopeShapes')}`
      case 'shapeText':
        return `${t('scopeTextIn')} ${extent.shapeName ?? t('scopeShape')}`
      case 'wordObject': {
        const noun = extent.objectKind === 'table' ? t('scopeTable') : extent.objectKind === 'chart' ? t('scopeChart') : t('scopeSmartArt')
        return `${noun} ${extent.objectIndex}`
      }
      case 'mailSelection':
        return extent.count === 1 ? extent.subject || t('scopeEmail') : `${extent.count} ${t('scopeEmails')}`
      default:
        return ''
    }
  }

  function refreshScopeHint(): void {
    if (lastSelection && lastSelection.hasSelection) {
      if (lastSelection.extent) {
        scopeHintLabel.textContent = buildExtentLabel(lastSelection.extent)
      } else {
        scopeHintLabel.textContent = t('scopeSelectionPrefix') + (lastSelection.preview ?? '') + '..."'
      }
    } else {
      scopeHintLabel.textContent = t(wholeScopeKey)
    }
  }

  function refreshModeLabel(): void {
    const selected = root.querySelector<HTMLElement>('.ai-mode-menu-item.selected')
    if (selected) modeBtnLabel.textContent = selected.querySelector('span')!.textContent
  }

  function setLang(l: Lang): void {
    dockEl.setAttribute('lang', l)
    dockEl.setAttribute('dir', l === 'he' ? 'rtl' : 'ltr')
    currentLang = l
    root.querySelectorAll<HTMLButtonElement>('.ai-lang-toggle button').forEach((b) => {
      b.classList.toggle('active', b.dataset.lang === l)
    })
    applyStrings()
    refreshScopeHint()
    refreshModeLabel()
    // FT-1 Task 6 Step 6: tool labels/descriptions come from config, not
    // STRINGS, so applyStrings() alone cannot relocalize them - re-render.
    renderToolsList()
    const existingEmpty = chatEl.querySelector('.ai-chat-empty')
    if (existingEmpty) {
      existingEmpty.remove()
      chatEl.insertAdjacentHTML('beforeend', emptyStateHtml(options, currentLang))
    }
  }

  applyStrings()

  let assistantBubble: HTMLDivElement | null = null
  let pendingLang: Lang = 'en'

  // "Thinking..." indicator: a shimmer label + pulsing dots. Shown whenever the
  // run is busy and no assistant text is currently streaming - so it stays up
  // through the send -> first-output gap, alongside a running tool (its
  // hourglass pulses too), and in every model-reasoning gap between tools. It
  // only hides while assistant text is actively streaming (the blinking caret
  // is the cue then).
  let thinkingBusy = false
  let thinkingEl: HTMLDivElement | null = null

  function refreshThinking(): void {
    const show = thinkingBusy && !assistantBubble
    if (show) {
      if (!thinkingEl) {
        thinkingEl = document.createElement('div')
        thinkingEl.className = 'ai-thinking'
        thinkingEl.innerHTML =
          `<span class="ai-typing-label" data-t="thinking">${escapeHtml(t('thinking'))}</span>` +
          `<span class="ai-typing-dots"><span></span><span></span><span></span></span>`
      }
      chatEl.appendChild(thinkingEl) // re-append keeps it as the last child
      scrollToBottom()
    } else if (thinkingEl) {
      thinkingEl.remove()
    }
  }

  function scrollToBottom(): void {
    chatEl.scrollTop = chatEl.scrollHeight
  }

  // Up/Down-arrow recall of previously sent messages, shell-style. Seeded from
  // any prior-conversation user turns replayed via showHistoric, then appended
  // to on every send. `historyNav` is null while the live draft is being
  // edited; the first ArrowUp stashes that draft so an ArrowDown past the
  // newest entry restores exactly what was being typed. Cleared by New chat.
  const sentHistory: string[] = []
  let historyNav: number | null = null
  let historyDraft = ''

  function pushSentHistory(text: string): void {
    if (sentHistory[sentHistory.length - 1] !== text) sentHistory.push(text)
    historyNav = null
    historyDraft = ''
  }

  function caretCollapsedAtFirstLine(): boolean {
    return textarea.selectionStart === textarea.selectionEnd &&
      textarea.value.lastIndexOf('\n', textarea.selectionStart - 1) === -1
  }
  function caretCollapsedAtLastLine(): boolean {
    return textarea.selectionStart === textarea.selectionEnd &&
      textarea.value.indexOf('\n', textarea.selectionEnd) === -1
  }
  function applyRecalledValue(v: string): void {
    textarea.value = v
    updateTextareaDir()
    textarea.setSelectionRange(v.length, v.length)
    textarea.scrollTop = textarea.scrollHeight
  }
  // Returns true when the key was consumed (so the caller preventDefaults).
  function recallPrevHistory(): boolean {
    if (sentHistory.length === 0) return false
    if (historyNav === null) {
      historyDraft = textarea.value
      historyNav = sentHistory.length - 1
    } else if (historyNav > 0) {
      historyNav--
    } else {
      return true // already at the oldest entry - swallow, don't move the caret
    }
    applyRecalledValue(sentHistory[historyNav]!)
    return true
  }
  function recallNextHistory(): boolean {
    if (historyNav === null) return false
    if (historyNav < sentHistory.length - 1) {
      historyNav++
      applyRecalledValue(sentHistory[historyNav]!)
    } else {
      historyNav = null
      applyRecalledValue(historyDraft)
    }
    return true
  }

  function doSend(): void {
    const text = textarea.value.trim()
    if (!text) return
    textarea.value = ''
    updateTextareaDir()
    pushSentHistory(text)
    options.onSend(text)
  }

  textarea.addEventListener('input', updateTextareaDir)

  chatEl.addEventListener('click', (e) => {
    const target = (e.target as HTMLElement).closest('.ai-starter')
    if (target) {
      textarea.value = target.textContent || ''
      updateTextareaDir()
      textarea.focus()
      textarea.setSelectionRange(textarea.value.length, textarea.value.length)
    }
  })

  // Post-hoc addition (2026-08-24, user-requested): a separate stop button
  // next to send (not send doubling as stop, per user feedback on the
  // first version of this) - shown only while busy. AgentLoop.cancel()
  // already existed and was fully wired end-to-end (bootstrap.ts's
  // onStop), just never reachable from any UI control before this.
  stopBtn.addEventListener('click', () => options.onStop?.())
  // Post-hoc change (2026-08-24, user-requested): send is now always
  // clickable, including while busy - the textarea also stays enabled
  // (see setBusy below), so the user can type and queue their next
  // message instead of being locked out until the current run finishes.
  // Queueing itself is bootstrap.ts's job (it owns run/busy state); this
  // layer just always relays "user hit send with this text".
  sendBtn.addEventListener('click', doSend)
  textarea.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      doSend()
      return
    }
    const plainKey = !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey
    if (e.key === 'ArrowUp' && plainKey && caretCollapsedAtFirstLine()) {
      if (recallPrevHistory()) e.preventDefault()
      return
    }
    if (e.key === 'ArrowDown' && plainKey && caretCollapsedAtLastLine()) {
      if (recallNextHistory()) e.preventDefault()
      return
    }
  })
  newChatBtn.addEventListener('click', () => options.onNewChat())

  // Note: the settings button's click handler lives further down (FT-1,
  // "one handler, not two rebound listeners") - it opens/closes this dropdown
  // in chat view and doubles as "back to conversation" in the settings view.
  // A second listener was briefly wired here too; removed - two listeners on
  // the same button each toggling .open cancelled each other out on every
  // click (confirmed repro: the button appeared completely unresponsive).
  root.querySelectorAll<HTMLButtonElement>('.ai-lang-toggle button').forEach((btn) => {
    btn.addEventListener('click', () => {
      pendingLang = btn.dataset.lang as 'en' | 'he'
      root.querySelectorAll('.ai-lang-toggle button').forEach((b) => b.classList.toggle('active', b === btn))
    })
  })
  root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.addEventListener('click', () => {
    setLang(pendingLang)
    options.onSettingsSave({
      provider: providerSelect.value as AiProviderId,
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
      skipTlsVerify: root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked,
      lang: pendingLang,
    })
    settingsPanel.classList.remove('open')
  })

  const testBtn = root.querySelector<HTMLButtonElement>('[data-t="testConnection"]')!
  const testResultEl = root.querySelector<HTMLDivElement>('#settingsTestResult')!
  if (!options.onSettingsTest) {
    testBtn.hidden = true
  } else {
    const onSettingsTest = options.onSettingsTest
    testBtn.addEventListener('click', () => {
      const settings: SettingsSavePayload = {
        provider: providerSelect.value as AiProviderId,
        baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
        apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
        model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
        skipTlsVerify: root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked,
        lang: pendingLang,
      }
      testBtn.disabled = true
      testResultEl.className = 'ai-settings-test-result'
      testResultEl.textContent = t('testTesting')
      onSettingsTest(settings)
        .then((msg) => {
          testResultEl.classList.add('success')
          testResultEl.textContent = msg || t('testSuccess')
        })
        .catch((err: unknown) => {
          testResultEl.classList.add('error')
          testResultEl.textContent = err instanceof Error ? err.message : String(err)
        })
        .finally(() => {
          testBtn.disabled = false
        })
    })
  }

  modeBtn.addEventListener('click', () => modeMenu.classList.toggle('open'))

  // FT-1 Task 3 Step 4: the composer's mode menu AND the settings view's scope
  // control must drive the exact same path, so the two controls cannot
  // diverge - both a menu-item click and a scope-option click call this.
  function selectMode(mode: EditingMode): void {
    root.querySelectorAll<HTMLElement>('.ai-mode-menu-item').forEach((el) => el.classList.toggle('selected', el.dataset.mode === mode))
    root.querySelectorAll<HTMLElement>('.ai-scope-option').forEach((el) => el.classList.toggle('selected', el.dataset.mode === mode))
    modeBtnLabel.textContent = t('mode' + mode[0]!.toUpperCase() + mode.slice(1))
    modeMenu.classList.remove('open')
    modeBtn.classList.toggle('accent', mode === 'trackChanges')
    options.onModeChange(mode)
  }

  root.querySelectorAll<HTMLElement>('.ai-mode-menu-item').forEach((item) => {
    item.addEventListener('click', () => selectMode(item.dataset.mode as EditingMode))
  })

  // ---- FT-1: scope control (settings view) - generated from MODES/STRINGS,
  // not re-listed, per Task 3 Step 1. ----
  const scopeControl = root.querySelector<HTMLDivElement>('#scopeControl')!
  scopeControl.innerHTML = menuModes.map((mode) => {
    const key = 'mode' + mode[0]!.toUpperCase() + mode.slice(1)
    return `<div class="ai-scope-option${mode === defaultMode ? ' selected' : ''}" data-mode="${mode}">
      <span data-t="${key}"></span><span class="desc" data-t="${key}Desc"></span>
    </div>`
  }).join('')
  root.querySelectorAll<HTMLElement>('.ai-scope-option').forEach((el) => {
    el.addEventListener('click', () => selectMode(el.dataset.mode as EditingMode))
  })

  // ---- FT-1 Task 4: tool registry model. Authoritative "which tools exist
  // with what labels" comes from options.tools (static, set at mount); "which
  // are in-scope right now" and "which are registered" are pushed in from the
  // host app (bootstrap.ts owns activeTools()/the editing-mode mapping) via
  // setToolScope - this component only renders whatever it was last told and
  // reports toggles back, it never computes scope-to-tool-name mappings itself. ----
  const toolsListEl = root.querySelector<HTMLDivElement>('#toolsList')!
  const toolsCountEl = root.querySelector<HTMLSpanElement>('#toolsCount')!
  const allTools = options.tools ?? []
  let availableToolNames: string[] = []
  let registeredToolNames: string[] = []

  function toolLabel(entry: { name: string; label: Record<Lang, string> }): string {
    return entry.label[currentLang] ?? entry.name
  }
  function toolDescription(entry: { description: Record<Lang, string> }): string {
    return entry.description[currentLang] ?? ''
  }

  function renderToolsList(): void {
    const availableSet = new Set(availableToolNames)
    const registeredSet = new Set(registeredToolNames)
    const available = allTools.filter((tl) => availableSet.has(tl.name))
    const outOfScope = allTools.filter((tl) => !availableSet.has(tl.name))

    function row(tl: ToolDisplayEntry, isAvailable: boolean): string {
      const checked = isAvailable && registeredSet.has(tl.name)
      const hint = isAvailable ? '' : ` title="${escapeHtml(t('toolOutOfScopeHint'))}"`
      return `<div class="ai-tool-row${isAvailable ? '' : ' out-of-scope'}" data-tool="${escapeHtml(tl.name)}"${hint}>
        <label>
          <input type="checkbox" class="ai-tool-toggle" ${checked ? 'checked' : ''} ${isAvailable ? '' : 'disabled'} />
          <div class="ai-tool-info">
            <div class="ai-tool-label">${escapeHtml(toolLabel(tl))}</div>
            <div class="ai-tool-desc">${escapeHtml(toolDescription(tl))}</div>
          </div>
        </label>
      </div>`
    }

    const groups: string[] = []
    if (available.length > 0) {
      groups.push(`<div class="ai-tools-group-label" data-t="toolsAvailable"></div>`)
      groups.push(available.map((tl) => row(tl, true)).join(''))
    }
    if (outOfScope.length > 0) {
      groups.push(`<div class="ai-tools-group-label" data-t="toolsOutOfScope"></div>`)
      groups.push(outOfScope.map((tl) => row(tl, false)).join(''))
    }
    toolsListEl.innerHTML = groups.join('')
    toolsCountEl.textContent = `(${registeredSet.size} / ${available.length})`
    applyStrings()

    toolsListEl.querySelectorAll<HTMLInputElement>('.ai-tool-toggle').forEach((cb) => {
      cb.addEventListener('change', () => {
        const name = cb.closest<HTMLElement>('.ai-tool-row')!.dataset.tool!
        const next = new Set(registeredToolNames)
        if (cb.checked) {
          next.add(name)
        } else {
          // Task 4 Step 4: never let the set go empty - the model would
          // otherwise have no tools at all and answer from thin air.
          if (next.size <= 1) {
            cb.checked = true
            return
          }
          next.delete(name)
        }
        registeredToolNames = Array.from(next)
        toolsCountEl.textContent = `(${next.size} / ${available.length})`
        markDirty()
        options.onToolRegistrationChange?.(registeredToolNames)
      })
    })
  }

  // ---- FT-1: view switching (Task 1) + doc system message (Task 8) + save
  // semantics (Task 9). ----
  const settingsPanelHomeParent = settingsPanel.parentElement!
  const settingsView = root.querySelector<HTMLDivElement>('#settingsView')!
  const connectionSlot = root.querySelector<HTMLDivElement>('#connectionSlot')!
  const moreSettingsBtn = root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!
  const panelInnerSaveBtn = settingsPanel.querySelector<HTMLButtonElement>('.ai-settings-actions .ai-btn-primary')!
  const settingsViewSaveBtn = root.querySelector<HTMLButtonElement>('#settingsViewSave')!
  const settingsSavedNote = root.querySelector<HTMLDivElement>('#settingsSavedNote')!
  const docMessageInput = root.querySelector<HTMLTextAreaElement>('#docMessageInput')!

  let inSettingsView = false
  let dirty = false
  function markDirty(): void { dirty = true }
  settingsView.addEventListener('input', markDirty)
  settingsView.addEventListener('change', markDirty)

  function openSettingsView(): void {
    settingsPanel.classList.remove('open') // in case it was showing as a dropdown
    moreSettingsBtn.hidden = true
    panelInnerSaveBtn.hidden = true // the view has its own consolidated Save below
    settingsPanel.classList.add('inline')
    connectionSlot.appendChild(settingsPanel)
    inSettingsView = true
    dockEl.querySelector('.ai-panel')!.classList.add('settings-open')
    settingsBtn.innerHTML = CHAT_ICON
    const settingsTitleTitle = t('backToChat')
    settingsBtn.title = settingsTitleTitle
    settingsBtn.setAttribute('aria-label', settingsTitleTitle)
    dirty = false
    settingsSavedNote.classList.remove('visible')
  }

  function closeSettingsView(): void {
    settingsPanel.classList.remove('inline')
    settingsPanelHomeParent.appendChild(settingsPanel)
    moreSettingsBtn.hidden = false
    panelInnerSaveBtn.hidden = false
    inSettingsView = false
    dockEl.querySelector('.ai-panel')!.classList.remove('settings-open')
    settingsBtn.innerHTML = GEAR_ICON
    const settingsTitle = t('settings')
    settingsBtn.title = settingsTitle
    settingsBtn.setAttribute('aria-label', settingsTitle)
  }

  moreSettingsBtn.addEventListener('click', () => openSettingsView())

  // The header gear button does double duty (Task 1 Step 4): opens/closes
  // the quick dropdown in chat view, is "back to conversation" in settings
  // view - one handler, not two rebound listeners.
  settingsBtn.addEventListener('click', () => {
    if (inSettingsView) {
      // Task 9 Step 3: unsaved-changes guard, shown inline rather than a
      // blocking native dialog.
      if (dirty && !window.confirm(t('discardChangesConfirm'))) return
      closeSettingsView()
    } else {
      settingsPanel.classList.toggle('open')
    }
  })

  // Post-hoc addition (2026-08-24, user-requested): closes the quick
  // settings dropdown on an outside click - only applies to the dropdown
  // ('open' class); the full inline settings VIEW has its own back/close
  // affordance (settingsBtn above) and unsaved-changes guard, so it is
  // deliberately untouched here.
  document.addEventListener('click', (e) => {
    if (!settingsPanel.classList.contains('open')) return
    const target = e.target as Node
    if (settingsPanel.contains(target) || settingsBtn.contains(target)) return
    settingsPanel.classList.remove('open')
  })

  settingsViewSaveBtn.addEventListener('click', () => {
    setLang(pendingLang)
    options.onSettingsSave({
      provider: providerSelect.value as AiProviderId,
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
      skipTlsVerify: root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked,
      lang: pendingLang,
      docSystemMessage: docMessageInput.value,
      registeredTools: registeredToolNames,
    })
    dirty = false
    settingsSavedNote.classList.add('visible')
    window.setTimeout(() => settingsSavedNote.classList.remove('visible'), 2000)
  })

  function renderMessage(role: 'user' | 'assistant', text: string): HTMLDivElement {
    const existingEmpty = chatEl.querySelector('.ai-chat-empty')
    if (existingEmpty) existingEmpty.remove()
    const div = document.createElement('div')
    div.className = role === 'user' ? 'ai-msg-user' : 'ai-msg-assistant'
    // dir="auto" lets the browser's own bidi algorithm pick this message's
    // paragraph direction from its own first strong-directional character,
    // instead of blindly inheriting the panel's UI-language dir (set by
    // setLang) - otherwise a Hebrew-first message typed while the panel is
    // in English mode (or vice versa) gets the wrong base direction and
    // mixed Hebrew/English word order renders scrambled.
    div.dir = 'auto'
    // Only assistant replies are markdown-rendered - the user's own typed
    // text is shown verbatim, not reinterpreted as markdown syntax.
    if (role === 'assistant') {
      div.innerHTML = renderMarkdown(text)
    } else {
      div.textContent = text
    }
    chatEl.appendChild(div)
    return div
  }

  // PP-2: closes out the currently-open assistant bubble, if any. Called
  // whenever a new bubble is about to be armed (beginAssistantMessage) or a
  // tool group is about to append below it (beginToolGroup) - this is what
  // keeps DOM order matching causal order across a multi-turn run (text,
  // tools, text, tools, text -> each contiguous text segment gets its own
  // bubble, sealed before the next thing appends).
  function sealAssistantBubble(): void {
    if (assistantBubble) {
      assistantBubble.classList.remove('streaming')
      // An empty bubble means the turn produced tool calls but no prose -
      // drop the element entirely rather than leaving an empty box in the transcript.
      if (!assistantBubble.textContent) assistantBubble.remove()
    }
    assistantBubble = null
  }

  return {
    addUserMessage(text) {
      renderMessage('user', text)
      scrollToBottom()
    },
    beginAssistantMessage() {
      // Deliberately does NOT append a bubble: the bubble is created lazily
      // on the first text delta (updateAssistantMessage), so a turn's
      // tool-call group - appended when its first tool starts - lands ABOVE
      // the text that depended on it. Pre-appending here is what made the
      // transcript read backwards (PP-2).
      sealAssistantBubble()
      refreshThinking()
      scrollToBottom()
    },
    updateAssistantMessage(cumulativeText) {
      if (!assistantBubble) {
        assistantBubble = renderMessage('assistant', '')
        assistantBubble.classList.add('streaming')
        refreshThinking() // the streaming bubble's blinking caret takes over
      }
      assistantBubble.innerHTML = renderMarkdown(cumulativeText)
      scrollToBottom()
    },
    endAssistantMessage(finalText) {
      // Covers a non-streaming transport (or a turn where the whole text
      // arrived after tools, with no updateAssistantMessage call) - without
      // this, such a run would show no answer at all.
      if (!assistantBubble && finalText) {
        assistantBubble = renderMessage('assistant', '')
      }
      if (assistantBubble) assistantBubble.innerHTML = renderMarkdown(finalText)
      sealAssistantBubble()
      // No refreshThinking here: endAssistantMessage is end-of-run only, and
      // the caller (onDone/onError) always calls setBusy(false) right after,
      // which clears the indicator.
      scrollToBottom()
    },
    endTurn() {
      // A tool-calling turn finished and the run continues: seal whatever text
      // this turn streamed (a "let me check X" bubble) so the next turn's
      // stream starts a fresh bubble instead of overwriting it, and let the
      // thinking dots reappear for the gap before the next turn.
      sealAssistantBubble()
      refreshThinking()
    },
    beginToolGroup() {
      // Close out any text streamed earlier in this run so the group appends
      // BELOW it and the next turn's text starts a fresh bubble BELOW the
      // group - preserving true chronological order across multi-turn runs.
      sealAssistantBubble()

      const groupEl = document.createElement('div')
      groupEl.className = 'ai-work-group running'
      groupEl.innerHTML = `<div class="ai-work-group-summary"><span class="caret">&#9656;</span><span class="label">Running tools...</span></div><div class="ai-work-group-body"><div class="steps"></div></div>`
      chatEl.appendChild(groupEl)
      groupEl.querySelector('.ai-work-group-summary')!.addEventListener('click', () => groupEl.classList.toggle('open'))
      const summaryEl = groupEl.querySelector<HTMLElement>('.label')!
      const stepsEl = groupEl.querySelector<HTMLElement>('.steps')!
      let count = 0
      refreshThinking() // keep the dots below the just-appended group
      scrollToBottom()
      return {
        addStep(toolName, input) {
          count++
          summaryEl.textContent = `Running ${count} tool${count > 1 ? 's' : ''}...`
          const rowEl = document.createElement('div')
          rowEl.className = 'ai-step-row'
          // .pending pulses the hourglass while the C# COM call runs (a
          // several-second window for a big Excel read / an Outlook GAL scan);
          // complete() swaps it for the check/cross and drops .pending.
          rowEl.innerHTML = `<div class="ai-step-icon pending">&#8987;</div><div class="ai-step-title">${escapeHtml(toolName)}(${escapeHtml(truncateForDisplay(JSON.stringify(input), 150))})</div>`
          // PP-3: hidden output region, populated (and unhidden via
          // .has-output/.output-open) in complete() below. <pre> preserves
          // the line structure of multi-line outputs without a formatter;
          // dir="auto" matches the per-message bidi rule used for chat bubbles.
          const outputEl = document.createElement('pre')
          outputEl.className = 'ai-step-output'
          outputEl.dir = 'auto'
          rowEl.appendChild(outputEl)
          stepsEl.appendChild(rowEl)
          scrollToBottom()

          // Per-row disclosure: the step title itself is the toggle, so a
          // group of 5 tools doesn't dump 5 outputs at once. stopPropagation
          // matters - the row lives inside the group summary's
          // click-to-toggle region, and without it this would also collapse
          // the whole group.
          const titleEl = rowEl.querySelector<HTMLElement>('.ai-step-title')!
          titleEl.addEventListener('click', (e) => {
            if (!rowEl.classList.contains('has-output')) return
            e.stopPropagation()
            rowEl.classList.toggle('output-open')
          })

          return {
            complete(result) {
              const iconEl = rowEl.querySelector<HTMLElement>('.ai-step-icon')!
              iconEl.classList.remove('pending')
              iconEl.textContent = result.isError ? '✗' : '✓'
              iconEl.classList.toggle('error', !!result.isError)
              if (result.mutated) {
                const tag = document.createElement('div')
                tag.className = 'ai-applied-tag'
                tag.textContent = '✓ Applied'
                stepsEl.appendChild(tag)
              }

              const text = result.output ?? ''
              if (text.length > 0) {
                rowEl.classList.add('has-output')
                const truncated = text.length > TOOL_OUTPUT_PREVIEW_CHARS
                outputEl.textContent = truncated ? text.slice(0, TOOL_OUTPUT_PREVIEW_CHARS) : text
                if (result.isError) outputEl.classList.add('error')
                // Error rows start expanded - an error the user cannot see
                // is the exact failure mode PP-3 exists to fix.
                if (result.isError) rowEl.classList.add('output-open')
                if (truncated) {
                  const more = document.createElement('button')
                  more.className = 'ai-step-output-more'
                  more.type = 'button'
                  more.textContent = `Show all (${text.length} chars)`
                  more.addEventListener('click', (e) => {
                    e.stopPropagation()
                    outputEl.textContent = text
                    more.remove()
                  })
                  rowEl.appendChild(more)
                }
              }

              scrollToBottom()
            },
          }
        },
        end() {
          summaryEl.textContent = `Ran ${count} tool${count === 1 ? '' : 's'}`
          groupEl.classList.remove('running')
        },
      }
    },
    setBusy(busy) {
      thinkingBusy = busy
      refreshThinking()
      // Post-hoc change (2026-08-24, user-requested): neither the textarea
      // nor the send button are disabled while busy any more - the user can
      // keep typing (and queue) their next message during a run instead of
      // being locked out. The stop button (separate from send) is the only
      // thing that toggles with busy state.
      stopBtn.hidden = !busy
      // Return focus to the textarea once the current run finishes, so the
      // user doesn't have to click back into it - harmless even if they're
      // mid-typing a queued message, since .focus() doesn't touch content
      // or cursor position on an already-focused element.
      if (!busy) textarea.focus()
    },
    showError(message) {
      const div = document.createElement('div')
      div.className = 'ai-msg-error'
      div.textContent = message
      chatEl.appendChild(div)
      scrollToBottom()
    },
    showNotice(kind, onContinue) {
      const div = document.createElement('div')
      div.className = 'ai-msg-notice'
      const textKey = kind === 'truncated' ? 'noticeTruncated' : 'noticeTurnLimit'
      const span = document.createElement('span')
      span.dataset.t = textKey
      span.textContent = t(textKey)
      div.appendChild(span)
      if (onContinue) {
        const btn = document.createElement('button')
        btn.className = 'ai-notice-action'
        btn.type = 'button'
        btn.dataset.t = 'noticeContinue'
        btn.textContent = t('noticeContinue')
        // Removes itself on click so it cannot be double-fired.
        btn.addEventListener('click', () => {
          btn.remove()
          onContinue()
        })
        div.appendChild(btn)
      }
      chatEl.appendChild(div)
      scrollToBottom()
    },
    translate(key) {
      return t(key)
    },
    resetToEmpty() {
      chatEl.innerHTML = emptyStateHtml(options, currentLang)
      // chatEl.innerHTML wiped the node; drop the stale ref and busy flag so a
      // New chat during a run doesn't leave a detached indicator behind.
      thinkingEl = null
      thinkingBusy = false
      sentHistory.length = 0
      historyNav = null
      historyDraft = ''
    },
    showHistoric(messages) {
      for (const m of messages) {
        renderMessage(m.role, m.text)
        if (m.role === 'user') pushSentHistory(m.text)
      }
      const sep = document.createElement('div')
      sep.className = 'ai-history-sep'
      sep.textContent = t('historySep')
      chatEl.appendChild(sep)
      chatEl.insertAdjacentHTML('beforeend', emptyStateHtml(options, currentLang))
      scrollToBottom()
    },
    setSelectionScope(selection) {
      lastSelection = selection
      refreshScopeHint()
    },
    setToolScope(available, registered) {
      availableToolNames = available
      registeredToolNames = registered
      renderToolsList()
    },
    setDocSystemMessage(message) {
      docMessageInput.value = message
    },
    openSettings() {
      openSettingsView()
    },
  }
}
