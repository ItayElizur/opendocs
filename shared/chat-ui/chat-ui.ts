import './chat-ui.css'

export type EditingMode = 'readOnly' | 'commentOnly' | 'trackChanges' | 'fullAutonomy'

export type Lang = 'en' | 'he'

const STRINGS: Record<string, Record<Lang, string>> = {
  panelTitle:           { en: 'Airchat Office', he: "איירצ'אט אופיס" },
  inputPlaceholder:     { en: 'Ask Airchat Office to edit this document...', he: 'בקש מ-Airchat Office לערוך את המסמך...' },
  send:                 { en: 'Send', he: 'שלח' },
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
}

export interface ChatUIOptions {
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; skipTlsVerify: boolean; lang: Lang }) => void
  starters: Array<{ en: string; he: string }>
  onCollapseChange: (collapsed: boolean) => void
  // Pre-fills the settings form on mount (e.g. from the host app's own
  // persisted storage) so the user doesn't see blank fields every time they
  // reopen the document, even though a value was saved previously.
  initialSettings?: { baseUrl?: string; apiKey?: string; model?: string; skipTlsVerify?: boolean }
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
  beginToolGroup(): ToolGroupHandle
  setBusy(busy: boolean): void
  showError(message: string): void
  resetToEmpty(): void
  showHistoric(messages: Array<{ role: 'user' | 'assistant'; text: string }>): void
  setSelectionScope(selection: { hasSelection: boolean; preview: string } | null): void
}

const MODES: EditingMode[] = ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']

function escapeHtml(s: string): string {
  const div = document.createElement('div')
  div.textContent = s
  return div.innerHTML
}

function truncateForDisplay(s: string, max: number): string {
  return s.length > max ? s.slice(0, max) + '…' : s
}

function emptyStateHtml(options: ChatUIOptions, currentLang: Lang): string {
  const pills = options.starters
    .map((s) => `<div class="ai-starter">${escapeHtml(s[currentLang])}</div>`)
    .join('')
  const title = escapeHtml(STRINGS.emptyTitle[currentLang])
  return `<div class="ai-chat-empty"><div class="ai-chat-empty-title" data-t="emptyTitle">${title}</div><div class="ai-starters">${pills}</div></div>`
}

export function mountChatUI(root: HTMLElement, options: ChatUIOptions): ChatUIHandle {
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
          <div class="ai-field"><label data-t="settingsBaseUrl">API Base URL</label><input data-field="baseUrl" type="text" /></div>
          <div class="ai-field"><label data-t="settingsApiKey">API Key</label><input data-field="apiKey" type="password" /></div>
          <div class="ai-field"><label data-t="settingsModel">Model name</label><input data-field="model" type="text" /></div>
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
          <div class="ai-settings-actions"><button class="ai-btn-primary" data-t="save">Save</button></div>
        </div>
      </div>
      <div class="ai-chat"></div>
      <div class="ai-composer">
        <div class="ai-input-box">
          <span class="ai-scope-hint"><span class="dot"></span><span class="label" id="scopeHintLabel">Whole document</span></span>
          <textarea class="ai-textarea" rows="1" placeholder="Ask Airchat Office to edit this document..." data-t-placeholder="inputPlaceholder"></textarea>
          <div class="ai-input-footer">
            <div style="position: relative;">
              <button class="ai-mode-btn"><span class="dot"></span><span id="modeBtnLabel">Full autonomy</span></button>
              <div class="ai-mode-menu" id="modeMenu">
                <div class="ai-mode-menu-item" data-mode="readOnly"><span data-t="modeReadOnly">Read only</span><span class="desc" data-t="modeReadOnlyDesc">AI can only read, never edit</span></div>
                <div class="ai-mode-menu-item" data-mode="commentOnly"><span data-t="modeCommentOnly">Comment only</span><span class="desc" data-t="modeCommentOnlyDesc">Adds comments, no content edits</span></div>
                <div class="ai-mode-menu-item" data-mode="trackChanges"><span data-t="modeTrackChanges">Track changes</span><span class="desc" data-t="modeTrackChangesDesc">Edits as reviewable revisions</span></div>
                <div class="ai-mode-menu-item selected" data-mode="fullAutonomy"><span data-t="modeFullAutonomy">Full autonomy</span><span class="desc" data-t="modeFullAutonomyDesc">Edits applied directly</span></div>
              </div>
            </div>
            <button class="ai-send-btn" data-t-title="send">&#10148;</button>
          </div>
        </div>
      </div>
      </div>
    </div>
  `

  const chatEl = root.querySelector<HTMLDivElement>('.ai-chat')!
  const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
  const sendBtn = root.querySelector<HTMLButtonElement>('.ai-send-btn')!
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

  if (options.initialSettings) {
    const init = options.initialSettings
    if (init.baseUrl) root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value = init.baseUrl
    if (init.apiKey) root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value = init.apiKey
    if (init.model) root.querySelector<HTMLInputElement>('[data-field="model"]')!.value = init.model
    if (init.skipTlsVerify) root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked = true
  }

  function setCollapsed(collapsed: boolean): void {
    dockEl.classList.toggle('collapsed', collapsed)
    options.onCollapseChange(collapsed)
  }

  collapseBtn.addEventListener('click', () => setCollapsed(true))
  railEl.addEventListener('click', () => setCollapsed(false))

  let currentLang: Lang = 'en'
  let lastSelection: { hasSelection: boolean; preview: string } | null = null

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
  }

  function refreshScopeHint(): void {
    if (lastSelection && lastSelection.hasSelection) {
      const prefix = currentLang === 'he' ? 'בחירה: "' : 'Selection: "'
      scopeHintLabel.textContent = prefix + lastSelection.preview + '..."'
    } else {
      scopeHintLabel.textContent = t('scopeWholeDoc')
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
    const existingEmpty = chatEl.querySelector('.ai-chat-empty')
    if (existingEmpty) {
      existingEmpty.remove()
      chatEl.insertAdjacentHTML('beforeend', emptyStateHtml(options, currentLang))
    }
  }

  applyStrings()

  let assistantBubble: HTMLDivElement | null = null
  let pendingLang: Lang = 'en'

  function scrollToBottom(): void {
    chatEl.scrollTop = chatEl.scrollHeight
  }

  function doSend(): void {
    const text = textarea.value.trim()
    if (!text) return
    textarea.value = ''
    options.onSend(text)
  }

  chatEl.addEventListener('click', (e) => {
    const target = (e.target as HTMLElement).closest('.ai-starter')
    if (target) textarea.value = target.textContent || ''
  })

  sendBtn.addEventListener('click', doSend)
  textarea.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      doSend()
    }
  })
  newChatBtn.addEventListener('click', () => options.onNewChat())

  settingsBtn.addEventListener('click', () => settingsPanel.classList.toggle('open'))
  root.querySelectorAll<HTMLButtonElement>('.ai-lang-toggle button').forEach((btn) => {
    btn.addEventListener('click', () => {
      pendingLang = btn.dataset.lang as 'en' | 'he'
      root.querySelectorAll('.ai-lang-toggle button').forEach((b) => b.classList.toggle('active', b === btn))
    })
  })
  root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.addEventListener('click', () => {
    setLang(pendingLang)
    options.onSettingsSave({
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
      skipTlsVerify: root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked,
      lang: pendingLang,
    })
    settingsPanel.classList.remove('open')
  })

  modeBtn.addEventListener('click', () => modeMenu.classList.toggle('open'))
  root.querySelectorAll<HTMLElement>('.ai-mode-menu-item').forEach((item) => {
    item.addEventListener('click', () => {
      const mode = item.dataset.mode as EditingMode
      root.querySelectorAll('.ai-mode-menu-item').forEach((el) => el.classList.toggle('selected', el === item))
      modeBtnLabel.textContent = item.querySelector('span')!.textContent
      modeMenu.classList.remove('open')
      modeBtn.classList.toggle('accent', mode === 'trackChanges')
      options.onModeChange(mode)
    })
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
    div.textContent = text
    chatEl.appendChild(div)
    return div
  }

  return {
    addUserMessage(text) {
      renderMessage('user', text)
      scrollToBottom()
    },
    beginAssistantMessage() {
      assistantBubble = renderMessage('assistant', '')
      assistantBubble.classList.add('streaming')
      scrollToBottom()
    },
    updateAssistantMessage(cumulativeText) {
      if (assistantBubble) assistantBubble.textContent = cumulativeText
      scrollToBottom()
    },
    endAssistantMessage(finalText) {
      if (assistantBubble) {
        assistantBubble.textContent = finalText
        assistantBubble.classList.remove('streaming')
      }
      assistantBubble = null
      scrollToBottom()
    },
    beginToolGroup() {
      const groupEl = document.createElement('div')
      groupEl.className = 'ai-work-group'
      groupEl.innerHTML = `<div class="ai-work-group-summary"><span class="caret">&#9656;</span><span class="label">Running tools...</span></div><div class="ai-work-group-body"><div class="steps"></div></div>`
      chatEl.appendChild(groupEl)
      groupEl.querySelector('.ai-work-group-summary')!.addEventListener('click', () => groupEl.classList.toggle('open'))
      const summaryEl = groupEl.querySelector<HTMLElement>('.label')!
      const stepsEl = groupEl.querySelector<HTMLElement>('.steps')!
      let count = 0
      scrollToBottom()
      return {
        addStep(toolName, input) {
          count++
          summaryEl.textContent = `Running ${count} tool${count > 1 ? 's' : ''}...`
          const rowEl = document.createElement('div')
          rowEl.className = 'ai-step-row'
          rowEl.innerHTML = `<div class="ai-step-icon">&#8987;</div><div class="ai-step-title">${escapeHtml(toolName)}(${escapeHtml(truncateForDisplay(JSON.stringify(input), 150))})</div>`
          stepsEl.appendChild(rowEl)
          scrollToBottom()
          return {
            complete(result) {
              const iconEl = rowEl.querySelector<HTMLElement>('.ai-step-icon')!
              iconEl.textContent = result.isError ? '✗' : '✓'
              iconEl.classList.toggle('error', !!result.isError)
              if (result.mutated) {
                const tag = document.createElement('div')
                tag.className = 'ai-applied-tag'
                tag.textContent = '✓ Applied'
                stepsEl.appendChild(tag)
              }
              scrollToBottom()
            },
          }
        },
        end() {
          summaryEl.textContent = `Ran ${count} tool${count === 1 ? '' : 's'}`
        },
      }
    },
    setBusy(busy) {
      sendBtn.disabled = busy
      textarea.disabled = busy
    },
    showError(message) {
      const div = document.createElement('div')
      div.className = 'ai-msg-error'
      div.textContent = message
      chatEl.appendChild(div)
      scrollToBottom()
    },
    resetToEmpty() {
      chatEl.innerHTML = emptyStateHtml(options, currentLang)
    },
    showHistoric(messages) {
      for (const m of messages) renderMessage(m.role, m.text)
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
  }
}
