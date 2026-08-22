import './chat-ui.css'

export type EditingMode = 'readOnly' | 'commentOnly' | 'trackChanges' | 'fullAutonomy'

export interface ChatUIOptions {
  title: string
  onSend: (text: string) => void
  onNewChat: () => void
  onModeChange: (mode: EditingMode) => void
  onSettingsSave: (settings: { baseUrl: string; apiKey: string; model: string; lang: 'en' | 'he' }) => void
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
  setScopeHint(label: string): void
}

const MODES: EditingMode[] = ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']

function escapeHtml(s: string): string {
  const div = document.createElement('div')
  div.textContent = s
  return div.innerHTML
}

function emptyStateHtml(): string {
  return `<div class="ai-chat-empty"><div class="ai-chat-empty-title">What can I help with?</div><div class="ai-starters"></div></div>`
}

export function mountChatUI(root: HTMLElement, options: ChatUIOptions): ChatUIHandle {
  root.innerHTML = `
    <div class="ai-panel">
      <div class="ai-panel-header">
        <div class="ai-panel-title"><span class="ai-logo">A</span><span>Airchat Office</span></div>
        <div class="ai-header-actions">
          <button class="ai-header-btn" data-t-title="newChat">+</button>
          <button class="ai-header-btn" data-t-title="settings">&#9881;</button>
          <button class="ai-header-btn" data-t-title="collapse">&#x276E;</button>
        </div>
        <div class="ai-settings-panel" id="settingsPanel">
          <h4>Airchat Office Settings</h4>
          <div class="ai-field"><label>API Base URL</label><input data-field="baseUrl" type="text" /></div>
          <div class="ai-field"><label>API Key</label><input data-field="apiKey" type="password" /></div>
          <div class="ai-field"><label>Model name</label><input data-field="model" type="text" /></div>
          <div class="ai-field">
            <label>Language</label>
            <div class="ai-lang-toggle">
              <button data-lang="en" class="active">English</button>
              <button data-lang="he">עברית</button>
            </div>
          </div>
          <div class="ai-settings-actions"><button class="ai-btn-primary">Save</button></div>
        </div>
      </div>
      <div class="ai-chat"></div>
      <div class="ai-composer">
        <div class="ai-input-box">
          <span class="ai-scope-hint"><span class="dot"></span><span class="label" id="scopeHintLabel">Whole document</span></span>
          <textarea class="ai-textarea" rows="1" placeholder="Ask Airchat Office to edit this document..."></textarea>
          <div class="ai-input-footer">
            <div style="position: relative;">
              <button class="ai-mode-btn"><span class="dot"></span><span id="modeBtnLabel">Full autonomy</span></button>
              <div class="ai-mode-menu" id="modeMenu">
                <div class="ai-mode-menu-item" data-mode="readOnly"><span>Read only</span><span class="desc">AI can only read, never edit</span></div>
                <div class="ai-mode-menu-item" data-mode="commentOnly"><span>Comment only</span><span class="desc">Adds comments, no content edits</span></div>
                <div class="ai-mode-menu-item" data-mode="trackChanges"><span>Track changes</span><span class="desc">Edits as reviewable revisions</span></div>
                <div class="ai-mode-menu-item selected" data-mode="fullAutonomy"><span>Full autonomy</span><span class="desc">Edits applied directly</span></div>
              </div>
            </div>
            <button class="ai-send-btn" data-t-title="send">&#10148;</button>
          </div>
        </div>
      </div>
    </div>
  `

  const chatEl = root.querySelector<HTMLDivElement>('.ai-chat')!
  chatEl.innerHTML = emptyStateHtml()
  const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
  const sendBtn = root.querySelector<HTMLButtonElement>('.ai-send-btn')!
  const newChatBtn = root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!
  const settingsBtn = root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!
  const settingsPanel = root.querySelector<HTMLDivElement>('#settingsPanel')!
  const modeBtn = root.querySelector<HTMLButtonElement>('.ai-mode-btn')!
  const modeMenu = root.querySelector<HTMLDivElement>('#modeMenu')!
  const modeBtnLabel = root.querySelector<HTMLSpanElement>('#modeBtnLabel')!
  const scopeHintLabel = root.querySelector<HTMLSpanElement>('#scopeHintLabel')!

  let assistantBubble: HTMLDivElement | null = null
  let pendingLang: 'en' | 'he' = 'en'

  function scrollToBottom(): void {
    chatEl.scrollTop = chatEl.scrollHeight
  }

  function doSend(): void {
    const text = textarea.value.trim()
    if (!text) return
    textarea.value = ''
    options.onSend(text)
  }

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
    options.onSettingsSave({
      baseUrl: root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value,
      apiKey: root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value,
      model: root.querySelector<HTMLInputElement>('[data-field="model"]')!.value,
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
          rowEl.innerHTML = `<div class="ai-step-icon">&#8987;</div><div class="ai-step-title">${escapeHtml(toolName)}(${escapeHtml(JSON.stringify(input))})</div>`
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
      chatEl.innerHTML = emptyStateHtml()
    },
    showHistoric(messages) {
      for (const m of messages) renderMessage(m.role, m.text)
      const sep = document.createElement('div')
      sep.className = 'ai-history-sep'
      sep.textContent = 'Earlier conversation'
      chatEl.appendChild(sep)
      chatEl.insertAdjacentHTML('beforeend', emptyStateHtml())
      scrollToBottom()
    },
    setScopeHint(label) {
      scopeHintLabel.textContent = label
    },
  }
}
