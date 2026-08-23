import { describe, expect, it, vi } from 'vitest'
import { mountChatUI } from './chat-ui'

function setup() {
  const root = document.createElement('div')
  document.body.appendChild(root)
  const onSend = vi.fn()
  const onModeChange = vi.fn()
  const onSettingsSave = vi.fn()
  const onNewChat = vi.fn()
  const handle = mountChatUI(root, {
    onSend, onModeChange, onSettingsSave, onNewChat,
    starters: [
      { en: 'Summarize this document', he: 'סכם את המסמך' },
      { en: 'Fix grammar issues', he: 'תקן שגיאות דקדוק' },
      { en: 'Improve conciseness', he: 'שפר תמציתיות' },
    ],
    onCollapseChange: vi.fn(),
  })
  return { root, onSend, onModeChange, onSettingsSave, onNewChat, handle }
}

describe('mountChatUI', () => {
  it('renders the title and no attachment button', () => {
    const { root } = setup()
    expect(root.textContent).toContain('Airchat Office')
    expect(root.querySelector('.ai-attach-btn')).toBeNull()
    expect(root.querySelector('input[type="file"]')).toBeNull()
  })

  it('sending calls onSend and clears the textarea', () => {
    const { root, onSend } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    textarea.value = 'do the thing'
    root.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
    expect(onSend).toHaveBeenCalledWith('do the thing')
    expect(textarea.value).toBe('')
  })

  it('clicking a mode menu item calls onModeChange with that mode and marks it selected', () => {
    const { root, onModeChange } = setup()
    root.querySelector<HTMLButtonElement>('.ai-mode-btn')!.click()
    root.querySelector<HTMLElement>('[data-mode="trackChanges"]')!.click()
    expect(onModeChange).toHaveBeenCalledWith('trackChanges')
    expect(root.querySelector('[data-mode="trackChanges"]')!.classList.contains('selected')).toBe(true)
  })

  it('settings only call onSettingsSave when Save is clicked, not on field input', () => {
    const { root, onSettingsSave } = setup()
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    const baseUrlInput = root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!
    baseUrlInput.value = 'http://localhost:9000/v1'
    baseUrlInput.dispatchEvent(new Event('input'))
    expect(onSettingsSave).not.toHaveBeenCalled()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ baseUrl: 'http://localhost:9000/v1' }))
  })

  it('the skipTlsVerify checkbox defaults unchecked and reports its state on Save', () => {
    const { root, onSettingsSave } = setup()
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    const checkbox = root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!
    expect(checkbox.checked).toBe(false)
    checkbox.checked = true
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ skipTlsVerify: true }))
  })

  it('initialSettings pre-fills the settings form so a returning user sees their saved values', () => {
    const root = document.createElement('div')
    document.body.appendChild(root)
    mountChatUI(root, {
      onSend: vi.fn(), onModeChange: vi.fn(), onSettingsSave: vi.fn(), onNewChat: vi.fn(),
      starters: [], onCollapseChange: vi.fn(),
      initialSettings: { baseUrl: 'https://internal-gateway.example/v1', apiKey: 'sk-existing', model: 'gpt-4o', skipTlsVerify: true },
    })
    expect(root.querySelector<HTMLInputElement>('[data-field="baseUrl"]')!.value).toBe('https://internal-gateway.example/v1')
    expect(root.querySelector<HTMLInputElement>('[data-field="apiKey"]')!.value).toBe('sk-existing')
    expect(root.querySelector<HTMLInputElement>('[data-field="model"]')!.value).toBe('gpt-4o')
    expect(root.querySelector<HTMLInputElement>('[data-field="skipTlsVerify"]')!.checked).toBe(true)
  })

  it('a tool group renders a step per addStep call and reflects completion, collapsed by default', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    expect(root.querySelector('.ai-work-group')!.classList.contains('open')).toBe(false)
    const step = group.addStep('insert_content', { text: 'hi' })
    step.complete({ output: 'Inserted text: hi', mutated: true })
    expect(root.querySelector('.ai-applied-tag')).not.toBeNull()
  })

  it('showHistoric renders messages above a divider with full opacity (no fade class)', () => {
    const { root, handle } = setup()
    handle.showHistoric([{ role: 'user', text: 'earlier question' }, { role: 'assistant', text: 'earlier answer' }])
    expect(root.querySelector('.ai-history-sep')).not.toBeNull()
    expect(root.textContent).toContain('earlier question')
    expect(root.querySelector('.ai-history-faded')).toBeNull()
  })

  it('setSelectionScope updates the hint label text for a live selection, and reverts to Whole document', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, preview: 'Q3 revenue grew' })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selection: "Q3 revenue grew..."')
    handle.setSelectionScope(null)
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Whole document')
  })

  it('saving settings with a language change updates panel strings via Save, not live', () => {
    const { root, onSettingsSave } = setup()
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    // Not yet applied - Hebrew string should not appear until Save.
    expect(root.querySelector('[data-t="panelTitle"]')!.textContent).toBe('Airchat Office')
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('[data-t="panelTitle"]')!.textContent).toBe("איירצ'אט אופיס")
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ lang: 'he' }))
  })

  it('resetToEmpty calls onNewChat is NOT implied - resetToEmpty just clears the DOM to the empty state', () => {
    const { root, handle } = setup()
    handle.addUserMessage('x')
    handle.resetToEmpty()
    expect(root.querySelector('.ai-msg-user')).toBeNull()
    expect(root.querySelector('.ai-chat-empty')).not.toBeNull()
  })

  it('shows starter pills in the empty state, and clicking one fills the textarea', () => {
    const { root } = setup()
    const pills = root.querySelectorAll<HTMLElement>('.ai-starter')
    expect(pills.length).toBe(3)
    expect(pills[0].textContent).toBe('Summarize this document')
    pills[0].click()
    expect(root.querySelector<HTMLTextAreaElement>('.ai-textarea')!.value).toBe('Summarize this document')
  })

  it('clicking collapse hides the panel and shows the rail; clicking the rail re-expands', () => {
    const onCollapseChange = vi.fn()
    const root = document.createElement('div')
    document.body.appendChild(root)
    mountChatUI(root, {
      onSend: vi.fn(), onModeChange: vi.fn(), onSettingsSave: vi.fn(), onNewChat: vi.fn(),
      starters: [], onCollapseChange,
    })
    root.querySelector<HTMLButtonElement>('[data-t-title="collapse"]')!.click()
    expect(root.querySelector('.ai-dock')!.classList.contains('collapsed')).toBe(true)
    expect(onCollapseChange).toHaveBeenCalledWith(true)
    root.querySelector<HTMLElement>('.ai-rail')!.click()
    expect(root.querySelector('.ai-dock')!.classList.contains('collapsed')).toBe(false)
    expect(onCollapseChange).toHaveBeenCalledWith(false)
  })

  it('RTL round trip: switching to Hebrew, sending a Hebrew message, and receiving a Hebrew reply all render inside the RTL-marked container', () => {
    const { root, handle } = setup()

    // Switch to Hebrew via Settings -> Save (language only takes effect on Save).
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('.ai-dock')!.getAttribute('dir')).toBe('rtl')

    // Send a Hebrew user message.
    const hebrewQuestion = 'סכם את המסמך הזה בבקשה'
    handle.addUserMessage(hebrewQuestion)

    // Simulate a streamed Hebrew assistant reply.
    handle.beginAssistantMessage()
    const hebrewAnswer = 'המסמך עוסק בתקציב הרבעוני ובתחזית המכירות.'
    handle.updateAssistantMessage(hebrewAnswer.slice(0, 10))
    handle.endAssistantMessage(hebrewAnswer)

    // Both messages must render intact (no mangling) and inside the
    // RTL-marked ancestor, so the CSS's [dir='rtl'] .ai-msg-user /
    // [dir='rtl'] .ai-msg-assistant alignment rules actually apply to them.
    const userMsg = root.querySelector<HTMLElement>('[dir="rtl"] .ai-msg-user')
    const assistantMsg = root.querySelector<HTMLElement>('[dir="rtl"] .ai-msg-assistant')
    expect(userMsg).not.toBeNull()
    expect(assistantMsg).not.toBeNull()
    expect(userMsg!.textContent).toBe(hebrewQuestion)
    expect(assistantMsg!.textContent).toBe(hebrewAnswer)

    // Switching back to English moves the container back to LTR and the
    // already-rendered Hebrew messages remain readable (dir is a container
    // attribute, not per-message - Unicode bidi still renders Hebrew
    // characters correctly regardless of the container's base direction).
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="en"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    expect(root.querySelector('.ai-dock')!.getAttribute('dir')).toBe('ltr')
    expect(root.querySelector('.ai-msg-user')!.textContent).toBe(hebrewQuestion)
  })

  it('mixed-script messages get their own bidi direction via dir="auto", independent of the panel\'s UI language', () => {
    const { root, handle } = setup()

    // Panel stays in English (default) - but the user types a message that
    // starts in Hebrew and switches to English mid-sentence. Without
    // dir="auto" on the message element, it would wrongly inherit the
    // panel's ltr dir and the mixed-script word order would render scrambled.
    const mixedHebrewFirst = 'שלום, please summarize this document'
    handle.addUserMessage(mixedHebrewFirst)
    const userMsg = root.querySelector<HTMLElement>('.ai-msg-user')!
    expect(userMsg.getAttribute('dir')).toBe('auto')
    expect(userMsg.textContent).toBe(mixedHebrewFirst)

    // Same for a streamed assistant reply that starts in English but
    // contains Hebrew.
    handle.beginAssistantMessage()
    const mixedEnglishFirst = 'Sure, the summary is: זהו סיכום המסמך.'
    handle.updateAssistantMessage(mixedEnglishFirst.slice(0, 10))
    handle.endAssistantMessage(mixedEnglishFirst)
    const assistantMsg = root.querySelector<HTMLElement>('.ai-msg-assistant')!
    expect(assistantMsg.getAttribute('dir')).toBe('auto')
    expect(assistantMsg.textContent).toBe(mixedEnglishFirst)
  })
})
