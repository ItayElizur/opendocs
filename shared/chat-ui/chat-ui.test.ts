import { describe, expect, it, vi } from 'vitest'
import { mountChatUI } from './chat-ui'

function setup() {
  const root = document.createElement('div')
  document.body.appendChild(root)
  const onSend = vi.fn()
  const onModeChange = vi.fn()
  const onSettingsSave = vi.fn()
  const onNewChat = vi.fn()
  const handle = mountChatUI(root, { title: 'Airchat Office', onSend, onModeChange, onSettingsSave, onNewChat })
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

  it('setScopeHint updates the hint label text', () => {
    const { root, handle } = setup()
    handle.setScopeHint('Selection: "Q3 revenue grew..."')
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selection: "Q3 revenue grew..."')
  })

  it('resetToEmpty calls onNewChat is NOT implied - resetToEmpty just clears the DOM to the empty state', () => {
    const { root, handle } = setup()
    handle.addUserMessage('x')
    handle.resetToEmpty()
    expect(root.querySelector('.ai-msg-user')).toBeNull()
    expect(root.querySelector('.ai-chat-empty')).not.toBeNull()
  })
})
