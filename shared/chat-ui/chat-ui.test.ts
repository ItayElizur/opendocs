import { describe, expect, it, vi } from 'vitest'
import { mountChatUI, type ToolDisplayEntry } from './chat-ui'

const TEST_TOOLS: ToolDisplayEntry[] = [
  { name: 'get_document_context', label: { en: 'Read document', he: 'קרא מסמך' }, description: { en: 'Reads a summary of the document.', he: 'קורא תקציר של המסמך.' } },
  { name: 'apply_commands', label: { en: 'Edit document', he: 'ערוך מסמך' }, description: { en: 'Applies formatting/editing commands.', he: 'מיישם פקודות עריכה.' } },
]

function setup(extra: Partial<Parameters<typeof mountChatUI>[1]> = {}) {
  const root = document.createElement('div')
  document.body.appendChild(root)
  const onSend = vi.fn()
  const onModeChange = vi.fn()
  const onSettingsSave = vi.fn()
  const onNewChat = vi.fn()
  const onToolRegistrationChange = vi.fn()
  const handle = mountChatUI(root, {
    onSend, onModeChange, onSettingsSave, onNewChat, onToolRegistrationChange,
    starters: [
      { en: 'Summarize this document', he: 'סכם את המסמך' },
      { en: 'Fix grammar issues', he: 'תקן שגיאות דקדוק' },
      { en: 'Improve conciseness', he: 'שפר תמציתיות' },
    ],
    onCollapseChange: vi.fn(),
    tools: TEST_TOOLS,
    ...extra,
  })
  return { root, onSend, onModeChange, onSettingsSave, onNewChat, onToolRegistrationChange, handle }
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

  it('a pending tool step pulses its hourglass, then drops .pending on completion', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('read_range', { address: 'A1:C10' })
    const icon = root.querySelector<HTMLElement>('.ai-step-icon')!
    expect(icon.classList.contains('pending')).toBe(true)
    step.complete({ output: 'rows...' })
    expect(icon.classList.contains('pending')).toBe(false)
    expect(icon.textContent).toBe('✓')
  })

  it('the thinking indicator shows in the send-to-first-output gap and hides once real content appears', () => {
    const { root, handle } = setup()
    const dots = () => root.querySelector('.ai-thinking')
    handle.setBusy(true)
    handle.beginAssistantMessage()
    expect(dots()).not.toBeNull()

    handle.updateAssistantMessage('Here') // first token -> streaming caret takes over
    expect(dots()).toBeNull()

    handle.endAssistantMessage('Here is the answer.')
    handle.setBusy(false)
    expect(dots()).toBeNull()
  })

  it('endTurn seals a midway reply so a later turn cannot overwrite it, and dots come back', () => {
    const { root, handle } = setup()
    const dots = () => root.querySelector('.ai-thinking')
    const bubbles = () => [...root.querySelectorAll('.ai-msg-assistant')].map((b) => b.textContent)
    handle.setBusy(true)
    handle.beginAssistantMessage()

    // turn 1: tool only, no text
    let g = handle.beginToolGroup()
    g.addStep('add_chart', {}).complete({ output: 'ok' })
    handle.endTurn()
    expect(dots()).not.toBeNull() // reasoning gap

    // turn 2: a midway reply, then another tool
    handle.updateAssistantMessage('Found the chart.')
    expect(dots()).toBeNull()
    g = handle.beginToolGroup()
    g.addStep('read_slide', {}).complete({ output: 'ok' })
    handle.endTurn()

    // turn 3: the final reply
    handle.updateAssistantMessage('All done.')
    handle.endAssistantMessage('All done.')
    handle.setBusy(false)

    expect(bubbles()).toContain('Found the chart.') // not overwritten
    expect(bubbles()).toContain('All done.')
    expect(dots()).toBeNull()
  })

  it('the thinking indicator stays through tool calls and reasoning, and only hides while text streams', () => {
    const { root, handle } = setup()
    const dots = () => root.querySelector('.ai-thinking')
    handle.setBusy(true)
    handle.beginAssistantMessage()
    expect(dots()).not.toBeNull() // send gap

    const group = handle.beginToolGroup()
    const s1 = group.addStep('add_chart', {})
    expect(dots()).not.toBeNull() // dots stay alongside the running tool
    s1.complete({ output: 'chart added' })
    expect(dots()).not.toBeNull() // model reasoning between tools

    handle.updateAssistantMessage('Done — chart is in.') // text streams -> caret only
    expect(dots()).toBeNull()

    handle.endAssistantMessage('Done — chart is in.')
    handle.setBusy(false)
    expect(dots()).toBeNull()
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

  it('shows starter pills in the empty state, and clicking one fills the textarea and focuses it', () => {
    const { root } = setup()
    const pills = root.querySelectorAll<HTMLElement>('.ai-starter')
    expect(pills.length).toBe(3)
    expect(pills[0].textContent).toBe('Summarize this document')
    pills[0].click()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    expect(textarea.value).toBe('Summarize this document')
    expect(document.activeElement).toBe(textarea)
    expect(textarea.selectionStart).toBe(textarea.value.length)
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

  it('the composer textarea shows a Hebrew placeholder as RTL when empty, then hands off to auto once text is typed', () => {
    const { root } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    expect(textarea.dir).toBe('ltr')

    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()
    // dir="auto" only ever looks at the value, never the placeholder, so an
    // empty box switched to Hebrew must be pinned to rtl explicitly or the
    // Hebrew placeholder text renders left-to-right.
    expect(textarea.dir).toBe('rtl')

    textarea.value = 'hello'
    textarea.dispatchEvent(new Event('input'))
    expect(textarea.dir).toBe('auto')

    textarea.value = ''
    textarea.dispatchEvent(new Event('input'))
    expect(textarea.dir).toBe('rtl')
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

  // ---- PP-2: chronological ordering ----

  it('renders a tool group above the assistant text that follows it', () => {
    const { root, handle } = setup()
    handle.addUserMessage('do the thing')
    handle.beginAssistantMessage()
    const group = handle.beginToolGroup()
    group.addStep('read_blocks', { startIndex: 0, endIndex: 5 }).complete({ output: 'ok' })
    group.end()
    handle.updateAssistantMessage('Here is the summary')
    handle.endAssistantMessage('Here is the summary')

    const nodes = Array.from(root.querySelectorAll('.ai-work-group, .ai-msg-assistant'))
    expect(nodes.map((n) => n.className.split(' ')[0])).toEqual(['ai-work-group', 'ai-msg-assistant'])
  })

  it('a multi-turn run (text, tools, text, tools, text) preserves that order in the DOM', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    handle.updateAssistantMessage('first')
    handle.endAssistantMessage('first')

    const group1 = handle.beginToolGroup()
    group1.addStep('read_blocks', {}).complete({ output: 'ok' })
    group1.end()

    handle.beginAssistantMessage()
    handle.updateAssistantMessage('second')
    handle.endAssistantMessage('second')

    const group2 = handle.beginToolGroup()
    group2.addStep('apply_commands', {}).complete({ output: 'ok' })
    group2.end()

    handle.beginAssistantMessage()
    handle.updateAssistantMessage('third')
    handle.endAssistantMessage('third')

    const nodes = Array.from(root.querySelectorAll('.ai-work-group, .ai-msg-assistant'))
    expect(nodes.map((n) => n.className.split(' ')[0])).toEqual([
      'ai-msg-assistant', 'ai-work-group', 'ai-msg-assistant', 'ai-work-group', 'ai-msg-assistant',
    ])
  })

  it('a turn with tool calls but no prose leaves no empty assistant bubble', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    const group = handle.beginToolGroup()
    group.addStep('insert_content', {}).complete({ output: 'ok', mutated: true })
    group.end()
    handle.endAssistantMessage('')
    expect(root.querySelector('.ai-msg-assistant')).toBeNull()
  })

  it('a non-streaming reply (no updateAssistantMessage call) still renders the final text', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    handle.endAssistantMessage('final')
    expect(root.querySelector('.ai-msg-assistant')!.textContent).toBe('final')
  })

  // ---- markdown rendering of assistant replies ----

  it('renders headers, bullet lists, and tables from the assistant reply as real markup', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    handle.endAssistantMessage(
      '## Summary\n\n- first point\n- second point\n\n| Name | Count |\n| --- | --- |\n| Alpha | 3 |\n| Beta | 5 |'
    )
    const msg = root.querySelector('.ai-msg-assistant')!

    const h2 = msg.querySelector('h2')!
    expect(h2.textContent).toBe('Summary')

    const items = msg.querySelectorAll('ul > li')
    expect(Array.from(items).map((li) => li.textContent)).toEqual(['first point', 'second point'])

    expect(msg.querySelector('table')).not.toBeNull()
    const headerCells = Array.from(msg.querySelectorAll('thead th')).map((c) => c.textContent)
    expect(headerCells).toEqual(['Name', 'Count'])
    const bodyRows = Array.from(msg.querySelectorAll('tbody tr')).map((tr) =>
      Array.from(tr.querySelectorAll('td')).map((td) => td.textContent)
    )
    expect(bodyRows).toEqual([['Alpha', '3'], ['Beta', '5']])
  })

  it('renders bold, italic, and inline code spans as real markup', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    handle.endAssistantMessage('This is **bold**, this is *italic*, and this is `code`.')
    const msg = root.querySelector('.ai-msg-assistant')!
    expect(msg.querySelector('strong')!.textContent).toBe('bold')
    expect(msg.querySelector('em')!.textContent).toBe('italic')
    expect(msg.querySelector('code')!.textContent).toBe('code')
  })

  it('never executes or injects raw HTML/script tags from the assistant reply - they render as visible escaped text', () => {
    const { root, handle } = setup()
    handle.beginAssistantMessage()
    handle.endAssistantMessage('Ignore this: <script>window.__pwned = true</script> and <img src=x onerror="window.__pwned = true">')
    const msg = root.querySelector('.ai-msg-assistant')!

    // No script/img element was actually created in the DOM...
    expect(msg.querySelector('script')).toBeNull()
    expect(msg.querySelector('img')).toBeNull()
    // ...the tags show up as literal, readable text instead...
    expect(msg.textContent).toContain('<script>window.__pwned = true</script>')
    expect(msg.textContent).toContain('<img src=x onerror="window.__pwned = true">')
    // ...and nothing actually ran.
    expect((window as unknown as { __pwned?: boolean }).__pwned).toBeUndefined()
  })

  it('does not markdown-process the user\'s own message text', () => {
    const { root, handle } = setup()
    handle.addUserMessage('**not bold** <b>not html</b>')
    const msg = root.querySelector('.ai-msg-user')!
    expect(msg.querySelector('strong')).toBeNull()
    expect(msg.querySelector('b')).toBeNull()
    expect(msg.textContent).toBe('**not bold** <b>not html</b>')
  })

  // ---- PP-3: inspectable tool output ----

  it('tool output renders in the step row, hidden until the title is clicked', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('get_document_context', {})
    step.complete({ output: 'Paragraphs: 3, Words: 40' })

    const row = root.querySelector<HTMLElement>('.ai-step-row')!
    const output = row.querySelector<HTMLElement>('.ai-step-output')!
    expect(output.textContent).toBe('Paragraphs: 3, Words: 40')
    expect(row.classList.contains('output-open')).toBe(false)

    row.querySelector<HTMLElement>('.ai-step-title')!.click()
    expect(row.classList.contains('output-open')).toBe(true)
  })

  it('tool output is rendered as text, never parsed as HTML (XSS guard)', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('read_blocks', {})
    const malicious = '<img src=x onerror=alert(1)>'
    step.complete({ output: malicious })

    const output = root.querySelector<HTMLElement>('.ai-step-output')!
    expect(output.querySelector('img')).toBeNull()
    expect(output.textContent).toBe(malicious)
  })

  it('a long tool output is truncated with a show-all toggle that reveals the rest', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('read_range', {})
    const long = 'x'.repeat(2500)
    step.complete({ output: long })

    const output = root.querySelector<HTMLElement>('.ai-step-output')!
    expect(output.textContent!.length).toBe(2000)
    const more = root.querySelector<HTMLButtonElement>('.ai-step-output-more')!
    expect(more).not.toBeNull()
    more.click()
    expect(output.textContent).toBe(long)
    expect(root.querySelector('.ai-step-output-more')).toBeNull()
  })

  it('an error result renders its output expanded by default, with the error class', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('read_blocks', {})
    step.complete({ output: 'boom', isError: true })

    const row = root.querySelector<HTMLElement>('.ai-step-row')!
    expect(row.classList.contains('output-open')).toBe(true)
    expect(row.querySelector('.ai-step-output')!.classList.contains('error')).toBe(true)
  })

  it('an empty tool output stays as quiet as today - no output region, no has-output class', () => {
    const { root, handle } = setup()
    const group = handle.beginToolGroup()
    const step = group.addStep('select_range', {})
    step.complete({ output: '' })

    const row = root.querySelector<HTMLElement>('.ai-step-row')!
    expect(row.classList.contains('has-output')).toBe(false)
    expect(row.querySelector('.ai-step-output')!.textContent).toBe('')
  })

  // ---- PP-4: truncation notice ----

  it('showNotice renders an informational row, and its continue button fires onContinue exactly once', () => {
    const { root, handle } = setup()
    const onContinue = vi.fn()
    handle.showNotice('truncated', onContinue)

    const notice = root.querySelector<HTMLElement>('.ai-msg-notice')!
    expect(notice).not.toBeNull()
    expect(notice.textContent).toContain('The reply was cut off by the length limit.')

    const btn = root.querySelector<HTMLButtonElement>('.ai-notice-action')!
    btn.click()
    expect(onContinue).toHaveBeenCalledTimes(1)
    expect(root.querySelector('.ai-notice-action')).toBeNull()
  })

  it('showNotice with no onContinue renders no action button', () => {
    const { root, handle } = setup()
    handle.showNotice('turnLimit')
    expect(root.querySelector('.ai-msg-notice')).not.toBeNull()
    expect(root.querySelector('.ai-notice-action')).toBeNull()
  })

  // ---- FT-1: full settings view ----

  it('More settings opens the settings view, the gear becomes "back", and back returns to chat', () => {
    const { root } = setup()
    expect(root.querySelector('#settingsView')!.classList.contains('open')).toBe(false)
    const gear = root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!
    expect(gear.title).toBe('Settings')

    const moreSettingsBtn = root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!
    moreSettingsBtn.click()
    expect(root.querySelector('.ai-panel')!.classList.contains('settings-open')).toBe(true)
    expect(gear.title).toBe('Back to conversation')
    // The "More settings" button makes no sense once already inside the
    // settings view it opens - this only checks the `hidden` IDL attribute
    // the code sets, which jsdom does not render; it does NOT catch a CSS
    // rule silently overriding [hidden]'s effect (confirmed repro: a
    // `display: block` rule on this exact button did exactly that - see the
    // fix in chat-ui.css). Real visual verification needs a real browser.
    expect(moreSettingsBtn.hidden).toBe(true)

    gear.click()
    expect(root.querySelector('.ai-panel')!.classList.contains('settings-open')).toBe(false)
    expect(gear.title).toBe('Settings')
    expect(moreSettingsBtn.hidden).toBe(false)
  })

  it('the gear button toggles the quick settings dropdown open and closed in chat view', () => {
    // Regression test: a duplicate click listener on this button (one from
    // before FT-1, one from FT-1 itself) each toggled .open independently,
    // so every click cancelled itself out and the button appeared totally
    // unresponsive - confirmed by real-world testing, not caught by any
    // existing test since the ones above interact with the panel's fields
    // directly without ever asserting the dropdown's own open/closed state.
    const { root } = setup()
    const gear = root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!
    const panel = root.querySelector<HTMLElement>('#settingsPanel')!
    expect(panel.classList.contains('open')).toBe(false)
    gear.click()
    expect(panel.classList.contains('open')).toBe(true)
    gear.click()
    expect(panel.classList.contains('open')).toBe(false)
  })

  it('chat messages stay in the DOM after opening and closing the settings view', () => {
    const { root, handle } = setup()
    handle.addUserMessage('remember this')
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    expect(root.querySelector('.ai-msg-user')!.textContent).toBe('remember this')
    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    expect(root.querySelector('.ai-msg-user')!.textContent).toBe('remember this')
  })

  it('the tools list renders one row per tool with localized labels that relocalize on language switch', () => {
    const { root, handle } = setup()
    handle.setToolScope(['get_document_context', 'apply_commands'], ['get_document_context', 'apply_commands'])
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    const rows = root.querySelectorAll<HTMLElement>('.ai-tool-row')
    expect(rows.length).toBe(2)
    expect(root.querySelector('.ai-tool-row[data-tool="get_document_context"] .ai-tool-label')!.textContent).toBe('Read document')
    expect(root.querySelector('.ai-tool-row[data-tool="apply_commands"] .ai-tool-label')!.textContent).toBe('Edit document')

    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    root.querySelector<HTMLButtonElement>('#settingsViewSave')!.click()

    expect(root.querySelector('.ai-tool-row[data-tool="get_document_context"] .ai-tool-label')!.textContent).toBe('קרא מסמך')
    expect(root.querySelector('.ai-tool-row[data-tool="apply_commands"] .ai-tool-label')!.textContent).toBe('ערוך מסמך')
  })

  it('out-of-scope tools render disabled with a hint, while in-scope tools can be toggled', () => {
    const { root, handle } = setup()
    handle.setToolScope(['get_document_context'], ['get_document_context'])
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()

    const inScopeRow = root.querySelector<HTMLElement>('.ai-tool-row[data-tool="get_document_context"]')!
    const outOfScopeRow = root.querySelector<HTMLElement>('.ai-tool-row[data-tool="apply_commands"]')!

    expect(inScopeRow.classList.contains('out-of-scope')).toBe(false)
    expect(outOfScopeRow.classList.contains('out-of-scope')).toBe(true)
    expect(outOfScopeRow.querySelector<HTMLInputElement>('.ai-tool-toggle')!.disabled).toBe(true)
    expect(outOfScopeRow.getAttribute('title')).toBe('Not available in the current edit scope - change the scope above to enable.')

    const inScopeToggle = inScopeRow.querySelector<HTMLInputElement>('.ai-tool-toggle')!
    expect(inScopeToggle.disabled).toBe(false)
  })

  it('setToolScope re-renders the list and reflects the newly passed registration, clearing any prior toggle state', () => {
    const { root, handle, onToolRegistrationChange } = setup()
    handle.setToolScope(['get_document_context', 'apply_commands'], ['get_document_context', 'apply_commands'])
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    expect(root.querySelector<HTMLInputElement>('.ai-tool-row[data-tool="apply_commands"] .ai-tool-toggle')!.checked).toBe(true)

    handle.setToolScope(['get_document_context', 'apply_commands'], ['get_document_context'])
    expect(root.querySelector<HTMLInputElement>('.ai-tool-row[data-tool="apply_commands"] .ai-tool-toggle')!.checked).toBe(false)
    expect(root.querySelector('#toolsCount')!.textContent).toBe('(1 / 2)')
    expect(onToolRegistrationChange).not.toHaveBeenCalled()
  })

  it('Save fires onSettingsSave once with docSystemMessage and registeredTools; field input alone does not fire it', () => {
    const { root, handle, onSettingsSave } = setup()
    handle.setToolScope(['get_document_context', 'apply_commands'], ['get_document_context', 'apply_commands'])
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    const docMessage = root.querySelector<HTMLTextAreaElement>('#docMessageInput')!
    docMessage.value = 'Always cite sources.'
    docMessage.dispatchEvent(new Event('input'))
    expect(onSettingsSave).not.toHaveBeenCalled()

    root.querySelector<HTMLButtonElement>('#settingsViewSave')!.click()
    expect(onSettingsSave).toHaveBeenCalledTimes(1)
    expect(onSettingsSave).toHaveBeenCalledWith(
      expect.objectContaining({ docSystemMessage: 'Always cite sources.', registeredTools: ['get_document_context', 'apply_commands'] }),
    )
  })

  it('a document guidelines message containing HTML renders as literal text (XSS guard), not markup', () => {
    const { root, handle, onSettingsSave } = setup()
    const malicious = '<img src=x onerror=alert(1)>'
    handle.setDocSystemMessage(malicious)
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    const docMessage = root.querySelector<HTMLTextAreaElement>('#docMessageInput')!
    expect(docMessage.value).toBe(malicious)
    expect(docMessage.querySelector('img')).toBeNull()

    root.querySelector<HTMLButtonElement>('#settingsViewSave')!.click()
    expect(onSettingsSave).toHaveBeenCalledWith(expect.objectContaining({ docSystemMessage: malicious }))
  })

  it('selecting a mode from the settings view updates the composer mode label and fires onModeChange exactly once', () => {
    const { root, onModeChange } = setup()
    root.querySelector<HTMLButtonElement>('#moreSettingsBtn')!.click()
    root.querySelector<HTMLElement>('.ai-scope-control [data-mode="trackChanges"]')!.click()

    expect(onModeChange).toHaveBeenCalledTimes(1)
    expect(onModeChange).toHaveBeenCalledWith('trackChanges')
    expect(root.querySelector('#modeBtnLabel')!.textContent).toBe('Track changes')
    expect(root.querySelector('.ai-scope-control [data-mode="trackChanges"]')!.classList.contains('selected')).toBe(true)
  })

  // ---- FT-2: selection context for Excel/PowerPoint ----

  it('an extent selection renders the address/dimensions with no quotes or ellipsis, unlike Word\'s quoted preview', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'range', address: 'B2:D40', rows: 39, cols: 3 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('B2:D40 · 39×3')

    handle.setSelectionScope({ hasSelection: true, preview: 'Q3 revenue grew' })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selection: "Q3 revenue grew..."')
  })

  it('a single-cell extent shows the bare address with no dimensions', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'cell', address: 'C7' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('C7')
  })

  it('whole-column/row extents report the data-bounded extent, or "empty" when there is none', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'wholeColumn', col: 'B', dataRows: 200 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('column B · 200 rows with data')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'wholeColumn', col: 'B', dataRows: null } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('column B · empty')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'wholeColumns', firstCol: 'B', lastCol: 'D', cols: 3, dataRows: 200 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('columns B–D · 200×3 with data')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'wholeRow', row: 5 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('row 5')
  })

  it('a multi-area extent reports the area and cell counts', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'multiArea', areaCount: 2, cellCount: 156 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('2 areas · 156 cells')
  })

  it('PowerPoint extents (slides/shapes/shapeText) render their own short forms', () => {
    const { root, handle } = setup()
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'slides', count: 3 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('3 slides')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'shapes', count: 1, primaryName: 'Revenue chart' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Revenue chart')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'shapes', count: 3, primaryName: null } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('3 shapes')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'shapeText', shapeName: 'Title 1' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('text in Title 1')
  })

  it('null reverts to the per-app whole-scope label - "Whole sheet"/"Whole deck" when scopeUnit is set, "Whole document" by default', () => {
    const { handle: wordHandle, root: wordRoot } = setup()
    wordHandle.setSelectionScope(null)
    expect(wordRoot.querySelector('#scopeHintLabel')!.textContent).toBe('Whole document')

    const { handle: excelHandle, root: excelRoot } = setup({ scopeUnit: 'sheet' })
    excelHandle.setSelectionScope({ hasSelection: true, extent: { kind: 'cell', address: 'C7' } })
    excelHandle.setSelectionScope(null)
    expect(excelRoot.querySelector('#scopeHintLabel')!.textContent).toBe('Whole sheet')

    const { handle: pptHandle, root: pptRoot } = setup({ scopeUnit: 'deck' })
    pptHandle.setSelectionScope(null)
    expect(pptRoot.querySelector('#scopeHintLabel')!.textContent).toBe('Whole deck')

    const { handle: outlookHandle, root: outlookRoot } = setup({ scopeUnit: 'mailbox' })
    outlookHandle.setSelectionScope(null)
    expect(outlookRoot.querySelector('#scopeHintLabel')!.textContent).toBe('Whole mailbox')
  })

  it('Outlook mailSelection extent renders the subject for one email and a count for many', () => {
    const { root, handle } = setup({ scopeUnit: 'mailbox' })
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'mailSelection', count: 1, subject: 'Q3 planning' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Q3 planning')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'mailSelection', count: 1, subject: '' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('Selected email')

    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'mailSelection', count: 4, subject: 'Q3 planning' } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('4 emails')
  })

  it('options.modes narrows the composer mode menu (Outlook: read only + full autonomy only)', () => {
    const { root } = setup({ modes: ['readOnly', 'fullAutonomy'] })
    const items = [...root.querySelectorAll<HTMLElement>('.ai-mode-menu-item')].map((el) => el.dataset.mode)
    expect(items).toEqual(['readOnly', 'fullAutonomy'])
    const scopeOptions = [...root.querySelectorAll<HTMLElement>('.ai-scope-option')].map((el) => el.dataset.mode)
    expect(scopeOptions).toEqual(['readOnly', 'fullAutonomy'])
  })

  // ---- Up/Down-arrow recall of previously sent messages ----

  function arrow(textarea: HTMLTextAreaElement, key: 'ArrowUp' | 'ArrowDown'): void {
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  }
  function sendText(root: HTMLElement, text: string): void {
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    textarea.value = text
    root.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
  }

  it('ArrowUp/ArrowDown cycles through previously sent messages and restores the live draft', () => {
    const { root } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    sendText(root, 'first message')
    sendText(root, 'second message')

    textarea.value = 'half-typed'
    textarea.setSelectionRange(textarea.value.length, textarea.value.length)

    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('second message')
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('first message')
    // Already at the oldest - stays put.
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('first message')

    arrow(textarea, 'ArrowDown')
    expect(textarea.value).toBe('second message')
    // Past the newest - the untouched draft comes back.
    arrow(textarea, 'ArrowDown')
    expect(textarea.value).toBe('half-typed')
  })

  it('ArrowUp only recalls when the caret is collapsed on the first line', () => {
    const { root } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    sendText(root, 'prior')

    textarea.value = 'line one\nline two'
    // Caret on the second line - ArrowUp should move within the textarea, not recall.
    textarea.setSelectionRange(textarea.value.length, textarea.value.length)
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('line one\nline two')

    // Caret at the very start (first line) - now it recalls.
    textarea.setSelectionRange(0, 0)
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('prior')
  })

  it('New chat clears the recall history', () => {
    const { root, handle } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    sendText(root, 'before reset')
    handle.resetToEmpty()

    textarea.value = ''
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('')
  })

  it('showHistoric seeds the recall history with prior-conversation user turns', () => {
    const { root, handle } = setup()
    const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
    handle.showHistoric([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
    arrow(textarea, 'ArrowUp')
    expect(textarea.value).toBe('earlier question')
  })

  it('the pill relocalizes an active extent selection on a language switch (Task 6 Step 3)', () => {
    const { root, handle } = setup({ scopeUnit: 'sheet' })
    handle.setSelectionScope({ hasSelection: true, extent: { kind: 'wholeColumn', col: 'B', dataRows: 200 } })
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('column B · 200 rows with data')

    root.querySelector<HTMLButtonElement>('[data-t-title="settings"]')!.click()
    root.querySelector<HTMLButtonElement>('[data-lang="he"]')!.click()
    root.querySelector<HTMLButtonElement>('.ai-btn-primary')!.click()

    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('עמודה B · 200 שורות עם נתונים')

    handle.setSelectionScope(null)
    expect(root.querySelector('#scopeHintLabel')!.textContent).toBe('כל הגיליון')
  })
})
