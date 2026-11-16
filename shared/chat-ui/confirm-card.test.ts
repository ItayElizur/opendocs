import { beforeEach, describe, expect, it, vi } from 'vitest'

const h = vi.hoisted(() => ({
  toolCalls: [] as Array<{ name: string; input: any; resolve: (r: unknown) => void; reject: (e: unknown) => void }>,
  streams: [] as Array<{ req: any; cb: any }>,
  confirmIrreversible: true,
}))

vi.mock('../web-src/app-shell/bridge', () => ({
  callDotNetTool: (name: string, input: unknown) =>
    new Promise((resolve, reject) => h.toolCalls.push({ name, input, resolve, reject })),
  initBridge: () => {},
  persistMessage: () => {},
  postCollapse: () => {},
  postMode: () => {},
  postNewChatDivider: () => {},
  postTlsBypass: () => {},
  requestDocSettings: () => {},
  requestHistory: () => {},
  requestOfficeLanguage: () => {},
  requestOfficeTheme: () => {},
  saveDocSettings: () => {},
}))

vi.mock('../web-src/app-shell/settings', () => ({
  MAX_TOKENS: 8192,
  getSettings: () => ({
    ai: { provider: 'anthropic', providers: { anthropic: { apiKey: 'k', model: 'm' } } },
    skipTlsVerify: false, theme: 'light', lang: 'en', reasoningEffort: 'default', developerTools: true,
    confirmIrreversible: h.confirmIrreversible,
  }),
  setSettings: () => {},
  makeTransport: () => ({
    stream: (req: unknown, cb: unknown) => {
      h.streams.push({ req, cb })
      return { cancel() {} }
    },
  }),
}))

import { startAddIn, type ConfirmRule } from '../web-src/app-shell/bootstrap'

const flush = () => new Promise((r) => setTimeout(r, 0))
const schema = { type: 'object', properties: { message_id: { type: 'string' }, permanent: { type: 'boolean' } } }
const display = (n: string) => ({ label: { en: n, he: n }, description: { en: n, he: n } })

const rules: ConfirmRule[] = [
  { name: 'delete_email', when: (i) => i.permanent === true, danger: true,
    describe: (i) => (typeof i.message_id === 'string' ? { tool: 'get_email', input: { message_id: i.message_id } } : null),
    title: { en: 'Delete forever?', he: 'למחוק לצמיתות?' } },
  { name: 'delete_email',
    describe: (i) => (typeof i.message_id === 'string' ? { tool: 'get_email', input: { message_id: i.message_id } } : null),
    title: { en: 'Move to Deleted Items?', he: 'להעביר לפריטים שנמחקו?' } },
]

function boot() {
  document.body.innerHTML = '<div id="root"></div>'
  h.toolCalls.length = 0
  h.streams.length = 0
  h.confirmIrreversible = true
  startAddIn({
    tools: ['get_email', 'delete_email', 'move_email'].map((name) => ({ name, description: name, inputSchema: schema })) as any,
    toolDisplay: { get_email: display('get'), delete_email: display('delete'), move_email: display('move') },
    systemPrompt: 'sys',
    skillId: 'test',
    starters: [],
    readOnlyTools: ['get_email'],
    confirmActions: rules,
  })
  const root = document.getElementById('root')!
  const textarea = root.querySelector<HTMLTextAreaElement>('.ai-textarea')!
  const send = (text: string) => {
    textarea.value = text
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  }
  return { root, send }
}

/** The model asks for one tool call; returns once the loop has started running it. */
async function modelCalls(send: (t: string) => void, name: string, input: Record<string, unknown>) {
  send('please')
  await flush()
  const s = h.streams.at(-1)!
  s.cb.onToolCall({ id: 't1', name, input })
  s.cb.onDone()
  await flush()
}

const bridgeNames = () => h.toolCalls.map((c) => c.name)
const card = (root: HTMLElement) => root.querySelector<HTMLElement>('.ai-confirm')
const press = (root: HTMLElement, cls: string) => root.querySelector<HTMLButtonElement>(`.${cls}`)!.click()

/** Answers the get_email preview the card waits for, then lets the card render. */
async function answerPreview(output = 'Subject: Quarterly report') {
  const preview = h.toolCalls.find((c) => c.name === 'get_email')!
  preview.resolve({ output, isError: false, summary: 'get_email' })
  await flush()
}

describe('Approve/Deny card for confirmActions', () => {
  beforeEach(() => vi.restoreAllMocks())

  it('shows what the id points at, and runs the tool only after Approve', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    expect(bridgeNames()).toEqual(['get_email'])
    await answerPreview()
    expect(card(root)!.textContent).toContain('Move to Deleted Items?')
    expect(card(root)!.textContent).toContain('Subject: Quarterly report')
    expect(bridgeNames()).toEqual(['get_email'])

    press(root, 'ai-confirm-approve')
    await flush()
    expect(bridgeNames()).toEqual(['get_email', 'delete_email'])
    expect(card(root)!.classList.contains('decided')).toBe(true)
  })

  it('on Deny never calls the tool and tells the model not to retry', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    press(root, 'ai-confirm-deny')
    await flush()
    expect(bridgeNames()).toEqual(['get_email'])
    expect(JSON.stringify(h.streams.at(-1)!.req.messages)).toContain('The user declined this action')
  })

  it('uses the danger wording for a permanent delete and the plain wording otherwise', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1', permanent: true })
    await answerPreview()
    expect(card(root)!.textContent).toContain('Delete forever?')
    expect(card(root)!.classList.contains('danger')).toBe(true)
  })

  it('denies when Escape is pressed or the run is stopped', async () => {
    const first = boot()
    await modelCalls(first.send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    card(first.root)!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await flush()
    expect(bridgeNames()).toEqual(['get_email'])

    const second = boot()
    await modelCalls(second.send, 'delete_email', { message_id: 'M2' })
    await answerPreview()
    second.root.querySelector<HTMLButtonElement>('.ai-stop-btn')!.click()
    await flush()
    expect(bridgeNames()).toEqual(['get_email'])
    expect(card(second.root)!.classList.contains('decided')).toBe(true)
  })

  it('denies a pending card when New chat starts', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!.click()
    await flush()
    expect(bridgeNames()).toEqual(['get_email'])
  })

  it('still asks, with the call itself, when the preview cannot be read', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    h.toolCalls.find((c) => c.name === 'get_email')!.reject(new Error('gone'))
    await flush()
    expect(card(root)!.textContent).toContain('delete_email({"message_id":"M1"})')
  })

  it('does not ask for a tool no rule names', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'move_email', { message_id: 'M1' })
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['move_email'])
  })

  it('does not ask when the setting is off', async () => {
    const { root, send } = boot()
    h.confirmIrreversible = false
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['delete_email'])
  })

  it('does not ask for a call the user typed with !', async () => {
    const { root, send } = boot()
    send('!delete_email({"message_id":"M1"})')
    await flush()
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['delete_email'])
  })
})
