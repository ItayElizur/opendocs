import { beforeEach, describe, expect, it, vi } from 'vitest'

const MODEL_KEY = 'MODELKEY-0123456789'

const h = vi.hoisted(() => ({
  toolCalls: [] as Array<{ name: string; input: unknown; resolve: (r: unknown) => void }>,
  persisted: [] as Array<{ role: string; text: string }>,
  streams: [] as Array<{ req: any; cb: any }>,
  disabledModes: [] as string[],
}))

vi.mock('../web-src/app-shell/bridge', () => ({
  callDotNetTool: (name: string, input: unknown) => new Promise((resolve) => h.toolCalls.push({ name, input, resolve })),
  initBridge: () => {},
  persistMessage: (role: string, text: string) => { h.persisted.push({ role, text }) },
  postCollapse: () => {},
  postMode: () => {},
  injectedDisabledModes: () => h.disabledModes,
  injectedApprovals: () => ({}),
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
    ai: { provider: 'anthropic', providers: { anthropic: { apiKey: MODEL_KEY, model: 'claude-sonnet-5' } } },
    skipTlsVerify: false, theme: 'light', lang: 'en', reasoningEffort: 'default', developerTools: true,
  }),
  setSettings: () => {},
  makeTransport: () => ({
    stream: (req: unknown, cb: unknown) => {
      h.streams.push({ req, cb })
      return { cancel() {} }
    },
  }),
}))

import { AgentLoop } from '@officeai/agent-core'
import { startAddIn } from '../web-src/app-shell/bootstrap'

const flush = () => new Promise((r) => setTimeout(r, 0))

function boot() {
  document.body.innerHTML = '<div id="root"></div>'
  h.toolCalls.length = 0
  h.persisted.length = 0
  h.streams.length = 0
  startAddIn({
    tools: [{ name: 'read_doc', description: 'reads', inputSchema: { type: 'object', properties: { text: { type: 'string' } } } }] as any,
    toolDisplay: { read_doc: { label: { en: 'Read', he: 'Read' }, description: { en: 'reads', he: 'reads' } } },
    systemPrompt: 'sys',
    skillId: 'test',
    starters: [],
    readOnlyTools: ['read_doc'],
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

const finishCall = (output: string) => h.toolCalls.at(-1)!.resolve({ output, isError: false, summary: 'read_doc' })
const lastMessages = () => JSON.stringify(h.streams.at(-1)!.req.messages)

describe('direct tool call in the app shell', () => {
  beforeEach(() => { vi.restoreAllMocks(); h.disabledModes = [] })

  it('refuses Continue while a direct call runs, and the call still reaches the model afterwards', async () => {
    const { root, send } = boot()
    send('hello')
    await flush()
    h.streams[0]!.cb.onDelta('partial')
    h.streams[0]!.cb.onStopReason('max_tokens')
    h.streams[0]!.cb.onDone()
    await flush()
    const continueBtn = root.querySelector<HTMLButtonElement>('.ai-notice-action')!

    send('!read_doc({})')
    continueBtn.click()
    await flush()
    expect(h.streams.length).toBe(1)

    finishCall('DIRECT-RESULT')
    await flush()
    send('what did the tool return?')
    await flush()
    expect(lastMessages()).toContain('DIRECT-RESULT')
  })

  it('drops a direct call that finishes after New chat, and lets the next one run', async () => {
    const { root, send } = boot()
    send('!read_doc({})')
    root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!.click()
    finishCall('OLD-CONVERSATION-RESULT')
    await flush()
    send('fresh question')
    await flush()
    expect(h.streams.at(-1)!.req.messages).toHaveLength(1)
    expect(h.persisted.filter((m) => m.role === 'assistant')).toEqual([])
    h.streams.at(-1)!.cb.onDelta('ok')
    h.streams.at(-1)!.cb.onDone()
    await flush()

    send('!read_doc({})')
    expect(h.toolCalls).toHaveLength(2)
  })

  it('stores the typed arguments as typed, and redacts only the model key from the typed line', async () => {
    const { send } = boot()
    send(`!read_doc({"text":"sk-ABCDEFGHIJKLMNOPQRSTUVWX ${MODEL_KEY}"})`)
    finishCall('ok')
    await flush()
    send('another message')
    await flush()
    const messages = h.streams.at(-1)!.req.messages
    expect(messages[0].text).toContain('[REDACTED_API_KEY]')
    expect(messages[0].text).not.toContain(MODEL_KEY)
    expect(messages[1].toolCalls[0].input.text).toBe(`sk-ABCDEFGHIJKLMNOPQRSTUVWX ${MODEL_KEY}`)
  })

  it('saves a reply after the typed line, so the call survives a reload', async () => {
    const { send } = boot()
    send('!read_doc({})')
    finishCall('DIRECT-RESULT')
    await flush()
    expect(h.persisted).toEqual([
      { role: 'user', text: '!read_doc({})' },
      { role: 'assistant', text: 'read_doc: DIRECT-RESULT' },
    ])
    const loop = new AgentLoop({ transport: { stream: () => ({ cancel() {} }) } as any, skill: { tools: [], executeTool: async () => ({ output: '', isError: false, summary: '' }) } as any })
    loop.restore(h.persisted as any)
    expect(JSON.stringify(loop.messages)).toContain('!read_doc')
  })
  it('does not run a message queued behind a direct call in the chat opened by New chat', async () => {
    const { root, send } = boot()
    send('!read_doc({})')
    send('queued behind the call')
    root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!.click()
    finishCall('OLD')
    await flush()
    send('fresh question')
    await flush()
    h.streams.at(-1)!.cb.onDelta('ok')
    h.streams.at(-1)!.cb.onDone()
    await flush()
    expect(h.streams).toHaveLength(1)
    expect(JSON.stringify(h.streams[0]!.req.messages)).not.toContain('queued behind the call')
  })
})

describe('editing modes switched off by policy.json', () => {
  beforeEach(() => { h.disabledModes = [] })

  const menuModes = (root: HTMLElement) => [...root.querySelectorAll<HTMLElement>('.ai-mode-menu-item')].map((el) => el.dataset.mode)

  it('leaves a disabled mode out of the menu and still offers the model the tools of the remaining modes', async () => {
    h.disabledModes = ['fullAutonomy']
    const { root, send } = boot()
    expect(menuModes(root)).toEqual(['readOnly', 'commentOnly', 'trackChanges'])
    send('hello')
    await flush()
    expect(h.streams[0]!.req.tools.map((t: { name: string }) => t.name)).toEqual(['read_doc'])
  })

  it('with every mode disabled, hides the menu, says why and offers no tools', async () => {
    h.disabledModes = ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy']
    const { root, send } = boot()
    expect(menuModes(root)).toEqual([])
    expect(root.querySelector<HTMLButtonElement>('.ai-mode-btn')!.hidden).toBe(true)
    expect(root.textContent).toContain('turned off every editing mode')
    send('hello')
    await flush()
    expect(h.streams[0]!.req.tools ?? []).toEqual([])
  })
})
