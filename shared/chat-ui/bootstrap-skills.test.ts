import { beforeEach, describe, expect, it, vi } from 'vitest'

const h = vi.hoisted(() => ({
  toolCalls: [] as Array<{ name: string; input: unknown; resolve: (r: unknown) => void }>,
  persisted: [] as Array<{ role: string; text: string }>,
  streams: [] as Array<{ req: any; cb: any }>,
  skills: [] as unknown[],
  approvals: {} as Record<string, string[]>,
  bridge: null as any,
}))

vi.mock('../web-src/app-shell/bridge', () => ({
  callDotNetTool: (name: string, input: unknown) => new Promise((resolve) => h.toolCalls.push({ name, input, resolve })),
  initBridge: (handlers: unknown) => { h.bridge = handlers },
  persistMessage: (role: string, text: string) => { h.persisted.push({ role, text }) },
  postCollapse: () => {},
  postMode: () => {},
  injectedDisabledModes: () => [],
  injectedApprovals: () => h.approvals,
  injectedSkills: () => h.skills,
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
    ai: { provider: 'anthropic', providers: { anthropic: { apiKey: 'k', model: 'claude-sonnet-5' } } },
    skipTlsVerify: false, theme: 'light', lang: 'en', reasoningEffort: 'default', developerTools: false,
  }),
  setSettings: () => {},
  makeTransport: () => ({
    stream: (req: unknown, cb: unknown) => {
      h.streams.push({ req, cb })
      return { cancel() {} }
    },
  }),
}))

import { startAddIn } from '../web-src/app-shell/bootstrap'

const flush = () => new Promise((r) => setTimeout(r, 0))

const REPORT = {
  name: 'report', description: 'Writes the weekly report', argumentHint: '[week]', dir: 'C:\\skills\\report',
  body: 'Report steps.', files: ['references/style.md'], modelInvocable: true, userInvocable: true, source: 'machine',
}

function boot() {
  document.body.innerHTML = '<div id="root"></div>'
  h.toolCalls.length = 0
  h.persisted.length = 0
  h.streams.length = 0
  startAddIn({
    tools: [
      { name: 'read_doc', description: 'reads', inputSchema: { type: 'object' } },
      { name: 'edit_doc', description: 'edits', inputSchema: { type: 'object' } },
    ] as any,
    toolDisplay: {
      read_doc: { label: { en: 'Read', he: 'Read' }, description: { en: 'reads', he: 'reads' } },
      edit_doc: { label: { en: 'Edit', he: 'Edit' }, description: { en: 'edits', he: 'edits' } },
    },
    systemPrompt: 'sys',
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
  const reply = async (text = 'ok') => {
    h.streams.at(-1)!.cb.onDelta(text)
    h.streams.at(-1)!.cb.onDone()
    await flush()
  }
  return { root, send, reply }
}

const userTexts = () => h.streams.at(-1)!.req.messages.filter((m: any) => m.role === 'user').map((m: any) => m.text as string)
const toolNames = () => h.streams.at(-1)!.req.tools.map((t: { name: string }) => t.name)

describe('skills in the app shell', () => {
  beforeEach(() => { h.skills = [REPORT]; h.approvals = {} })

  it('sends the skill for a typed /name, but shows and saves what was typed', async () => {
    const { root, send } = boot()
    send('/report week 12')
    await flush()
    expect(userTexts()[0]).toMatch(/^Base directory for this skill: C:\\skills\\report\n\nReport steps\./)
    expect(userTexts()[0]).toContain('ARGUMENTS: week 12')
    expect(h.persisted[0]).toEqual({ role: 'user', text: '/report week 12' })
    expect(root.querySelector('.ai-msg-user, .ai-user')?.textContent ?? root.textContent).toContain('/report week 12')
  })

  it('sends a re-invocation note the second time', async () => {
    const { send, reply } = boot()
    send('/report week 12')
    await flush()
    await reply()
    send('/report week 13')
    await flush()
    expect(userTexts()[1]).toBe('(Re-invocation of /report: the skill instructions were previously loaded; the arguments below are new.)\n\nARGUMENTS: week 13')
  })

  it('expands saved /name lines again when a chat is restored', async () => {
    const { send } = boot()
    h.bridge.onHistoryLoaded([
      { role: 'user', text: '/report week 12' },
      { role: 'assistant', text: 'done' },
      { role: 'user', text: '/report week 13' },
      { role: 'assistant', text: 'done again' },
    ])
    send('thanks')
    await flush()
    const texts = userTexts()
    expect(texts[0]).toContain('Base directory for this skill: C:\\skills\\report')
    expect(texts[1]).toContain('(Re-invocation of /report')
    expect(texts[2]).toBe('thanks')
  })

  it('lists the skill in the system prompt and answers load_skill without the app or an Approve', async () => {
    h.approvals = { '*': ['trackChanges'] }
    const { root, send } = boot()
    send('write the weekly report')
    await flush()
    expect(h.streams[0]!.req.system).toContain('- report: Writes the weekly report')
    expect(toolNames()).toEqual(['read_doc', 'edit_doc', 'load_skill', 'read_skill_file'])
    h.streams[0]!.cb.onToolCall({ id: 't1', name: 'load_skill', input: { name: 'report' } })
    h.streams[0]!.cb.onDone()
    await flush()
    expect(root.querySelector('.ai-confirm-card, [data-confirm]')).toBeNull()
    expect(h.toolCalls).toEqual([])
    const results = h.streams.at(-1)!.req.messages.filter((m: any) => m.role === 'tool')
    expect(results[0].results[0].output).toContain('Base directory for this skill: C:\\skills\\report')
    expect(results[0].results[0].output).toContain('- references/style.md')
  })

  it('sends read_skill_file to the app, also in Read only', async () => {
    const { root, send } = boot()
    root.querySelector<HTMLElement>('.ai-mode-menu-item[data-mode="readOnly"]')!.click()
    send('read the style guide')
    await flush()
    expect(toolNames()).toEqual(['read_doc', 'load_skill', 'read_skill_file'])
    h.streams[0]!.cb.onToolCall({ id: 't1', name: 'read_skill_file', input: { skill: 'report', path: 'references/style.md' } })
    h.streams[0]!.cb.onDone()
    await flush()
    expect(h.toolCalls.map((c) => c.name)).toEqual(['read_skill_file'])
  })

  it('without skills, the tools and prompt are what the app configured and /text is sent as typed', async () => {
    h.skills = []
    const { send } = boot()
    send('/report week 12')
    await flush()
    expect(toolNames()).toEqual(['read_doc', 'edit_doc'])
    expect(h.streams[0]!.req.system).not.toContain('Skills:')
    expect(userTexts()[0]).toBe('/report week 12')
  })
})
