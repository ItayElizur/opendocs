import { beforeEach, describe, expect, it, vi } from 'vitest'

const h = vi.hoisted(() => ({
  toolCalls: [] as Array<{ name: string; input: any; resolve: (r: unknown) => void; reject: (e: unknown) => void }>,
  streams: [] as Array<{ req: any; cb: any }>,
  approvals: {} as Record<string, string[]>,
}))

vi.mock('../web-src/app-shell/bridge', () => ({
  callDotNetTool: (name: string, input: unknown) =>
    new Promise((resolve, reject) => h.toolCalls.push({ name, input, resolve, reject })),
  initBridge: () => {},
  persistMessage: () => {},
  postCollapse: () => {},
  postMode: () => {},
  injectedDisabledModes: () => [],
  injectedApprovals: () => h.approvals,
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
import { OUTLOOK_CONFIRM_ACTIONS } from '../../OutlookAiAddIn/web-src/confirm'

const flush = () => new Promise((r) => setTimeout(r, 0))
const schema = { type: 'object', properties: { message_id: { type: 'string' }, permanent: { type: 'boolean' } } }
const display = (n: string) => ({ label: { en: n, he: n }, description: { en: n, he: n } })

const rules: ConfirmRule[] = [
  { name: 'delete_email', when: (i) => i.permanent === true, danger: true, preview: true,
    title: { en: 'Delete forever?', he: 'למחוק לצמיתות?' } },
  { name: 'delete_email', preview: true, title: { en: 'Move to Deleted Items?', he: 'להעביר לפריטים שנמחקו?' } },
  { name: 'send_email', danger: true,
    fields: (i) => [{ key: 'to', value: [i.to].flat().join(', ') }, { key: 'subject', value: String(i.subject ?? '') }],
    title: { en: 'Send now?', he: 'לשלוח עכשיו?' } },
]

// The policy the tests run under unless they pass their own: the panel starts in Track changes.
const POLICY = { delete_email: ['trackChanges'], send_email: ['trackChanges'] }

function boot(approvals: Record<string, string[]> = POLICY) {
  document.body.innerHTML = '<div id="root"></div>'
  h.toolCalls.length = 0
  h.streams.length = 0
  h.approvals = approvals
  startAddIn({
    tools: ['get_email', 'delete_email', 'move_email', 'send_email'].map((name) => ({ name, description: name, inputSchema: schema })) as any,
    toolDisplay: { get_email: display('get'), delete_email: display('delete'), move_email: display('move'), send_email: display('send') },
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

const PREVIEW = { type: 'mail', subject: 'Quarterly report', from: 'Dana Levi', date: '2026-11-02 09:30' }

/** Answers the preview_action call the card waits for, then lets the card render. */
async function answerPreview(fields: Record<string, string> = PREVIEW) {
  const preview = h.toolCalls.find((c) => c.name === 'preview_action')!
  preview.resolve({ output: JSON.stringify(fields), isError: false, summary: 'preview_action' })
  await flush()
}

describe('Approve/Deny card', () => {
  beforeEach(() => vi.restoreAllMocks())

  it('shows what the id points at, and runs the tool only after Approve', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    expect(bridgeNames()).toEqual(['preview_action'])
    expect(h.toolCalls[0]!.input).toEqual({ tool: 'delete_email', input: { message_id: 'M1' } })
    await answerPreview()
    const text = card(root)!.textContent!
    expect(text).toContain('Move to Deleted Items?')
    for (const shown of ['Email', 'Quarterly report', 'Dana Levi', '2026-11-02 09:30']) expect(text).toContain(shown)
    expect(text).not.toContain('delete_email(')
    expect(text).not.toContain('M1')
    expect(bridgeNames()).toEqual(['preview_action'])

    press(root, 'ai-confirm-approve')
    await flush()
    expect(bridgeNames()).toEqual(['preview_action', 'delete_email'])
    expect(card(root)!.classList.contains('decided')).toBe(true)
  })

  it('on Deny never calls the tool and tells the model not to retry', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    press(root, 'ai-confirm-deny')
    await flush()
    expect(bridgeNames()).toEqual(['preview_action'])
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
    expect(bridgeNames()).toEqual(['preview_action'])

    const second = boot()
    await modelCalls(second.send, 'delete_email', { message_id: 'M2' })
    await answerPreview()
    second.root.querySelector<HTMLButtonElement>('.ai-stop-btn')!.click()
    await flush()
    expect(bridgeNames()).toEqual(['preview_action'])
    expect(card(second.root)!.classList.contains('decided')).toBe(true)
  })

  it('denies a pending card when New chat starts', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!.click()
    await flush()
    expect(bridgeNames()).toEqual(['preview_action'])
  })

  it('shows no card, and says why, when the app says the call would be refused', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1', permanent: true })
    await answerPreview({ blocked: 'Blocked: "Quarterly report" is not in Deleted Items.' })
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['preview_action'])
    expect(JSON.stringify(h.streams.at(-1)!.req.messages)).toContain('is not in Deleted Items')
  })

  it('still asks, with only the question, when the preview cannot be read', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    h.toolCalls.find((c) => c.name === 'preview_action')!.reject(new Error('gone'))
    await flush()
    expect(card(root)!.textContent).toContain('Move to Deleted Items?')
    expect(card(root)!.querySelector('.ai-confirm-rows')).toBeNull()
  })

  it('shows rows from the call own arguments, never the body or the call text', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'send_email', { to: ['a@x.com', 'b@x.com'], subject: 'Hello', body: 'SECRET BODY' })
    const text = card(root)!.textContent!
    expect(bridgeNames()).toEqual([])
    expect(text).toContain('a@x.com, b@x.com')
    expect(text).toContain('Hello')
    expect(text).not.toContain('SECRET BODY')
    expect(text).not.toContain('send_email(')
  })

  it('hides the stop button after New chat during a pending card', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    await answerPreview()
    expect(root.querySelector<HTMLElement>('.ai-stop-btn')!.hidden).toBe(false)
    root.querySelector<HTMLButtonElement>('[data-t-title="newChat"]')!.click()
    await flush()
    expect(root.querySelector<HTMLElement>('.ai-stop-btn')!.hidden).toBe(true)
  })

  it('does not ask for a tool the policy does not name', async () => {
    const { root, send } = boot()
    await modelCalls(send, 'move_email', { message_id: 'M1' })
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['move_email'])
  })

  it('does not ask in a mode the policy does not list for the tool', async () => {
    const { root, send } = boot({ delete_email: ['commentOnly', 'fullAutonomy'] })
    await modelCalls(send, 'delete_email', { message_id: 'M1' })
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['delete_email'])
  })

  it('asks with a plain card for a tool with no card rule, which "*" covers', async () => {
    const { root, send } = boot({ '*': ['trackChanges'] })
    await modelCalls(send, 'move_email', { message_id: 'M1', folder: 'Archive' })
    const text = card(root)!.textContent!
    expect(text).toContain('Allow this action?')
    for (const shown of ['move', 'message_id', 'M1', 'Archive']) expect(text).toContain(shown)
    expect(card(root)!.classList.contains('danger')).toBe(false)
    press(root, 'ai-confirm-approve')
    await flush()
    expect(bridgeNames()).toEqual(['move_email'])
  })

  it('"*" leaves read-only tools alone, and a tool own entry overrides it', async () => {
    for (const tool of ['get_email', 'move_email']) {
      const { root, send } = boot({ '*': ['trackChanges'], move_email: [] })
      await modelCalls(send, tool, { message_id: 'M1' })
      expect(card(root)).toBeNull()
      expect(bridgeNames()).toEqual([tool])
    }
  })

  it('fails closed on a tool name the app does not have: every change asks, in every mode, and the chat says why', async () => {
    const { root, send } = boot({ delete_emial: ['commentOnly'] })
    expect(root.textContent).toContain('delete_emial')
    await modelCalls(send, 'move_email', { message_id: 'M1' })
    expect(card(root)!.textContent).toContain('Allow this action?')
    expect(bridgeNames()).toEqual([])
  })

  it('does not ask for a call the user typed with !', async () => {
    const { root, send } = boot()
    send('!delete_email({"message_id":"M1"})')
    await flush()
    expect(card(root)).toBeNull()
    expect(bridgeNames()).toEqual(['delete_email'])
  })
})

describe("Outlook's card rules (confirm.ts)", () => {
  const ruleFor = (name: string, input: Record<string, unknown>) =>
    OUTLOOK_CONFIRM_ACTIONS.find((r) => r.name === name && (r.when?.(input) ?? true))!

  it('marks a permanent delete as dangerous and a plain delete as not', () => {
    expect(ruleFor('delete_email', { permanent: true }).danger).toBe(true)
    expect(ruleFor('delete_email', { permanent: true }).title.en).toContain('permanently')
    expect(ruleFor('delete_email', {}).danger).toBeFalsy()
    expect(ruleFor('delete_email', { permanent: false }).title.en).toContain('Deleted Items')
  })

  it('uses the cancel wording only for a cancel', () => {
    expect(ruleFor('send_event', { action: 'cancel' }).title.en).toContain('Cancel')
    expect(ruleFor('send_event', { action: 'new' }).title.en).not.toContain('Cancel')
  })
})
