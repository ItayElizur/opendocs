import { beforeEach, describe, expect, it, vi } from 'vitest'

// Outlook's entry.ts picks the item window's modes and tools from window.__opendocsPaneKind (ItemPaneHost.cs).
const h = vi.hoisted(() => ({ configs: [] as any[] }))
vi.mock('@officeai/app-shell', () => ({ startAddIn: (config: unknown) => { h.configs.push(config) } }))

async function boot(paneKind: string | undefined) {
  ;(globalThis as { __opendocsPaneKind?: string }).__opendocsPaneKind = paneKind
  vi.resetModules()
  h.configs.length = 0
  await import('../../OutlookAiAddIn/web-src/entry')
  expect(h.configs).toHaveLength(1)
  return h.configs[0]
}

const MAILBOX_ACTIONS = ['move_email', 'delete_email', 'mark_email_read', 'flag_email_important', 'draft_email', 'draft_event',
  'respond_meeting', 'send_email', 'send_event', 'undo_last_action', 'redo_last_action', 'apply_search', 'create_task', 'set_event_categories']

describe('Outlook item window', () => {
  beforeEach(() => { delete (globalThis as { __opendocsPaneKind?: string }).__opendocsPaneKind })

  it.each(['mail', 'appointment'])('a %s window offers two modes and no mailbox action', async (kind) => {
    const c = await boot(kind)
    expect(c.availableModes).toEqual(['readOnly', 'commentOnly'])
    expect(c.defaultMode).toBe('commentOnly')
    const names: string[] = c.tools.map((t: { name: string }) => t.name)
    expect(names.filter((n) => MAILBOX_ACTIONS.includes(n))).toEqual([])
    expect(names).toEqual(expect.arrayContaining(['get_current_item', 'edit_current_body', 'set_current_fields', 'read_main_chat', 'search_emails']))
    expect(new Set(names)).toEqual(new Set([...c.readOnlyTools, ...c.commentOnlyExtraTools]))
    expect(c.commentOnlyExtraTools).toEqual(['edit_current_body', 'set_current_fields'])
    for (const n of names) expect(c.toolDisplay[n], n).toBeDefined()
    expect(c.modeOverrides.commentOnly.label.en).toBe(kind === 'mail' ? 'Edit this draft' : 'Edit this meeting')
  })

  it('takes the item-window versions of find_meeting_slots and of the recurrence argument', async () => {
    const c = await boot('appointment')
    const slots = c.tools.find((t: { name: string }) => t.name === 'find_meeting_slots')
    expect(slots.inputSchema.required).toEqual(['duration_minutes'])
    expect(slots.description).toContain('set_current_fields')
    const fields = c.tools.find((t: { name: string }) => t.name === 'set_current_fields')
    expect(fields.inputSchema.properties.recurrence.description).not.toContain('New events only')
  })

  it('the main window keeps its four modes and the mailbox tools', async () => {
    const c = await boot(undefined)
    expect(c.availableModes).toEqual(['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy'])
    const names: string[] = c.tools.map((t: { name: string }) => t.name)
    expect(names).toContain('move_email')
    expect(names).not.toContain('edit_current_body')
  })
})
