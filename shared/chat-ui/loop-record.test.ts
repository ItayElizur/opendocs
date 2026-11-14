import { describe, expect, it } from 'vitest'
import { AgentLoop, COMPLETED_VIA_TOOLS_TEXT, type AgentSkill, type AgentTransport } from '@officeai/agent-core'

const skill = { name: 'test', tools: [], executeTool: async () => ({ output: '', isError: false, summary: '' }), getContext: async () => '' } as unknown as AgentSkill
const idle = { stream: () => ({ cancel() {} }) } as unknown as AgentTransport
const call = { id: 'c1', name: 'edit_table', input: { kind: 'set_cell' } }
const execution = { output: 'ok', isError: false, summary: 'edit_table' }

describe('AgentLoop.recordToolExchange', () => {
  it('adds the typed text, the call, its result and a closing assistant turn, in the order a model run produces', () => {
    const loop = new AgentLoop({ transport: idle, skill })
    loop.recordToolExchange('!edit_table({"kind":"set_cell"})', call, execution)
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'tool', 'assistant'])
    expect(loop.messages[1]).toMatchObject({ toolCalls: [call] })
    expect(loop.messages[2]).toMatchObject({ results: [{ id: 'c1', name: 'edit_table', output: 'ok', isError: false }] })
    expect(loop.messages[3]).toMatchObject({ text: COMPLETED_VIA_TOOLS_TEXT })
  })

  it('replaces a trailing user message that never got a reply', () => {
    const loop = new AgentLoop({ transport: idle, skill })
    loop.recordToolExchange('first', call, execution)
    loop.recordToolExchange('second', call, execution)
    expect(loop.messages.filter((m) => m.role === 'user')).toHaveLength(2)
  })

  it('does nothing while a model run is in flight', () => {
    const loop = new AgentLoop({ transport: idle, skill })
    loop.run('hello')
    expect(loop.busy).toBe(true)
    const before = loop.messages.length
    loop.recordToolExchange('!x()', call, execution)
    expect(loop.messages.length).toBe(before)
  })
})
