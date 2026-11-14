import { describe, expect, it } from 'vitest'
import { filterTools, scanArgs, paletteState, parseDirectCall, schemaParams, suggestTool, validateArgs, type PaletteTool } from './tool-palette'

const tool = (name: string, inputSchema?: Record<string, unknown>): PaletteTool => ({
  name, label: { en: name, he: name }, description: { en: '', he: '' }, inputSchema,
})

const EDIT_TABLE = {
  type: 'object',
  properties: {
    kind: { type: 'string', enum: ['set_cell', 'insert_row'] },
    row: { type: 'number' },
    borders: { type: 'boolean', description: 'Draws borders.' },
  },
  required: ['kind'],
}

describe('paletteState', () => {
  it('opens the menu only for a message that starts with !', () => {
    expect(paletteState('!')).toEqual({ kind: 'menu', query: '' })
    expect(paletteState('!edit_')).toEqual({ kind: 'menu', query: 'edit_' })
    expect(paletteState('wow!')).toEqual({ kind: 'off' })
    expect(paletteState('hello !edit')).toEqual({ kind: 'off' })
  })

  it('switches to the signature once a parenthesis follows a name', () => {
    expect(paletteState('!edit_table({"kind"')).toEqual({ kind: 'signature', name: 'edit_table', argsText: '{"kind"' })
    expect(paletteState('!edit_table(')).toEqual({ kind: 'signature', name: 'edit_table', argsText: '' })
  })

  it('closes on text that cannot be a tool name', () => {
    expect(paletteState('!edit table')).toEqual({ kind: 'off' })
    expect(paletteState('!(')).toEqual({ kind: 'off' })
    expect(paletteState('!9x(')).toEqual({ kind: 'off' })
  })
})

describe('filterTools', () => {
  const tools = [tool('edit_table'), tool('edit_smartart'), tool('read_object')]
  it('matches the start of the name, case-insensitively, keeping the app order', () => {
    expect(filterTools(tools, 'edit_').map((t) => t.name)).toEqual(['edit_table', 'edit_smartart'])
    expect(filterTools(tools, 'EDIT_S').map((t) => t.name)).toEqual(['edit_smartart'])
    expect(filterTools(tools, '').length).toBe(3)
  })
  it('lists read tools before the rest so each group is contiguous', () => {
    const mixed = [{ ...tool('write_a'), minMode: 'trackChanges' as const }, { ...tool('write_b'), minMode: 'trackChanges' as const }, { ...tool('read_a'), minMode: 'readOnly' as const }, tool('x')]
    expect(filterTools(mixed, '').map((t) => t.name)).toEqual(['read_a', 'write_a', 'write_b', 'x'])
  })
  it('does not match in the middle of a name', () => {
    expect(filterTools(tools, 'table')).toEqual([])
  })
})

describe('schemaParams', () => {
  it('lists type, requiredness, enum values and description', () => {
    expect(schemaParams(EDIT_TABLE)).toEqual([
      { name: 'kind', type: 'string', required: true, enumValues: ['set_cell', 'insert_row'] },
      { name: 'row', type: 'number', required: false },
      { name: 'borders', type: 'boolean', required: false, description: 'Draws borders.' },
    ])
  })
  it('is empty without a schema', () => {
    expect(schemaParams(undefined)).toEqual([])
  })
})

describe('parseDirectCall', () => {
  it('reads a JSON object', () => {
    expect(parseDirectCall('!edit_table({"kind":"set_cell","row":0})')).toEqual({ ok: true, name: 'edit_table', args: { kind: 'set_cell', row: 0 } })
  })
  it('takes empty parentheses as no arguments', () => {
    expect(parseDirectCall('!get_headings()')).toEqual({ ok: true, name: 'get_headings', args: {} })
  })
  it('keeps parentheses inside strings', () => {
    expect(parseDirectCall('!insert_content({"text":"a (b)"})')).toEqual({ ok: true, name: 'insert_content', args: { text: 'a (b)' } })
  })
  it('rejects a missing closing parenthesis, bad JSON and non-objects', () => {
    expect(parseDirectCall('!edit_table({"kind":"set_cell"}').ok).toBe(false)
    expect(parseDirectCall('!edit_table({kind})').ok).toBe(false)
    expect(parseDirectCall('!edit_table([1])').ok).toBe(false)
    expect(parseDirectCall('!edit_table').ok).toBe(false)
  })
})

describe('validateArgs', () => {
  it('passes valid arguments', () => {
    expect(validateArgs(EDIT_TABLE, { kind: 'set_cell', row: 1 })).toEqual([])
  })
  it('reports a missing required key, a wrong type and a bad enum value', () => {
    expect(validateArgs(EDIT_TABLE, { row: 1 })).toEqual(['kind is required.'])
    expect(validateArgs(EDIT_TABLE, { kind: 'set_cell', row: '1' })).toEqual(['row must be a number.'])
    expect(validateArgs(EDIT_TABLE, { kind: 'nope' })).toEqual(['kind must be one of: set_cell, insert_row.'])
  })
  it('lets keys the schema does not list through', () => {
    expect(validateArgs(EDIT_TABLE, { kind: 'set_cell', extra: 1 })).toEqual([])
  })
  it('accepts anything when there is no schema', () => {
    expect(validateArgs(undefined, { a: 1 })).toEqual([])
  })
})

describe('scanArgs', () => {
  it('is in the value of the last key until a comma', () => {
    expect(scanArgs('{"kind": "set_cell", "row": ')).toMatchObject({ at: 'value', key: 'row', keys: ['kind', 'row'] })
    expect(scanArgs('{"kind": "set_cell", "row": 1')).toMatchObject({ at: 'value', key: 'row' })
  })
  it('lets go of the previous key after a comma, and reads the name being typed', () => {
    expect(scanArgs('{"kind": "set_cell",')).toMatchObject({ at: 'key', key: null, partial: '' })
    expect(scanArgs('{"kind": "set_cell", "ro')).toMatchObject({ at: 'key', key: null, partial: 'ro' })
  })
  it('follows the caret inside earlier text', () => {
    const text = '{"kind": "set_cell", "row": 1}'
    expect(scanArgs(text, 10)).toMatchObject({ at: 'value', key: 'kind' })
  })
  it('ignores commas, colons and keys inside strings and nested values', () => {
    expect(scanArgs('{"text": "a, \\"b\\": c", "opts": {"x": 1, "y": 2}, ')).toMatchObject({ at: 'key', keys: ['text', 'opts'] })
  })
  it('is empty before any key', () => {
    expect(scanArgs('{')).toMatchObject({ at: 'key', keys: [], key: null })
    expect(scanArgs('')).toMatchObject({ at: null, keys: [] })
  })
})

describe('suggestTool', () => {
  const tools = [tool('edit_table'), tool('edit_smartart'), tool('read_object')]
  it('finds a near typo', () => {
    expect(suggestTool(tools, 'edit_tabel')).toBe('edit_table')
  })
  it('gives nothing for a name that is far from every tool', () => {
    expect(suggestTool(tools, 'zzzzzzzz')).toBeNull()
  })
  it('reads an escaped key by the name it parses to', () => {
    expect(scanArgs('{"kind": "set_cell", "\u0074ext": 1')).toMatchObject({ keys: ['kind', 'text'] })
  })
})
