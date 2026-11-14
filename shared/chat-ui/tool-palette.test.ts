import { describe, expect, it } from 'vitest'
import { filterTools, keyAtCaret, paletteState, parseDirectCall, schemaParams, suggestTool, validateArgs, type PaletteTool } from './tool-palette'

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

describe('keyAtCaret', () => {
  const text = '{"kind": "set_cell", "row": '
  it('is the last key before the caret', () => {
    expect(keyAtCaret(text, text.length)).toBe('row')
    expect(keyAtCaret(text, 10)).toBe('kind')
  })
  it('is null before any key', () => {
    expect(keyAtCaret('{', 1)).toBeNull()
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
})
