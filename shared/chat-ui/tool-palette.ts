import type { EditingMode } from './chat-ui'

// Pure logic behind the "!" direct-call palette: which text opens it, how the
// tool list filters, what a signature looks like, and whether typed arguments
// fit the tool's schema. No DOM here so it can be unit-tested; chat-ui.ts owns
// the popup and the keyboard.

export interface PaletteTool {
  name: string
  label: { en: string; he: string }
  description: { en: string; he: string }
  /** JSON Schema (object) the model reads; the source of the signature. */
  inputSchema?: Record<string, unknown> | undefined
  /** Lowest editing mode that offers this tool to the model. */
  minMode?: EditingMode | undefined
}

export interface SchemaParam {
  name: string
  type: string
  required: boolean
  enumValues?: string[]
  description?: string
}

export type PaletteState =
  | { kind: 'off' }
  | { kind: 'menu'; query: string }
  | { kind: 'signature'; name: string; argsText: string }

/** The palette is for a message that starts with "!", so "wow!" never opens it. */
export function paletteState(text: string): PaletteState {
  if (!text.startsWith('!')) return { kind: 'off' }
  const open = text.indexOf('(')
  if (open < 0) return /^!\S*$/.test(text) ? { kind: 'menu', query: text.slice(1) } : { kind: 'off' }
  const name = text.slice(1, open).trim()
  if (!/^[A-Za-z_]\w*$/.test(name)) return { kind: 'off' }
  return { kind: 'signature', name, argsText: text.slice(open + 1) }
}

/** Prefix match on the tool name, in the app's own order. */
export function filterTools(tools: readonly PaletteTool[], query: string): PaletteTool[] {
  const q = query.toLowerCase()
  return tools.filter((t) => t.name.toLowerCase().startsWith(q))
}

export function schemaParams(schema: Record<string, unknown> | undefined): SchemaParam[] {
  const props = (schema?.properties ?? {}) as Record<string, Record<string, unknown>>
  const required = new Set(Array.isArray(schema?.required) ? (schema!.required as string[]) : [])
  return Object.entries(props).map(([name, p]) => {
    const param: SchemaParam = { name, type: typeof p.type === 'string' ? p.type : 'any', required: required.has(name) }
    if (Array.isArray(p.enum)) param.enumValues = p.enum.map(String)
    if (typeof p.description === 'string') param.description = p.description
    return param
  })
}

export type ParsedCall = { ok: true; name: string; args: Record<string, unknown> } | { ok: false; error: string }

/** `!name(...)` with an empty body or a JSON object; nothing may follow the closing parenthesis. */
export function parseDirectCall(text: string): ParsedCall {
  const state = paletteState(text.trim())
  if (state.kind !== 'signature') return { ok: false, error: 'Write the call as !tool_name({ ...json... }).' }
  const body = state.argsText.trimEnd()
  if (!body.endsWith(')')) return { ok: false, error: 'Missing the closing parenthesis.' }
  const json = body.slice(0, -1).trim()
  if (json === '') return { ok: true, name: state.name, args: {} }
  let parsed: unknown
  try {
    parsed = JSON.parse(json)
  } catch (e) {
    return { ok: false, error: `Arguments are not valid JSON: ${e instanceof Error ? e.message : String(e)}` }
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    return { ok: false, error: 'Arguments must be one JSON object.' }
  }
  return { ok: true, name: state.name, args: parsed as Record<string, unknown> }
}

function typeMatches(type: string, value: unknown): boolean {
  switch (type) {
    case 'string': return typeof value === 'string'
    case 'number': return typeof value === 'number'
    case 'integer': return typeof value === 'number' && Number.isInteger(value)
    case 'boolean': return typeof value === 'boolean'
    case 'array': return Array.isArray(value)
    case 'object': return typeof value === 'object' && value !== null && !Array.isArray(value)
    default: return true
  }
}

/** Required keys, JSON types and enum values. Keys the schema does not list are let through. */
export function validateArgs(schema: Record<string, unknown> | undefined, args: Record<string, unknown>): string[] {
  const errors: string[] = []
  for (const p of schemaParams(schema)) {
    const value = args[p.name]
    if (value === undefined) {
      if (p.required) errors.push(`${p.name} is required.`)
      continue
    }
    if (!typeMatches(p.type, value)) errors.push(`${p.name} must be a ${p.type}.`)
    else if (p.enumValues && !p.enumValues.includes(String(value))) errors.push(`${p.name} must be one of: ${p.enumValues.join(', ')}.`)
  }
  return errors
}

/** The argument key the caret is inside, for highlighting it in the signature. */
export function keyAtCaret(argsText: string, caret: number): string | null {
  let key: string | null = null
  for (const m of argsText.slice(0, caret).matchAll(/"([^"\\]+)"\s*:/g)) key = m[1]!
  return key
}

function distance(a: string, b: string): number {
  const row = Array.from({ length: b.length + 1 }, (_, i) => i)
  for (let i = 1; i <= a.length; i++) {
    let prev = row[0]!
    row[0] = i
    for (let j = 1; j <= b.length; j++) {
      const tmp = row[j]!
      row[j] = Math.min(row[j]! + 1, row[j - 1]! + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1))
      prev = tmp
    }
  }
  return row[b.length]!
}

/** The closest tool name, when one is near enough to be a typo. */
export function suggestTool(tools: readonly PaletteTool[], name: string): string | null {
  let best: string | null = null
  let bestDistance = 3
  for (const t of tools) {
    const d = distance(name.toLowerCase(), t.name.toLowerCase())
    if (d < bestDistance) {
      best = t.name
      bestDistance = d
    }
  }
  return best
}
