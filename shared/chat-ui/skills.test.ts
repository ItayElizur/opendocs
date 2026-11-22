import { describe, expect, it } from 'vitest'
import type { AgentMessage } from '@officeai/agent-core'
import { catalogPrompt, invocationText, isLoaded, loadSkill, parseInvocation, skillBlock, toModelText, type Skill } from '../web-src/app-shell/skills'

const skill = (name: string, extra: Partial<Skill> = {}): Skill => ({
  name,
  description: `does ${name}`,
  argumentHint: '',
  dir: `C:\\skills\\${name}`,
  body: `Steps for ${name}.`,
  files: [],
  modelInvocable: true,
  userInvocable: true,
  source: 'machine',
  ...extra,
})

const user = (text: string): AgentMessage => ({ role: 'user', text })

describe('skills', () => {
  it('lists only model-invocable skills in the prompt, and nothing without skills', () => {
    const prompt = catalogPrompt([skill('report'), skill('secret', { modelInvocable: false })])
    expect(prompt).toContain('- report: does report')
    expect(prompt).not.toContain('secret')
    expect(catalogPrompt([skill('x', { modelInvocable: false })])).toBe('')
    expect(catalogPrompt([])).toBe('')
  })

  it('starts the block with the base directory and lists the files', () => {
    const block = skillBlock(skill('brand', { files: ['assets/logo.png', 'references/style.md'] }))
    expect(block.startsWith('Base directory for this skill: C:\\skills\\brand\n\nSteps for brand.')).toBe(true)
    expect(block).toContain('- assets/logo.png\n- references/style.md')
  })

  it('parses /name args only for user-invocable skills', () => {
    const skills = [skill('report'), skill('hidden', { userInvocable: false })]
    expect(parseInvocation('/report  last week ', skills)).toEqual({ skill: skills[0], args: 'last week' })
    expect(parseInvocation('/report', skills)?.args).toBe('')
    expect(parseInvocation('/hidden x', skills)).toBeNull()
    expect(parseInvocation('/unknown', skills)).toBeNull()
    expect(parseInvocation('see /report', skills)).toBeNull()
    expect(parseInvocation('/etc/passwd', skills)).toBeNull()
  })

  it('appends the arguments, or puts them where $ARGUMENTS is', () => {
    expect(invocationText(skill('a'), 'x y', [])).toBe('Base directory for this skill: C:\\skills\\a\n\nSteps for a.\n\nARGUMENTS: x y')
    expect(invocationText(skill('a'), '', [])).toBe('Base directory for this skill: C:\\skills\\a\n\nSteps for a.')
    const templ = skill('t', { body: 'Summarize $ARGUMENTS in one line.' })
    expect(invocationText(templ, 'the intro', [])).toContain('Summarize the intro in one line.')
  })

  it('sends a re-invocation note once the skill is in the history', () => {
    const a = skill('a')
    const history = [user(invocationText(a, 'first', [])), { role: 'assistant', text: 'done' } as AgentMessage]
    expect(invocationText(a, 'second', history)).toBe(
      '(Re-invocation of /a: the skill instructions were previously loaded; the arguments below are new.)\n\nARGUMENTS: second',
    )
  })

  it('counts a load_skill result and a $ARGUMENTS-filled copy as loaded', () => {
    const templ = skill('t', { body: 'Summarize $ARGUMENTS in one line.' })
    expect(isLoaded(templ, [user(invocationText(templ, 'the intro', []))])).toBe(true)
    const tool: AgentMessage = { role: 'tool', results: [{ id: '1', name: 'load_skill', output: loadSkill([templ], 't').output }] }
    expect(isLoaded(templ, [tool])).toBe(true)
  })

  it('treats a cut-down copy (compaction digest, squashed tool output) as not loaded', () => {
    const long = skill('long', { body: 'x'.repeat(3000) + ' end' })
    const full = invocationText(long, '', [])
    expect(isLoaded(long, [user(`[Summary of earlier conversation]\n- User: ${full.slice(0, 200)}`)])).toBe(false)
    const squashed: AgentMessage = { role: 'tool', results: [{ id: '1', name: 'load_skill', output: full.slice(0, 1000) + '\n…(output truncated: too long)' }] }
    expect(isLoaded(long, [squashed])).toBe(false)
    expect(invocationText(long, '', [squashed])).toBe(full)
  })

  it('does not confuse two skills whose folders share a prefix', () => {
    const a = skill('a', { dir: 'C:\\s\\a', body: 'same' })
    const ab = skill('ab', { dir: 'C:\\s\\ab', body: 'same' })
    expect(isLoaded(a, [user(invocationText(ab, '', []))])).toBe(false)
  })

  it('leaves other text alone', () => {
    expect(toModelText('hello', [skill('a')], [])).toBe('hello')
    expect(toModelText('/a go', [skill('a')], [])).toContain('ARGUMENTS: go')
  })

  it('load_skill refuses unknown and user-only skills', () => {
    const skills = [skill('a'), skill('manual', { modelInvocable: false })]
    expect(loadSkill(skills, 'a').isError).toBe(false)
    const refused = loadSkill(skills, 'manual')
    expect(refused.isError).toBe(true)
    expect(refused.output).toContain('Skills: a.')
  })
})
