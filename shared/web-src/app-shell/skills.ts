import type { AgentMessage } from '@officeai/agent-core'

// Skills (Agent Skills format, read by OfficeAi.Shared/SkillCatalog.cs). Every piece of skill text the model sees is
// built here, so load_skill and a typed "/name" produce the same block, and "is it already loaded" is one check.

export interface Skill {
  name: string
  description: string
  argumentHint: string
  /** absolute base directory */
  dir: string
  body: string
  /** the skill's other files, relative '/' paths */
  files: string[]
  modelInvocable: boolean
  userInvocable: boolean
  source: 'machine' | 'user'
}

export const LOAD_SKILL = 'load_skill'
export const READ_SKILL_FILE = 'read_skill_file'

const ARGS_TOKEN = '$ARGUMENTS'

/** The system-prompt section listing the skills the model may load; '' when there are none. */
export function catalogPrompt(skills: readonly Skill[]): string {
  const listed = skills.filter((s) => s.modelInvocable)
  if (listed.length === 0) return ''
  return (
    '\n\nSkills: each skill below holds instructions for one kind of task. When the request matches a skill\'s description, ' +
    `call ${LOAD_SKILL} with its name before doing the task, then follow the instructions it returns. ` +
    `Read a file the instructions mention with ${READ_SKILL_FILE}. If a skill was already loaded in this conversation, do not load it again.\n` +
    listed.map((s) => `- ${s.name}: ${s.description}`).join('\n')
  )
}

/** The skill as the model reads it, starting with the base-directory line. May still hold $ARGUMENTS. */
export function skillBlock(skill: Skill): string {
  const files = skill.files.length
    ? `\n\nFiles in this skill (read one with ${READ_SKILL_FILE}; a file's full path is the base directory plus its path):\n` +
      skill.files.map((f) => `- ${f}`).join('\n')
    : ''
  return `Base directory for this skill: ${skill.dir}\n\n${skill.body}${files}`
}

/**
 * True when the history still holds this skill's full text. It reads the history itself, not a flag, because
 * compaction, the history trim and the stale-tool-output squash can each remove or cut a loaded skill; a cut copy
 * no longer matches, so the skill is sent in full again.
 */
export function isLoaded(skill: Skill, history: readonly AgentMessage[]): boolean {
  const parts = skillBlock(skill).split(ARGS_TOKEN)
  const texts = history.flatMap((m) => (m.role === 'tool' ? m.results.map((r) => r.output) : [m.role === 'user' ? m.text : '']))
  return texts.some((t) => containsInOrder(t, parts))
}

function containsInOrder(text: string, parts: readonly string[]): boolean {
  let at = 0
  for (const part of parts) {
    const i = text.indexOf(part, at)
    if (i < 0) return false
    at = i + part.length
  }
  return true
}

/** `/name args` naming a skill the user may call; null for anything else (the text is then sent as typed). */
export function parseInvocation(text: string, skills: readonly Skill[]): { skill: Skill; args: string } | null {
  const m = /^\/([a-z0-9]+(?:-[a-z0-9]+)*)(?:\s+([\s\S]*))?$/.exec(text.trim())
  if (!m) return null
  const skill = skills.find((s) => s.name === m[1] && s.userInvocable)
  return skill ? { skill, args: (m[2] ?? '').trim() } : null
}

/** What the model gets for a typed `/name args`: the full skill, or a short note when it is already loaded. */
export function invocationText(skill: Skill, args: string, history: readonly AgentMessage[]): string {
  const argsLine = args ? `\n\nARGUMENTS: ${args}` : ''
  if (isLoaded(skill, history)) {
    return `(Re-invocation of /${skill.name}: the skill instructions were previously loaded; the arguments below are new.)${argsLine}`
  }
  const block = skillBlock(skill)
  return block.includes(ARGS_TOKEN) ? block.split(ARGS_TOKEN).join(args) : block + argsLine
}

/** The model-facing text for a user message: a skill invocation expanded, anything else unchanged. */
export function toModelText(text: string, skills: readonly Skill[], history: readonly AgentMessage[]): string {
  const call = parseInvocation(text, skills)
  return call ? invocationText(call.skill, call.args, history) : text
}

/** load_skill's result for the model. */
export function loadSkill(skills: readonly Skill[], name: unknown): { output: string; isError: boolean } {
  const skill = skills.find((s) => s.name === name && s.modelInvocable)
  if (skill) return { output: skillBlock(skill).split(ARGS_TOKEN).join(''), isError: false }
  const names = skills.filter((s) => s.modelInvocable).map((s) => s.name)
  return { output: `There is no skill named '${String(name)}'. Skills: ${names.join(', ') || '(none)'}.`, isError: true }
}
