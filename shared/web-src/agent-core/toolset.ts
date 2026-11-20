import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'

/**
 * Everything one add-in gives the agent loop: its system prompt, its tools,
 * per-turn context, and the tool executor.
 */
export interface Toolset {
  systemPrompt: string
  tools: AgentToolDef[]
  /**
   * Fresh context sections attached to every user turn (e.g. document
   * skeleton + selection). Return '' when there is nothing to attach.
   */
  buildContext?(): string
  /**
   * signal: aborted when the user hits stop. Long-running tools (e.g.
   * generate_deck with internal LLM calls) should check signal.aborted in
   * their loops and stop promptly.
   */
  executeTool(call: AgentToolCall, signal?: AbortSignal): ToolExecution | Promise<ToolExecution>
}
