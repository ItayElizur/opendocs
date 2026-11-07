import { startAddIn, SHARED_TOOLS, SHARED_TOOL_DISPLAY } from '@officeai/app-shell'
import { ALL_WORD_TOOLS } from './tools'
import { WORD_TOOL_DISPLAY, STARTERS } from './display'
import { SYSTEM_PROMPT } from './prompt'

// Word add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell (@officeai/app-shell), which holds everything else - the WebView2 bridge,
// settings, transport, chat-UI mount and AgentLoop event plumbing.

startAddIn({
  skillId: 'word-tools',
  tools: [...ALL_WORD_TOOLS, ...SHARED_TOOLS],
  toolDisplay: { ...WORD_TOOL_DISPLAY, ...SHARED_TOOL_DISPLAY },
  systemPrompt: SYSTEM_PROMPT,
  starters: STARTERS,
  readOnlyTools: ['get_document_context', 'read_blocks', 'find_text', 'get_headings', 'read_object'],
  commentOnlyExtraTools: ['add_comment'],
  useSelectionContext: true,
})
