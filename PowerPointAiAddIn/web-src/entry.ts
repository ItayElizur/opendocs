import { startAddIn, SHARED_TOOLS, SHARED_TOOL_DISPLAY } from '@officeai/app-shell'
import { READER_TOOLS, ALL_TOOLS } from './tools'
import { POWERPOINT_TOOL_DISPLAY, STARTERS } from './display'
import { SYSTEM_PROMPT } from './prompt'

// PowerPoint add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell (@officeai/app-shell), which holds everything else - the WebView2 bridge,
// settings, transport, chat-UI mount and AgentLoop event plumbing.

startAddIn({
  skillId: 'powerpoint-tools',
  tools: [...ALL_TOOLS, ...SHARED_TOOLS],
  toolDisplay: { ...POWERPOINT_TOOL_DISPLAY, ...SHARED_TOOL_DISPLAY },
  systemPrompt: SYSTEM_PROMPT,
  starters: STARTERS,
  readOnlyTools: READER_TOOLS.map((t) => t.name),
  // FT-2 Task 4/5: the current slide/shape/text selection is injected into
  // per-turn context as slideIndex/shapeIndex (see PowerPointAiAddIn/
  // TaskPaneHost.cs's OnSelectionChanged) and the scope-hint pill reads
  // "Whole deck" rather than Word's "Whole document" when nothing is selected.
  useSelectionContext: true,
  scopeUnit: 'deck',
})
