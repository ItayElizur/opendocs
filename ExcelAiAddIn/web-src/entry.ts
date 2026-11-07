import { startAddIn, SHARED_TOOLS, SHARED_TOOL_DISPLAY } from '@officeai/app-shell'
import { ALL_TOOLS, READ_ONLY_TOOL_NAMES } from './tools'
import { EXCEL_TOOL_DISPLAY, STARTERS } from './display'
import { SYSTEM_PROMPT } from './prompt'

// Excel add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell (@officeai/app-shell), which holds everything else - the WebView2 bridge,
// settings, transport, chat-UI mount and AgentLoop event plumbing.

startAddIn({
  skillId: 'excel-tools',
  tools: [...ALL_TOOLS, ...SHARED_TOOLS],
  toolDisplay: { ...EXCEL_TOOL_DISPLAY, ...SHARED_TOOL_DISPLAY },
  systemPrompt: SYSTEM_PROMPT,
  starters: STARTERS,
  // Excel has no add_comment-equivalent tool yet, so Comment Only mode allows
  // the same read-only set as Read Only mode (documented gap) - leaving
  // commentOnlyExtraTools unset gives the same result as before this
  // migration, where both modes mapped to READ_ONLY_TOOLS.
  readOnlyTools: READ_ONLY_TOOL_NAMES,
  // FT-2 Task 4/5: the current cell selection is injected into per-turn
  // context (as an A1 address, not values - see ExcelAiAddIn/TaskPaneHost.cs's
  // OnSelectionChanged) and the scope-hint pill reads "Whole sheet" rather
  // than Word's "Whole document" when nothing is selected.
  useSelectionContext: true,
  scopeUnit: 'sheet',
})
