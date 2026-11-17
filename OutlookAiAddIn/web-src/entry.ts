import { startAddIn } from '@officeai/app-shell'
import { OUTLOOK_CONFIRM_ACTIONS } from './confirm'
import { ALL_OUTLOOK_TOOLS } from './tools'
import { OUTLOOK_TOOL_DISPLAY, STARTERS } from './display'
import { SYSTEM_PROMPT } from './prompt'

// Outlook add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell, which holds everything else.

startAddIn({
  skillId: 'outlook-tools',
  tools: ALL_OUTLOOK_TOOLS,
  toolDisplay: OUTLOOK_TOOL_DISPLAY,
  systemPrompt: SYSTEM_PROMPT,
  starters: STARTERS,
  readOnlyTools: [
    'search_emails',
    'get_email',
    'open_email',
    'get_attachment',
    'list_folders',
    'search_contacts',
    'list_events',
    'get_event',
    'find_meeting_slots',
    'list_color_categories',
    'list_tasks',
  ],
  // Tier 2 ("Draft only") on top of the read-only set above - every tool
  // that mutates the mailbox or opens a draft but never sends/creates
  // unreviewed. apply_search never mutates data but is here for the same
  // reason as OutlookTools.cs's DraftTierTools comment explains: it visibly
  // takes over the user's real Outlook window, which "Read only" is
  // supposed to never do. Must stay in sync with OutlookTools.cs's
  // DraftTierTools.
  commentOnlyExtraTools: [
    'mark_email_read',
    'flag_email_important',
    'move_email',
    'delete_email',
    'create_task',
    'update_task',
    'set_reminder',
    'set_email_reminder',
    'draft_email',
    'draft_event',
    'set_event_categories',
    'set_category_color',
    'set_event_availability',
    'apply_search',
    'undo_last_action',
    'redo_last_action',
  ],
  // Tier 3 ("Actions with approval"), on top of tier 2: everything that sends - a meeting response, an email, an
  // event - so this tier offers every tool, like Excel/PowerPoint's "Edits with approval". The panel asks before
  // each call as the administrator's policy.json sets. Full autonomy adds no tools; it only never asks.
  // Must stay in sync with OutlookTools.cs's ApprovalTierTools.
  trackChangesExtraTools: ['respond_meeting', 'send_email', 'send_event'],
  useSelectionContext: true,
  scopeUnit: 'mailbox',
  availableModes: ['readOnly', 'commentOnly', 'trackChanges', 'fullAutonomy'],
  defaultMode: 'commentOnly',
  modeOverrides: {
    commentOnly: {
      label: { en: 'Draft only', he: 'טיוטות בלבד' },
      description: { en: 'Drafts, flags, moves, and manages mail/tasks for your review - nothing sends.', he: 'מכין טיוטות, מסמן, מעביר ומנהל דואר/משימות לבדיקתך - שום דבר לא נשלח.' },
    },
    trackChanges: {
      label: { en: 'Actions with approval', he: 'פעולות באישור' },
      description: { en: "Everything, including sending mail and answering meetings; asks for your approval as your organization's policy sets.", he: 'הכול, כולל שליחת דואר ומענה לפגישות; מבקש את אישורך לפי מדיניות הארגון.' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Everything, including sending in your name, with no approval step.', he: 'הכול, כולל שליחה בשמך, ללא שלב אישור.' },
    },
  },
  autoSendTools: ['send_email', 'send_event'],
  confirmActions: OUTLOOK_CONFIRM_ACTIONS,
})
