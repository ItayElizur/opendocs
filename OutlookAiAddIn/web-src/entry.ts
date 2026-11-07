import { startAddIn } from '@officeai/app-shell'
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
  // Tier 3 ("Automate approvals"), on top of tier 2 - respond_meeting
  // already auto-notifies the organizer via resp.Send(), so they get their
  // own tier rather than hiding in Draft only or Full autonomy. Must stay
  // in sync with OutlookTools.cs's ApprovalTierTools.
  trackChangesExtraTools: ['respond_meeting'],
  // send_email/send_event
  // are deliberately in neither list above -
  // that omission alone confines them to tier 4 (Full autonomy), which shows
  // every tool.
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
      label: { en: 'Automate approvals', he: 'אוטומציית אישורים' },
      description: { en: 'Everything in Draft only, plus auto-accepting/declining/tentatively-responding to meeting invites (notifies the organizer).', he: 'כל מה שיש בטיוטות בלבד, בתוספת אישור/דחייה/תגובה זמנית אוטומטיים של הזמנות לפגישה (מודיע למארגן).' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Everything above, plus sending emails and creating/sending/rescheduling calendar invites in your name.', he: 'כל מה שלמעלה, בתוספת שליחת הודעות ויצירה/שליחה/שינוי מועד של הזמנות יומן בשמך.' },
    },
  },
  autoSendTools: ['send_email', 'send_event'],
})
