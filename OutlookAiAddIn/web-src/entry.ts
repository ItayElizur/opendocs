import { startAddIn, type ConfirmRule } from '@officeai/app-shell'
import { ALL_OUTLOOK_TOOLS } from './tools'
import { OUTLOOK_TOOL_DISPLAY, STARTERS } from './display'
import { SYSTEM_PROMPT } from './prompt'

// Outlook add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell, which holds everything else.

// Asked before they run (Approve/Deny card; the user can turn this off in settings). The first matching rule wins,
// so the specific delete_email and send_event rules come before the general ones. `preview` asks the add-in for the item
// behind the call (type, subject, sender, date) and skips the card when the call would be refused; `fields` adds rows from
// the call's own arguments when there is no item yet. The body is never shown.
const text = (v: unknown): string => (Array.isArray(v) ? v.join(', ') : typeof v === 'string' ? v : '')
const confirmActions: ConfirmRule[] = [
  { name: 'delete_email', when: (i) => i.permanent === true, danger: true, preview: true,
    title: { en: 'Delete this message permanently? This cannot be undone.', he: 'למחוק את ההודעה לצמיתות? אי אפשר לבטל.' } },
  { name: 'delete_email', preview: true,
    title: { en: 'Move this message to Deleted Items?', he: 'להעביר את ההודעה לפריטים שנמחקו?' } },
  { name: 'send_email', danger: true, preview: true,
    fields: (i) => [{ key: 'to', value: text(i.to) }, { key: 'subject', value: text(i.subject) }],
    title: { en: 'Send this email now? A sent email cannot be taken back.', he: 'לשלוח את ההודעה עכשיו? אי אפשר להחזיר הודעה שנשלחה.' } },
  { name: 'send_event', when: (i) => i.action === 'cancel', danger: true, preview: true,
    title: { en: 'Cancel this event? Attendees may be notified.', he: 'לבטל את האירוע? ייתכן שהמשתתפים יקבלו הודעה.' } },
  { name: 'send_event', danger: true, preview: true,
    fields: (i) => (i.event_id ? [] : [{ key: 'subject', value: text(i.subject) }, { key: 'start', value: text(i.start) }, { key: 'end', value: text(i.end) }]),
    title: { en: 'Create or update this event now? Attendees may be notified.', he: 'ליצור או לעדכן את האירוע עכשיו? ייתכן שהמשתתפים יקבלו הודעה.' } },
  { name: 'respond_meeting', danger: true, preview: true,
    fields: (i) => [{ key: 'response', value: text(i.response) }],
    title: { en: 'Send this response to the organizer?', he: 'לשלוח את התגובה למארגן?' } },
]

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
  confirmActions,
})
