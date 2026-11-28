import { startAddIn } from '@officeai/app-shell'
import { OUTLOOK_CONFIRM_ACTIONS } from './confirm'
import { ALL_OUTLOOK_TOOLS, ITEM_WINDOW_TOOLS } from './tools'
import { ITEM_STARTERS, OUTLOOK_TOOL_DISPLAY, STARTERS } from './display'
import { ITEM_SYSTEM_PROMPT, SYSTEM_PROMPT } from './prompt'

// Outlook add-in entry point: wires this app's tools (tools.ts), UI text (display.ts) and system prompt
// (prompt.ts) into the shared app shell, which holds everything else. The main window's pane and an item window's pane
// (a draft, or an appointment or meeting: ItemPaneHost.cs, which sets window.__opendocsPaneKind) get different modes
// and tools.

const paneKind = (globalThis as { __opendocsPaneKind?: string | null }).__opendocsPaneKind

if (paneKind !== 'mail' && paneKind !== 'appointment') {
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
  // Tier 3 ("Actions with approval"), on top of tier 2: respond_meeting, which notifies the organizer (plus a
  // permanent delete, gated by its argument in C#). The panel asks before each call as the administrator's
  // policy.json sets. Must stay in sync with OutlookTools.cs's ApprovalTierTools.
  trackChangesExtraTools: ['respond_meeting'],
  // send_email/send_event are deliberately in no list above: that alone confines them to Full autonomy, which shows
  // every tool and never asks. The reviewed way to send is a draft.
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
      description: { en: "Everything in Draft only, plus answering meetings and permanent deletes; asks for your approval as your organization's policy sets.", he: 'כל מה שיש בטיוטות בלבד, בתוספת מענה לפגישות ומחיקה לצמיתות; מבקש את אישורך לפי מדיניות הארגון.' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Everything above, plus sending mail and events immediately in your name, with no approval step.', he: 'כל מה שלמעלה, בתוספת שליחת דואר ואירועים מיידית בשמך, ללא שלב אישור.' },
    },
  },
  autoSendTools: ['send_email', 'send_event'],
  confirmActions: OUTLOOK_CONFIRM_ACTIONS,
})
} else {
  // An item window: the item's own tools and the mailbox reads, in two modes. Every mailbox action stays in the main
  // window. Must stay in sync with ItemPaneHost.cs's ItemReadTools.
  const itemReadOnlyTools = [
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
    'get_current_item',
    'read_main_chat',
  ]
  // "Edit this draft" / "Edit this meeting", on the CommentOnly slot: every change stays in the open window until the
  // user sends it. Must stay in sync with ItemPaneHost.cs's ItemEditTools.
  const itemEditExtraTools = [
    'edit_current_body',
    'set_current_fields',
  ]
  const offered = new Set([...itemReadOnlyTools, ...itemEditExtraTools])
  const itemOwn = new Set(ITEM_WINDOW_TOOLS.map((t) => t.name))
  const meeting = paneKind === 'appointment'
  startAddIn({
    skillId: 'outlook-item-tools',
    tools: [...ITEM_WINDOW_TOOLS, ...ALL_OUTLOOK_TOOLS.filter((t) => !itemOwn.has(t.name))].filter((t) => offered.has(t.name)),
    toolDisplay: OUTLOOK_TOOL_DISPLAY,
    systemPrompt: ITEM_SYSTEM_PROMPT,
    starters: ITEM_STARTERS[paneKind],
    readOnlyTools: itemReadOnlyTools,
    commentOnlyExtraTools: itemEditExtraTools,
    useSelectionContext: true,
    scopeUnit: 'mailbox',
    availableModes: ['readOnly', 'commentOnly'],
    defaultMode: 'commentOnly',
    modeOverrides: {
      readOnly: {
        label: { en: 'Read only', he: 'קריאה בלבד' },
        description: meeting
          ? { en: 'Reads this meeting and your mailbox, and suggests changes - changes nothing.', he: 'קורא את הפגישה ואת תיבת הדואר ומציע שינויים - לא משנה דבר.' }
          : { en: 'Reads this draft and your mailbox, and suggests changes - changes nothing.', he: 'קורא את הטיוטה ואת תיבת הדואר ומציע שינויים - לא משנה דבר.' },
      },
      commentOnly: meeting
        ? {
            label: { en: 'Edit this meeting', he: 'עריכת הפגישה' },
            description: { en: 'Changes this meeting in its window - attendees, times, text. Nothing is sent until you press Send.', he: 'משנה את הפגישה בחלון שלה - משתתפים, זמנים, טקסט. שום דבר לא נשלח עד שתלחץ על שליחה.' },
          }
        : {
            label: { en: 'Edit this draft', he: 'עריכת הטיוטה' },
            description: { en: 'Changes this draft in its window - text, recipients, subject. Nothing is sent until you press Send.', he: 'משנה את הטיוטה בחלון שלה - טקסט, נמענים, נושא. שום דבר לא נשלח עד שתלחץ על שליחה.' },
          },
    },
  })
}
