import { startAddIn } from '@officeai/app-shell'

// Outlook add-in: Explorer-only, per-mailbox chat. Tool schemas mirror
// C:\dev\mcp-outlook as closely as the COM object model allows. Everything
// else (WebView2 bridge, settings, transport, chat-UI, AgentLoop) is shared.

const FOLDER = {
  type: 'string',
  description: 'Folder name: "inbox" (default), "sent", "drafts", "deleted", "junk", or any custom folder name from list_folders.',
}

const MESSAGE_ID = { type: 'string', description: 'Outlook EntryID from list_emails/search_emails or the current selection.' }

const ALL_OUTLOOK_TOOLS = [
  {
    name: 'list_emails',
    description:
      'Lists recent emails (newest first) from a folder using Outlook\'s fast table API. unread_only limits to unread. Returns message_id (EntryID), subject, sender, received time, unread flag, and has_attachments.',
    inputSchema: {
      type: 'object',
      properties: {
        folder: FOLDER,
        limit: { type: 'number', description: 'Max messages (default 20).' },
        unread_only: { type: 'boolean' },
      },
      required: [],
    },
  },
  {
    name: 'search_emails',
    description:
      'Searches a folder server-side (Outlook Restrict/DASL) by text (subject+body), date range, and/or sender; falls back to a capped scan only if the filter cannot be pushed down. recipient is matched client-side. Newest first.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text matched against subject and body (contains).' },
        folder: FOLDER,
        start_date: { type: 'string', description: 'Only messages on/after this date (YYYY-MM-DD).' },
        end_date: { type: 'string', description: 'Only messages on/before this date (YYYY-MM-DD).' },
        sender: { type: 'string', description: 'Sender email or display-name fragment.' },
        recipient: { type: 'string', description: 'To/CC email or name fragment (client-side filter, capped scan).' },
        limit: { type: 'number', description: 'Max messages (default 20).' },
        unread_only: { type: 'boolean' },
      },
      required: [],
    },
  },
  {
    name: 'get_email',
    description:
      'Full message: body (capped), To/CC recipients, conversation id/topic, importance, and an attachments list with each attachment\'s 1-based index, name, type, and size (feed the index to get_attachment).',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'get_attachment',
    description:
      'Saves one attachment to a local file and returns its path. For text files and Office documents (.docx/.xlsx/.pptx) it also returns extracted_text. PDFs and images return the path only. OLE and linked attachments cannot be fetched.',
    inputSchema: {
      type: 'object',
      properties: {
        message_id: MESSAGE_ID,
        attachment_index: { type: 'number', description: '1-based index from get_email\'s attachments list.' },
        folder: FOLDER,
        save_dir: { type: 'string', description: 'Optional local directory to save into.' },
      },
      required: ['message_id', 'attachment_index'],
    },
  },
  {
    name: 'list_folders',
    description: 'Lists mail folders (with item and unread counts) across all stores, for use as folder / destination arguments.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'search_contacts',
    description:
      'Searches contacts by name or email fragment. By default it covers the Global Address List plus every contact folder across every store (custom folders, shared mailboxes, subfolders). Pass folder to restrict to one named contact folder. Returns {name, email} entries.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string' },
        folder: { type: 'string', description: 'Optional: limit the search to this contact folder by name (also skips the GAL).' },
        limit: { type: 'number', description: 'Default 10.' },
      },
      required: ['query'],
    },
  },
  {
    name: 'list_events',
    description:
      'Lists calendar events in a date range (expands recurring meetings). Returns event_id, subject, start/end, location, organizer, and your response status. Recurring instances share the master event_id.',
    inputSchema: {
      type: 'object',
      properties: {
        start_date: { type: 'string', description: 'Range start, inclusive (YYYY-MM-DD). Default today.' },
        end_date: { type: 'string', description: 'Range end, inclusive (YYYY-MM-DD). Default +7 days.' },
        limit: { type: 'number', description: 'Default 50.' },
      },
      required: [],
    },
  },
  {
    name: 'get_event',
    description: 'Full calendar event: body, required/optional attendees, location, organizer, response status.',
    inputSchema: { type: 'object', properties: { event_id: { type: 'string' } }, required: ['event_id'] },
  },
  {
    name: 'find_meeting_slots',
    description:
      'Finds open meeting times for you plus one or more attendees, using their Outlook free/busy. Ranked by how many people are free (so a best partial match still comes back if nobody is free for the whole group). Defaults to this Sunday-Thursday work week, 09:00-18:00. Feed a returned slot to draft_event.',
    inputSchema: {
      type: 'object',
      properties: {
        attendees: { type: 'string', description: 'Comma-separated emails (or "Name <email>") besides yourself.' },
        duration_minutes: { type: 'number' },
        start_date: { type: 'string', description: 'Range start (YYYY-MM-DD). Defaults to the work-week rule above.' },
        end_date: { type: 'string', description: 'Range end (YYYY-MM-DD).' },
        start_hour: { type: 'number', description: 'Earliest hour to consider (default 9).' },
        end_hour: { type: 'number', description: 'Latest hour, exclusive (default 18).' },
        limit: { type: 'number', description: 'Max slots to return (default 5).' },
      },
      required: ['attendees', 'duration_minutes'],
    },
  },
  {
    name: 'list_tasks',
    description: 'Lists tasks (open only by default) via Outlook\'s table API. Returns task_id, subject, due/start dates, status, percent complete.',
    inputSchema: {
      type: 'object',
      properties: { limit: { type: 'number', description: 'Default 50.' }, include_completed: { type: 'boolean' } },
      required: [],
    },
  },
  {
    name: 'mark_email_read',
    description: 'Marks a message as read.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'mark_email_unread',
    description: 'Marks a message as unread.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'flag_email_important',
    description: 'Sets a message\'s importance to High (important:true, default) or Normal (important:false).',
    inputSchema: {
      type: 'object',
      properties: { message_id: MESSAGE_ID, important: { type: 'boolean' }, folder: FOLDER },
      required: ['message_id'],
    },
  },
  {
    name: 'move_email',
    description: 'Moves a message to another folder. Returns the NEW message_id (EntryID changes on move).',
    inputSchema: {
      type: 'object',
      properties: { message_id: MESSAGE_ID, destination: { type: 'string', description: 'Target folder name (see list_folders).' }, folder: FOLDER },
      required: ['message_id', 'destination'],
    },
  },
  {
    name: 'delete_email',
    description: 'Moves a message to Deleted Items (permanent:true also removes it from there).',
    inputSchema: {
      type: 'object',
      properties: { message_id: MESSAGE_ID, permanent: { type: 'boolean' }, folder: FOLDER },
      required: ['message_id'],
    },
  },
  {
    name: 'accept_meeting',
    description: 'Accepts a meeting invitation and notifies the organizer.',
    inputSchema: { type: 'object', properties: { event_id: { type: 'string' } }, required: ['event_id'] },
  },
  {
    name: 'decline_meeting',
    description: 'Declines a meeting invitation and notifies the organizer.',
    inputSchema: { type: 'object', properties: { event_id: { type: 'string' } }, required: ['event_id'] },
  },
  {
    name: 'create_task',
    description: 'Creates a task in the default Tasks folder.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        body: { type: 'string' },
        due_date: { type: 'string', description: 'YYYY-MM-DD or a full date-time.' },
        start_date: { type: 'string' },
        reminder_time: { type: 'string', description: 'Date-time for a reminder.' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'] },
      },
      required: ['subject'],
    },
  },
  {
    name: 'update_task',
    description: 'Updates an existing task. Only the fields you pass change. mark_complete:true completes it.',
    inputSchema: {
      type: 'object',
      properties: {
        task_id: { type: 'string' },
        subject: { type: 'string' },
        due_date: { type: 'string' },
        start_date: { type: 'string' },
        status: { type: 'string', enum: ['notStarted', 'inProgress', 'waiting', 'deferred', 'complete'] },
        percent_complete: { type: 'number' },
        mark_complete: { type: 'boolean' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'set_reminder',
    description: 'Sets (or clears) a reminder on an appointment or task, by its EntryID.',
    inputSchema: {
      type: 'object',
      properties: {
        item_id: { type: 'string', description: 'EntryID of an appointment or task.' },
        reminder_time: { type: 'string', description: 'Date-time. Required unless clear:true.' },
        clear: { type: 'boolean' },
      },
      required: ['item_id'],
    },
  },
  {
    name: 'set_email_reminder',
    description: 'Flags an email for follow-up with an optional due date and reminder time.',
    inputSchema: {
      type: 'object',
      properties: {
        message_id: MESSAGE_ID,
        due_date: { type: 'string', description: 'YYYY-MM-DD or date-time.' },
        reminder_time: { type: 'string', description: 'Date-time for the reminder pop-up.' },
        mark_interval: { type: 'string', enum: ['today', 'tomorrow', 'thisWeek', 'nextWeek', 'noDate'] },
        folder: FOLDER,
      },
      required: ['message_id'],
    },
  },
  {
    name: 'draft_email',
    description: 'Opens a new compose window in Outlook, pre-filled. The user reviews and sends it - this never sends directly.',
    inputSchema: {
      type: 'object',
      properties: { to: { type: 'string' }, subject: { type: 'string' }, body: { type: 'string' } },
      required: [],
    },
  },
  {
    name: 'reply_email',
    description: 'Opens a reply (to the sender only) in Outlook, pre-filled with your text above the quoted original. The user sends it.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, body: { type: 'string' }, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'reply_all_email',
    description: 'Opens a reply-to-all in Outlook, pre-filled. The user reviews the full recipient list and sends it.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, body: { type: 'string' }, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'forward_email',
    description: 'Opens a forward in Outlook, pre-filled. The user sends it.',
    inputSchema: {
      type: 'object',
      properties: { message_id: MESSAGE_ID, to: { type: 'string' }, body: { type: 'string' }, folder: FOLDER },
      required: ['message_id'],
    },
  },
  {
    name: 'draft_event',
    description:
      'Opens a new appointment/meeting window in Outlook, pre-filled. With attendees it becomes a meeting request. The user reviews and sends/saves it.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        start: { type: 'string', description: 'Date-time, e.g. "2026-09-01T14:00".' },
        end: { type: 'string', description: 'Date-time.' },
        location: { type: 'string' },
        body: { type: 'string' },
        required_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>".' },
        optional_attendees: { type: 'string' },
      },
      required: [],
    },
  },
]

const d = (en: string, he: string, den: string, dhe: string) => ({ label: { en, he }, description: { en: den, he: dhe } })

const OUTLOOK_TOOL_DISPLAY: Record<string, ReturnType<typeof d>> = {
  list_emails: d('List emails', 'רשימת הודעות', 'Lists recent messages from a folder.', 'מציג הודעות אחרונות מתיקייה.'),
  search_emails: d('Search emails', 'חיפוש הודעות', 'Searches a folder by text, date, or sender.', 'מחפש בתיקייה לפי טקסט, תאריך או שולח.'),
  get_email: d('Read email', 'קריאת הודעה', 'Reads one message in full, including its attachment list.', 'קורא הודעה אחת במלואה, כולל רשימת הקבצים המצורפים.'),
  get_attachment: d('Get attachment', 'קבלת קובץ מצורף', 'Saves an attachment and extracts text from documents.', 'שומר קובץ מצורף ומחלץ טקסט ממסמכים.'),
  list_folders: d('List folders', 'רשימת תיקיות', 'Lists the available mail folders.', 'מציג את תיקיות הדואר הזמינות.'),
  search_contacts: d('Search contacts', 'חיפוש אנשי קשר', 'Finds people by name or email.', 'מוצא אנשים לפי שם או דוא"ל.'),
  list_events: d('List calendar events', 'רשימת אירועים', 'Lists calendar events in a date range.', 'מציג אירועי יומן בטווח תאריכים.'),
  get_event: d('Read event', 'קריאת אירוע', 'Reads one calendar event in full.', 'קורא אירוע יומן אחד במלואו.'),
  find_meeting_slots: d('Find meeting times', 'מציאת זמני פגישה', 'Finds open times for you and the attendees, ranked by availability.', 'מוצא זמנים פנויים עבורך והמוזמנים, מדורגים לפי זמינות.'),
  list_tasks: d('List tasks', 'רשימת משימות', 'Lists tasks and their due dates.', 'מציג משימות ותאריכי יעד.'),
  mark_email_read: d('Mark read', 'סימון כנקרא', 'Marks a message as read.', 'מסמן הודעה כנקראה.'),
  mark_email_unread: d('Mark unread', 'סימון כלא נקרא', 'Marks a message as unread.', 'מסמן הודעה כלא נקראה.'),
  flag_email_important: d('Flag importance', 'סימון חשיבות', 'Sets a message to High or Normal importance.', 'מגדיר חשיבות גבוהה או רגילה להודעה.'),
  move_email: d('Move email', 'העברת הודעה', 'Moves a message to another folder.', 'מעביר הודעה לתיקייה אחרת.'),
  delete_email: d('Delete email', 'מחיקת הודעה', 'Moves a message to Deleted Items.', 'מעביר הודעה לפריטים שנמחקו.'),
  accept_meeting: d('Accept meeting', 'אישור פגישה', 'Accepts a meeting invitation.', 'מאשר הזמנה לפגישה.'),
  decline_meeting: d('Decline meeting', 'דחיית פגישה', 'Declines a meeting invitation.', 'דוחה הזמנה לפגישה.'),
  create_task: d('Create task', 'יצירת משימה', 'Creates a task with an optional due date and reminder.', 'יוצר משימה עם תאריך יעד ותזכורת אופציונליים.'),
  update_task: d('Update task', 'עדכון משימה', 'Updates or completes an existing task.', 'מעדכן או משלים משימה קיימת.'),
  set_reminder: d('Set reminder', 'הגדרת תזכורת', 'Sets a reminder on an appointment or task.', 'מגדיר תזכורת לפגישה או משימה.'),
  set_email_reminder: d('Flag with reminder', 'סימון עם תזכורת', 'Flags an email for follow-up with a reminder.', 'מסמן הודעה למעקב עם תזכורת.'),
  draft_email: d('Draft email', 'טיוטת הודעה', 'Opens a pre-filled compose window to review and send.', 'פותח חלון חיבור מלא מראש לבדיקה ושליחה.'),
  reply_email: d('Draft reply', 'טיוטת תשובה', 'Opens a pre-filled reply to review and send.', 'פותח תשובה מלאה מראש לבדיקה ושליחה.'),
  reply_all_email: d('Draft reply all', 'טיוטת תשובה לכולם', 'Opens a pre-filled reply-to-all to review and send.', 'פותח תשובה-לכולם מלאה מראש לבדיקה ושליחה.'),
  forward_email: d('Draft forward', 'טיוטת העברה', 'Opens a pre-filled forward to review and send.', 'פותח העברה מלאה מראש לבדיקה ושליחה.'),
  draft_event: d('Draft event', 'טיוטת אירוע', 'Opens a pre-filled appointment/meeting to review and send.', 'פותח פגישה/אירוע מלא מראש לבדיקה ושליחה.'),
}

startAddIn({
  skillId: 'outlook-tools',
  tools: ALL_OUTLOOK_TOOLS,
  toolDisplay: OUTLOOK_TOOL_DISPLAY,
  systemPrompt:
    'You are an AI assistant embedded in Microsoft Outlook via the Airchat Office add-in. You work from the main Outlook window (Explorer). ' +
    'You can read and search mail, read attachments, triage messages (mark read/unread, flag importance, move, delete), manage the calendar (list/read events, accept/decline invitations), ' +
    'manage tasks and reminders, and draft replies/forwards/new mail and calendar events. ' +
    'Drafting tools open a normal Outlook compose or appointment window pre-filled - you never send mail or create events directly; the user reviews and sends. ' +
    'message_id / event_id / task_id values are Outlook EntryIDs. When the user has one or more messages selected, that selection (with its message_id) is in your context - prefer it over searching. ' +
    'Prefer list_emails / search_emails / list_tasks (fast, server-side) over reading items one by one. ' +
    "Your available tools depend on the user's editing mode: in Read only you can read and search but not change anything; switch to Full autonomy for triage, drafts, tasks, reminders, and invitation responses.",
  starters: [
    { en: 'Summarize my unread emails', he: 'סכם את ההודעות שלא קראתי' },
    { en: 'Draft a reply to the selected email', he: 'נסח תשובה להודעה שנבחרה' },
    { en: "What's on my calendar this week?", he: 'מה יש ביומן שלי השבוע?' },
  ],
  readOnlyTools: [
    'list_emails',
    'search_emails',
    'get_email',
    'get_attachment',
    'list_folders',
    'search_contacts',
    'list_events',
    'get_event',
    'find_meeting_slots',
    'list_tasks',
  ],
  useSelectionContext: true,
  scopeUnit: 'mailbox',
  availableModes: ['readOnly', 'fullAutonomy'],
})
