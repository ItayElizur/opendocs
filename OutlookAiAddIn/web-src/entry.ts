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
    name: 'apply_search',
    description:
      "Applies a search to the user's actual Outlook window - navigates to the folder and runs the search there, so the user sees the same results you found. Use after search_emails/list_emails once you know what's relevant; reuses the same query/date/sender filters.",
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text matched against subject and body (contains).' },
        folder: FOLDER,
        start_date: { type: 'string', description: 'Only messages on/after this date (YYYY-MM-DD).' },
        end_date: { type: 'string', description: 'Only messages on/before this date (YYYY-MM-DD).' },
        sender: { type: 'string', description: 'Sender email or display-name fragment.' },
        scope: {
          type: 'string',
          enum: ['current_folder', 'subfolders', 'mailbox', 'all_mailboxes'],
          description:
            'How far to search beyond the given folder. "current_folder" (default) - just that folder. "subfolders" - that folder and its subfolders. "mailbox" - every folder in the current mailbox. "all_mailboxes" - every account configured in Outlook.',
        },
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
      'Resolves a name or email fragment against Exchange via server-side ambiguous-name resolution (EWS ResolveName): your personal Contacts first, then the Global Address List. Returns {name, email} entries. On-prem Exchange only; returns a clear error if Exchange cannot be reached.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Name or email fragment (2+ characters recommended).' },
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
      "Finds open meeting times for you plus one or more attendees, using their Outlook free/busy. Ranked by how many people are free (so a best partial match still comes back if nobody is free for the whole group). Defaults to your mailbox's own configured work week/hours (read from Exchange; falls back to Sun-Thu 09:00-18:00 if that can't be read). Feed a returned slot to draft_event.",
    inputSchema: {
      type: 'object',
      properties: {
        attendees: { type: 'string', description: 'Comma-separated emails (or "Name <email>") besides yourself.' },
        duration_minutes: { type: 'number' },
        start_date: { type: 'string', description: 'Range start (YYYY-MM-DD). Defaults to the work-week rule above.' },
        end_date: { type: 'string', description: 'Range end (YYYY-MM-DD).' },
        start_hour: { type: 'number', description: "Earliest hour to consider. Defaults to your mailbox's configured work-day start (or 9 if that can't be read)." },
        end_hour: { type: 'number', description: "Latest hour, exclusive. Defaults to your mailbox's configured work-day end (or 18 if that can't be read)." },
        limit: { type: 'number', description: 'Max slots to return (default 5).' },
      },
      required: ['attendees', 'duration_minutes'],
    },
  },
  {
    name: 'list_color_categories',
    description:
      'Lists the color tags (Outlook "Categories") available to apply to events, mail, or tasks - each with its name and color. Use this to see valid names before calling set_event_categories, or valid colors before set_category_color.',
    inputSchema: { type: 'object', properties: {}, required: [] },
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
    name: 'set_event_categories',
    description:
      'Colors a calendar event with one or more color tags (Outlook "Categories"), shown as a colored block on the event. Pass names from list_color_categories, comma-separated for more than one; an empty/omitted categories clears all tags from the event. A name not yet in the master list is auto-added with an arbitrary color - call set_category_color first to control it.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'string' },
        categories: { type: 'string', description: 'Comma-separated color tag name(s), e.g. "Urgent, Travel". Omit or pass "" to clear.' },
      },
      required: ['event_id'],
    },
  },
  {
    name: 'set_category_color',
    description: 'Creates a new color tag, or changes an existing one\'s color, in the shared master list used by set_event_categories.',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Color tag name, e.g. "Urgent".' },
        color: {
          type: 'string',
          description: 'One of: None, Red, Orange, Peach, Yellow, Green, Teal, Olive, Blue, Purple, Maroon, Steel, Dark Steel, Gray, Dark Gray, Black, Dark Red, Dark Orange, Dark Peach, Dark Yellow, Dark Green, Dark Teal, Dark Olive, Dark Blue, Dark Purple, Dark Maroon.',
        },
      },
      required: ['name', 'color'],
    },
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
  {
    name: 'send_email',
    description:
      'Sends an email immediately - NO review window, no draft. Only available in Full autonomy. Prefer draft_email unless the user clearly wants this sent right now, with no chance to review it first.',
    inputSchema: {
      type: 'object',
      properties: { to: { type: 'string' }, subject: { type: 'string' }, body: { type: 'string' } },
      required: ['to'],
    },
  },
  {
    name: 'send_reply',
    description: 'Sends a reply (to the sender only) immediately - NO review window. Only available in Full autonomy. Prefer reply_email unless the user clearly wants this sent right now.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, body: { type: 'string' }, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'send_reply_all',
    description: 'Sends a reply-to-all immediately - NO review window. Only available in Full autonomy. Prefer reply_all_email unless the user clearly wants this sent right now.',
    inputSchema: { type: 'object', properties: { message_id: MESSAGE_ID, body: { type: 'string' }, folder: FOLDER }, required: ['message_id'] },
  },
  {
    name: 'send_forward',
    description: 'Forwards a message immediately - NO review window. Only available in Full autonomy. Prefer forward_email unless the user clearly wants this sent right now.',
    inputSchema: {
      type: 'object',
      properties: { message_id: MESSAGE_ID, to: { type: 'string' }, body: { type: 'string' }, folder: FOLDER },
      required: ['message_id', 'to'],
    },
  },
  {
    name: 'create_event',
    description:
      'Creates a calendar event immediately - NO review window. With attendees, sends the meeting invite right away (notifies them). Only available in Full autonomy. Prefer draft_event unless the user clearly wants this created/sent right now.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        start: { type: 'string', description: 'Date-time, e.g. "2026-09-01T14:00".' },
        end: { type: 'string', description: 'Date-time.' },
        location: { type: 'string' },
        body: { type: 'string' },
        required_attendees: { type: 'string', description: 'Comma-separated emails or "Name <email>". Presence of attendees sends the invite instead of just saving the event.' },
        optional_attendees: { type: 'string' },
      },
      required: ['start', 'end'],
    },
  },
]

const d = (en: string, he: string, den: string, dhe: string) => ({ label: { en, he }, description: { en: den, he: dhe } })

const OUTLOOK_TOOL_DISPLAY: Record<string, ReturnType<typeof d>> = {
  list_emails: d('List emails', 'רשימת הודעות', 'Lists recent messages from a folder.', 'מציג הודעות אחרונות מתיקייה.'),
  search_emails: d('Search emails', 'חיפוש הודעות', 'Searches a folder by text, date, or sender.', 'מחפש בתיקייה לפי טקסט, תאריך או שולח.'),
  apply_search: d('Show search in Outlook', 'הצגת חיפוש ב-Outlook', 'Applies the search to the Outlook window itself.', 'מיישם את החיפוש בחלון Outlook עצמו.'),
  get_email: d('Read email', 'קריאת הודעה', 'Reads one message in full, including its attachment list.', 'קורא הודעה אחת במלואה, כולל רשימת הקבצים המצורפים.'),
  get_attachment: d('Get attachment', 'קבלת קובץ מצורף', 'Saves an attachment and extracts text from documents.', 'שומר קובץ מצורף ומחלץ טקסט ממסמכים.'),
  list_folders: d('List folders', 'רשימת תיקיות', 'Lists the available mail folders.', 'מציג את תיקיות הדואר הזמינות.'),
  search_contacts: d('Search contacts', 'חיפוש אנשי קשר', 'Finds people by name or email.', 'מוצא אנשים לפי שם או דוא"ל.'),
  list_events: d('List calendar events', 'רשימת אירועים', 'Lists calendar events in a date range.', 'מציג אירועי יומן בטווח תאריכים.'),
  get_event: d('Read event', 'קריאת אירוע', 'Reads one calendar event in full.', 'קורא אירוע יומן אחד במלואו.'),
  find_meeting_slots: d('Find meeting times', 'מציאת זמני פגישה', 'Finds open times for you and the attendees, ranked by availability.', 'מוצא זמנים פנויים עבורך והמוזמנים, מדורגים לפי זמינות.'),
  list_color_categories: d('List color tags', 'רשימת תגיות צבע', 'Lists the available color tags and their colors.', 'מציג את תגיות הצבע הזמינות והצבעים שלהן.'),
  list_tasks: d('List tasks', 'רשימת משימות', 'Lists tasks and their due dates.', 'מציג משימות ותאריכי יעד.'),
  mark_email_read: d('Mark read', 'סימון כנקרא', 'Marks a message as read.', 'מסמן הודעה כנקראה.'),
  mark_email_unread: d('Mark unread', 'סימון כלא נקרא', 'Marks a message as unread.', 'מסמן הודעה כלא נקראה.'),
  flag_email_important: d('Flag importance', 'סימון חשיבות', 'Sets a message to High or Normal importance.', 'מגדיר חשיבות גבוהה או רגילה להודעה.'),
  move_email: d('Move email', 'העברת הודעה', 'Moves a message to another folder.', 'מעביר הודעה לתיקייה אחרת.'),
  delete_email: d('Delete email', 'מחיקת הודעה', 'Moves a message to Deleted Items.', 'מעביר הודעה לפריטים שנמחקו.'),
  accept_meeting: d('Accept meeting', 'אישור פגישה', 'Accepts a meeting invitation.', 'מאשר הזמנה לפגישה.'),
  decline_meeting: d('Decline meeting', 'דחיית פגישה', 'Declines a meeting invitation.', 'דוחה הזמנה לפגישה.'),
  set_event_categories: d('Color event', 'צביעת אירוע', 'Applies or clears color tags on a calendar event.', 'מחיל או מנקה תגיות צבע על אירוע יומן.'),
  set_category_color: d('Set tag color', 'הגדרת צבע תגית', 'Creates or recolors a color tag.', 'יוצר או משנה צבע של תגית.'),
  create_task: d('Create task', 'יצירת משימה', 'Creates a task with an optional due date and reminder.', 'יוצר משימה עם תאריך יעד ותזכורת אופציונליים.'),
  update_task: d('Update task', 'עדכון משימה', 'Updates or completes an existing task.', 'מעדכן או משלים משימה קיימת.'),
  set_reminder: d('Set reminder', 'הגדרת תזכורת', 'Sets a reminder on an appointment or task.', 'מגדיר תזכורת לפגישה או משימה.'),
  set_email_reminder: d('Flag with reminder', 'סימון עם תזכורת', 'Flags an email for follow-up with a reminder.', 'מסמן הודעה למעקב עם תזכורת.'),
  draft_email: d('Draft email', 'טיוטת הודעה', 'Opens a pre-filled compose window to review and send.', 'פותח חלון חיבור מלא מראש לבדיקה ושליחה.'),
  reply_email: d('Draft reply', 'טיוטת תשובה', 'Opens a pre-filled reply to review and send.', 'פותח תשובה מלאה מראש לבדיקה ושליחה.'),
  reply_all_email: d('Draft reply all', 'טיוטת תשובה לכולם', 'Opens a pre-filled reply-to-all to review and send.', 'פותח תשובה-לכולם מלאה מראש לבדיקה ושליחה.'),
  forward_email: d('Draft forward', 'טיוטת העברה', 'Opens a pre-filled forward to review and send.', 'פותח העברה מלאה מראש לבדיקה ושליחה.'),
  draft_event: d('Draft event', 'טיוטת אירוע', 'Opens a pre-filled appointment/meeting to review and send.', 'פותח פגישה/אירוע מלא מראש לבדיקה ושליחה.'),
  send_email: d('Send email (auto-send)', 'שליחת הודעה (שליחה אוטומטית)', 'Sends an email immediately, no review window.', 'שולח הודעה מיידית, ללא חלון בדיקה.'),
  send_reply: d('Send reply (auto-send)', 'שליחת תשובה (שליחה אוטומטית)', 'Sends a reply immediately, no review window.', 'שולח תשובה מיידית, ללא חלון בדיקה.'),
  send_reply_all: d('Send reply all (auto-send)', 'שליחת תשובה לכולם (שליחה אוטומטית)', 'Sends a reply-to-all immediately, no review window.', 'שולח תשובה-לכולם מיידית, ללא חלון בדיקה.'),
  send_forward: d('Send forward (auto-send)', 'שליחת העברה (שליחה אוטומטית)', 'Forwards a message immediately, no review window.', 'מעביר הודעה מיידית, ללא חלון בדיקה.'),
  create_event: d('Create event (auto-send)', 'יצירת אירוע (שליחה אוטומטית)', 'Creates/sends a calendar event immediately, no review window.', 'יוצר/שולח אירוע יומן מיידית, ללא חלון בדיקה.'),
}

startAddIn({
  skillId: 'outlook-tools',
  tools: ALL_OUTLOOK_TOOLS,
  toolDisplay: OUTLOOK_TOOL_DISPLAY,
  systemPrompt:
    'You are an AI assistant embedded in Microsoft Outlook via the Airchat Office add-in. You work from the main Outlook window (Explorer). ' +
    'You can read and search mail, read attachments, triage messages (mark read/unread, flag importance, move, delete), manage the calendar (list/read events, accept/decline invitations, color events with tags via list_color_categories/set_event_categories/set_category_color), ' +
    'manage tasks and reminders, and draft replies/forwards/new mail and calendar events. ' +
    'Drafting tools (draft_email, reply_email, reply_all_email, forward_email, draft_event) open a normal Outlook compose or appointment window pre-filled - they never send or create directly; the user reviews and sends. ' +
    'send_email/send_reply/send_reply_all/send_forward/create_event are different: they send or create IMMEDIATELY, with no review window at all - only available in Full autonomy, and only worth using when the user has clearly asked for something to go out right now with no chance to check it first. Default to the drafting tools otherwise. ' +
    'message_id / event_id / task_id values are Outlook EntryIDs. When the user has one or more messages selected, that selection (with its message_id) is in your context - prefer it over searching. ' +
    'Prefer list_emails / search_emails / list_tasks (fast, server-side) over reading items one by one. ' +
    "Once you've found the relevant messages, apply_search can show the same results in the user's own Outlook window instead of only listing them in chat. " +
    "Your available tools depend on the user's editing mode, from least to most permissive: Read only (read/search only) -> Draft only (also triage, tasks, reminders, and drafting replies/forwards/new mail/events) -> Automate approvals (also auto-accept/decline meeting invitations, which notifies the organizer) -> Full autonomy (also send_email/send_reply/send_reply_all/send_forward/create_event, which send/create immediately).",
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
    'mark_email_unread',
    'flag_email_important',
    'move_email',
    'delete_email',
    'create_task',
    'update_task',
    'set_reminder',
    'set_email_reminder',
    'draft_email',
    'reply_email',
    'reply_all_email',
    'forward_email',
    'draft_event',
    'set_event_categories',
    'set_category_color',
    'apply_search',
  ],
  // Tier 3 ("Automate approvals"), on top of tier 2 - accept/decline
  // already auto-notify the organizer via resp.Send(), so they get their
  // own tier rather than hiding in Draft only or Full autonomy. Must stay
  // in sync with OutlookTools.cs's ApprovalTierTools.
  trackChangesExtraTools: ['accept_meeting', 'decline_meeting'],
  // send_email/send_reply/send_reply_all/send_forward/create_event are
  // deliberately in neither list above - that omission alone confines them
  // to tier 4 (Full autonomy), which shows every tool.
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
      description: { en: 'Everything in Draft only, plus auto-accepting/declining meeting invites (notifies the organizer).', he: 'כל מה שיש בטיוטות בלבד, בתוספת אישור/דחייה אוטומטיים של הזמנות לפגישה (מודיע למארגן).' },
    },
    fullAutonomy: {
      label: { en: 'Full autonomy', he: 'אוטונומיה מלאה' },
      description: { en: 'Everything above, plus sending emails and creating/sending calendar invites in your name.', he: 'כל מה שלמעלה, בתוספת שליחת הודעות ויצירה/שליחה של הזמנות יומן בשמך.' },
    },
  },
  autoSendTools: ['send_email', 'send_reply', 'send_reply_all', 'send_forward', 'create_event'],
})
