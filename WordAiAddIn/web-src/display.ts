import type { ToolDisplayInfo } from '@officeai/app-shell'

export const WORD_TOOL_DISPLAY = {
  get_document_context: {
    label: { en: 'Read document', he: 'קרא מסמך' },
    description: { en: "Reads the document's structure and a short preview of its text.", he: 'קורא את מבנה המסמך ותצוגה מקדימה קצרה של הטקסט.' },
  },
  read_object: {
    label: { en: 'Read document objects', he: 'קריאת אובייקטים במסמך' },
    description: {
      en: 'Reads tables, charts, SmartArt, text boxes and shapes, content controls, headers/footers, footnotes, and tables of contents.',
      he: 'קורא טבלאות, תרשימים, SmartArt, תיבות טקסט וצורות, בקרי תוכן, כותרות עליונות/תחתונות, הערות שוליים ותוכן עניינים.',
    },
  },
  insert_content: {
    label: { en: 'Insert content', he: 'הוספת תוכן' },
    description: { en: 'Inserts new text or formatted content at a chosen position in the document.', he: 'מוסיף טקסט חדש או תוכן מעוצב במיקום נבחר במסמך.' },
  },
  edit_chart: {
    label: { en: 'Create or edit chart', he: 'יצירה/עריכת תרשים' },
    description: { en: 'Creates a new chart or edits an existing one, with categories and data series.', he: 'יוצר תרשים חדש או עורך תרשים קיים, כולל קטגוריות וסדרות נתונים.' },
  },
  find_text: {
    label: { en: 'Search document', he: 'חיפוש במסמך' },
    description: { en: 'Searches the document\'s text for a word or phrase.', he: 'מחפש מילה או ביטוי בטקסט המסמך.' },
  },
  get_headings: {
    label: { en: 'List headings', he: 'רשימת כותרות' },
    description: { en: 'Lists the document\'s headings, like the Navigation Pane.', he: 'מציג את כותרות המסמך, בדומה לחלונית הניווט.' },
  },
  read_blocks: {
    label: { en: 'Read paragraphs', he: 'קריאת פסקאות' },
    description: { en: 'Reads a range of paragraphs from the document.', he: 'קורא טווח פסקאות מהמסמך.' },
  },
  replace_blocks: {
    label: { en: 'Replace paragraphs', he: 'החלפת פסקאות' },
    description: { en: 'Replaces a range of paragraphs with new content.', he: 'מחליף טווח פסקאות בתוכן חדש.' },
  },
  add_image: {
    label: { en: 'Insert image', he: 'הוספת תמונה' },
    description: { en: 'Inserts an image from a local file into the document.', he: 'מוסיף תמונה מקובץ מקומי אל המסמך.' },
  },
  add_table: {
    label: { en: 'Insert table', he: 'הוספת טבלה' },
    description: { en: 'Adds a new table, optionally pre-filled with text.', he: 'מוסיף טבלה חדשה, ניתן למלא מראש בטקסט.' },
  },
  edit_table: {
    label: { en: 'Edit table', he: 'עריכת טבלה' },
    description: { en: 'Edits an existing table\'s cells, rows/columns, or style.', he: 'עורך את התאים, השורות/העמודות או העיצוב של טבלה קיימת.' },
  },
  add_smartart: {
    label: { en: 'Insert SmartArt', he: 'הוספת SmartArt' },
    description: { en: 'Adds a SmartArt diagram (list, process, cycle, hierarchy, pyramid, matrix, or Venn).', he: 'מוסיף דיאגרמת SmartArt (רשימה, תהליך, מעגל, היררכיה, פירמידה, מטריצה או ון).' },
  },
  edit_smartart: {
    label: { en: 'Edit SmartArt', he: 'עריכת SmartArt' },
    description: { en: 'Edits an existing SmartArt diagram\'s node text.', he: 'עורך את טקסט הצמתים של דיאגרמת SmartArt קיימת.' },
  },
  apply_commands: {
    label: { en: 'Edit and format', he: 'עריכה ועיצוב' },
    description: { en: 'Applies formatting and editing commands such as bold, headings, bullets, and find/replace.', he: 'מיישם פקודות עיצוב ועריכה כגון הדגשה, כותרות, תבליטים וחיפוש/החלפה.' },
  },
  add_comment: {
    label: { en: 'Add comment', he: 'הוספת הערה' },
    description: { en: 'Adds a comment anchored to text in the document, without changing its content.', he: 'מוסיף הערה מעוגנת לטקסט במסמך, מבלי לשנות את תוכנו.' },
  },
  edit_content_control: {
    label: { en: 'Edit content controls', he: 'עריכת בקרי תוכן' },
    description: { en: 'Adds, edits, or removes content controls such as text fields, dropdowns, checkboxes, and date pickers.', he: 'מוסיף, עורך או מסיר בקרי תוכן כגון שדות טקסט, רשימות נפתחות, תיבות סימון ובוחרי תאריך.' },
  },
  edit_header_footer: {
    label: { en: 'Edit header/footer', he: 'עריכת כותרת עליונה/תחתונה' },
    description: { en: 'Sets or clears the header or footer of a section.', he: 'קובע או מנקה את הכותרת העליונה או התחתונה של מקטע.' },
  },
  edit_note: {
    label: { en: 'Edit footnotes/endnotes', he: 'עריכת הערות שוליים/סיום' },
    description: { en: 'Adds, edits, or deletes footnotes and endnotes.', he: 'מוסיף, עורך או מוחק הערות שוליים והערות סיום.' },
  },
  edit_toc: {
    label: { en: 'Edit table of contents', he: 'עריכת תוכן עניינים' },
    description: { en: 'Adds, refreshes, adjusts, or removes a table of contents.', he: 'מוסיף, מרענן, מתאים או מסיר תוכן עניינים.' },
  },
  edit_textbox: {
    label: { en: 'Edit text boxes', he: 'עריכת תיבות טקסט' },
    description: { en: 'Adds, edits, moves, or deletes floating text boxes.', he: 'מוסיף, עורך, מזיז או מוחק תיבות טקסט צפות.' },
  },
  undo_last_action: {
    label: { en: 'Undo', he: 'ביטול' },
    description: { en: "Reverses the document's last action, like Ctrl+Z.", he: 'מבטל את הפעולה האחרונה במסמך, כמו Ctrl+Z.' },
  },
  redo_last_action: {
    label: { en: 'Redo', he: 'ביצוע חוזר' },
    description: { en: 'Re-applies the last undone action, like Ctrl+Y.', he: 'מבצע מחדש את הפעולה האחרונה שבוטלה, כמו Ctrl+Y.' },
  },
} satisfies Record<string, ToolDisplayInfo>

export const STARTERS = [
  { en: 'Summarize the key points of this document', he: 'סכם את הנקודות העיקריות במסמך' },
  { en: 'Polish the whole document for a more professional tone', he: 'לטש את כל המסמך לטון מקצועי יותר' },
  { en: 'Continue writing from where the document leaves off', he: 'המשך לכתוב מהיכן שהמסמך מסתיים' },
]
