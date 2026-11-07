import type { ToolDisplayInfo } from '@officeai/app-shell'

export const POWERPOINT_TOOL_DISPLAY = {
  get_deck_context: {
    label: { en: 'Read deck outline', he: 'קרא תקציר מצגת' },
    description: { en: 'Reads a one-line summary of every slide in the presentation.', he: 'קורא תקציר שורה אחת לכל שקופית במצגת.' },
  },
  read_slide: {
    label: { en: 'Read slide', he: 'קריאת שקופית' },
    description: { en: 'Reads the full text of every shape on one slide.', he: 'קורא את הטקסט המלא של כל האובייקטים בשקופית.' },
  },
  find_text: {
    label: { en: 'Search deck', he: 'חיפוש במצגת' },
    description: { en: 'Searches every slide\'s text and speaker notes for a word or phrase.', he: 'מחפש מילה או ביטוי בטקסט של כל השקופיות ובהערות הדובר.' },
  },
  replace_text: {
    label: { en: 'Find and replace', he: 'חיפוש והחלפה' },
    description: { en: 'Replaces every occurrence of a word or phrase across the whole deck.', he: 'מחליף כל מופע של מילה או ביטוי בכל המצגת.' },
  },
  set_element_text: {
    label: { en: 'Edit shape text', he: 'עריכת טקסט באובייקט' },
    description: { en: 'Replaces the text content of a shape.', he: 'מחליף את תוכן הטקסט של אובייקט.' },
  },
  format_element: {
    label: { en: 'Format shape', he: 'עיצוב אובייקט' },
    description: { en: 'Changes a shape\'s text style, fill, outline, or opacity.', he: 'משנה את סגנון הטקסט, המילוי, המתאר או השקיפות של אובייקט.' },
  },
  set_element_transform: {
    label: { en: 'Move/resize/reorder shape', he: 'הזזה/שינוי גודל/סדר של אובייקט' },
    description: { en: 'Moves, resizes, or rotates a shape, or changes its stacking order.', he: 'מזיז, משנה גודל או מסובב אובייקט, או משנה את סדר הערימה שלו.' },
  },
  add_shape: {
    label: { en: 'Add shape or text box', he: 'הוספת צורה או תיבת טקסט' },
    description: { en: 'Creates a new shape (such as a rectangle, ellipse, or arrow) or a text box.', he: 'יוצר צורה חדשה (כגון מלבן, אליפסה או חץ) או תיבת טקסט.' },
  },
  delete_element: {
    label: { en: 'Delete shape', he: 'מחיקת אובייקט' },
    description: { en: 'Deletes a shape from a slide.', he: 'מוחק אובייקט משקופית.' },
  },
  copy_element: {
    label: { en: 'Copy or move shape', he: 'העתקה או העברה של אובייקט' },
    description: { en: 'Copies a shape on the same slide or to another slide, or moves it to another slide, with full native formatting.', he: 'מעתיק אובייקט באותה שקופית או לשקופית אחרת, או מעביר אותו לשקופית אחרת, כולל כל העיצוב המקורי.' },
  },
  manage_slide: {
    label: { en: 'Manage slide', he: 'ניהול שקופית' },
    description: { en: 'Adds, duplicates, moves, or deletes a slide.', he: 'מוסיף, משכפל, מעביר או מוחק שקופית.' },
  },
  set_slide: {
    label: { en: 'Update slide', he: 'עדכון שקופית' },
    description: { en: "Changes a slide's layout, background, transition, or speaker notes.", he: 'משנה את פריסת השקופית, הרקע, המעבר או הערות הדובר.' },
  },
  edit_animation: {
    label: { en: 'Animation', he: 'אנימציה' },
    description: { en: 'Adds an entrance or exit animation to a shape, or deletes, retimes, or reorders an existing one.', he: 'מוסיף אנימציית כניסה או יציאה לאובייקט, או מוחק, משנה תזמון או מסדר מחדש אנימציה קיימת.' },
  },
  copy_element_style: {
    label: { en: 'Copy shape formatting', he: 'העתקת עיצוב אובייקט' },
    description: { en: 'Copies text, fill, and outline formatting from one shape onto others (like a format painter).', he: 'מעתיק עיצוב טקסט, מילוי ומתאר מאובייקט אחד לאובייקטים אחרים (כמו מכחול עיצוב).' },
  },
  group_element: {
    label: { en: 'Group or ungroup shapes', he: 'קיבוץ או פירוק קבוצת אובייקטים' },
    description: { en: 'Combines shapes into a group, or breaks a group into its parts.', he: 'מאחד אובייקטים לקבוצה, או מפרק קבוצה לחלקיה.' },
  },
  edit_table: {
    label: { en: 'Table', he: 'טבלה' },
    description: { en: 'Adds a table, or edits a table\'s cells, rows and columns, or styling.', he: 'מוסיף טבלה, או עורך את התאים, השורות והעמודות שלה או את העיצוב.' },
  },
  edit_chart: {
    label: { en: 'Chart', he: 'תרשים' },
    description: { en: 'Adds a native, editable chart, or modifies an existing chart\'s type, title, legend, or labels.', he: 'מוסיף תרשים מובנה וניתן לעריכה, או משנה סוג, כותרת, מקרא או תוויות של תרשים קיים.' },
  },
  edit_smartart: {
    label: { en: 'SmartArt', he: 'SmartArt' },
    description: { en: 'Adds a SmartArt diagram, or edits an existing one\'s text, nodes, style, layout, or flow direction.', he: 'מוסיף דיאגרמת SmartArt, או עורך טקסט, צמתים, עיצוב, פריסה או כיוון זרימה של דיאגרמה קיימת.' },
  },
  edit_picture: {
    label: { en: 'Edit picture', he: 'עריכת תמונה' },
    description: { en: 'Adds a picture from a file, crops a picture, or swaps its image for another file.', he: 'מוסיף תמונה מקובץ, חותך תמונה, או מחליף את התמונה בקובץ אחר.' },
  },
  set_headers_footers: {
    label: { en: 'Slide numbers, footer & date', he: 'מספור שקופיות, כותרת תחתונה ותאריך' },
    description: { en: 'Turns slide numbers, footer text, and date/time on or off, deck-wide or for one slide.', he: 'מפעיל או מכבה מספור שקופיות, טקסט כותרת תחתונה ותאריך, בכל המצגת או בשקופית אחת.' },
  },
  edit_master_element: {
    label: { en: 'Edit master element', he: 'עריכת רכיב באב' },
    description: { en: 'Adds, moves, or removes an element on the slide master or one layout.', he: 'מוסיף, מזיז או מסיר רכיב באב השקופיות או בפריסה אחת.' },
  },
  read_master: {
    label: { en: 'Read slide master', he: 'קריאת אב השקופיות' },
    description: { en: 'Lists the slide master\'s or a layout\'s shapes, or the deck\'s layout names.', he: 'מציג את האובייקטים באב השקופיות או בפריסה, או את שמות הפריסות במצגת.' },
  },
  undo_last_action: {
    label: { en: 'Undo', he: 'ביטול' },
    description: { en: "Reverses the presentation's last action, like Ctrl+Z.", he: 'מבטל את הפעולה האחרונה במצגת, כמו Ctrl+Z.' },
  },
  redo_last_action: {
    label: { en: 'Redo', he: 'ביצוע חוזר' },
    description: { en: 'Re-applies the last undone action, like Ctrl+Y.', he: 'מבצע מחדש את הפעולה האחרונה שבוטלה, כמו Ctrl+Y.' },
  },
} satisfies Record<string, ToolDisplayInfo>

export const STARTERS = [
  { en: "Improve this slide's title and copy", he: 'שפר את הכותרת והטקסט של השקופית' },
  { en: "Make this slide's bullets more concise", he: 'קצר את התבליטים בשקופית' },
  { en: 'Check the whole deck for typos and fix them', he: 'בדוק שגיאות כתיב בכל המצגת ותקן אותן' },
  { en: 'Turn on slide numbers, but skip the title slide', he: 'הפעל מספור שקופיות, אך דלג על שקופית הכותרת' },
  { en: 'Add our logo to the bottom-right corner of every slide', he: 'הוסף את הלוגו שלנו לפינה הימנית התחתונה בכל שקופית' },
]
