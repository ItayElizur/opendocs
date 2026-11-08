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
  read_group: {
    label: { en: 'Read group', he: 'קריאת קבוצה' },
    description: { en: 'Lists the shapes inside a group, including nested groups.', he: 'מציג את האובייקטים בתוך קבוצה, כולל קבוצות מקוננות.' },
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
  set_element_style: {
    label: { en: 'Format shape text', he: 'עיצוב טקסט באובייקט' },
    description: { en: 'Changes the text formatting of a shape, such as bold, color, or font size.', he: 'משנה את עיצוב הטקסט של אובייקט, כגון הדגשה, צבע או גודל גופן.' },
  },
  set_element_transform: {
    label: { en: 'Move/resize shape', he: 'הזזה/שינוי גודל של אובייקט' },
    description: { en: 'Moves, resizes, or rotates a shape.', he: 'מזיז, משנה גודל או מסובב אובייקט.' },
  },
  set_element_order: {
    label: { en: 'Reorder shape', he: 'שינוי סדר אובייקט' },
    description: { en: 'Brings a shape forward or backward in the stacking order.', he: 'מקדם או מרחיק אובייקט בסדר הערימה.' },
  },
  add_text_box: {
    label: { en: 'Add text box', he: 'הוספת תיבת טקסט' },
    description: { en: 'Creates a new text box on a slide.', he: 'יוצר תיבת טקסט חדשה בשקופית.' },
  },
  add_shape: {
    label: { en: 'Add shape', he: 'הוספת צורה' },
    description: { en: 'Creates a new shape, such as a rectangle, ellipse, or arrow.', he: 'יוצר צורה חדשה, כגון מלבן, אליפסה או חץ.' },
  },
  delete_element: {
    label: { en: 'Delete shape', he: 'מחיקת אובייקט' },
    description: { en: 'Deletes a shape from a slide.', he: 'מוחק אובייקט משקופית.' },
  },
  duplicate_element: {
    label: { en: 'Duplicate shape', he: 'שכפול אובייקט' },
    description: { en: 'Creates a copy of a shape on the same slide.', he: 'יוצר עותק של אובייקט באותה שקופית.' },
  },
  copy_element: {
    label: { en: 'Copy shape to another slide', he: 'העתקת אובייקט לשקופית אחרת' },
    description: { en: 'Copies a shape to a different slide, with full native formatting.', he: 'מעתיק אובייקט לשקופית אחרת, כולל כל העיצוב המקורי.' },
  },
  manage_slide: {
    label: { en: 'Manage slide', he: 'ניהול שקופית' },
    description: { en: 'Adds, duplicates, moves, or deletes a slide.', he: 'מוסיף, משכפל, מעביר או מוחק שקופית.' },
  },
  set_slide: {
    label: { en: 'Update slide', he: 'עדכון שקופית' },
    description: { en: "Changes a slide's layout, background, transition, or speaker notes.", he: 'משנה את פריסת השקופית, הרקע, המעבר או הערות הדובר.' },
  },
  move_element: {
    label: { en: 'Move shape to another slide', he: 'העברת אובייקט לשקופית אחרת' },
    description: { en: 'Moves a shape to a different slide, removing it from its original slide.', he: 'מעביר אובייקט לשקופית אחרת, ומסיר אותו מהשקופית המקורית.' },
  },
  add_animation: {
    label: { en: 'Add animation', he: 'הוספת אנימציה' },
    description: { en: 'Adds an entrance or exit animation to a shape.', he: 'מוסיף אנימציית כניסה או יציאה לאובייקט.' },
  },
  read_animations: {
    label: { en: 'Read animations', he: 'קריאת אנימציות' },
    description: { en: "Reads a slide's animations and their order.", he: 'קורא את אנימציות השקופית וסדרן.' },
  },
  edit_animation: {
    label: { en: 'Edit animation', he: 'עריכת אנימציה' },
    description: { en: "Deletes, retimes, or reorders an existing animation.", he: 'מוחק, משנה תזמון או מסדר מחדש אנימציה קיימת.' },
  },
  set_element_fill: {
    label: { en: 'Set shape fill', he: 'צביעת אובייקט' },
    description: { en: "Sets or removes a shape's fill color.", he: 'מגדיר או מסיר את צבע המילוי של אובייקט.' },
  },
  set_element_stroke: {
    label: { en: 'Set shape outline', he: 'מתאר אובייקט' },
    description: { en: "Sets or removes a shape's outline color and width.", he: 'מגדיר או מסיר את צבע ועובי המתאר של אובייקט.' },
  },
  copy_element_style: {
    label: { en: 'Copy shape formatting', he: 'העתקת עיצוב אובייקט' },
    description: { en: 'Copies text, fill, and outline formatting from one shape onto others (like a format painter).', he: 'מעתיק עיצוב טקסט, מילוי ומתאר מאובייקט אחד לאובייקטים אחרים (כמו מכחול עיצוב).' },
  },
  ungroup_element: {
    label: { en: 'Ungroup shapes', he: 'פירוק קבוצת אובייקטים' },
    description: { en: 'Breaks a grouped shape into its individual parts.', he: 'מפרק אובייקט מקובץ לחלקיו הבודדים.' },
  },
  group_element: {
    label: { en: 'Group shapes', he: 'קיבוץ אובייקטים' },
    description: { en: 'Combines two or more shapes into a single group.', he: 'מאחד שני אובייקטים או יותר לקבוצה אחת.' },
  },
  add_table: {
    label: { en: 'Add table', he: 'הוספת טבלה' },
    description: { en: 'Adds a table to a slide, optionally pre-filled with data.', he: 'מוסיף טבלה לשקופית, ניתן למלא אותה מראש בנתונים.' },
  },
  edit_table_cell: {
    label: { en: 'Edit table cell', he: 'עריכת תא בטבלה' },
    description: { en: 'Replaces the text in one table cell.', he: 'מחליף את הטקסט בתא אחד בטבלה.' },
  },
  edit_table_structure: {
    label: { en: 'Edit table rows/columns', he: 'עריכת מבנה הטבלה' },
    description: { en: 'Inserts or deletes a row or column in a table.', he: 'מוסיף או מוחק שורה או עמודה בטבלה.' },
  },
  edit_table_style: {
    label: { en: 'Format table', he: 'עיצוב טבלה' },
    description: { en: 'Applies styling to a table, such as header row, banding, shading, and borders.', he: 'מיישם עיצוב על טבלה, כגון שורת כותרת, פסים, הצללה ומסגרות.' },
  },
  add_chart: {
    label: { en: 'Add chart', he: 'הוספת תרשים' },
    description: { en: 'Adds a native, editable chart with categories and data series.', he: 'מוסיף תרשים מובנה וניתן לעריכה עם קטגוריות וסדרות נתונים.' },
  },
  edit_chart: {
    label: { en: 'Edit chart', he: 'עריכת תרשים' },
    description: { en: "Modifies an existing chart's type, title, legend, or labels.", he: 'משנה סוג, כותרת, מקרא או תוויות של תרשים קיים.' },
  },
  edit_smartart: {
    label: { en: 'Edit SmartArt', he: 'עריכת SmartArt' },
    description: { en: 'Edits an existing SmartArt diagram\'s text, nodes, style, or layout.', he: 'עורך טקסט, צמתים, עיצוב או פריסה של דיאגרמת SmartArt קיימת.' },
  },
  read_smartart: {
    label: { en: 'Read SmartArt', he: 'קריאת SmartArt' },
    description: { en: 'Reads the node text of SmartArt diagrams on a slide.', he: 'קורא את טקסט הצמתים של דיאגרמות SmartArt בשקופית.' },
  },
  add_smartart: {
    label: { en: 'Add SmartArt', he: 'הוספת SmartArt' },
    description: { en: 'Adds a SmartArt diagram built from a list of items.', he: 'מוסיף תרשים SmartArt הבנוי מרשימת פריטים.' },
  },
  crop_image: {
    label: { en: 'Crop image', he: 'חיתוך תמונה' },
    description: { en: 'Crops a picture without permanently discarding the cropped-out parts.', he: 'חותך תמונה מבלי למחוק לצמיתות את החלקים שנחתכו.' },
  },
  set_picture_opacity: {
    label: { en: 'Set image opacity', he: 'שקיפות תמונה' },
    description: { en: 'Sets how transparent or opaque a picture appears.', he: 'מגדיר עד כמה התמונה תהיה שקופה או אטומה.' },
  },
  replace_image: {
    label: { en: 'Replace image', he: 'החלפת תמונה' },
    description: { en: "Swaps a picture's content for a different image file, keeping its position and size.", he: 'מחליף את תוכן התמונה בקובץ תמונה אחר, תוך שמירה על מיקום וגודל.' },
  },
  set_headers_footers: {
    label: { en: 'Slide numbers, footer & date', he: 'מספור שקופיות, כותרת תחתונה ותאריך' },
    description: { en: 'Turns slide numbers, footer text, and date/time on or off, deck-wide or for one slide.', he: 'מפעיל או מכבה מספור שקופיות, טקסט כותרת תחתונה ותאריך, בכל המצגת או בשקופית אחת.' },
  },
  add_master_element: {
    label: { en: 'Add logo/label to every slide', he: 'הוספת לוגו/תווית לכל השקופיות' },
    description: { en: 'Pins an icon or small text label to a corner of the Slide Master so it appears on every slide.', he: 'מצמיד סמל או תווית טקסט קטנה לפינה של השקופית האב, כך שתופיע בכל השקופיות.' },
  },
  set_master_element_transform: {
    label: { en: 'Move/resize master element', he: 'הזזה/שינוי גודל של רכיב בשקופית האב' },
    description: { en: 'Moves, resizes, or rotates an element already on the Slide Master.', he: 'מזיז, משנה גודל או מסובב רכיב שכבר נמצא בשקופית האב.' },
  },
  remove_master_element: {
    label: { en: 'Remove master element', he: 'הסרת רכיב מהשקופית האב' },
    description: { en: 'Deletes a shape that was added to the Slide Master.', he: 'מוחק אובייקט שנוסף לשקופית האב.' },
  },
  read_master_elements: {
    label: { en: 'Read master elements', he: 'קריאת רכיבי השקופית האב' },
    description: { en: 'Lists the shapes placed on the Slide Master.', he: 'מציג רשימה של האובייקטים הממוקמים על השקופית האב.' },
  },
  list_layouts: {
    label: { en: 'List layouts', he: 'רשימת פריסות' },
    description: { en: "Lists every layout name in the presentation's theme.", he: 'מציג רשימה של כל שמות הפריסות בעיצוב המצגת.' },
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
