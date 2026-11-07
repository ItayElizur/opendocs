import type { ToolDisplayInfo } from '@officeai/app-shell'

export const EXCEL_TOOL_DISPLAY = {
  get_workbook_context: {
    label: { en: 'Read workbook', he: 'קרא חוברת עבודה' },
    description: { en: "Reads the active sheet's name, used range, and current selection.", he: 'קורא את שם הגיליון הפעיל, הטווח בשימוש והבחירה הנוכחית.' },
  },
  read_range: {
    label: { en: 'Read range', he: 'קריאת טווח' },
    description: { en: 'Reads the values in a rectangular range of cells.', he: 'קורא את הערכים בטווח תאים מלבני.' },
  },
  read_cells: {
    label: { en: 'Read specific cells', he: 'קריאת תאים ספציפיים' },
    description: { en: 'Reads the values of specific, possibly non-adjacent cells.', he: 'קורא ערכים של תאים ספציפיים, לאו דווקא סמוכים.' },
  },
  select_range: {
    label: { en: 'Select range', he: 'בחירת טווח' },
    description: { en: 'Navigates to and selects a range on a sheet, without changing any data.', he: 'עובר ובוחר טווח בגיליון, מבלי לשנות נתונים.' },
  },
  read_formats: {
    label: { en: 'Read formatting', he: 'קריאת עיצוב' },
    description: { en: 'Reads which cells in a range have explicit formatting, such as bold or a number format.', he: 'קורא אילו תאים בטווח מעוצבים במפורש, כגון הדגשה או פורמט מספרי.' },
  },
  read_sheet_features: {
    label: { en: 'Read sheet settings', he: 'קריאת הגדרות גיליון' },
    description: { en: "Reports a sheet's filters, freeze panes, conditional formatting, defined names, and similar settings.", he: 'מדווח על מסננים, הקפאת חלוניות, עיצוב מותנה, שמות מוגדרים והגדרות דומות בגיליון.' },
  },
  find_cells: {
    label: { en: 'Search workbook', he: 'חיפוש בחוברת העבודה' },
    description: { en: 'Searches cell values or formulas across the workbook, or finds cells with errors.', he: 'מחפש ערכים או נוסחאות בתאים בכל החוברת, או מאתר תאים עם שגיאות.' },
  },
  trace_precedents: {
    label: { en: 'Trace precedents', he: 'מעקב אחר תאי מקור' },
    description: { en: 'Lists the cells a formula directly reads from.', he: 'מציג את התאים שנוסחה קוראת מהם ישירות.' },
  },
  trace_dependents: {
    label: { en: 'Trace dependents', he: 'מעקב אחר תאים תלויים' },
    description: { en: 'Lists the formulas that read a given cell.', he: 'מציג את הנוסחאות שקוראות מתא נתון.' },
  },
  propose_operations: {
    label: { en: 'Edit spreadsheet', he: 'עריכת גיליון אלקטרוני' },
    description: { en: 'Applies changes to the spreadsheet, such as cell edits, formatting, sorting, charts, and tables.', he: 'מבצע שינויים בגיליון, כגון עריכת תאים, עיצוב, מיון, תרשימים וטבלאות.' },
  },
} satisfies Record<string, ToolDisplayInfo>

export const STARTERS = [
  { en: 'Summarize this sheet', he: 'סכם את הגיליון הזה' },
  { en: 'Add a totals row', he: 'הוסף שורת סיכום' },
  { en: 'Check the formulas', he: 'בדוק את הנוסחאות' },
]
