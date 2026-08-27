import { startAddIn } from '@officeai/app-shell'

// Word add-in: tool definitions, system prompt, and starters are the only
// app-specific pieces left here - everything else (WebView2 bridge, settings,
// transport, chat-UI mount, AgentLoop event plumbing) lives once in
// @officeai/app-shell (PP-0), shared with ExcelAiAddIn and PowerPointAiAddIn.

// PP-5: structural per-kind schema for apply_commands, replacing the old
// `items: { type: 'object' }` (no structure at all) + one prose paragraph.
// This is documentation the model reads, not a validator that runs (not
// every provider enforces oneOf/const) - the actual guarantee is the
// required-field precheck in WordTools.cs's ApplyCommands (PP-5 Task 4).
// Every kind below matches WordTools.cs's ApplyCommands switch exactly
// (cross-checked in docs/ai-tool-surface.md) - do not add a kind here
// without a matching case there, or vice versa.

const TARGET_SCHEMA = {
  type: 'object',
  description:
    'Selects paragraphs. Fields are AND-combined; at least one of nodeType/containsText/blockIndexes is required.',
  properties: {
    nodeType: { type: 'string', enum: ['heading', 'paragraph', 'listItem'] },
    headingLevel: { type: 'number', minimum: 1, maximum: 6 },
    containsText: { type: 'string' },
    matchCase: { type: 'boolean' },
    blockIndexes: { type: 'array', items: { type: 'number' }, description: '0-based paragraph indices' },
    scope: { type: 'string', enum: ['document', 'selection'] },
  },
}

// Exactly the 10 keys UpdateTextStyle (WordTools.cs) checks against `fields` -
// PP-12 implemented `highlight` (was deliberately absent pre-PP-12).
const TEXT_STYLE_FIELDS = ['bold', 'italic', 'underline', 'strike', 'sizeHalfPoints', 'font', 'color', 'baselineOffset', 'link', 'highlight']

// Word's HighlightColors palette (WordTools.cs) - a fixed 16-entry
// WdColorIndex palette, NOT arbitrary RGB like `color` above.
const HIGHLIGHT_COLORS = [
  'none', 'yellow', 'brightGreen', 'turquoise', 'pink', 'blue', 'red', 'darkBlue',
  'teal', 'green', 'violet', 'darkRed', 'darkYellow', 'gray50', 'gray25', 'black', 'white',
]

const TEXT_STYLE_SCHEMA = {
  type: 'object',
  description: 'Only keys also listed in `fields` are applied. Any key not in fields/here errors instead of silently no-opping.',
  properties: {
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean' },
    strike: { type: 'boolean' },
    sizeHalfPoints: { type: 'number', description: 'Font size in half-points (e.g. 24 = 12pt).' },
    font: { type: 'string' },
    color: { type: 'string', description: 'Hex color, e.g. "#FF0000".' },
    baselineOffset: { type: 'string', enum: ['SUPERSCRIPT', 'SUBSCRIPT', 'NONE'] },
    link: { type: 'object', properties: { url: { type: 'string' } }, required: ['url'] },
    highlight: { type: 'string', enum: HIGHLIGHT_COLORS, description: 'A fixed palette name, not a hex color - hex is rejected.' },
  },
}

// Exactly the 10 keys UpdateParagraphStyle (WordTools.cs) checks against `fields`.
const PARAGRAPH_STYLE_FIELDS = [
  'align', 'lineSpacing', 'indentLeft', 'indentRight', 'indentFirstLine',
  'spaceBefore', 'spaceAfter', 'pageBreakBefore', 'shadingFill', 'borders',
]

const PARAGRAPH_STYLE_SCHEMA = {
  type: 'object',
  description: 'Only keys also listed in `fields` are applied.',
  properties: {
    align: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
    lineSpacing: { type: 'number' },
    indentLeft: { type: 'number' },
    indentRight: { type: 'number' },
    indentFirstLine: { type: 'number' },
    spaceBefore: { type: 'number' },
    spaceAfter: { type: 'number' },
    pageBreakBefore: { type: 'boolean' },
    shadingFill: { type: 'string', description: 'Hex color, e.g. "#FFFF00".' },
    borders: { type: 'boolean', description: 'true = single-line borders on all sides; false = no borders.' },
  },
}

// Exactly the 3 keys UpdateImageProperties (WordTools.cs) checks against
// `fields` - note align here is left/center/right only (no "justify"),
// unlike updateParagraphStyle's align - the handler's own switch has no
// justify case for images.
const IMAGE_PROPERTIES_FIELDS = ['widthPx', 'heightPx', 'align']

const IMAGE_PROPERTIES_SCHEMA = {
  type: 'object',
  description: 'Only keys also listed in `fields` are applied. Setting only one of widthPx/heightPx scales the other proportionally.',
  properties: {
    widthPx: { type: 'number' },
    heightPx: { type: 'number' },
    align: { type: 'string', enum: ['left', 'center', 'right'] },
  },
}

const WORD_COMMAND_SCHEMAS = [
  {
    type: 'object',
    properties: {
      kind: { const: 'set_bold' },
      startIndex: { type: 'number', description: '0-based paragraph index' },
      endIndex: { type: 'number', description: '0-based paragraph index, inclusive' },
      value: { type: 'boolean' },
    },
    required: ['kind', 'startIndex', 'endIndex', 'value'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'set_italic' },
      startIndex: { type: 'number', description: '0-based paragraph index' },
      endIndex: { type: 'number', description: '0-based paragraph index, inclusive' },
      value: { type: 'boolean' },
    },
    required: ['kind', 'startIndex', 'endIndex', 'value'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'set_heading' },
      index: { type: 'number', description: '0-based paragraph index' },
      level: { type: 'number', minimum: 0, maximum: 9, description: '0 = Normal style, 1-9 = Heading 1-9' },
    },
    required: ['kind', 'index', 'level'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'find_replace' },
      find: { type: 'string' },
      replace: { type: 'string' },
      matchCase: { type: 'boolean' },
    },
    required: ['kind', 'find', 'replace'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'updateTextStyle' },
      target: TARGET_SCHEMA,
      style: TEXT_STYLE_SCHEMA,
      fields: { type: 'array', items: { type: 'string', enum: TEXT_STYLE_FIELDS } },
    },
    required: ['kind', 'target', 'style', 'fields'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'updateParagraphStyle' },
      target: TARGET_SCHEMA,
      style: PARAGRAPH_STYLE_SCHEMA,
      fields: { type: 'array', items: { type: 'string', enum: PARAGRAPH_STYLE_FIELDS } },
    },
    required: ['kind', 'target', 'style', 'fields'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'deleteBlocks' },
      target: TARGET_SCHEMA,
    },
    required: ['kind', 'target'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'moveBlocks' },
      blockIndexes: { type: 'array', items: { type: 'number' }, description: '0-based paragraph indices to move' },
      afterBlockIndex: { type: 'number', description: '0-based paragraph index to insert after; -1 = start of document. Cannot be one of blockIndexes.' },
    },
    required: ['kind', 'blockIndexes', 'afterBlockIndex'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'createParagraphBullets' },
      target: TARGET_SCHEMA,
      bulletPreset: {
        type: 'string',
        enum: ['BULLET_DISC_CIRCLE_SQUARE', 'BULLET_DIAMOND_X', 'BULLET_CHECKBOX', 'NUMBERED_DECIMAL', 'NUMBERED_DECIMAL_ALPHA_ROMAN', 'NUMBERED_UPPERALPHA', 'NUMBERED_UPPERROMAN'],
        description: 'Default plain bullet if omitted. An unrecognized value errors (PP-12) rather than silently collapsing to a generic bullet.',
      },
    },
    required: ['kind', 'target'],
  },
  {
    type: 'object',
    description:
      'Turns bullets on or off in one command - the set_X-shaped alias for createParagraphBullets/deleteParagraphBullets, ' +
      'matching set_bold/set_italic/set_heading. value:true (or omitted) adds bullets, value:false removes them.',
    properties: {
      kind: { const: 'set_bullet' },
      target: TARGET_SCHEMA,
      value: { type: 'boolean', description: 'Default true. true = add bullets, false = remove them.' },
      bulletPreset: {
        type: 'string',
        enum: ['BULLET_DISC_CIRCLE_SQUARE', 'BULLET_DIAMOND_X', 'BULLET_CHECKBOX', 'NUMBERED_DECIMAL', 'NUMBERED_DECIMAL_ALPHA_ROMAN', 'NUMBERED_UPPERALPHA', 'NUMBERED_UPPERROMAN'],
        description: 'Only meaningful with value:true. Same presets as createParagraphBullets.',
      },
    },
    required: ['kind', 'target'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'deleteParagraphBullets' },
      target: TARGET_SCHEMA,
    },
    required: ['kind', 'target'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'updateImageProperties' },
      imageIndex: { type: 'number', description: '0-based index into the document\'s inline images' },
      properties: IMAGE_PROPERTIES_SCHEMA,
      fields: { type: 'array', items: { type: 'string', enum: IMAGE_PROPERTIES_FIELDS } },
    },
    required: ['kind', 'imageIndex', 'properties', 'fields'],
  },
  {
    type: 'object',
    properties: {
      kind: { const: 'insertToc' },
      afterBlockIndex: { type: 'number', description: '0-based paragraph index to insert after; -1 = start of document. Requires at least one Heading-styled paragraph in the document.' },
    },
    required: ['kind', 'afterBlockIndex'],
  },
]

const ALL_WORD_TOOLS = [
    {
      name: 'get_document_context',
      description: "Reads the active Word document's paragraph/word count and a text preview.",
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'insert_content',
      description:
        'Inserts content into the active Word document. Supply exactly one of text (plain; newlines create separate paragraphs) ' +
        'or html (restricted subset: <p> <h1>-<h3> <ul>/<ol>/<li> <b>/<strong> <i>/<em> <u> <br/> - must be well-formed XHTML, no attributes, no nesting of lists). ' +
        'afterBlockIndex (0-based paragraph index) anchors the insertion after that paragraph; -1 = start of document; omit = end of document (the original behavior).',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string' },
          html: { type: 'string' },
          afterBlockIndex: { type: 'number' },
        },
        required: [],
      },
    },
    {
      name: 'edit_chart',
      description:
        'Creates or edits a native Word chart. Supply categories + one or more named series for a labeled chart. ' +
        'chartIndex addresses an existing chart (0-based, inline shapes first then floating shapes, in document order); omit it to edit the first chart, ' +
        'or pass create:true to always add a new one. With no afterBlockIndex, a newly created chart is a floating shape at document origin (unchanged legacy behavior); ' +
        'with afterBlockIndex, it is inserted inline (flows with the text) after that paragraph.',
      inputSchema: {
        type: 'object',
        properties: {
          title: { type: 'string' },
          chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
          categories: { type: 'array', items: { type: 'string' } },
          series: {
            type: 'array',
            items: {
              type: 'object',
              properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } },
              required: ['values'],
            },
          },
          values: { type: 'array', items: { type: 'number' }, description: 'Legacy single-series shorthand; prefer series.' },
          chartIndex: { type: 'number' },
          create: { type: 'boolean' },
          afterBlockIndex: { type: 'number', description: '0-based paragraph index to anchor a NEW chart after; -1 = start.' },
        },
        required: [],
      },
    },
    {
      name: 'read_chart',
      description:
        'Reads an existing chart\'s current title, type, categories, and per-series names/values. ' +
        'Call this before an incremental edit_chart change (e.g. removing or renaming one category or series) - ' +
        'edit_chart REPLACES the whole dataset when categories/series is given, so you need the current data to correctly resend everything you are keeping. ' +
        'chartIndex addresses an existing chart (0-based, inline shapes first then floating shapes, in document order); omit it to read the first chart.',
      inputSchema: {
        type: 'object',
        properties: { chartIndex: { type: 'number' } },
        required: [],
      },
    },
    {
      name: 'find_text',
      description:
        'Searches the document\'s paragraph text for a substring (or, with regex:true, a .NET regular expression) - read-only, never modifies the document. ' +
        'Returns each match as "[index] full paragraph text" - that index is the EXACT SAME 0-based paragraph index read_blocks\'/replace_blocks\' ' +
        'startIndex/endIndex and apply_commands\' Target.blockIndexes use, so pass it straight through with no translation (e.g. a hit "[42] ..." means ' +
        'read_blocks({startIndex: 42, endIndex: 42}) or a Target of {blockIndexes: [42]}). Use this to locate the paragraphs you actually need instead of ' +
        'reading a large range blindly or guessing indices before find_replace.',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string' },
          regex: { type: 'boolean' },
          matchCase: { type: 'boolean' },
          max_results: { type: 'number' },
        },
        required: ['query', 'max_results'],
      },
    },
    {
      name: 'get_headings',
      description:
        'Lists every heading-styled paragraph in the document, in order, like Word\'s Navigation Pane - each line is "[index] H<level>: text". ' +
        'Use this to see the document\'s outline/structure in one call instead of reading every paragraph via read_blocks.',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'read_blocks',
      description:
        'Reads paragraphs [startIndex, endIndex] (0-based, inclusive) of the active document, one per line prefixed with its index - capped at 1000 paragraphs per call. ' +
        'format:"html" emits the same restricted HTML subset insert_content/replace_blocks accept (headings/bold/italic/underline/list membership survive; ' +
        'anything outside that subset, e.g. font color, does not) - capped lower, at 100 paragraphs per call (per-paragraph formatting reads are slower); strip the leading "[i] " markers before feeding the fragment back into html. ' +
        'Prefer find_text to locate the paragraph indices you actually need instead of reading a large range blindly.',
      inputSchema: {
        type: 'object',
        properties: { startIndex: { type: 'number' }, endIndex: { type: 'number' }, format: { type: 'string', enum: ['text', 'html'] } },
        required: ['startIndex', 'endIndex'],
      },
    },
    {
      name: 'replace_blocks',
      description:
        'Replaces paragraphs [startIndex, endIndex] (0-based, inclusive) with new content. Supply exactly one of text or html (same restricted subset as insert_content). ' +
        'Empty text deletes the range. preserveFormatting (default true) reapplies the first replaced paragraph\'s style (e.g. Heading 2) to the result - ' +
        'pass false for the old strip-everything behavior. preserveFormatting has no effect when html is given (the fragment\'s own tags dictate style).',
      inputSchema: {
        type: 'object',
        properties: {
          startIndex: { type: 'number' },
          endIndex: { type: 'number' },
          text: { type: 'string' },
          html: { type: 'string' },
          preserveFormatting: { type: 'boolean' },
        },
        required: ['startIndex', 'endIndex'],
      },
    },
    {
      name: 'add_image',
      description:
        'Inserts an image from a LOCAL FILE PATH into the document (no URLs - this deployment is air-gapped). ' +
        'Inserts inline in the text flow by default, after the paragraph given by afterBlockIndex (0-based; -1 = start; omit = end of document). ' +
        'Inline images are addressable afterwards by apply_commands/updateImageProperties via their 0-based index, which this tool returns; floating images (floating:true) are not.',
      inputSchema: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          afterBlockIndex: { type: 'number' },
          floating: { type: 'boolean' },
          widthPoints: { type: 'number' },
          heightPoints: { type: 'number' },
          altText: { type: 'string' },
        },
        required: ['path'],
      },
    },
    {
      name: 'add_table',
      description:
        'Adds a native Word table, optionally pre-filled with cell text (row-major array of arrays; extra cells beyond rows/cols are ignored). ' +
        'afterBlockIndex is the 0-based paragraph index to insert after (-1 = start of document; omit = end of document).',
      inputSchema: {
        type: 'object',
        properties: {
          rows: { type: 'number' },
          cols: { type: 'number' },
          cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
          afterBlockIndex: { type: 'number' },
        },
        required: ['rows', 'cols'],
      },
    },
    {
      name: 'edit_table',
      description:
        'Edits an existing table. kind: "set_cell" (row,col,text), "insert_row"/"delete_row"/"insert_col"/"delete_col" (index,before?), ' +
        '"set_style" (styleName?,headerRow?,bandedRows?,borders?,borderColor?), "set_shading" (scope,color,row?,col?) fills cell background color. ' +
        'borders:true draws a single-line border on every table edge (default color black, override with borderColor); borders:false removes all table borders. ' +
        'set_shading scope: "cell" (needs row+col), "row" (needs row, fills the whole row), "col" (needs col, fills the whole column), "table" (fills every cell) - color is a required hex string, e.g. "#FFFF00". ' +
        'tableIndex addresses the table (0-based, document order); omit to target the first table. ' +
        'Structural edits shift later indices - re-read the table (read_table) before a second structural edit in the same run.',
      inputSchema: {
        type: 'object',
        properties: {
          tableIndex: { type: 'number' },
          kind: { type: 'string', enum: ['set_cell', 'insert_row', 'delete_row', 'insert_col', 'delete_col', 'set_style', 'set_shading'] },
          row: { type: 'number' },
          col: { type: 'number' },
          text: { type: 'string' },
          index: { type: 'number' },
          before: { type: 'boolean' },
          styleName: { type: 'string' },
          headerRow: { type: 'boolean' },
          bandedRows: { type: 'boolean' },
          borders: { type: 'boolean' },
          borderColor: { type: 'string', description: 'Hex color, e.g. "#000000". Only applied when borders:true.' },
          scope: { type: 'string', enum: ['cell', 'row', 'col', 'table'], description: 'Required for kind:"set_shading".' },
          color: { type: 'string', description: 'Hex color, e.g. "#FFFF00". Required for kind:"set_shading".' },
        },
        required: ['kind'],
      },
    },
    {
      name: 'read_table',
      description: 'Reads an existing table\'s cell contents, one row per line. tableIndex addresses the table (0-based, document order); omit to read the first table.',
      inputSchema: { type: 'object', properties: { tableIndex: { type: 'number' } }, required: [] },
    },
    {
      name: 'add_smartart',
      description:
        'Adds a shape-composed SmartArt diagram. layout: "list"|"process"|"cycle"|"hierarchy"|"pyramid"|"matrix"|"venn". items are flat node texts, one per top-level node. ' +
        'afterBlockIndex (0-based paragraph index, -1 = start, omit = end of document) inserts inline; omitting both x/y and afterBlockIndex places a floating shape at a default position.',
      inputSchema: {
        type: 'object',
        properties: {
          layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'] },
          items: { type: 'array', items: { type: 'string' } },
          x: { type: 'number' },
          y: { type: 'number' },
          w: { type: 'number' },
          h: { type: 'number' },
          afterBlockIndex: { type: 'number' },
        },
        required: ['layout', 'items'],
      },
    },
    {
      name: 'edit_smartart',
      description:
        'Edits an existing SmartArt diagram. kind: "set_text" (nodeIndex,text), "add_node" (text?), "delete_node" (nodeIndex), ' +
        '"set_style" (colorName?,quickStyleName? - free-text, matched by substring against this Office install\'s actual color/quick-style gallery; ' +
        'an unmatched name errors listing the real available names to retry with, e.g. try "Colorful" or "Accent" for colorName, "Intense" or "Simple" for quickStyleName), ' +
        '"set_layout" (layout - same layout keys as add_smartart: list/process/cycle/hierarchy/pyramid/matrix/venn; changes an existing diagram\'s layout, keeping its current node text). ' +
        'smartArtIndex addresses the diagram (0-based, document order); omit to target the first one. ' +
        'delete_node shifts later node indices - re-read (read_smartart) before another node edit in the same run.',
      inputSchema: {
        type: 'object',
        properties: {
          smartArtIndex: { type: 'number' },
          kind: { type: 'string', enum: ['set_text', 'add_node', 'delete_node', 'set_style', 'set_layout'] },
          nodeIndex: { type: 'number' },
          text: { type: 'string' },
          colorName: { type: 'string' },
          quickStyleName: { type: 'string' },
          layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'] },
        },
        required: ['kind'],
      },
    },
    {
      name: 'read_smartart',
      description: 'Reads SmartArt diagram node texts, one per line. smartArtIndex addresses a specific diagram (0-based, document order); omit to read every diagram in the document in one call.',
      inputSchema: { type: 'object', properties: { smartArtIndex: { type: 'number' } }, required: [] },
    },
    {
      name: 'apply_commands',
      description:
        'Applies a batch of formatting/editing commands. Each command is one of the kinds described by the schema below - see each branch\'s properties for its exact fields.',
      inputSchema: {
        type: 'object',
        properties: { commands: { type: 'array', items: { oneOf: WORD_COMMAND_SCHEMAS } } },
        required: ['commands'],
      },
    },
    {
      name: 'add_comment',
      description:
        'Adds a Word comment anchored to the first occurrence of the given text, without changing document content. Available in every editing mode.',
      inputSchema: {
        type: 'object',
        properties: { anchorText: { type: 'string' }, commentText: { type: 'string' } },
        required: ['anchorText', 'commentText'],
      },
    },
  ]

// FT-1 Task 5: settings-screen labels/descriptions, one entry per tool above.
// These are short, user-facing sentences for a non-technical reader - NOT the
// model-facing `description` strings above, which enumerate parameter shapes
// and are the wrong register for this screen (Task 5 Step 2).
const WORD_TOOL_DISPLAY = {
  get_document_context: {
    label: { en: 'Read document', he: 'קרא מסמך' },
    description: { en: "Reads the document's structure and a short preview of its text.", he: 'קורא את מבנה המסמך ותצוגה מקדימה קצרה של הטקסט.' },
  },
  insert_content: {
    label: { en: 'Insert content', he: 'הוספת תוכן' },
    description: { en: 'Inserts new text or formatted content at a chosen position in the document.', he: 'מוסיף טקסט חדש או תוכן מעוצב במיקום נבחר במסמך.' },
  },
  edit_chart: {
    label: { en: 'Create or edit chart', he: 'יצירה/עריכת תרשים' },
    description: { en: 'Creates a new chart or edits an existing one, with categories and data series.', he: 'יוצר תרשים חדש או עורך תרשים קיים, כולל קטגוריות וסדרות נתונים.' },
  },
  read_chart: {
    label: { en: 'Read chart data', he: 'קריאת נתוני תרשים' },
    description: { en: 'Reads an existing chart\'s title, type, categories, and series values.', he: 'קורא את כותרת התרשים, סוגו, הקטגוריות וערכי הסדרות שלו.' },
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
  read_table: {
    label: { en: 'Read table', he: 'קריאת טבלה' },
    description: { en: 'Reads an existing table\'s cell contents.', he: 'קורא את תוכן התאים של טבלה קיימת.' },
  },
  add_smartart: {
    label: { en: 'Insert SmartArt', he: 'הוספת SmartArt' },
    description: { en: 'Adds a SmartArt diagram (list, process, cycle, hierarchy, pyramid, matrix, or Venn).', he: 'מוסיף דיאגרמת SmartArt (רשימה, תהליך, מעגל, היררכיה, פירמידה, מטריצה או ון).' },
  },
  edit_smartart: {
    label: { en: 'Edit SmartArt', he: 'עריכת SmartArt' },
    description: { en: 'Edits an existing SmartArt diagram\'s node text.', he: 'עורך את טקסט הצמתים של דיאגרמת SmartArt קיימת.' },
  },
  read_smartart: {
    label: { en: 'Read SmartArt', he: 'קריאת SmartArt' },
    description: { en: 'Reads an existing SmartArt diagram\'s node text.', he: 'קורא את טקסט הצמתים של דיאגרמת SmartArt קיימת.' },
  },
  apply_commands: {
    label: { en: 'Edit and format', he: 'עריכה ועיצוב' },
    description: { en: 'Applies formatting and editing commands such as bold, headings, bullets, and find/replace.', he: 'מיישם פקודות עיצוב ועריכה כגון הדגשה, כותרות, תבליטים וחיפוש/החלפה.' },
  },
  add_comment: {
    label: { en: 'Add comment', he: 'הוספת הערה' },
    description: { en: 'Adds a comment anchored to text in the document, without changing its content.', he: 'מוסיף הערה מעוגנת לטקסט במסמך, מבלי לשנות את תוכנו.' },
  },
}

startAddIn({
  skillId: 'word-tools',
  tools: ALL_WORD_TOOLS,
  toolDisplay: WORD_TOOL_DISPLAY,
  systemPrompt:
    'You are an AI assistant embedded in Microsoft Word via the Airchat Office add-in. ' +
    'You can read the document, insert content at any position (plain text or a restricted HTML subset - headings, bold/italic/underline, bulleted/numbered lists), ' +
    'read and non-destructively replace paragraph ranges, apply formatting and find/replace commands, insert images from local file paths, add comments, ' +
    'search the document\'s text read-only with find_text, and list the document\'s heading outline with get_headings (like the Navigation Pane). ' +
    'When you need to find something in the document rather than read it start-to-finish, call find_text FIRST rather than reading a large range with ' +
    'read_blocks blindly (or worse, paging through the whole document) - the "[index]" it returns is the exact same paragraph index read_blocks/' +
    'replace_blocks\' startIndex/endIndex and apply_commands\' Target.blockIndexes use, so pass it straight through with no translation. ' +
    'create, read, or edit a native Word chart with labeled categories and named multi-series, create/read/edit native Word tables, and create/read/edit SmartArt diagrams. ' +
    'edit_chart REPLACES the whole categories/series dataset when given, so before an incremental change to an existing chart (e.g. removing or renaming one category), call read_chart first to see the current data and resend everything you are keeping. ' +
    'edit_table\'s insert_row/delete_row/insert_col/delete_col and edit_smartart\'s delete_node shift later row/column/node indices - re-read (read_table/read_smartart) before a second structural edit to the same table or diagram in the same run. ' +
    "Your available tools depend on the user's current editing mode (Read only, Comment only, Track changes, or Full autonomy); only call tools that are currently offered to you. " +
    'If the user has selected text in the document, it will be included in your context as "Content selected by the user."',
  starters: [
    { en: 'Summarize the key points of this document', he: 'סכם את הנקודות העיקריות במסמך' },
    { en: 'Polish the whole document for a more professional tone', he: 'לטש את כל המסמך לטון מקצועי יותר' },
    { en: 'Continue writing from where the document leaves off', he: 'המשך לכתוב מהיכן שהמסמך מסתיים' },
  ],
  readOnlyTools: ['get_document_context', 'read_blocks', 'find_text', 'get_headings', 'read_chart', 'read_table', 'read_smartart'],
  commentOnlyExtraTools: ['add_comment'],
  useSelectionContext: true,
})
