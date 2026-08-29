import { startAddIn } from '@officeai/app-shell'

// Excel add-in: tool definitions, system prompt, and starters are the only
// app-specific pieces left here - everything else (WebView2 bridge, settings,
// transport, chat-UI mount, AgentLoop event plumbing) lives once in
// @officeai/app-shell (PP-0), shared with WordAiAddIn and PowerPointAiAddIn.
// 9 standalone read tools plus propose_operations (a 50-operation-kind
// gateway covering writing, formatting, layout, sheet structure, charts/
// visuals, native tables, pivot tables, and misc data operations).

// PP-5: table-driven schema generation for propose_operations, replacing the
// old `items: { type: 'object' }` (no structure at all) + one hand-written
// ~30-line description string covering ~50 kinds by prose alone. EXCEL_OPS is
// the single source of truth for BOTH the JSON Schema (opSchemas) and the
// human-readable description (opsDescription) - editing one edits both, so
// they cannot drift apart the way the old hand-maintained pair could.
// This is documentation the model reads, not a validator that runs (not
// every provider enforces oneOf/const) - the actual guarantee is the
// required-field precheck in ExcelTools.cs's ProposeOperations (PP-5 Task 4).
// Every kind below matches ExcelTools.cs's ProposeOperations switch exactly
// (cross-checked in docs/ai-tool-surface.md) - do not add a kind here
// without a matching case there, or vice versa.

interface OpSpec {
  kind: string
  group: 'Writing' | 'Formatting' | 'Layout' | 'Structure' | 'Charts/visuals' | 'Tables' | 'Pivot' | 'Data'
  /** JSON Schema properties, excluding `kind` and the shared optional `sheet` */
  props: Record<string, unknown>
  required?: string[]
  note?: string
}

// Shared "format" shape applied by add_conditional_format's number/text/
// blank/duplicate/top10/formula rule kinds (AddConditionalFormat, ExcelTools.cs)
// - colorScale/dataBar do NOT get this (they carry their own visual and the
// handler returns before reaching the shared format block).
const CF_FORMAT_SCHEMA = {
  type: 'object',
  properties: { bold: { type: 'boolean' }, fontColor: { type: 'string' }, fillColor: { type: 'string' } },
  description: 'Applied when the rule matches.',
}

// add_conditional_format's rule.kind - 8 branches, matching AddConditionalFormat's
// switch (ExcelTools.cs) field-for-field.
const CF_RULE_SCHEMA = {
  type: 'object',
  description: 'Selects a conditional-formatting rule kind and its fields.',
  oneOf: [
    {
      type: 'object',
      properties: {
        kind: { const: 'number' },
        operator: { type: 'string', enum: ['greaterThan', 'lessThan', 'equal', 'notEqual', 'greaterEqual', 'lessEqual', 'between', 'notBetween'] },
        value: { type: 'number' },
        value2: { type: 'number', description: 'Upper bound; required for between/notBetween.' },
        format: CF_FORMAT_SCHEMA,
      },
      required: ['kind', 'operator', 'value'],
    },
    {
      type: 'object',
      properties: {
        kind: { const: 'text' },
        text: { type: 'string' },
        match: { type: 'string', enum: ['contains', 'notContains', 'beginsWith', 'endsWith'], description: 'Default contains.' },
        format: CF_FORMAT_SCHEMA,
      },
      required: ['kind', 'text'],
    },
    { type: 'object', properties: { kind: { const: 'blank' }, format: CF_FORMAT_SCHEMA }, required: ['kind'] },
    {
      type: 'object',
      properties: {
        kind: { const: 'duplicate' },
        mode: { type: 'string', enum: ['duplicate', 'unique'], description: 'Default duplicate.' },
        format: CF_FORMAT_SCHEMA,
      },
      required: ['kind'],
    },
    {
      type: 'object',
      properties: {
        kind: { const: 'top10' },
        rank: { type: 'number', description: 'Default 10.' },
        percent: { type: 'boolean' },
        bottom: { type: 'boolean' },
        format: CF_FORMAT_SCHEMA,
      },
      required: ['kind'],
    },
    {
      type: 'object',
      properties: {
        kind: { const: 'formula' },
        formula: { type: 'string', description: 'Excel formula relative to the range\'s top-left cell, e.g. "=$C1>100" - a common source of "matches the wrong cells" if not written relative to that anchor.' },
        format: CF_FORMAT_SCHEMA,
      },
      required: ['kind', 'formula'],
    },
    {
      type: 'object',
      properties: {
        kind: { const: 'colorScale' },
        minColor: { type: 'string' },
        midColor: { type: 'string' },
        maxColor: { type: 'string' },
      },
      required: ['kind'],
      description: 'No `format` - the color scale itself is the visual.',
    },
    {
      type: 'object',
      properties: { kind: { const: 'dataBar' }, color: { type: 'string' } },
      required: ['kind'],
      description: 'No `format` - the data bar itself is the visual.',
    },
  ],
}

// set_data_validation's validation.kind - 5 supported branches, matching
// SetDataValidation's switch (ExcelTools.cs). "checkbox" is deliberately NOT
// a branch here: the handler always throws for it (this PIA has no
// checkbox-cell Data Validation type - CheckBox is a form control, not a
// Validation kind) - listing it as valid would be false advertising.
const DATA_VALIDATION_SCHEMA = {
  type: 'object',
  description: 'Pass null (on the parent operation\'s `validation` field) to remove existing validation instead of setting a new rule.',
  oneOf: [
    { type: 'object', properties: { kind: { const: 'list' }, values: { type: 'array', items: { type: 'string' } } }, required: ['kind', 'values'] },
    { type: 'object', properties: { kind: { const: 'listRef' }, range: { type: 'string' } }, required: ['kind', 'range'] },
    { type: 'object', properties: { kind: { const: 'numberBetween' }, min: { type: 'number' }, max: { type: 'number' } }, required: ['kind', 'min', 'max'] },
    { type: 'object', properties: { kind: { const: 'dateBetween' }, start: { type: 'string' }, end: { type: 'string' } }, required: ['kind', 'start', 'end'] },
    { type: 'object', properties: { kind: { const: 'formula' }, formula: { type: 'string' } }, required: ['kind', 'formula'] },
  ],
}

const EXCEL_OPS: OpSpec[] = [
  // --- Writing ---
  { kind: 'set_cell', group: 'Writing', props: { address: { type: 'string' }, value: {} }, required: ['address', 'value'] },
  { kind: 'set_formula', group: 'Writing', props: { address: { type: 'string' }, formula: { type: 'string' } }, required: ['address', 'formula'] },
  { kind: 'set_range', group: 'Writing', props: { address: { type: 'string' }, values: { type: 'array', items: { type: 'array' }, description: 'Row-major 2D array of cell values.' } }, required: ['address', 'values'] },
  { kind: 'clear_cell', group: 'Writing', props: { address: { type: 'string' } }, required: ['address'] },
  { kind: 'clear_range', group: 'Writing', props: { range: { type: 'string' } }, required: ['range'] },
  {
    kind: 'find_replace', group: 'Writing',
    props: {
      find: { type: 'string' },
      replace: { type: 'string' },
      regex: { type: 'boolean' },
      matchCase: { type: 'boolean' },
      sheetId: { type: 'string' },
      allSheets: { type: 'boolean' },
    },
    required: ['find', 'replace'],
    note: 'matches find_cells\' scoping: active sheet only by default (like Ctrl+H\'s "Within: Sheet") - pass allSheets:true for the whole workbook, or sheetId to name one specific sheet. Only replaces within literal cell VALUES, never formulas - use find_cells to locate formulas, then set_formula to edit them. Reports how many cells actually changed.',
  },

  // --- Formatting ---
  // PP-13: widened from bold/italic/numberFormat/fillColor to full parity
  // with genoffice's format_range. Every property is optional and additive -
  // an absent property leaves the cell's current value alone.
  {
    kind: 'format_range', group: 'Formatting',
    props: {
      address: { type: 'string' },
      bold: { type: 'boolean' },
      italic: { type: 'boolean' },
      numberFormat: { type: 'string' },
      fillColor: { type: 'string', description: 'Hex color, e.g. "#FFFF00".' },
      fontName: { type: 'string' },
      fontSize: { type: 'number' },
      fontColor: { type: 'string', description: 'Hex color, e.g. "#FF0000".' },
      strikethrough: { type: 'boolean' },
      underline: {
        description: 'true = single underline, false = none, or a specific style name.',
        oneOf: [{ type: 'boolean' }, { type: 'string', enum: ['none', 'single', 'double', 'singleAccounting', 'doubleAccounting'] }],
      },
      horizontalAlignment: { type: 'string', enum: ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerAcrossSelection', 'distributed'] },
      verticalAlignment: { type: 'string', enum: ['top', 'center', 'bottom', 'justify', 'distributed'] },
      wrapText: { type: 'boolean' },
      textRotation: { type: 'number', minimum: -90, maximum: 90, description: 'Degrees, -90 to 90.' },
      indent: { type: 'number', minimum: 0, maximum: 15 },
      borders: {
        type: 'object',
        description: 'preset is applied first, then edges refines it (not mutually exclusive). Omit both for no border change.',
        properties: {
          preset: { type: 'string', enum: ['none', 'outline', 'all', 'thick-outline'] },
          edges: { type: 'array', items: { type: 'string', enum: ['top', 'bottom', 'left', 'right', 'insideHorizontal', 'insideVertical', 'diagonalDown', 'diagonalUp'] } },
          style: { type: 'string', enum: ['thin', 'medium', 'thick', 'double', 'dotted', 'dashed', 'none'] },
          color: { type: 'string', description: 'Hex color, e.g. "#RRGGBB".' },
        },
      },
    },
    required: ['address'],
  },

  // --- Layout ---
  {
    kind: 'sort_range', group: 'Layout',
    props: {
      range: { type: 'string' },
      byColumn: { type: 'string', description: 'Column letter to sort by, e.g. "A".' },
      order: { type: 'string', enum: ['asc', 'desc'] },
      hasHeader: { type: 'boolean' },
    },
    required: ['range', 'byColumn', 'order'],
  },
  { kind: 'merge_cells', group: 'Layout', props: { range: { type: 'string' } }, required: ['range'] },
  { kind: 'unmerge_cells', group: 'Layout', props: { range: { type: 'string' } }, required: ['range'] },
  {
    kind: 'set_row_height', group: 'Layout',
    props: { row: { type: 'number', description: '1-based' }, count: { type: 'number', description: 'Default 1.' }, heightPoints: { type: 'number' } },
    required: ['row', 'heightPoints'],
  },
  {
    kind: 'set_col_width', group: 'Layout',
    props: { column: { type: 'number', description: '1-based' }, count: { type: 'number', description: 'Default 1.' }, widthPx: { type: 'number' } },
    required: ['column', 'widthPx'],
  },
  {
    kind: 'set_rows_hidden', group: 'Layout',
    props: { row: { type: 'number', description: '1-based' }, count: { type: 'number', description: 'Default 1.' }, hidden: { type: 'boolean' } },
    required: ['row', 'hidden'],
  },
  {
    kind: 'set_cols_hidden', group: 'Layout',
    props: { column: { type: 'number', description: '1-based' }, count: { type: 'number', description: 'Default 1.' }, hidden: { type: 'boolean' } },
    required: ['column', 'hidden'],
  },
  {
    kind: 'set_freeze', group: 'Layout',
    props: { rows: { type: 'number' }, columns: { type: 'number' } },
    required: ['rows', 'columns'],
    note: 'rows:0, columns:0 unfreezes.',
  },
  {
    kind: 'set_page_setup', group: 'Layout',
    props: {
      orientation: { type: 'string', enum: ['portrait', 'landscape'] },
      scale: { type: 'number', minimum: 10, maximum: 400, description: 'Mutually exclusive with fitToWidth/fitToHeight (Excel\'s own UI rule) - combining them now errors instead of silently dropping scale.' },
      fitToWidth: { type: 'number', description: 'Pages wide. Mutually exclusive with scale.' },
      fitToHeight: { type: 'number', description: 'Pages tall. 0 means unlimited (fit to width only) - the common "fit on one page wide" request. Mutually exclusive with scale.' },
      printGridlines: { type: 'boolean' },
      printHeadings: { type: 'boolean' },
      printArea: { type: 'string' },
      margins: { type: 'string', enum: ['normal', 'wide', 'narrow'] },
    },
  },

  // --- Structure ---
  { kind: 'insert_rows', group: 'Structure', props: { startRow: { type: 'number', description: '1-based' }, count: { type: 'number' } }, required: ['startRow', 'count'] },
  { kind: 'delete_rows', group: 'Structure', props: { startRow: { type: 'number', description: '1-based' }, count: { type: 'number' } }, required: ['startRow', 'count'] },
  { kind: 'insert_cols', group: 'Structure', props: { startCol: { type: 'number', description: '1-based' }, count: { type: 'number' } }, required: ['startCol', 'count'] },
  { kind: 'delete_cols', group: 'Structure', props: { startCol: { type: 'number', description: '1-based' }, count: { type: 'number' } }, required: ['startCol', 'count'] },
  { kind: 'add_sheet', group: 'Structure', props: { name: { type: 'string' } }, required: ['name'] },
  { kind: 'delete_sheet', group: 'Structure', props: {} },
  { kind: 'duplicate_sheet', group: 'Structure', props: { name: { type: 'string' } } },
  { kind: 'set_sheet_hidden', group: 'Structure', props: { hidden: { type: 'boolean' } }, required: ['hidden'] },
  { kind: 'move_sheet', group: 'Structure', props: { position: { type: 'number', description: '1-based target position' } }, required: ['position'] },
  { kind: 'protect_sheet', group: 'Structure', props: { protected: { type: 'boolean' } }, required: ['protected'] },
  { kind: 'rename_sheet', group: 'Structure', props: { name: { type: 'string' } }, required: ['name'] },

  // --- Charts/visuals ---
  {
    kind: 'add_chart', group: 'Charts/visuals',
    props: {
      dataRange: { type: 'string', description: 'The series values to plot (e.g. "B1:C10" including header row for series names). If the range meant for the x-axis is text (dates/names), you can instead include it as the first column of this same range and Excel will usually pick it up as categories automatically. If that column is numeric, Excel cannot tell it apart from another series - use categoryRange to force it onto the x-axis instead of a plain 1,2,3 index.' },
      categoryRange: { type: 'string', description: 'A single column or row of cells (data only, no header, e.g. "A2:A10") to use as the x-axis/category labels, overriding Excel\'s auto-detected categories. Use this whenever the intended x-axis column is numeric (Excel otherwise treats it as just another series and falls back to a plain row-index x-axis).' },
      chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'], description: 'Default column. Unrecognized values error rather than silently becoming column.' },
      title: { type: 'string' },
      name: { type: 'string', description: 'Optional stable chart name; omit to use Excel\'s auto-assigned name (returned by this call - use it with edit_chart\'s chartPath). Legend/colors/data-label styling is NOT accepted here - set those afterwards via edit_chart.' },
    },
    required: ['dataRange'],
  },
  {
    kind: 'edit_chart', group: 'Charts/visuals',
    props: {
      chartPath: { type: 'string', description: 'The chart\'s name (as shown in Excel\'s Name Box / Selection Pane, or returned by add_chart / read_sheet_features).' },
      chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
      dataRange: { type: 'string', description: 'Rebinds the chart to this range (e.g. after a table grows). Applied before chartType/title/legend/dataLabels/seriesColors/seriesData in the same call, so one call can rebind and restyle together.' },
      dataSheet: { type: 'string', description: 'Sheet dataRange lives on, if different from the chart\'s own sheet.' },
      plotBy: { type: 'string', enum: ['rows', 'columns'], description: 'Only meaningful together with dataRange; omit to let Excel infer from the range shape.' },
      categoryRange: { type: 'string', description: 'A single column or row of cells (data only, no header, e.g. "A2:A10") to force onto the x-axis/category labels for every series - use this if the chart is showing a plain 1,2,3 index instead of one of your data columns (typically because that column is numeric, which Excel cannot auto-distinguish from another series). Independent of dataRange/plotBy; can be set on its own.' },
      title: { type: 'string' },
      legend: { type: 'string', enum: ['none', 'right', 'top', 'left', 'bottom'] },
      dataLabels: { type: 'string', enum: ['none', 'value', 'percent'] },
      seriesColors: { type: 'object', description: 'Map of 0-based series index (as a string key) to hex color.' },
      seriesData: { type: 'array', items: { type: 'object', properties: { name: { type: 'string', description: 'A literal name, or a cell reference formula like "=Sheet1!$B$1" so the legend follows that cell.' } } } },
    },
    required: ['chartPath'],
  },
  { kind: 'delete_visual', group: 'Charts/visuals', props: { visualId: { type: 'string', description: 'Shape or chart name.' } }, required: ['visualId'] },
  {
    kind: 'add_sparkline', group: 'Charts/visuals',
    props: {
      dataRange: { type: 'string' },
      targetCell: { type: 'string', description: 'Required - must not overlap dataRange. A single cell (one sparkline for the whole range) or a range matching dataRange\'s row count (one sparkline per row).' },
      type: { type: 'string', enum: ['line', 'column', 'stacked'], description: 'Default line.' },
      color: { type: 'string' },
    },
    required: ['dataRange', 'targetCell'],
  },
  {
    kind: 'add_shape', group: 'Charts/visuals',
    props: {
      shapeType: {
        type: 'string',
        enum: [
          'textbox', 'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'parallelogram', 'trapezoid',
          'diamond', 'pentagon', 'hexagon', 'octagon', 'pie', 'chord', 'donut', 'foldedCorner', 'heart',
          'lightningBolt', 'sun', 'moon', 'cloud', 'arc', 'star5', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
        ],
      },
      anchorCell: { type: 'string' },
      fillColor: { type: 'string' },
      text: { type: 'string' },
      name: { type: 'string', description: 'Optional stable name to set on the shape, so a later edit_shape/delete_visual in the same batch can address it without needing this call\'s returned name.' },
    },
    required: ['shapeType', 'anchorCell'],
  },
  {
    kind: 'edit_shape', group: 'Charts/visuals',
    props: {
      visualId: { type: 'string' },
      text: { type: 'string' },
      fillColor: { type: 'string' },
      anchorCell: { type: 'string' },
    },
    required: ['visualId'],
  },
  {
    kind: 'add_image', group: 'Charts/visuals',
    props: { path: { type: 'string', description: 'LOCAL FILE PATH ONLY - no URLs (air-gapped deployment).' }, anchorCell: { type: 'string' } },
    required: ['path', 'anchorCell'],
  },

  // --- Tables ---
  {
    kind: 'add_table', group: 'Tables',
    props: { range: { type: 'string' }, name: { type: 'string' }, style: { type: 'string' }, bandedRows: { type: 'boolean' } },
    required: ['range'],
  },
  {
    kind: 'add_table_row', group: 'Tables',
    props: { tableName: { type: 'string' }, row: { type: 'number', description: '0-based; omit to append.' }, count: { type: 'number', description: 'Default 1.' } },
    required: ['tableName'],
  },
  {
    kind: 'add_table_column', group: 'Tables',
    props: { tableName: { type: 'string' }, columnName: { type: 'string' }, column: { type: 'number', description: '0-based; omit to append.' }, count: { type: 'number', description: 'Default 1.' } },
    required: ['tableName', 'columnName'],
  },
  {
    kind: 'delete_table_row', group: 'Tables',
    props: { tableName: { type: 'string' }, row: { type: 'number', description: '0-based' }, count: { type: 'number', description: 'Default 1.' } },
    required: ['tableName', 'row'],
  },
  {
    kind: 'delete_table_column', group: 'Tables',
    props: { tableName: { type: 'string' }, column: { type: 'number', description: '0-based' }, count: { type: 'number', description: 'Default 1.' } },
    required: ['tableName', 'column'],
  },
  {
    kind: 'delete_table', group: 'Tables',
    props: {
      tableName: { type: 'string' },
      deleteData: { type: 'boolean', description: 'Default false: converts the table back to a plain range (Excel\'s Unlist), keeping all data/formatting. Pass true to also remove the cells.' },
      shift: { type: 'string', enum: ['up', 'left', 'none'], description: 'Only used with deleteData:true. Default up. "none" clears the cells in place instead of shifting neighbors.' },
    },
    required: ['tableName'],
    note: 'Default converts the table to a plain range and keeps all data - "delete" alone does not remove cells.',
  },

  // --- Pivot ---
  {
    kind: 'add_pivot', group: 'Pivot',
    props: {
      sourceRange: { type: 'string' },
      targetCell: { type: 'string' },
      targetSheetId: { type: 'string' },
      name: { type: 'string' },
      rowFields: { description: 'A single field name or an array of field names.' },
      columnField: { type: 'string' },
      pageFields: { type: 'array', items: { type: 'string' } },
      values: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            field: { type: 'string' },
            agg: { type: 'string', enum: ['sum', 'count', 'average', 'max', 'min'], description: 'Default sum.' },
            formula: { type: 'string', description: 'Adds a calculated field instead of aggregating an existing one.' },
            numFmt: { type: 'string' },
          },
          required: ['field'],
        },
      },
    },
    required: ['sourceRange', 'targetCell', 'values'],
  },
  { kind: 'refresh_pivot', group: 'Pivot', props: {} },

  // --- Data ---
  {
    kind: 'set_hyperlink', group: 'Data',
    props: { address: { type: 'string' }, target: { type: 'string', description: 'Omit (or null) to remove the existing hyperlink.' } },
    required: ['address'],
  },
  {
    kind: 'set_note', group: 'Data',
    props: { address: { type: 'string' }, text: { type: 'string', description: 'Omit (or null) to remove the existing note.' } },
    required: ['address'],
  },
  {
    kind: 'add_defined_name', group: 'Data',
    props: {
      name: { type: 'string' },
      ref: { type: 'string', description: 'e.g. "B2:B20" or "Sheet1!$B$2:$B$20". An unqualified range is auto-qualified to the target sheet (scope:"sheet"\'s sheet, or sheet?, or the active sheet) - pass a qualified reference to be explicit.' },
      scope: { type: 'string', enum: ['workbook', 'sheet'], description: 'Default workbook. read_sheet_features reports both scopes\' existing names, labeled, so check there before acting.' },
      overwrite: { type: 'boolean', description: 'Default false: an existing name in the target scope errors rather than being silently replaced.' },
    },
    required: ['name', 'ref'],
  },
  {
    kind: 'delete_defined_name', group: 'Data',
    props: {
      name: { type: 'string' },
      scope: { type: 'string', enum: ['workbook', 'sheet'], description: 'Default workbook.' },
    },
    required: ['name'],
  },
  { kind: 'set_filter', group: 'Data', props: { range: { type: 'string' } }, required: ['range'] },
  { kind: 'clear_filter', group: 'Data', props: {} },
  {
    kind: 'set_filter_criteria', group: 'Data',
    props: {
      column: { type: 'number', description: '0-based, relative to the AutoFilter range\'s first column.' },
      values: { type: 'array', items: { type: 'string' }, description: 'Omit (or null) to clear this column\'s filter.' },
    },
    required: ['column'],
  },
  {
    kind: 'add_conditional_format', group: 'Data',
    props: { range: { type: 'string' }, rule: CF_RULE_SCHEMA },
    required: ['range', 'rule'],
  },
  { kind: 'clear_conditional_formats', group: 'Data', props: {} },
  {
    kind: 'set_data_validation', group: 'Data',
    props: { range: { type: 'string' }, validation: DATA_VALIDATION_SCHEMA },
    required: ['range'],
    note: '"checkbox" is NOT supported (this PIA has no checkbox-cell Data Validation type) - it will error if attempted.',
  },
]

// Task 5: measured JSON.stringify(ALL_TOOLS) growth from adding full oneOf
// detail for all 51 kinds was ~4,870 tokens (cl100k_base) - over the plan's
// ~4k-token threshold. Per the plan's fallback, only the highest-traffic,
// highest-ambiguity kinds get full oneOf branches; every other kind collapses
// into one permissive branch carrying just its name (kind enum) + sheet. The
// generated prose (opsDescription) still documents every kind's exact fields
// regardless of which side of this split it's on - the schema-side cut is
// purely a token-budget optimization, not a documentation gap. Runtime
// validation in ExcelTools.cs (PP-5 Task 4) is unaffected either way, since
// it never consulted this TS schema in the first place.
const DETAILED_KINDS = new Set(['format_range', 'add_conditional_format', 'add_chart', 'edit_chart', 'add_shape', 'set_data_validation', 'add_pivot'])

function opSchemas(ops: OpSpec[]) {
  const detailed = ops.filter((o) => DETAILED_KINDS.has(o.kind))
  const collapsed = ops.filter((o) => !DETAILED_KINDS.has(o.kind))

  const detailedSchemas = detailed.map((o) => ({
    type: 'object',
    properties: { kind: { const: o.kind }, sheet: { type: 'string' }, ...o.props },
    required: ['kind', ...(o.required ?? [])],
  }))

  const collapsedSchema = {
    type: 'object',
    description:
      'Covers every operation kind not detailed above. Fields are not structurally validated here (kept out of the schema to control its size) - see the tool description for each kind\'s exact fields.',
    properties: {
      kind: { type: 'string', enum: collapsed.map((o) => o.kind) },
      sheet: { type: 'string' },
    },
    required: ['kind'],
  }

  return [...detailedSchemas, collapsedSchema]
}

function propsSummary(o: OpSpec): string {
  const keys = Object.keys(o.props)
  if (keys.length === 0) return 'sheet?'
  return 'sheet?, ' + keys.map((k) => (o.required?.includes(k) ? k : k + '?')).join(', ')
}

function opsDescription(ops: OpSpec[]): string {
  const order: OpSpec['group'][] = ['Writing', 'Formatting', 'Layout', 'Structure', 'Charts/visuals', 'Tables', 'Pivot', 'Data']
  const lines: string[] = []
  for (const group of order) {
    const specs = ops.filter((o) => o.group === group)
    if (specs.length === 0) continue
    const kindLines = specs.map((o) => `"${o.kind}" (${propsSummary(o)})` + (o.note ? ` - ${o.note}` : ''))
    lines.push(`${group}: ` + kindLines.join(', ') + '.')
  }
  return lines.join('\n')
}

const ALL_TOOLS = [
  {
    name: 'get_workbook_context',
    description: "Reads the active sheet's name, used range, and current selection address, plus every sheet in the workbook with its used range. Call this FIRST, before read_range/read_cells - it does not return cell values, but its used-range addresses tell you exactly how many rows/columns actually have data, so you can size your read_range calls precisely instead of guessing a range and over-fetching or missing rows.",
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_range',
    description: 'Reads cell values in a rectangular range (e.g. "A1:C10"), max 2000 cells. Optional sheet name defaults to the active sheet. Call get_workbook_context first to learn the sheet\'s actual used range rather than guessing how many rows to request.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_cells',
    description: 'Reads specific scattered cell addresses (e.g. ["A1","C5"]).',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, addresses: { type: 'array', items: { type: 'string' } } }, required: ['addresses'] },
  },
  {
    name: 'select_range',
    description: 'Activates a sheet and selects/navigates to a range - UI navigation only, no data change.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_formats',
    description: 'Reads only explicitly-formatted cells (bold/italic/underline/strikethrough/font name+size/number format/horizontal+vertical alignment/wrap/text rotation/indent/has-border) in a range, max 200 cells. Cells with entirely default formatting are omitted.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'read_sheet_features',
    description: 'Reports a sheet\'s AutoFilter range, freeze panes, conditional-format rule count, defined names (sheet- and workbook-scoped), hidden/protected state, shape/image count, and the addresses of every contiguous data block on the sheet (tables separated by at least one fully blank row or column) - use this to find multiple tables stacked or side-by-side on one sheet before reading them individually.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' } } },
  },
  {
    name: 'find_cells',
    description:
      'Searches cell values/formulas for a substring or regex, or scans for error-valued cells (#REF!, #DIV/0!, etc.) via errors_only. ' +
      'Searches the ACTIVE sheet only by default (like Ctrl+F\'s "Within: Sheet") - pass allSheets:true to search the whole workbook instead ("Within: Workbook"), or sheetId to name one specific sheet. ' +
      'look_in: "values"|"formulas"|"both" (default "both"). Requires query or errors_only. Capped at max_results.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string' }, regex: { type: 'boolean' }, look_in: { type: 'string' },
        sheetId: { type: 'string' }, allSheets: { type: 'boolean' }, errors_only: { type: 'boolean' }, max_results: { type: 'number' },
      },
      required: ['max_results'],
    },
  },
  {
    name: 'trace_precedents',
    description: 'Lists the cells a formula directly reads from (same-sheet only), flagging any that hold an error value.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'trace_dependents',
    description: 'Lists every formula on the same sheet that reads a given cell.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['address'] },
  },
  {
    name: 'propose_operations',
    description: 'Applies a batch of spreadsheet operations. Each operation is one of the kinds below - see the schema for exact per-kind fields.\n' + opsDescription(EXCEL_OPS),
    inputSchema: {
      type: 'object',
      properties: { operations: { type: 'array', items: { oneOf: opSchemas(EXCEL_OPS) } } },
      required: ['operations'],
    },
  },
]

const READ_ONLY_TOOL_NAMES = [
  'get_workbook_context', 'read_range', 'read_cells',
  'select_range', 'read_formats', 'read_sheet_features', 'find_cells', 'trace_precedents', 'trace_dependents',
]

// FT-1 Task 5: settings-screen labels/descriptions, one entry per tool above -
// short, user-facing sentences, not the model-facing `description`s above.
const EXCEL_TOOL_DISPLAY = {
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
}

startAddIn({
  skillId: 'excel-tools',
  tools: ALL_TOOLS,
  toolDisplay: EXCEL_TOOL_DISPLAY,
  systemPrompt:
    'You are an assistant running inside a VSTO Excel add-in. You can help the user with their active workbook. ' +
    'You have read tools (workbook context, ranges, individual cells, explicit cell formats, sheet features like ' +
    'filters/freeze-panes/defined-names, workbook-wide search including a native error-cell scan, and formula ' +
    'precedent/dependent tracing) and a single propose_operations gateway for all writes - covering cell/range edits, ' +
    'formatting, sorting, row/column sizing and visibility, freeze panes, page setup, sheet management (add/delete/' +
    'duplicate/hide/move/protect/rename), charts, sparklines, shapes and local images, native Excel Tables, native ' +
    'pivot tables, hyperlinks, cell notes, defined names, AutoFilter, conditional formatting, and data validation. ' +
    'Your available tools depend on the current editing mode (Read Only/Comment Only allow only the read tools; ' +
    'Track Changes/Full Autonomy allow everything) - only call tools currently offered to you. ' +
    'Use the tools when asked to inspect or modify the spreadsheet. ' +
    "Before reading cell data, call get_workbook_context first - it's cheap and tells you each sheet's actual " +
    'used range, so you can size read_range/read_cells calls to the real data instead of guessing a range that ' +
    'reads too many or too few rows. ' +
    'When a request needs a computation over sheet data (totals, averages, differences, growth rates, lookups, ' +
    "etc.), prefer writing a real Excel formula into an unused cell (propose_operations' set_formula) over " +
    "computing the number yourself and only stating it in chat - that way the result stays live in the sheet and " +
    "recalculates if the data changes. Confirm a cell is actually empty first (get_workbook_context's used range, " +
    'or read_sheet_features\' data-block addresses) before writing into it - e.g. an empty column right after the ' +
    'data, or empty rows right below it - so you never overwrite existing content.',
  starters: [
    { en: 'Summarize this sheet', he: 'סכם את הגיליון הזה' },
    { en: 'Add a totals row', he: 'הוסף שורת סיכום' },
    { en: 'Check the formulas', he: 'בדוק את הנוסחאות' },
  ],
  // Excel has no add_comment-equivalent tool yet, so Comment Only mode allows
  // the same read-only set as Read Only mode (documented gap) - leaving
  // commentOnlyExtraTools unset gives the same result as before this
  // migration, where both modes mapped to READ_ONLY_TOOLS.
  readOnlyTools: READ_ONLY_TOOL_NAMES,
  // FT-2 Task 4/5: the current cell selection is injected into per-turn
  // context (as an A1 address, not values - see ExcelAiAddIn/TaskPaneHost.cs's
  // OnSelectionChanged) and the scope-hint pill reads "Whole sheet" rather
  // than Word's "Whole document" when nothing is selected.
  useSelectionContext: true,
  scopeUnit: 'sheet',
})
