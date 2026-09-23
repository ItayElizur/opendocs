import { startAddIn } from '@officeai/app-shell'

// PowerPoint add-in: tool definitions, system prompt, and starters are the
// only app-specific pieces left here - everything else (WebView2 bridge,
// settings, transport, chat-UI mount, AgentLoop event plumbing) lives once in
// @officeai/app-shell (PP-0), shared with WordAiAddIn and ExcelAiAddIn.
//
// Note (PP-0 Task 6): tool filtering by editing mode is now handled by the
// shell's shared live-getter pattern (readOnlyTools below), the same
// mechanism Word and Excel use - replacing this file's previous
// onModeChange-driven `powerPointSkill.tools = ...` reassignment. Comment
// Only has no PowerPoint-specific extra tool, so (matching prior behavior)
// it gets the same set as Read Only - commentOnlyExtraTools is left unset.

const READER_TOOLS = [
  {
    name: 'get_deck_context',
    description: 'Reads a one-line-per-slide outline: slide index and a text preview of its shapes.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_slide',
    description: 'Reads full text of every shape on one slide (0-based index), plus its layout, transition, animation count, and speaker notes (if any). Shapes are listed back-to-front (z-order/stacking order) - use set_element_order to change it. A group shape is shown as one line with a child count - call read_group to see inside it.',
    inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
  },
  {
    name: 'read_group',
    description:
      'Lists the shapes inside a group, recursively (nested groups are expanded in place). shapeIndex points at the group - a 0-based number, or a dotted path like "3.1" for a group nested inside another group. ' +
      'Each child line is prefixed with its dotted path (e.g. "3.1.0"); pass that path as shapeIndex to set_element_text / set_element_style / set_element_fill / set_element_stroke to edit a child in place. ' +
      'Positional or structural edits (set_element_transform, set_element_order, delete_element, add_animation) require ungroup_element on the top-level group first.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: ['number', 'string'] } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'read_animations',
    description: 'Reads a slide\'s animations in play order, one per line (shape, effect, entrance/exit, trigger, timing). Needed before edit_animation, since animationIndex addresses this same order.',
    inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
  },
  {
    name: 'find_text',
    description:
      'Searches every slide\'s shape text (text boxes, placeholders, table cells, SmartArt node text) and speaker notes for a substring, or with regex:true a .NET regular expression - read-only, never modifies the deck. ' +
      'Returns each match as "[slide i, shape j] text" or "[slide i, notes] text", capped at max_results. Use this to find slideIndex/shapeIndex before set_element_text/replace_text instead of reading every slide via read_slide.',
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
    name: 'read_smartart',
    description:
      'Reads the node text of SmartArt diagrams on a slide, one "[i] text" line per node. ' +
      'Omit smartArtIndex to read every diagram on the slide in one call. ' +
      'Call this before an incremental edit_smartart change so you are working from current node indices.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        smartArtIndex: { type: 'number', description: 'Omit to read all diagrams on the slide.' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'list_layouts',
    description:
      'Lists every layout name in this presentation (e.g. "Title Slide", "Title and Content"), plus how many slides currently use each one - across every design/theme in the deck if it has more than one, not just the default. ' +
      'Call this before passing layoutName to add_master_element/read_master_elements/remove_master_element/set_master_element_transform/set_slide_layout, instead of guessing a name. ' +
      'A layoutName matching more than one layout (e.g. two designs both have a layout with that name) errors instead of guessing - pass a more specific or exact name.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_master_elements',
    description:
      'Lists the shapes placed directly on the Slide Master (index, name, kind, position/size, text if any) - not the shapes on any one slide. ' +
      'Pass layoutName to instead list one specific layout\'s own shapes (e.g. "Title Slide", "Title and Content") - matched against this presentation\'s real layout names (exact match preferred, else substring; matching more than one errors, e.g. two designs both having a layout with that name), across every design/theme in the deck; omit it for the Slide Master (the default, and what every slide inherits from unless its own layout overrides a shape). ' +
      'Shapes added here with add_master_element appear on every slide whose layout doesn\'t hide master shapes. Call this before remove_master_element/set_master_element_transform to find the right masterShapeIndex.',
    inputSchema: {
      type: 'object',
      properties: {
        layoutName: { type: 'string', description: 'Optional. A specific layout\'s name (substring match), instead of the Slide Master.' },
      },
    },
  },
]

const MUTATION_TOOLS = [
  {
    name: 'replace_text',
    description:
      'Replaces every occurrence of find with replace across the whole deck - every text-frame shape (text boxes, title/body placeholders) on every slide, plus speaker notes unless includeNotes:false. ' +
      'NOT table cells or SmartArt node text - use edit_table_cell / edit_smartart for those. regex:true treats find as a .NET regular expression (replace can use $1-style backreferences). Reports the number of occurrences actually replaced.',
    inputSchema: {
      type: 'object',
      properties: {
        find: { type: 'string' },
        replace: { type: 'string' },
        regex: { type: 'boolean' },
        matchCase: { type: 'boolean' },
        includeNotes: { type: 'boolean', description: 'Default true.' },
      },
      required: ['find', 'replace'],
    },
  },
  {
    name: 'set_element_text',
    description:
      'Replaces the text content of one shape (0-based slideIndex, 0-based shapeIndex within that slide). ' +
      'Never type a literal bullet character ("•", "-", "*") at the start of a line - many placeholders (e.g. a "Title and Content" layout\'s body) already render their own native bullet per paragraph, so a literal one produces two bullets per line. ' +
      'Pass bulleted:true/false to explicitly turn PowerPoint\'s real bullets on or off instead; omit it to leave the shape\'s existing bullet setting untouched (each "\\n"-separated line becomes its own paragraph either way).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_group).' },
        text: { type: 'string' },
        bulleted: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex', 'text'],
    },
  },
  {
    name: 'set_slide_notes',
    description: 'Replaces a slide\'s speaker notes (0-based slideIndex). Overwrites any existing notes text; read_slide shows the current notes first.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, text: { type: 'string' } },
      required: ['slideIndex', 'text'],
    },
  },
  {
    name: 'set_element_style',
    description:
      'Changes text formatting of one shape (its whole text range) without changing its text. ' +
      'Only the fields you provide are changed; anything omitted is left as-is.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_group).' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        shadow: { type: 'boolean' },
        fontSize: { type: 'number' },
        fontName: { type: 'string' },
        color: { type: 'string', description: 'Hex color, e.g. "#FF0000".' },
        alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
        baselineOffset: { type: 'string', enum: ['SUPERSCRIPT', 'SUBSCRIPT', 'NONE'] },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_element_transform',
    description: 'Moves/resizes/rotates one shape (values in points; rotation in degrees).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: 'number' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_element_order',
    description:
      'Changes a shape\'s stacking order (z-order) - which shapes appear in front of/behind others, not its position or size (use set_element_transform for that). ' +
      'kind: "bringToFront"/"sendToBack" (moves to the very top/bottom of the stack), "bringForward"/"sendBackward" (moves one step relative to overlapping shapes). ' +
      'Shifts other shapes\' shapeIndex on the same slide - re-read the slide (read_slide) before addressing another shape by index in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: 'number' },
        kind: { type: 'string', enum: ['bringToFront', 'sendToBack', 'bringForward', 'sendBackward'] },
      },
      required: ['slideIndex', 'shapeIndex', 'kind'],
    },
  },
  {
    name: 'add_text_box',
    description:
      'Creates a new text box on the given slide. A plain text box has no bullets by default (unlike a layout\'s content placeholder) - ' +
      'pass bulleted:true for a real bulleted list instead of typing a literal "•"/"-"/"*" at the start of each line.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        text: { type: 'string' },
        bulleted: { type: 'boolean' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_group - use a short, meaningful label (e.g. "CTA button").' },
      },
      required: ['slideIndex', 'left', 'top', 'width', 'height', 'text'],
    },
  },
  {
    name: 'add_shape',
    description:
      'Creates a shape with optional text. shapeType is one of the names below - "rect"/"ellipse" ' +
      '(canonical) or "rectangle"/"oval" (aliases, same shapes) plus 24 more presets. ' +
      'No fill/line parameters here - use set_element_fill/set_element_stroke afterward for those.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeType: {
          type: 'string',
          enum: [
            'rect', 'rectangle', 'roundRect', 'ellipse', 'oval', 'triangle', 'rtTriangle',
            'parallelogram', 'trapezoid', 'diamond', 'pentagon', 'hexagon', 'octagon',
            'pie', 'chord', 'donut', 'foldedCorner', 'heart', 'lightningBolt', 'sun', 'moon',
            'cloud', 'arc', 'star5', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
          ],
        },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        text: { type: 'string' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_group - use a short, meaningful label (e.g. "Arrow to Q3").' },
      },
      required: ['slideIndex', 'shapeType', 'left', 'top', 'width', 'height'],
    },
  },
  {
    name: 'delete_element',
    description: 'Deletes one shape from a slide.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_slide',
    description: 'Clones an existing slide\'s layout as a new blank (or templated) slide inserted right after it.',
    inputSchema: {
      type: 'object',
      properties: { sourceIndex: { type: 'number' }, clearText: { type: 'boolean' } },
      required: ['sourceIndex'],
    },
  },
  {
    name: 'delete_slide',
    description: 'Deletes one slide (0-based slideIndex). Slides after it shift DOWN by one index - re-read the deck with get_deck_context before deleting another slide in the same run. Cannot delete the last remaining slide.',
    inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
  },
  {
    name: 'move_slide',
    description: 'Moves a slide to a new 0-based position; other slides shift accordingly.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, toIndex: { type: 'number' } },
      required: ['slideIndex', 'toIndex'],
    },
  },
  {
    name: 'duplicate_slide',
    description: 'Inserts a copy of a slide (content included) directly after it. Use add_slide instead to create a new slide from a slide\'s layout without its content.',
    inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
  },
  {
    name: 'set_slide_layout',
    description:
      'Changes a slide\'s layout. kind:"classic" (default) uses layout, a fixed set: title, titleOnly, blank, text, twoColumnText, object, objectAndText, textAndObject, twoObjects, twoObjectsAndText, fourObjects, table, chart, sectionHeader, comparison, contentWithCaption, pictureWithCaption. ' +
      'kind:"custom" uses layoutName instead - free text, matched against this slide\'s own theme layouts (exact match preferred, else substring; an unmatched name errors listing the real available names, and a name matching more than one layout errors asking for a more specific one).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        kind: { type: 'string', enum: ['classic', 'custom'] },
        layout: {
          type: 'string',
          enum: ['title', 'titleOnly', 'blank', 'text', 'twoColumnText', 'object', 'objectAndText', 'textAndObject', 'twoObjects', 'twoObjectsAndText', 'fourObjects', 'table', 'chart', 'sectionHeader', 'comparison', 'contentWithCaption', 'pictureWithCaption'],
        },
        layoutName: { type: 'string' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'set_slide_transition',
    description:
      'Sets or removes a slide\'s entry transition. effect:"none" removes it. durationSeconds controls how long the transition animation itself takes. ' +
      'advanceOnClick (default true in PowerPoint) and advanceAfterSeconds (sets an automatic timed advance) are independent - both can be on at once.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        effect: {
          type: 'string',
          enum: ['none', 'cut', 'fade', 'dissolve', 'random', 'wipeLeft', 'wipeRight', 'wipeUp', 'wipeDown', 'pushLeft', 'pushRight', 'pushUp', 'pushDown', 'coverLeft', 'coverRight', 'coverUp', 'coverDown', 'uncoverLeft', 'uncoverRight', 'uncoverUp', 'uncoverDown', 'zoomIn', 'zoomOut', 'zoomCenter', 'circle', 'diamond', 'splitHorizontal', 'splitVertical', 'wheel', 'blindsHorizontal', 'blindsVertical', 'checkerboard'],
        },
        durationSeconds: { type: 'number' },
        advanceOnClick: { type: 'boolean' },
        advanceAfterSeconds: { type: 'number' },
      },
      required: ['slideIndex', 'effect'],
    },
  },
  {
    name: 'add_animation',
    description:
      'Adds an entrance (default) or exit animation to a shape. effect: appear, fade, fly, flashOnce, wipe, zoom, dissolve, bounce, spiral, swivel, wheel, split, box, circle, diamond, plus, checkerboard, randomBars, growAndTurn, riseUp. ' +
      'trigger (default "onClick"): "onClick" starts on its own click during the slideshow, "withPrevious" starts together with the animation before it, "afterPrevious" starts automatically once the animation before it finishes. ' +
      'Does not support directional variants (e.g. "wipe from the left") - only the base effect. Returns the new animationIndex (0-based) for read_animations/edit_animation.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: 'number' },
        effect: { type: 'string', enum: ['appear', 'fade', 'fly', 'flashOnce', 'wipe', 'zoom', 'dissolve', 'bounce', 'spiral', 'swivel', 'wheel', 'split', 'box', 'circle', 'diamond', 'plus', 'checkerboard', 'randomBars', 'growAndTurn', 'riseUp'] },
        exit: { type: 'boolean' },
        trigger: { type: 'string', enum: ['onClick', 'withPrevious', 'afterPrevious'] },
        durationSeconds: { type: 'number' },
        delaySeconds: { type: 'number' },
      },
      required: ['slideIndex', 'shapeIndex', 'effect'],
    },
  },
  {
    name: 'edit_animation',
    description:
      'Edits an existing animation. kind: "delete", "set_timing" (durationSeconds?,delaySeconds?,trigger?), "reorder" (toIndex - 0-based new position in the play sequence). ' +
      'animationIndex addresses the animation (0-based, current play order - call read_animations first). delete/reorder shift later indices - re-read before another edit in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        animationIndex: { type: 'number' },
        kind: { type: 'string', enum: ['delete', 'set_timing', 'reorder'] },
        durationSeconds: { type: 'number' },
        delaySeconds: { type: 'number' },
        trigger: { type: 'string', enum: ['onClick', 'withPrevious', 'afterPrevious'] },
        toIndex: { type: 'number' },
      },
      required: ['slideIndex', 'animationIndex', 'kind'],
    },
  },
  {
    name: 'set_element_fill',
    description: 'Sets a shape\'s solid fill color, or "none" to remove its fill.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_group).' },
        fill: { type: 'string' },
      },
      required: ['slideIndex', 'shapeIndex', 'fill'],
    },
  },
  {
    name: 'set_element_stroke',
    description: 'Sets a shape\'s outline/stroke color and width, or removes it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndex: { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_group).' },
        color: { type: 'string' }, widthPt: { type: 'number' }, remove: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_slide_background',
    description: 'Sets a solid background color for one slide, or slideIndex=-1 for every slide in the deck.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, color: { type: 'string' } },
      required: ['slideIndex', 'color'],
    },
  },
  {
    name: 'ungroup_element',
    description: 'Promotes a group shape\'s direct children to top-level shapes. Shape indices change after this call - re-read the slide before addressing the promoted shapes.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: ['number', 'string'] } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'group_element',
    description:
      'Groups two or more top-level shapes on a slide into a single group shape. shapeIndexes is an array of 0-based indices (as shown by read_slide). ' +
      'Every other shape\'s index on the slide shifts afterward - re-read the slide before addressing another shape by index in the same run. Use ungroup_element to reverse it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        shapeIndexes: { type: 'array', items: { type: 'number' }, description: 'Two or more distinct 0-based top-level shape indices.' },
        name: { type: 'string', description: 'Optional. Sets the new group\'s name, shown in read_slide/read_group - use a short, meaningful label (e.g. "Logo lockup").' },
      },
      required: ['slideIndex', 'shapeIndexes'],
    },
  },
  {
    name: 'add_table',
    description: 'Adds a native PowerPoint table, optionally pre-filled with cell text (row-major array of arrays).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, rows: { type: 'number' }, cols: { type: 'number' },
        cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_group.' },
      },
      required: ['slideIndex', 'rows', 'cols'],
    },
  },
  {
    name: 'edit_table_cell',
    description: 'Replaces one table cell\'s text (0-based row/col).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, row: { type: 'number' }, col: { type: 'number' }, paragraphs: { type: 'string' } },
      required: ['slideIndex', 'shapeIndex', 'row', 'col', 'paragraphs'],
    },
  },
  {
    name: 'edit_table_structure',
    description: 'Inserts or deletes a table row/column. index (0-based) addresses an EXISTING row/column; before decides which side the new one goes on for insert kinds. Deleting/inserting shifts every later row/column\'s index - re-read the table before a second structural edit in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        kind: { type: 'string', enum: ['insert-row', 'delete-row', 'insert-col', 'delete-col'] },
        index: { type: 'number' }, before: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex', 'kind', 'index'],
    },
  },
  {
    name: 'edit_table_style',
    description: 'Applies granular table styling: firstRow/bandRow (header row / banded rows), shadingColor (all cells), borderColor/borderWidthPt/borderPreset ("all" = every cell edge, "outline" = only the table\'s outer perimeter, "none" = remove borders).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        firstRow: { type: 'boolean' }, bandRow: { type: 'boolean' }, shadingColor: { type: 'string' },
        borderColor: { type: 'string' }, borderWidthPt: { type: 'number' },
        borderPreset: { type: 'string', enum: ['all', 'outline', 'none'] },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_chart',
    description: 'Adds a native, editable PowerPoint chart, returning its shapeIndex for a follow-up edit_chart call. Every series\' values array must have exactly one value per category.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        kind: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
        title: { type: 'string' },
        categories: { type: 'array', items: { type: 'string' } },
        series: { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } } } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        name: { type: 'string', description: 'Optional. Sets the chart shape\'s name, shown in read_slide/read_group.' },
      },
      required: ['slideIndex', 'kind', 'categories', 'series'],
    },
  },
  {
    name: 'edit_chart',
    description: 'Modifies an existing chart\'s type/title/legend position/data labels/gridlines. Only properties provided are changed; the result names exactly what applied.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' }, shapeIndex: { type: 'number' },
        chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
        title: { type: 'string' },
        legendPos: { type: 'string', enum: ['none', 'right', 'top', 'left', 'bottom', 'r', 't', 'l', 'b'] },
        dataLabels: { type: 'string', enum: ['none', 'value', 'percent'] },
        gridlines: { type: 'boolean', description: 'Errors with a clear message on chart types with no value axis (pie/doughnut).' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_smartart',
    description: 'Adds a shape-composed SmartArt diagram. items are flat node texts, one per top-level node.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'] },
        items: { type: 'array', items: { type: 'string' } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_group.' },
      },
      required: ['slideIndex', 'layout', 'items'],
    },
  },
  {
    name: 'edit_smartart',
    description:
      'Edits an existing SmartArt diagram on a slide. kind is one of: set_text (needs nodeIndex + text), add_node (optional text, appends), ' +
      'delete_node (needs nodeIndex), set_style (needs colorName and/or quickStyleName, matched loosely against this Office install\'s gallery), ' +
      'set_layout (needs layout). smartArtIndex is 0-based WITHIN THAT SLIDE and defaults to 0. ' +
      'delete_node shifts later node indices - call read_smartart again before a second node edit on the same diagram in one run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number' },
        smartArtIndex: { type: 'number', description: '0-based among the SmartArt diagrams on that slide. Default 0.' },
        kind: { type: 'string', enum: ['set_text', 'add_node', 'delete_node', 'set_style', 'set_layout'] },
        nodeIndex: { type: 'number', description: '0-based. Required for set_text and delete_node.' },
        text: { type: 'string' },
        colorName: { type: 'string', description: 'set_style: substring-matched against the gallery, e.g. "Colorful".' },
        quickStyleName: { type: 'string', description: 'set_style: substring-matched, e.g. "Intense".' },
        layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'] },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'crop_image',
    description: 'Non-destructively crops a picture shape. l/t/r/b are 0..1 fractions of the current on-slide image size cut from each edge; all zero clears the crop.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, l: { type: 'number' }, t: { type: 'number' }, r: { type: 'number' }, b: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'l', 't', 'r', 'b'],
    },
  },
  {
    name: 'set_picture_opacity',
    description: 'Sets a picture shape\'s overall opacity, 0 (invisible) to 1 (fully opaque).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, opacity: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'opacity'],
    },
  },
  {
    name: 'replace_image',
    description: 'Swaps a picture shape\'s image content in place from a local file path, keeping position/size/rotation/approximate z-order.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, localPath: { type: 'string' }, keepCrop: { type: 'boolean' } },
      required: ['slideIndex', 'shapeIndex', 'localPath'],
    },
  },
  {
    name: 'set_headers_footers',
    description:
      'Sets slide number, footer text, and date/time - the same 3 toggles as PowerPoint\'s native "Insert Header and Footer" dialog (there is no per-slide "header", only slide number/footer/date). ' +
      'slideIndex:-1 applies to every slide in the deck; a specific 0-based slideIndex applies to just that one. Only the fields you provide are changed. ' +
      'footerText implies footerVisible:true unless you also pass footerVisible explicitly; dateText implies dateMode:"fixed" the same way. ' +
      'dateMode defaults to and stays "fixed" unless you pass dateMode:"auto" explicitly - it is never inferred from dateFormat. ' +
      'dateMode "auto" shows a self-updating today\'s-date (in dateFormat, default "M/d/yy" - requires dateMode:"auto" to have any effect), "fixed" shows dateText verbatim. ' +
      'skipTitleSlide and startNumber are both ALWAYS deck-wide, regardless of slideIndex - PowerPoint has no per-slide version of either (skipTitleSlide sets the Slide/Title Master\'s "Don\'t show on title slide" flag, startNumber sets the deck\'s first slide number). ' +
      'If nothing appears after enabling a toggle, the slide\'s layout may simply not include that placeholder - try set_slide_layout or check read_master_elements.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '0-based, or -1 for the whole deck.' },
        slideNumberVisible: { type: 'boolean' },
        footerVisible: { type: 'boolean' },
        footerText: { type: 'string' },
        dateVisible: { type: 'boolean' },
        dateMode: { type: 'string', enum: ['auto', 'fixed'] },
        dateText: { type: 'string', description: 'Required when dateMode is "fixed". Implies dateMode:"fixed" if dateMode is omitted.' },
        dateFormat: {
          type: 'string',
          enum: ['M/d/yy', 'dddd, MMMM dd, yyyy', 'd MMMM, yyyy', 'MMMM d, yyyy', 'd-MMM-yy', 'MMMM yy', 'MM/yy'],
          description: 'Only applies when dateMode is "auto" - you must pass dateMode:"auto" explicitly, it is not inferred. Default "M/d/yy" if omitted.',
        },
        skipTitleSlide: { type: 'boolean', description: 'Always deck-wide (see description) - not scoped by slideIndex.' },
        startNumber: { type: 'number', description: 'Always deck-wide (see description) - not scoped by slideIndex.' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'add_master_element',
    description:
      'Adds a persistent icon (picture) or small text label to the Slide Master, so it shows up in the same spot on every slide (e.g. a logo/watermark/confidentiality label) - not just one slide. ' +
      'Pass layoutName to add it to one specific layout instead (e.g. "Title Slide") - it then only shows on slides using that layout, not deck-wide; matched against this presentation\'s real layout names across every design/theme in the deck (exact match preferred, else substring; matching more than one, e.g. two designs both having a layout with that name, errors instead of guessing). ' +
      'Use corner for a quick topLeft/topRight/bottomLeft/bottomRight placement (inset by marginPt, default 12pt), or left+top together for an exact position (left/top must both be given, or neither). ' +
      'For kind:"icon", width/height are optional - omit one and it scales proportionally from the image\'s natural size; omit both to use the natural size. ' +
      'Known limits: with no layoutName, only this presentation\'s default (first) Slide Master is affected - a deck combining more than one theme has more than one, and layoutName is the way to reach a non-default one\'s layout directly; and a layout with "Hide Background Graphics" enabled won\'t show a Slide-Master-level element.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['icon', 'text'] },
        layoutName: { type: 'string', description: 'Optional. Adds to one specific layout (matched across every design/theme in the deck) instead of the Slide Master.' },
        corner: { type: 'string', enum: ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'], description: 'Ignored if left+top are both given. Required if they aren\'t.' },
        left: { type: 'number', description: 'Must be given together with top (or neither) - a single coordinate alone is rejected.' },
        top: { type: 'number', description: 'Must be given together with left (or neither) - a single coordinate alone is rejected.' },
        width: { type: 'number' },
        height: { type: 'number' },
        marginPt: { type: 'number', description: 'Default 12. Inset from the slide edge, only used with corner.' },
        localPath: { type: 'string', description: 'Required for kind:"icon" - a local image file path.' },
        text: { type: 'string', description: 'Required for kind:"text".' },
        fontSize: { type: 'number', description: 'kind:"text" only. Default 10.' },
        color: { type: 'string', description: 'kind:"text" only. Hex color, e.g. "#808080" (the default).' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_master_elements.' },
      },
      required: ['kind'],
    },
  },
  {
    name: 'set_master_element_transform',
    description:
      'Moves/resizes/rotates an existing Slide Master element by masterShapeIndex (from read_master_elements). Only the fields you provide are changed. ' +
      'Works on a theme placeholder too (e.g. repositioning where the footer shows) - unlike remove_master_element, this is not refused, since moving/resizing a placeholder is a normal, reversible edit. ' +
      'Pass layoutName to target one specific layout\'s shape instead of the Slide Master\'s - must match what you used (or omitted) in read_master_elements, since masterShapeIndex is only valid within that same target. Matched across every design/theme in the deck (exact match preferred, else substring; matching more than one errors).',
    inputSchema: {
      type: 'object',
      properties: {
        masterShapeIndex: { type: 'number' },
        layoutName: { type: 'string', description: 'Optional. Targets one specific layout (matched across every design/theme in the deck) instead of the Slide Master.' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
      },
      required: ['masterShapeIndex'],
    },
  },
  {
    name: 'remove_master_element',
    description:
      'Deletes one shape from the Slide Master by the 0-based masterShapeIndex read_master_elements reports. Refuses (with an error) to delete a theme placeholder - read_master_elements marks those "(placeholder)"; only delete a shape add_master_element actually created. ' +
      'To turn a slide number/footer/date placeholder off instead of deleting it, use set_headers_footers with the matching *Visible:false field - that\'s the correct way to "remove" one, matching PowerPoint\'s own native behavior. ' +
      'Pass layoutName to target one specific layout instead of the Slide Master - must match what you used (or omitted) in read_master_elements. Matched across every design/theme in the deck (exact match preferred, else substring; matching more than one errors). Other master shapes\' indices may shift afterward - re-read before another edit in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        masterShapeIndex: { type: 'number' },
        layoutName: { type: 'string', description: 'Optional. Targets one specific layout (matched across every design/theme in the deck) instead of the Slide Master.' },
      },
      required: ['masterShapeIndex'],
    },
  },
]

const ALL_TOOLS = [...READER_TOOLS, ...MUTATION_TOOLS]

// FT-1 Task 5: settings-screen labels/descriptions, one entry per tool above -
// short, user-facing sentences, not the model-facing `description`s above.
const POWERPOINT_TOOL_DISPLAY = {
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
  add_slide: {
    label: { en: 'Add slide', he: 'הוספת שקופית' },
    description: { en: "Adds a new slide, based on an existing slide's layout.", he: 'מוסיף שקופית חדשה, בהתבסס על פריסה של שקופית קיימת.' },
  },
  delete_slide: {
    label: { en: 'Delete slide', he: 'מחיקת שקופית' },
    description: { en: 'Deletes a slide from the presentation.', he: 'מוחק שקופית מהמצגת.' },
  },
  move_slide: {
    label: { en: 'Move slide', he: 'העברת שקופית' },
    description: { en: 'Moves a slide to a different position in the presentation.', he: 'מעביר שקופית למיקום אחר במצגת.' },
  },
  duplicate_slide: {
    label: { en: 'Duplicate slide', he: 'שכפול שקופית' },
    description: { en: 'Inserts a copy of a slide, including its content.', he: 'מוסיף עותק של שקופית, כולל תוכנה.' },
  },
  set_slide_layout: {
    label: { en: 'Change slide layout', he: 'שינוי פריסת שקופית' },
    description: { en: "Changes a slide's layout (e.g. title, blank, two content).", he: 'משנה את פריסת השקופית (למשל כותרת, ריקה, שני תכנים).' },
  },
  set_slide_transition: {
    label: { en: 'Set slide transition', he: 'הגדרת מעבר שקופית' },
    description: { en: 'Sets or removes the transition effect shown when this slide appears.', he: 'מגדיר או מסיר את אפקט המעבר המוצג בכניסה לשקופית.' },
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
  set_slide_background: {
    label: { en: 'Set slide background', he: 'רקע שקופית' },
    description: { en: 'Sets a solid background color for one slide or the whole deck.', he: 'מגדיר צבע רקע אחיד לשקופית אחת או לכל המצגת.' },
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
}

startAddIn({
  skillId: 'powerpoint-tools',
  tools: ALL_TOOLS,
  toolDisplay: POWERPOINT_TOOL_DISPLAY,
  systemPrompt:
    'You are an AI assistant running inside a VSTO PowerPoint add-in. ' +
    'You can read the deck outline (get_deck_context) and the full text of any slide (read_slide). ' +
    'You can search every slide\'s text and speaker notes for a word or phrase with find_text (read-only), and replace every occurrence across the whole deck with replace_text - use find_text first to confirm what you\'re about to change and locate slideIndex/shapeIndex, instead of reading every slide via read_slide. ' +
    'You can edit text and shapes: set_element_text, set_element_style, set_element_transform, set_element_order (stacking/z-order - which shape is drawn in front of which), add_text_box, add_shape, and delete_element. ' +
    'For bulleted text, use set_element_text/add_text_box\'s bulleted:true/false parameter - NEVER type a literal "•"/"-"/"*" at the start of a line yourself. Many placeholders (e.g. a "Title and Content" layout\'s body) already render their own native bullet per paragraph, so typing one too produces two bullets per line. ' +
    'You can manage slides and shape styling: add_slide, delete_slide, move_slide, duplicate_slide, set_element_fill, set_element_stroke, and set_slide_background. ' +
    'You can group and ungroup shapes: group_element (two or more shapeIndexes into one group) and ungroup_element. ' +
    'read_slide shows a group as one line; call read_group to list its contents recursively. Each child line gives a dotted path (e.g. "3.1.0") that you can pass as shapeIndex to set_element_text/set_element_style/set_element_fill/set_element_stroke to edit that child in place. To move, resize, reorder, delete, or animate a shape inside a group, call ungroup_element on the top-level group first. ' +
    'Whenever you add a shape (add_text_box, add_shape, add_table, add_chart, add_smartart) or create a group, pass a short, meaningful name so later read_slide/read_group output is self-explanatory. ' +
    'You can add and edit tables: add_table, edit_table_cell, edit_table_structure, and edit_table_style. ' +
    'You can create and edit charts: add_chart and edit_chart. ' +
    'You can add, read and edit SmartArt diagrams: add_smartart, read_smartart, edit_smartart. ' +
    'edit_smartart\'s delete_node shifts later node indices - call read_smartart again before a second node edit on the same diagram in one run. ' +
    'You can work with images: crop_image, set_picture_opacity, and replace_image. ' +
    'You can change a slide\'s layout (set_slide_layout), set or remove its transition (set_slide_transition), and add, read, or edit shape animations (add_animation, read_animations, edit_animation). ' +
    'edit_animation\'s delete/reorder shift later animation indices - call read_animations again before a second animation edit on the same slide in the same run. ' +
    'You can manage Slide Master basics: set_headers_footers (slide number/footer/date, deck-wide or per slide - mirrors PowerPoint\'s native "Insert Header and Footer" dialog), ' +
    'and add_master_element/read_master_elements/set_master_element_transform/remove_master_element to pin a logo, icon, or small text label to a corner of every slide via the Slide Master rather than editing each slide individually. ' +
    'Any of those four tools can also target one specific layout instead of the whole Slide Master, via an optional layoutName (matched across every design/theme in the deck) - call list_layouts first to see the real names, or check the selected slide\'s layoutName already given in your context. ' +
    'Your available tools depend on the current editing mode (Read Only, Comment Only, Track Changes, or Full Autonomy) - only call tools that are currently offered to you.',
  starters: [
    { en: "Improve this slide's title and copy", he: 'שפר את הכותרת והטקסט של השקופית' },
    { en: "Make this slide's bullets more concise", he: 'קצר את התבליטים בשקופית' },
    { en: 'Check the whole deck for typos and fix them', he: 'בדוק שגיאות כתיב בכל המצגת ותקן אותן' },
    { en: 'Turn on slide numbers, but skip the title slide', he: 'הפעל מספור שקופיות, אך דלג על שקופית הכותרת' },
    { en: 'Add our logo to the bottom-right corner of every slide', he: 'הוסף את הלוגו שלנו לפינה הימנית התחתונה בכל שקופית' },
  ],
  readOnlyTools: READER_TOOLS.map((t) => t.name),
  // FT-2 Task 4/5: the current slide/shape/text selection is injected into
  // per-turn context as slideIndex/shapeIndex (see PowerPointAiAddIn/
  // TaskPaneHost.cs's OnSelectionChanged) and the scope-hint pill reads
  // "Whole deck" rather than Word's "Whole document" when nothing is selected.
  useSelectionContext: true,
  scopeUnit: 'deck',
})
