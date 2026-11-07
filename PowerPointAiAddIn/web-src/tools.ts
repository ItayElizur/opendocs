const SLIDE_INDEX = { type: 'number' };
const createRect = (defaultHeight: number) => ({
  left: { type: 'number', description: 'create only. Default 100.' },
  top: { type: 'number', description: 'create only. Default 100.' },
  width: { type: 'number', description: 'create only. Default 400.' },
  height: { type: 'number', description: `create only. Default ${defaultHeight}.` },
})
const SHAPE_PATH = { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_slide (view=group)).' };

export const READER_TOOLS = [
  {
    name: 'get_deck_context',
    description: 'Reads a one-line-per-slide outline: slide index and a text preview of its shapes.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_slide',
    description:
      'Reads one slide (0-based slideIndex); view picks what: "slide" (default), "group", "animations", or "smartart". ' +
      '"slide": the text of every shape plus the layout, transition, animation count and speaker notes. Shapes are listed back-to-front (z-order); change it with set_element_transform order. A group appears as one line with a child count. ' +
      '"group" (needs shapeIndex, a number or a dotted path like "3.1" for a nested group): lists the group\'s children recursively, each with its dotted path (e.g. "3.1.0") - pass that path as shapeIndex to set_element_text / format_element to edit a child in place. Moving, resizing, reordering, deleting or animating a child needs group_element with action "ungroup" first. ' +
      '"animations": the slide\'s animations in play order (shape, effect, entrance/exit, trigger, timing); animationIndex in edit_animation addresses this order, so read it first. ' +
      '"smartart" (optional smartArtIndex, omit for all): node text as "[i] text" lines; read it before an edit_smartart so the node indices are current.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        view: { type: 'string', enum: ['slide', 'group', 'animations', 'smartart'], description: 'Default "slide".' },
        shapeIndex: { type: ['number', 'string'], description: 'view "group" only: the group, a 0-based number or a dotted path like "3.1" for a group nested in another.' },
        smartArtIndex: { type: 'number', description: 'view "smartart" only. Omit to read all diagrams on the slide.' },
      },
      required: ['slideIndex'],
    },
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
    name: 'read_master',
    description:
      'Reads the Slide Master; view: "elements" (default) or "layouts". ' +
      '"elements": lists the shapes placed directly on the Slide Master (index, name, kind, position/size, text) - not any one slide\'s shapes. Pass layoutName to list one layout\'s own shapes instead (e.g. "Title Slide"): exact name preferred, else substring, matched across every design/theme; an ambiguous name errors. Elements added with edit_master_element appear on every slide whose layout doesn\'t hide master shapes. Call this before edit_master_element to find the masterShapeIndex. ' +
      '"layouts": lists every layout name in the deck, across all designs/themes, with how many slides use each. Call it before passing a layoutName anywhere instead of guessing; an ambiguous layoutName errors, so pass a more specific or exact name.',
    inputSchema: {
      type: 'object',
      properties: {
        view: { type: 'string', enum: ['elements', 'layouts'], description: 'Default "elements".' },
        layoutName: { type: 'string', description: 'view "elements" only. Optional. A specific layout\'s name (substring match), instead of the Slide Master.' },
      },
    },
  },
]

const MUTATION_TOOLS = [
  {
    name: 'replace_text',
    description:
      'Replaces every occurrence of find with replace across the whole deck - every text-frame shape (text boxes, title/body placeholders) on every slide, plus speaker notes unless includeNotes:false. ' +
      'NOT table cells or SmartArt node text - use edit_table / edit_smartart for those. regex:true treats find as a .NET regular expression (replace can use $1-style backreferences). Reports the number of occurrences actually replaced.',
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
        slideIndex: SLIDE_INDEX,
        shapeIndex: SHAPE_PATH,
        text: { type: 'string' },
        bulleted: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex', 'text'],
    },
  },
  {
    name: 'manage_slide',
    description:
      'Adds, duplicates, moves, or deletes one slide (0-based slideIndex). ' +
      'action:"add" inserts a new slide right after slideIndex with the same layout and every text frame cleared. ' +
      'action:"duplicate" inserts a full copy (content included) right after it. ' +
      'action:"move" moves it to the 0-based position toIndex; other slides shift accordingly. ' +
      'action:"delete" removes it; slides after it shift DOWN by one index, so re-read the deck with get_deck_context before deleting or editing another slide by index in the same run. The last remaining slide cannot be deleted.',
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['add', 'duplicate', 'move', 'delete'] },
        slideIndex: SLIDE_INDEX,
        toIndex: { type: 'number', description: 'For action:"move" only: the 0-based destination position.' },
      },
      required: ['action', 'slideIndex'],
    },
  },
  {
    name: 'set_slide',
    description:
      'Changes ONE thing about a slide per call (0-based slideIndex): layout, background, entry transition, or speaker notes. Fields from more than one of these error - make a separate call for each. ' +
      'Layout: layout (a classic preset) or layoutName (free text matched against the slide\'s theme layouts, exact preferred else substring; no match lists the real names, several matches asks for a more specific one) - one, not both. ' +
      'Background: a solid color; slideIndex -1 applies it to every slide (the only change that allows -1). ' +
      'Notes: replaces the speaker notes. ' +
      'Transition: effect sets or removes the transition ("none" removes it); durationSeconds, advanceOnClick (PowerPoint default true) and advanceAfterSeconds (timed advance) are independent, may be combined, and can change without restating effect. With both advance options off the slide advances only from the keyboard or menu.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '0-based, or -1 for the whole deck (background only).' },
        layout: {
          type: 'string',
          enum: ['title', 'titleOnly', 'blank', 'text', 'twoColumnText', 'object', 'objectAndText', 'textAndObject', 'twoObjects', 'twoObjectsAndText', 'fourObjects', 'table', 'chart', 'sectionHeader', 'comparison', 'contentWithCaption', 'pictureWithCaption'],
        },
        layoutName: { type: 'string' },
        background: { type: 'string', description: 'Solid color, e.g. "#F2F2F2".' },
        notes: { type: 'string' },
        effect: {
          type: 'string',
          enum: ['none', 'cut', 'fade', 'dissolve', 'random', 'wipeLeft', 'wipeRight', 'wipeUp', 'wipeDown', 'pushLeft', 'pushRight', 'pushUp', 'pushDown', 'coverLeft', 'coverRight', 'coverUp', 'coverDown', 'uncoverLeft', 'uncoverRight', 'uncoverUp', 'uncoverDown', 'zoomIn', 'zoomOut', 'zoomCenter', 'circle', 'diamond', 'splitHorizontal', 'splitVertical', 'wheel', 'blindsHorizontal', 'blindsVertical', 'checkerboard'],
        },
        durationSeconds: { type: 'number' },
        advanceOnClick: { type: 'boolean' },
        advanceAfterSeconds: { type: 'number' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'format_element',
    description:
      'Changes the formatting of one shape, one kind of formatting per call: give fields from only ONE of the groups below (mixing groups errors - make a separate call for each). Only the fields you provide are changed. ' +
      'Text style (the shape\'s whole text range, text itself unchanged): bold, italic, underline, shadow, fontSize, fontName, color (the TEXT color), alignment, baselineOffset. ' +
      'Fill: fill = a solid hex color, or "none" to remove the fill. ' +
      'Stroke (outline): strokeColor and strokeWidthPt, or removeStroke:true to remove it. ' +
      'Opacity: opacity from 0 (invisible) to 1 (fully opaque), for a picture\'s overall opacity.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: SHAPE_PATH,
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        shadow: { type: 'boolean' },
        fontSize: { type: 'number' },
        fontName: { type: 'string' },
        color: { type: 'string', description: 'Text color, hex e.g. "#FF0000".' },
        alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
        baselineOffset: { type: 'string', enum: ['SUPERSCRIPT', 'SUBSCRIPT', 'NONE'] },
        fill: { type: 'string', description: 'Solid fill hex color, or "none".' },
        strokeColor: { type: 'string', description: 'Outline hex color.' },
        strokeWidthPt: { type: 'number' },
        removeStroke: { type: 'boolean' },
        opacity: { type: 'number', description: '0 (invisible) to 1 (fully opaque).' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'set_element_transform',
    description:
      'Changes one shape\'s geometry or stacking, one per call (giving both errors - make a separate call). ' +
      'Position/size/rotation: any of left, top, width, height (points) and rotation (degrees). ' +
      'Order: order = "bringToFront"/"sendToBack" (moves to the very top/bottom of the stack) or "bringForward"/"sendBackward" (one step relative to overlapping shapes) - changes which shapes appear in front of/behind others, not position or size. ' +
      'Changing the order shifts other shapes\' shapeIndex on the same slide - re-read the slide (read_slide) before addressing another shape by index in the same run. Not for a shape inside a group - use group_element with action:"ungroup" first.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        left: { type: 'number' },
        top: { type: 'number' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
        order: { type: 'string', enum: ['bringToFront', 'sendToBack', 'bringForward', 'sendBackward'] },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'add_shape',
    description:
      'Creates a shape or a text box on the given slide. ' +
      'With shapeType, creates that preset shape with optional text; shapeType is one of the names below - "rect"/"ellipse" (canonical) or "rectangle"/"oval" (aliases, same shapes) plus 24 more presets. ' +
      'Without shapeType, creates a plain text box and text is required; a text box has no bullets by default (unlike a layout\'s content placeholder) - pass bulleted:true for a real bulleted list instead of typing a literal "•"/"-"/"*" at the start of each line. ' +
      'No fill/line parameters here - use format_element afterward for those.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
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
        bulleted: { type: 'boolean', description: 'Text boxes only (no shapeType).' },
        name: { type: 'string', description: 'Optional shape name, shown by read_slide; use a short, meaningful label.' },
      },
      required: ['slideIndex', 'left', 'top', 'width', 'height'],
    },
  },
  {
    name: 'delete_element',
    description: 'Deletes one shape from a slide.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'copy_element',
    description:
      'Copies one shape, leaving the original. Without targetSlideIndex (or with the same slide) it copies on the SAME slide, offsetX/offsetY points away (default 12/12) or at left/top. ' +
      'With a targetSlideIndex on a DIFFERENT slide it uses PowerPoint\'s native copy/paste, so every shape kind comes across with full fidelity (pictures, video/audio, OLE, freeform, tables with merged cells, charts, SmartArt, groups of these); left/top then override where the paste landed. ' +
      'move:true (different slide only) also removes the original, but only after the copy fully succeeded. The copy gets a unique name (e.g. "Rectangle 5 2") unless you pass name.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        targetSlideIndex: { type: 'number', description: 'Destination slide for a cross-slide copy or move. Omit to copy on the same slide.' },
        move: { type: 'boolean', description: 'Remove the original after copying. Needs a different targetSlideIndex.' },
        offsetX: { type: 'number', description: 'Same-slide copies only. Points. Default 12. Ignored if left/top is given.' },
        offsetY: { type: 'number', description: 'Same-slide copies only. Points. Default 12. Ignored if left/top is given.' },
        left: { type: 'number', description: 'Exact destination position.' },
        top: { type: 'number', description: 'Exact destination position.' },
        name: { type: 'string', description: 'Optional name for the new shape, shown by read_slide.' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'edit_animation',
    description:
      'Adds or edits a shape animation. kind: "add", "delete", "set_timing", "reorder". ' +
      'add: an entrance (default) or exit (exit:true) animation on shapeIndex; effect is one of appear, fade, fly, flashOnce, wipe, zoom, dissolve, bounce, spiral, swivel, wheel, split, box, circle, diamond, plus, checkerboard, randomBars, growAndTurn, riseUp (base effects only, no directional variants like "wipe from the left"); optional trigger, durationSeconds, delaySeconds; returns the new animationIndex. trigger defaults to "onClick" (its own click); "withPrevious" starts with the animation before it, "afterPrevious" once that one finishes. ' +
      'delete, set_timing (durationSeconds/delaySeconds/trigger) and reorder (toIndex, the new 0-based position) act on animationIndex, the 0-based play order - read it with read_slide (view=animations) first. delete and reorder shift later indices; re-read before another edit.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        kind: { type: 'string', enum: ['add', 'delete', 'set_timing', 'reorder'] },
        shapeIndex: { type: 'number', description: 'add only: the shape to animate.' },
        effect: { type: 'string', enum: ['appear', 'fade', 'fly', 'flashOnce', 'wipe', 'zoom', 'dissolve', 'bounce', 'spiral', 'swivel', 'wheel', 'split', 'box', 'circle', 'diamond', 'plus', 'checkerboard', 'randomBars', 'growAndTurn', 'riseUp'], description: 'add only.' },
        exit: { type: 'boolean', description: 'add only: true for an exit animation.' },
        animationIndex: { type: 'number', description: 'delete, set_timing and reorder: which animation.' },
        durationSeconds: { type: 'number' },
        delaySeconds: { type: 'number' },
        trigger: { type: 'string', enum: ['onClick', 'withPrevious', 'afterPrevious'] },
        toIndex: { type: 'number', description: 'reorder only.' },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'copy_element_style',
    description:
      'Format painter: applies PowerPoint\'s own native format painter (Shape.PickUp/Apply, NOT the clipboard) for broad, generic fidelity (fill, outline, shadow, 3-D/bevel, and anything else native formatting covers), plus targeted copies for the specifics this needs guaranteed - text formatting PER RUN (bold/italic/size/color/font/underline/shadow/alignment/superscript-subscript - reproduces mixed formatting within the text exactly, not just a single sampled value), ' +
      'text outline/strikethrough/glow/reflection/bevel, text box vertical+horizontal anchor and margins, gradient/patterned fill, dash/arrowhead outlines, rotation, flip, and AutoShape adjustment handles - from one shape onto one or more others, ' +
      'not position or size. Each target can be on ANY slide, including the source\'s own.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number', description: 'The source shape to copy formatting FROM.' },
        targets: {
          type: 'array',
          items: {
            type: 'object',
            properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' } },
            required: ['slideIndex', 'shapeIndex'],
          },
          description: 'One or more {slideIndex, shapeIndex} shapes to copy formatting TO - can span multiple slides.',
        },
      },
      required: ['slideIndex', 'shapeIndex', 'targets'],
    },
  },
  {
    name: 'group_element',
    description:
      'Groups shapes or ungroups a group. ' +
      'action:"group" groups two or more top-level shapes on a slide into a single group shape; shapeIndexes is an array of 0-based indices (as shown by read_slide). ' +
      'action:"ungroup" promotes a group shape\'s direct children to top-level shapes; shapeIndex points at the group. ' +
      'Either way every other shape\'s index on the slide shifts afterward - re-read the slide before addressing another shape by index in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['group', 'ungroup'] },
        slideIndex: SLIDE_INDEX,
        shapeIndexes: { type: 'array', items: { type: 'number' }, description: 'For "group": two or more distinct 0-based top-level shape indices.' },
        shapeIndex: { type: ['number', 'string'], description: 'For "ungroup": the group to break up.' },
        name: { type: 'string', description: 'For "group", optional name, shown by read_slide; use a short, meaningful label.' },
      },
      required: ['slideIndex', 'action'],
    },
  },
  {
    name: 'edit_table',
    description:
      'Creates or edits a native table. kind: ' +
      '"create" (rows, cols; optional cells = row-major array of arrays of text; optional left/top/width/height; returns its shapeIndex), ' +
      '"set_cell" (shapeIndex, 0-based row/col, paragraphs = the cell text), ' +
      '"insert-row"/"delete-row"/"insert-col"/"delete-col" (shapeIndex, index = a 0-based EXISTING row/column; before picks the side for inserts), ' +
      '"style" (firstRow/bandRow, shadingColor for all cells, borderColor/borderWidthPt, borderPreset "all"/"outline"/"none"). ' +
      'Row 0 is simply the first physical row, even if it holds headers; firstRow only changes its look, not its index. Inserts and deletes shift later indices - re-read the table before another structural edit.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        kind: { type: 'string', enum: ['create', 'set_cell', 'insert-row', 'delete-row', 'insert-col', 'delete-col', 'style'] },
        shapeIndex: { type: 'number', description: 'Every kind except create: the table to edit.' },
        rows: { type: 'number', description: 'create only.' },
        cols: { type: 'number', description: 'create only.' },
        cells: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'create only: optional pre-filled cell text.' },
        ...createRect(200),
        name: { type: 'string', description: 'create only, optional. Shape name, shown by read_slide.' },
        row: { type: 'number', description: 'set_cell only.' },
        col: { type: 'number', description: 'set_cell only.' },
        paragraphs: { type: 'string', description: 'set_cell only: the cell text.' },
        index: { type: 'number', description: 'insert/delete kinds only.' },
        before: { type: 'boolean', description: 'insert kinds only.' },
        firstRow: { type: 'boolean' }, bandRow: { type: 'boolean' }, shadingColor: { type: 'string' },
        borderColor: { type: 'string' }, borderWidthPt: { type: 'number' },
        borderPreset: { type: 'string', enum: ['all', 'outline', 'none'] },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'edit_chart',
    description:
      'Creates or edits a native, editable chart. ' +
      'kind "create": adds a chart and returns its shapeIndex (chartType, categories, series; optional title and left/top/width/height). Every series\' values array needs exactly one value per category. ' +
      'kind "update": changes ONE property of the chart at shapeIndex per call - exactly one of chartType, title, legendPos, dataLabels or gridlines; more than one errors, so make a separate call for each. The result names what applied.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        kind: { type: 'string', enum: ['create', 'update'] },
        shapeIndex: { type: 'number', description: 'update only: the chart to edit.' },
        chartType: { type: 'string', enum: ['column', 'columnStacked', 'bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'], description: 'Required for create; optional for update.' },
        title: { type: 'string' },
        categories: { type: 'array', items: { type: 'string' }, description: 'create only.' },
        series: { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, values: { type: 'array', items: { type: 'number' } } } }, description: 'create only.' },
        ...createRect(300),
        name: { type: 'string', description: 'create only, optional. Shape name, shown by read_slide.' },
        legendPos: { type: 'string', enum: ['none', 'right', 'top', 'left', 'bottom', 'r', 't', 'l', 'b'], description: 'update only.' },
        dataLabels: { type: 'string', enum: ['none', 'value', 'percent'], description: 'update only.' },
        gridlines: { type: 'boolean', description: 'update only. Errors with a clear message on chart types with no value axis (pie/doughnut).' },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'edit_smartart',
    description:
      'Creates or edits a shape-composed SmartArt diagram. kind is one of: create (layout + items - flat node texts, one per top-level node; optional position/size; returns the diagram), set_text (needs nodeIndex + text), add_node (optional text, appends), ' +
      'delete_node (needs nodeIndex), set_style (needs colorName and/or quickStyleName, matched loosely against this Office install\'s gallery), ' +
      'set_layout (needs layout). For every kind except create, smartArtIndex is 0-based WITHIN THAT SLIDE and defaults to 0. ' +
      'delete_node shifts later node indices - call read_slide (view=smartart) again before a second node edit on the same diagram in one run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        smartArtIndex: { type: 'number', description: '0-based among the SmartArt diagrams on that slide. Default 0. Not used by create.' },
        kind: { type: 'string', enum: ['create', 'set_text', 'add_node', 'delete_node', 'set_style', 'set_layout'] },
        nodeIndex: { type: 'number', description: '0-based. Required for set_text and delete_node.' },
        text: { type: 'string' },
        colorName: { type: 'string', description: 'set_style: substring-matched against the gallery, e.g. "Colorful".' },
        quickStyleName: { type: 'string', description: 'set_style: substring-matched, e.g. "Intense".' },
        layout: { type: 'string', enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'], description: 'Required for create and set_layout.' },
        items: { type: 'array', items: { type: 'string' }, description: 'create only.' },
        ...createRect(300),
        name: { type: 'string', description: 'create only, optional. Shape name, shown by read_slide.' },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'edit_picture',
    description:
      'kind: "add", "crop", or "replace". ' +
      'add: inserts a new picture from localPath, a full local path, and returns its shapeIndex. It keeps the natural size (shrunk to fit the slide), scales proportionally if you give only width or height, and stretches if you give both; omitted left/top centers it. Optional name. ' +
      'crop and replace need a picture at shapeIndex - any other shape errors and nothing changes; replace also refuses a placeholder, since deleting it leaves an empty one behind. ' +
      'crop: l/t/r/b are 0..1 fractions of the current on-slide size cut from each edge (all 0 clears the crop). ' +
      'replace: swaps the image in place from localPath, keeping position/size/rotation/approximate z-order; keepCrop keeps the crop.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number', description: 'crop and replace only: the picture.' },
        kind: { type: 'string', enum: ['add', 'crop', 'replace'] },
        l: { type: 'number', description: 'crop only.' },
        t: { type: 'number', description: 'crop only.' },
        r: { type: 'number', description: 'crop only.' },
        b: { type: 'number', description: 'crop only.' },
        localPath: { type: 'string', description: 'add and replace: a full local drive path to the image file.' },
        keepCrop: { type: 'boolean', description: 'replace only.' },
        left: { type: 'number', description: 'add only. Omit to center horizontally.' },
        top: { type: 'number', description: 'add only. Omit to center vertically.' },
        width: { type: 'number', description: 'add only.' },
        height: { type: 'number', description: 'add only.' },
        name: { type: 'string', description: 'add only, optional. Shape name, shown by read_slide.' },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'set_headers_footers',
    description:
      'Sets slide number, footer text and date/time, the same 3 toggles as PowerPoint\'s Insert Header and Footer dialog. slideIndex -1 = every slide, otherwise that 0-based slide. Only the fields you give change. ' +
      'footerText implies footerVisible:true unless you pass footerVisible; dateText implies dateMode:"fixed" the same way. dateMode is never inferred from dateFormat: "fixed" (default) shows dateText verbatim, "auto" shows today\'s date in dateFormat (default "M/d/yy", only used with "auto"). ' +
      'skipTitleSlide ("Don\'t show on title slide") and startNumber (the first slide number) are ALWAYS deck-wide, whatever slideIndex is. If nothing appears after a toggle, the slide\'s layout may not include that placeholder - try set_slide or check read_master.',
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
    name: 'edit_master_element',
    description:
      'Adds, moves, or removes one element on the Slide Master, or on one specific layout. kind: "add", "move", or "remove". ' +
      'add (contentType "icon" = a picture from localPath, or "text" = a small label): pins it in the same spot on every slide (logo, watermark, confidentiality label). Place it with corner (topLeft/topRight/bottomLeft/bottomRight, inset by marginPt, default 12) or left+top together (both or neither). For an icon, width/height are optional: one scales proportionally, neither keeps the natural size. ' +
      'move: changes the position/size/rotation of the element at masterShapeIndex (from read_master); only the fields you give change. It also works on theme placeholders (e.g. repositioning the footer), so double-check the index. ' +
      'remove: deletes the element at masterShapeIndex. It refuses theme placeholders (read_master marks them "(placeholder)") and only deletes shapes added here; to switch a slide number, footer or date off, use set_headers_footers with the matching *Visible:false. Other indices may shift afterward - re-read before another edit. ' +
      'layoutName (any kind) targets one layout instead of the Slide Master: exact name preferred, else substring, matched across every design/theme; an ambiguous name errors. For move/remove it must match what you used in read_master, since masterShapeIndex is only valid within that target. Without layoutName only the default (first) Slide Master is affected, and a layout set to "Hide Background Graphics" will not show a Slide Master element.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['add', 'move', 'remove'] },
        contentType: { type: 'string', enum: ['icon', 'text'], description: 'add only.' },
        masterShapeIndex: { type: 'number', description: 'move and remove: the element, from read_master.' },
        layoutName: { type: 'string', description: 'Optional. Targets one specific layout (matched across every design/theme in the deck) instead of the Slide Master. For move/remove it must match what you used (or omitted) in read_master.' },
        corner: { type: 'string', enum: ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'], description: 'add only. Ignored if left+top are both given. Required if they aren\'t.' },
        left: { type: 'number', description: 'add: must be given together with top (or neither) - a single coordinate alone is rejected. move: the new left.' },
        top: { type: 'number', description: 'add: must be given together with left (or neither) - a single coordinate alone is rejected. move: the new top.' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number', description: 'move only.' },
        marginPt: { type: 'number', description: 'add only. Default 12. Inset from the slide edge, only used with corner.' },
        localPath: { type: 'string', description: 'add with contentType "icon": a local image file path.' },
        text: { type: 'string', description: 'add with contentType "text".' },
        fontSize: { type: 'number', description: 'add with contentType "text". Default 10.' },
        color: { type: 'string', description: 'add with contentType "text". Hex color, e.g. "#808080" (the default).' },
        name: { type: 'string', description: 'add only, optional. Name, shown by read_master.' },
      },
      required: ['kind'],
    },
  },
  {
    name: 'undo_last_action',
    description:
      "Reverses the presentation's last action, like Ctrl+Z (dispatched via PowerPoint's own ribbon Undo command, since the object model has no direct Undo method). " +
      'Reports honestly whether anything was actually undone ("Nothing to undo." if the undo history is empty).',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'redo_last_action',
    description:
      "Re-applies the presentation's last undone action, like Ctrl+Y (dispatched via PowerPoint's own ribbon Redo command). " +
      'Reports honestly whether anything was actually redone ("Nothing to redo." if there is nothing to reapply).',
    inputSchema: { type: 'object', properties: {} },
  },
]

export const ALL_TOOLS = [...READER_TOOLS, ...MUTATION_TOOLS]

// FT-1 Task 5: settings-screen labels/descriptions, one entry per tool above -
// short, user-facing sentences, not the model-facing `description`s above.
