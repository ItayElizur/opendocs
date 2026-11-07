const SLIDE_INDEX = { type: 'number' };
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
      'view selects what to read (default "slide"): "slide", "group", "animations", or "smartart". Every view needs slideIndex. ' +
      'view "slide": Reads full text of every shape on one slide (0-based index), plus its layout, transition, animation count, and speaker notes (if any). Shapes are listed back-to-front (z-order/stacking order) - use set_element_transform with order to change it. A group shape is shown as one line with a child count - call read_slide (view=group) to see inside it.' +
      'view "group" (also needs shapeIndex): Lists the shapes inside a group, recursively (nested groups are expanded in place). shapeIndex points at the group - a 0-based number, or a dotted path like "3.1" for a group nested inside another group. ' +
      'Each child line is prefixed with its dotted path (e.g. "3.1.0"); pass that path as shapeIndex to set_element_text / format_element to edit a child in place. ' +
      'Positional or structural edits (set_element_transform, delete_element, edit_animation) require group_element with action:"ungroup" on the top-level group first.' +
      'view "animations": Reads a slide\'s animations in play order, one per line (shape, effect, entrance/exit, trigger, timing). Needed before edit_animation, since animationIndex addresses this same order.' +
      'view "smartart" (optional smartArtIndex): Reads the node text of SmartArt diagrams on a slide, one "[i] text" line per node. ' +
      'Omit smartArtIndex to read every diagram on the slide in one call. ' +
      'Call this before an incremental edit_smartart change so you are working from current node indices.',
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
      'view selects what to read (default "elements"): "elements" or "layouts". ' +
      'view "elements": Lists the shapes placed directly on the Slide Master (index, name, kind, position/size, text if any) - not the shapes on any one slide. ' +
      'Pass layoutName to instead list one specific layout\'s own shapes (e.g. "Title Slide", "Title and Content") - matched against this presentation\'s real layout names (exact match preferred, else substring; matching more than one errors, e.g. two designs both having a layout with that name), across every design/theme in the deck; omit it for the Slide Master (the default, and what every slide inherits from unless its own layout overrides a shape). ' +
      'Shapes added here with edit_master_element appear on every slide whose layout doesn\'t hide master shapes. Call this before edit_master_element to find the right masterShapeIndex.' +
      'view "layouts": Lists every layout name in this presentation (e.g. "Title Slide", "Title and Content"), plus how many slides currently use each one - across every design/theme in the deck if it has more than one, not just the default. ' +
      'Call this before passing layoutName to edit_master_element/read_master/set_slide, instead of guessing a name. ' +
      'A layoutName matching more than one layout (e.g. two designs both have a layout with that name) errors instead of guessing - pass a more specific or exact name.',
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
      'Changes one thing about one slide per call (0-based slideIndex): its layout, background, entry transition, or speaker notes. Giving fields from more than one of these errors - make a separate call for each. ' +
      'Layout: layout (classic, a fixed set) or layoutName (custom: free text matched against this slide\'s own theme layouts - exact match preferred, else substring; an unmatched name errors listing the real available names, and a name matching more than one layout errors asking for a more specific one) - give one, not both. ' +
      'Background: a solid color; slideIndex=-1 applies it to every slide in the deck (the only change that allows -1). ' +
      'Notes: replaces the speaker notes (read_slide shows the current notes first). ' +
      'Transition: effect sets or removes the entry transition ("none" removes it); durationSeconds is how long the transition animation takes; advanceOnClick (default true in PowerPoint) and advanceAfterSeconds (an automatic timed advance) are independent - both can be on at once, and they can be changed without restating effect. These count as one change, so they may be combined. When both advance options are off the slide only advances from the keyboard or menu.',
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
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_slide (view=group) - use a short, meaningful label (e.g. "CTA button").' },
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
      'Copies one shape, leaving the original in place. ' +
      'Without targetSlideIndex (or with the same slide) it makes a copy on the SAME slide - every shape kind, including groups/pictures/tables/charts/SmartArt - placed offsetX/offsetY points from the original (default 12/12) unless left/top gives an exact position. ' +
      'With a targetSlideIndex on a DIFFERENT slide it copies there using PowerPoint\'s own native copy/paste - so EVERY shape kind is supported with full native fidelity, exactly as it already exists (pictures, video/audio, OLE objects, freeform, tables with merged cells/per-cell shading/borders, charts, SmartArt, and groups of any of these, including a group nesting a SmartArt - since this clones the existing shape tree rather than rebuilding it); left/top then override the position PowerPoint\'s own paste chose. ' +
      'move:true (needs a different targetSlideIndex) also removes the shape from its original slide once the copy has fully succeeded; on any error nothing is changed. ' +
      'Without a name the copy is given a unique name (e.g. "Rectangle 5" -> "Rectangle 5 2") instead of colliding with the original\'s.',
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
        name: { type: 'string', description: 'Optional. Sets the new shape\'s name, shown in read_slide/read_slide (view=group).' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'edit_animation',
    description:
      'Adds or edits a shape animation. kind:"add" adds an entrance (default) or exit animation to a shape (shapeIndex, effect; optional exit, trigger, durationSeconds, delaySeconds) and returns the new animationIndex. ' +
      'effect: appear, fade, fly, flashOnce, wipe, zoom, dissolve, bounce, spiral, swivel, wheel, split, box, circle, diamond, plus, checkerboard, randomBars, growAndTurn, riseUp - base effects only, no directional variants (e.g. "wipe from the left"). ' +
      'trigger (default "onClick"): "onClick" starts on its own click during the slideshow, "withPrevious" starts together with the animation before it, "afterPrevious" starts automatically once the animation before it finishes. ' +
      'Existing animations: kind "delete", "set_timing" (durationSeconds?,delaySeconds?,trigger?), "reorder" (toIndex - 0-based new position in the play sequence); animationIndex addresses the animation (0-based, current play order - call read_slide (view=animations) first). delete/reorder shift later indices - re-read before another edit in the same run.',
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
        name: { type: 'string', description: 'For "group", optional. Sets the new group\'s name, shown in read_slide/read_slide (view=group) - use a short, meaningful label (e.g. "Logo lockup").' },
      },
      required: ['slideIndex', 'action'],
    },
  },
  {
    name: 'edit_table',
    description:
      'Creates or edits a native PowerPoint table. kind selects the operation: ' +
      '"create" adds a table (rows, cols; optional cells = row-major array of arrays of cell text; optional position/size; returns its shapeIndex); ' +
      '"set_cell" replaces one cell\'s text (shapeIndex, 0-based row/col, paragraphs); ' +
      '"insert-row"/"delete-row"/"insert-col"/"delete-col" change the structure (shapeIndex, index = a 0-based EXISTING row/column; before decides which side the new one goes on for insert kinds); ' +
      '"style" applies firstRow/bandRow (header row / banded rows), shadingColor (all cells), borderColor/borderWidthPt/borderPreset ("all" = every cell edge, "outline" = only the table\'s outer perimeter, "none" = remove borders). ' +
      'Row 0 is just the first physical row, including when it holds column headers - there is no separate header concept in the index space (style\'s firstRow only changes that row\'s visual style, not its index). ' +
      'Inserting/deleting shifts every later row/column\'s index - re-read the table before a second structural edit in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        kind: { type: 'string', enum: ['create', 'set_cell', 'insert-row', 'delete-row', 'insert-col', 'delete-col', 'style'] },
        shapeIndex: { type: 'number', description: 'Every kind except create: the table to edit.' },
        rows: { type: 'number', description: 'create only.' },
        cols: { type: 'number', description: 'create only.' },
        cells: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'create only: optional pre-filled cell text.' },
        left: { type: 'number', description: 'create only. Default 100.' },
        top: { type: 'number', description: 'create only. Default 100.' },
        width: { type: 'number', description: 'create only. Default 400.' },
        height: { type: 'number', description: 'create only. Default 200.' },
        name: { type: 'string', description: 'create only, optional. Sets the shape\'s name, shown in read_slide/read_slide (view=group).' },
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
      'Creates or edits a native, editable PowerPoint chart. kind:"create" adds a chart and returns its shapeIndex (chartType, categories, series; optional title and position/size) - every series\' values array must have exactly one value per category. ' +
      'kind:"update" modifies an existing chart (shapeIndex) - chartType/title/legendPos/dataLabels/gridlines; only the properties provided are changed and the result names exactly what applied.',
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
        left: { type: 'number', description: 'create only. Default 100.' },
        top: { type: 'number', description: 'create only. Default 100.' },
        width: { type: 'number', description: 'create only. Default 400.' },
        height: { type: 'number', description: 'create only. Default 300.' },
        name: { type: 'string', description: 'create only, optional. Sets the chart shape\'s name, shown in read_slide/read_slide (view=group).' },
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
        left: { type: 'number', description: 'create only. Default 100.' },
        top: { type: 'number', description: 'create only. Default 100.' },
        width: { type: 'number', description: 'create only. Default 400.' },
        height: { type: 'number', description: 'create only. Default 300.' },
        name: { type: 'string', description: 'create only, optional. Sets the shape\'s name, shown in read_slide/read_slide (view=group).' },
      },
      required: ['slideIndex', 'kind'],
    },
  },
  {
    name: 'edit_picture',
    description:
      'kind selects the operation: "crop" or "replace". Both work only on a picture - any other shape errors and nothing is changed (a replace also refuses a placeholder, since replacing deletes the old shape and PowerPoint would leave an empty placeholder behind). ' +
      'kind "crop": Non-destructively crops a picture shape. l/t/r/b are 0..1 fractions of the current on-slide image size cut from each edge; all zero clears the crop.' +
      'kind "replace": Swaps a picture shape\'s image content in place from a local file path, keeping position/size/rotation/approximate z-order.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        kind: { type: 'string', enum: ['crop', 'replace'] },
        l: { type: 'number', description: 'crop only.' },
        t: { type: 'number', description: 'crop only.' },
        r: { type: 'number', description: 'crop only.' },
        b: { type: 'number', description: 'crop only.' },
        localPath: { type: 'string', description: 'replace only: the new image file.' },
        keepCrop: { type: 'boolean', description: 'replace only.' },
      },
      required: ['slideIndex', 'shapeIndex', 'kind'],
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
      'If nothing appears after enabling a toggle, the slide\'s layout may simply not include that placeholder - try set_slide or check read_master.',
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
      'Adds, moves/resizes/rotates, or removes one element on the Slide Master (or on one specific layout). kind selects the operation: "add", "move", or "remove". ' +
      'kind "add" (contentType "icon" or "text"): Adds a persistent icon (picture) or small text label to the Slide Master, so it shows up in the same spot on every slide (e.g. a logo/watermark/confidentiality label) - not just one slide. ' +
      'Pass layoutName to add it to one specific layout instead (e.g. "Title Slide") - it then only shows on slides using that layout, not deck-wide; matched against this presentation\'s real layout names across every design/theme in the deck (exact match preferred, else substring; matching more than one, e.g. two designs both having a layout with that name, errors instead of guessing). ' +
      'Use corner for a quick topLeft/topRight/bottomLeft/bottomRight placement (inset by marginPt, default 12pt), or left+top together for an exact position (left/top must both be given, or neither). ' +
      'For contentType:"icon", width/height are optional - omit one and it scales proportionally from the image\'s natural size; omit both to use the natural size. ' +
      'Known limits: with no layoutName, only this presentation\'s default (first) Slide Master is affected - a deck combining more than one theme has more than one, and layoutName is the way to reach a non-default one\'s layout directly; and a layout with "Hide Background Graphics" enabled won\'t show a Slide-Master-level element.' +
      'kind "move": Moves/resizes/rotates an existing Slide Master element by masterShapeIndex (from read_master). Only the fields you provide are changed. ' +
      'Works on a theme placeholder too (e.g. repositioning where the footer shows) - unlike edit_master_element, this is not refused, since moving/resizing a placeholder is a normal, reversible edit. ' +
      'Pass layoutName to target one specific layout\'s shape instead of the Slide Master\'s - must match what you used (or omitted) in read_master, since masterShapeIndex is only valid within that same target. Matched across every design/theme in the deck (exact match preferred, else substring; matching more than one errors).' +
      'kind "remove": Deletes one shape from the Slide Master by the 0-based masterShapeIndex read_master reports. Refuses (with an error) to delete a theme placeholder - read_master marks those "(placeholder)"; only delete a shape edit_master_element actually created. ' +
      'To turn a slide number/footer/date placeholder off instead of deleting it, use set_headers_footers with the matching *Visible:false field - that\'s the correct way to "remove" one, matching PowerPoint\'s own native behavior. ' +
      'Pass layoutName to target one specific layout instead of the Slide Master - must match what you used (or omitted) in read_master. Matched across every design/theme in the deck (exact match preferred, else substring; matching more than one errors). Other master shapes\' indices may shift afterward - re-read before another edit in the same run.',
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
        name: { type: 'string', description: 'add only, optional. Sets the shape\'s name, shown in read_master.' },
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
