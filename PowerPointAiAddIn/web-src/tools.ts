const SLIDE_INDEX = { type: 'number' };
const SHAPE_PATH = { type: ['number', 'string'], description: '0-based shape index, or a dotted path like "3.1.0" for a shape inside a group (from read_group).' };

export const READER_TOOLS = [
  {
    name: 'get_deck_context',
    description: 'Reads a one-line-per-slide outline: slide index and a text preview of its shapes.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'read_slide',
    description: 'Reads full text of every shape on one slide (0-based index), plus its layout, transition, animation count, and speaker notes (if any). Shapes are listed back-to-front (z-order/stacking order) - use set_element_order to change it. A group shape is shown as one line with a child count - call read_group to see inside it.',
    inputSchema: { type: 'object', properties: { slideIndex: SLIDE_INDEX }, required: ['slideIndex'] },
  },
  {
    name: 'read_group',
    description:
      'Lists the shapes inside a group, recursively (nested groups are expanded in place). shapeIndex points at the group - a 0-based number, or a dotted path like "3.1" for a group nested inside another group. ' +
      'Each child line is prefixed with its dotted path (e.g. "3.1.0"); pass that path as shapeIndex to set_element_text / set_element_style / set_element_fill / set_element_stroke to edit a child in place. ' +
      'Positional or structural edits (set_element_transform, set_element_order, delete_element, add_animation) require ungroup_element on the top-level group first.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: ['number', 'string'] } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'read_animations',
    description: 'Reads a slide\'s animations in play order, one per line (shape, effect, entrance/exit, trigger, timing). Needed before edit_animation, since animationIndex addresses this same order.',
    inputSchema: { type: 'object', properties: { slideIndex: SLIDE_INDEX }, required: ['slideIndex'] },
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
        slideIndex: SLIDE_INDEX,
        smartArtIndex: { type: 'number', description: 'Omit to read all diagrams on the slide.' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'list_layouts',
    description:
      'Lists every layout name in this presentation (e.g. "Title Slide", "Title and Content"), plus how many slides currently use each one - across every design/theme in the deck if it has more than one, not just the default. ' +
      'Call this before passing layoutName to add_master_element/read_master_elements/remove_master_element/set_master_element_transform/set_slide, instead of guessing a name. ' +
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
      'Sets any combination of one slide\'s layout, background, entry transition, and speaker notes in one call (0-based slideIndex). Give only the fields you want to change. ' +
      'layout (classic, a fixed set) or layoutName (custom: free text matched against this slide\'s own theme layouts - exact match preferred, else substring; an unmatched name errors listing the real available names, and a name matching more than one layout errors asking for a more specific one) - give one, not both. ' +
      'background: a solid color; slideIndex=-1 applies it to every slide in the deck (background only - the other fields need a real slideIndex). ' +
      'notes: replaces the speaker notes (read_slide shows the current notes first). ' +
      'Transition: effect sets or removes the entry transition ("none" removes it); durationSeconds is how long the transition animation takes; advanceOnClick (default true in PowerPoint) and advanceAfterSeconds (an automatic timed advance) are independent - both can be on at once, and either can be changed without restating effect. When both advance options are off the slide only advances from the keyboard or menu.',
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
    name: 'set_element_style',
    description:
      'Changes text formatting of one shape (its whole text range) without changing its text. ' +
      'Only the fields you provide are changed; anything omitted is left as-is.',
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX,
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
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'duplicate_element',
    description:
      'Creates a copy of a shape on the SAME slide - works for every shape kind, including groups/pictures/tables/charts/SmartArt. For a DIFFERENT slide, use copy_element/move_element instead. ' +
      'Positions the duplicate offsetX/offsetY points from the original (default 12/12) unless left/top is given, which places it at an exact position instead. ' +
      'If you don\'t pass name, the duplicate is automatically given a unique name (e.g. "Rectangle 5" -> "Rectangle 5 2") instead of colliding with the original\'s name.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        offsetX: { type: 'number', description: 'Points. Default 12. Ignored if left/top is given.' },
        offsetY: { type: 'number', description: 'Points. Default 12. Ignored if left/top is given.' },
        left: { type: 'number', description: 'Exact destination position - overrides offsetX/offsetY.' },
        top: { type: 'number', description: 'Exact destination position - overrides offsetX/offsetY.' },
        name: { type: 'string', description: 'Optional. Sets the duplicate\'s name, shown in read_slide/read_group.' },
      },
      required: ['slideIndex', 'shapeIndex'],
    },
  },
  {
    name: 'copy_element',
    description:
      'Copies a shape to a DIFFERENT slide, leaving the original in place, using PowerPoint\'s own native copy/paste - so EVERY shape kind is supported with full native fidelity, exactly as it already exists (pictures, video/audio, OLE objects, freeform, tables with merged cells/per-cell shading/borders, charts, SmartArt, and groups of any of these, including a group nesting a SmartArt - since this clones the existing shape tree rather than rebuilding it). ' +
      'For a same-slide copy, use duplicate_element instead - this tool refuses targetSlideIndex equal to slideIndex. ' +
      'Inherits the source shape\'s own name (auto-suffixed if it collides on the destination slide) unless you pass an explicit name. Optional left/top override the position PowerPoint\'s own paste chose.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        targetSlideIndex: { type: 'number' },
        left: { type: 'number', description: 'Destination position. Defaults to the source shape\'s own position.' },
        top: { type: 'number', description: 'Destination position. Defaults to the source shape\'s own position.' },
        name: { type: 'string', description: 'Optional. Sets the new shape\'s name (otherwise inherits the source\'s, deduped).' },
      },
      required: ['slideIndex', 'shapeIndex', 'targetSlideIndex'],
    },
  },
  {
    name: 'move_element',
    description:
      'Same as copy_element, but also removes the shape from its original slide - equivalent to copy_element followed by deleting the original. ' +
      'The original is only removed once the copy has fully succeeded; on any error nothing is changed. Same kind support, formatting fidelity, and limitations as copy_element.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX,
        shapeIndex: { type: 'number' },
        targetSlideIndex: { type: 'number' },
        left: { type: 'number', description: 'Destination position. Defaults to the source shape\'s own position.' },
        top: { type: 'number', description: 'Destination position. Defaults to the source shape\'s own position.' },
        name: { type: 'string', description: 'Optional. Sets the new shape\'s name (otherwise inherits the source\'s, deduped).' },
      },
      required: ['slideIndex', 'shapeIndex', 'targetSlideIndex'],
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX,
        shapeIndex: SHAPE_PATH,
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
        slideIndex: SLIDE_INDEX,
        shapeIndex: SHAPE_PATH,
        color: { type: 'string' }, widthPt: { type: 'number' }, remove: { type: 'boolean' },
      },
      required: ['slideIndex', 'shapeIndex'],
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
    name: 'ungroup_element',
    description: 'Promotes a group shape\'s direct children to top-level shapes. Shape indices change after this call - re-read the slide before addressing the promoted shapes.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: ['number', 'string'] } },
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX, rows: { type: 'number' }, cols: { type: 'number' },
        cells: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
        x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' },
        name: { type: 'string', description: 'Optional. Sets the shape\'s name, shown in read_slide/read_group.' },
      },
      required: ['slideIndex', 'rows', 'cols'],
    },
  },
  {
    name: 'edit_table_cell',
    description: 'Replaces one table cell\'s text (0-based row/col). Row 0 is just the first physical row, including when it holds column headers - there is no separate header concept in the index space (edit_table_style\'s firstRow only changes that row\'s visual style, not its index).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' }, row: { type: 'number' }, col: { type: 'number' }, paragraphs: { type: 'string' } },
      required: ['slideIndex', 'shapeIndex', 'row', 'col', 'paragraphs'],
    },
  },
  {
    name: 'edit_table_structure',
    description: 'Inserts or deletes a table row/column. index (0-based) addresses an EXISTING row/column, including row 0 even if it holds column headers - there is no separate header concept in the index space (edit_table_style\'s firstRow only changes that row\'s visual style, not its index). before decides which side the new one goes on for insert kinds. Deleting/inserting shifts every later row/column\'s index - re-read the table before a second structural edit in the same run.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' },
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
        slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' },
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' },
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
        slideIndex: SLIDE_INDEX,
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
        slideIndex: SLIDE_INDEX,
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
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' }, l: { type: 'number' }, t: { type: 'number' }, r: { type: 'number' }, b: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'l', 't', 'r', 'b'],
    },
  },
  {
    name: 'set_picture_opacity',
    description: 'Sets a picture shape\'s overall opacity, 0 (invisible) to 1 (fully opaque).',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' }, opacity: { type: 'number' } },
      required: ['slideIndex', 'shapeIndex', 'opacity'],
    },
  },
  {
    name: 'replace_image',
    description: 'Swaps a picture shape\'s image content in place from a local file path, keeping position/size/rotation/approximate z-order.',
    inputSchema: {
      type: 'object',
      properties: { slideIndex: SLIDE_INDEX, shapeIndex: { type: 'number' }, localPath: { type: 'string' }, keepCrop: { type: 'boolean' } },
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
      'If nothing appears after enabling a toggle, the slide\'s layout may simply not include that placeholder - try set_slide or check read_master_elements.',
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
