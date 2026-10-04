# PowerPointTools.Read.cs

## `ShapeText`

PP-23 Task 8 (post-hoc fix): previously only checked HasTextFrame,
so read_slide/get_deck_context reported table and SmartArt shapes
with no content at all - the model could add_table/add_smartart
and then have no way to see what it just created. Table reading
reuses the same statically-typed Table.Cell(r,c).Shape.TextFrame
pattern AddTable/EditTableCell already use in this file; SmartArt
is dynamic (HasSmartArt/.SmartArt aren't on the statically-typed
Shape interface), matching AddSmartArt's own existing pattern.

## `ReadSlide` - layout/transition/animation surfacing

PP-24: surfaces layout/transition/animation-count so the model
isn't blind to state set_slide_layout/set_slide_transition/
add_animation just created - same "the model can add
something but then can't see it" gap PP-23's read_chart/
read_table/read_smartart all exist to close.

## `ReadSlide` - guarded HeadersFooters read

Review finding: unguarded, unlike every other HeadersFooters
access PowerPointTools.Master.cs added elsewhere in this same
PR - that file documents several real, previously-unknown COM
states where reading/writing a slide's HeadersFooters throws
"HeaderFooter (unknown member)". read_slide is a core, always-
allowed read tool; if the same restriction ever hits a read
(not just the write paths already fixed), don't let it discard
everything already built into `sb` above.

## `ReadSlide` - explicit z-order note

Post-hoc addition (2026-08-24, user-requested: "see the order
between objects"): slide.Shapes is already ordered back-to-
front by z-order (confirmed via reflection: Shape.ZOrderPosition
is a get-only int matching this same collection order) - the
shapeIndex below was already exactly this order, just never
stated explicitly. No new read tool needed; making the
existing order's meaning explicit is enough.
