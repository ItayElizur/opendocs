# PowerPointTools.Elements.cs

## `MakeUniqueNameOnSlide`

siblingShapes defaults to null, meaning "resolve it from the shape's
own slide" (the common case for an ordinary slide shape, and the only
case CrossSlide.cs/DuplicateElement ever need). Pass it explicitly
(e.g. AddMasterElement passes target.Shapes) for a shape whose parent
is never a PowerPoint.Slide - a Slide Master/layout shape - where
ShapeSlide can never resolve one automatically; without this, dedup
silently no-ops for those shapes. Review finding: this used to be two
separate overloads and a caller could pick the wrong one (the 2-arg
overload silently skipped dedup for a master/layout shape for a full
round of development before being caught) - one method with a
defaultable parameter makes the common case's default safe
automatically instead of relying on the caller remembering which
overload fits.

## `ApplyAutoDirection`

Post-hoc fix (2026-08-26, user-reported): PowerPoint's TextRange
never auto-flips paragraph direction/alignment based on typed
content the way Word's editor does with "detect language
automatically" - every write here always came out left-to-right/
left-aligned, even for Hebrew. Same bidi mismatch class as
chat-ui.ts's dir="auto" fix, but PowerPoint's COM object model has
no built-in "auto" direction - it has to be decided per write from
the text's own script mix. IsRtlMajority itself now lives in
TextUtil (Phase 0) since it's free of COM types.

## `ApplyBulletSetting`

User-reported bug: a placeholder from a layout like "Title and
Content" already renders its own native bullet per paragraph, so a
model that ALSO types a literal bullet character ("•"/"-"/"*") at
the start of each line ends up with two bullets per line. This
gives the model a real on/off switch instead, so it never needs to
embed a literal bullet character in the text. bulleted omitted
(the JSON property absent, not merely false) leaves the shape's
existing bullet setting untouched - matches prior behavior for
every existing caller that doesn't pass it.

## `SetElementStyle` - strikethrough omission

PP-20 Task 1 Step 2: strikethrough deliberately NOT implemented.
TextFrame (the older text model this file uses everywhere else)
has no Strikethrough member on this PIA. Excel's interop PIA
exposes TextFrame2/TextRange2 (the newer "DrawingML" text model,
used by ExcelTools.cs's chart formatting) with a Strikethrough
property, but Microsoft.Office.Interop.PowerPoint has no
TextFrame2 type or Shape.TextFrame2 member at all (confirmed:
absent from this PIA's own XML docs, and CS0234 - "TextRange2
does not exist in the namespace" - on a direct attempt). Per
this plan's own instruction: omit rather than ship a schema
field the handler can't back, or a `dynamic` call this
environment cannot runtime-verify.

## `ZOrderMap`

Post-hoc addition (2026-08-24, user-requested: "change the order
of an element" - stacking/z-order, distinct from set_element_
transform's position/size). Confirmed via reflection: Shape.ZOrder
(MsoZOrderCmd) is the real relative-move method; MsoZOrderCmd has
6 values total, of which the first 4 apply to slide shapes
(msoBringInFrontOfText/msoSendBehindText are for a shape's
position relative to body text, not meaningful for a slide's flat
z-order stack) - only those 4 are exposed.

## `DuplicateElement`

Shape.Duplicate() is a native in-place COM clone independent of
Shape.Type - unlike copy_element/move_element (PowerPointTools.
CrossSlide.cs), it needs no shape-kind dispatch at all and works
uniformly for every kind, including groups/pictures/tables/charts/
SmartArt. Positioning of the duplicate by Duplicate() itself is
unverified, so this always computes the new position from the
ORIGINAL shape's Left/Top, never from the duplicate's own
post-Duplicate() position.

## `DuplicateElement` - independent offset axes

Each axis is independent: an explicit left/top is an exact
coordinate; an omitted one falls back to the DEFAULT OFFSET on
that axis, not to the original's exact coordinate. Review
finding: the previous either/or branching treated "left given,
top omitted" as "use exact left, but exact (unoffset) top too" -
a duplicate with only left set landed fully overlapping the
original vertically instead of keeping the normal offsetY gap.

## `DuplicateElement` - name collision

Real-user-confirmed (2026-09-22, live testing): Shape.Duplicate()
keeps the EXACT source Name (unlike a UI Ctrl+D/paste, which
auto-renames) - two shapes end up identically named, confusing
read_slide/read_group output. Dedupe it the same way an explicit
name would be, using the duplicate's own (source-inherited) name
as the "desired" one.
