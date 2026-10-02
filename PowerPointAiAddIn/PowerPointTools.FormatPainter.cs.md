# PowerPointTools.FormatPainter.cs

## File header

Plain Shape-to-Shape native-value copies (no JSON parsing, no hex
round trip) - power copy_element_style below. copy_element/
move_element (PowerPointTools.CrossSlide.cs) used to share these
too, back when they reconstructed a destination shape property by
property; that approach was replaced (2026-09-23) with native
Shape.Copy()/Slide.Shapes.Paste(), which needs none of this.

## `CopyViaPickUpApply`

User-asked (2026-09-23): "isn't there a way to just copy all
formatting generically instead of hardcoding every property?" -
yes: Shape.PickUp()/Shape.Apply() is PowerPoint's own native format
painter (confirmed via reflection to exist on both Shape and
ShapeRange, no arguments) - it's an internal Application-level
"last picked up formatting" slot, NOT the Windows clipboard (a
different mechanism from Shape.Copy()/Shapes.Paste(), which this
codebase avoids for the documented clobbering/racing reason). This
is called FIRST, as a broad baseline pass, before any of the
targeted Copy*Formatting calls below - it should already cover
fill/line/shadow/3-D/adjustments generically (including things we
haven't individually hardcoded, like shape bevel), and the
targeted calls that follow then refine/guarantee the specific
properties this tool documents (most importantly per-RUN text
formatting, which a cursor-position-based format painter cannot
reproduce - PickUp/Apply on a whole shape samples one formatting
state, not each run individually). Kept as a supplement, not a
replacement, for exactly that reason.

## `CopyTextFormatting`

User-requested (2026-09-22) upgrade from the original "sample the
first character only" design to real per-run fidelity. There is no
bulk "list the formatting runs" API - confirmed via reflection,
TextRange.Runs(start,length) is an indexed accessor exactly like
Characters/Paragraphs, not an enumerable collection - so run
boundaries are found by scanning: read every character's Font,
compare to the previous one, and a difference in ANY tracked
property starts a new run. Paragraph Alignment is copied
per-paragraph (paragraph boundaries found by scanning the string
itself for '\r', PowerPoint's own paragraph separator when reading
TextRange.Text - avoids relying on Paragraphs(start,length)'s own
indexing semantics, which were never independently verified).
Replaying onto the SAME character offsets on the destination text
is a best-effort match, not a guaranteed one, since the target is
an independently-existing shape with its own (possibly shorter,
possibly longer) text - each write is individually caught
(ApplyFontRun) so a length mismatch just stops applying formatting
past the target's own text length rather than throwing.

Real cost, not hidden: this is O(text length) COM round trips (one
Font read per character, up to 9 properties each) to detect
boundaries, plus O(run count) writes. Not capped - a hard cap would
silently lose fidelity on long text, contradicting the point of
this change - but a very long text body will be proportionally
slower. Strikethrough still isn't copied (unchanged pre-existing
gap - no TextFrame2/TextRange2 on this PIA's Shape, same as
set_element_style's own documented limitation).

## `CopyTextFormatting` - table crash gate

Real-user-confirmed (2026-09-23): copying a TABLE crashed the
whole reconstruction with "The specified value is out of
range" - traced to THIS gate being unguarded. A table's outer
graphic-frame Shape doesn't carry text the way a text box/
autoshape does (text lives in the individual cells' own
shapes), and querying HasTextFrame/TextFrame on it can throw
instead of just reporting false - every other Copy* method in
this file already treats its own entry gate as best-effort;
this one didn't.

## `CopyTextFormatting` - text-frame-level layout

Text-frame-level layout (real-user-confirmed gap, 2026-09-23:
vertical anchor - top/middle/bottom - wasn't copied at all).
Confirmed via reflection: PowerPoint.TextFrame has plain
read/write VerticalAnchor/HorizontalAnchor/margins/WordWrap -
no TextFrame2 needed for these.

## `CopyTextEffects`

User-requested (2026-09-23): text outline and "text effects" -
NOT accessible via the classic Font object CopyTextFormatting
above uses (confirmed via reflection: no Line/Glow/Reflection/
Strikethrough members on PowerPoint.Font at all). They ARE
accessible via Shape.TextFrame2 -> Core.TextRange2 -> Core.Font2 -
confirmed via reflection this session, CORRECTING an earlier,
wrong claim elsewhere in this codebase (set_element_style's own
comment) that TextFrame2 doesn't exist on this PIA's Shape at all;
it does, just needs Microsoft.Office.Core's own TextRange2/Font2
types, not a PowerPoint-namespace TextRange2 (which is what a
previous, unqualified attempt actually failed to find).

Deliberately sampled from the FIRST character only (uniform
across the whole shape), NOT per-run like CopyTextFormatting -
text outline/glow are rarely mixed within one text box in
practice, and giving this its own full per-run scan would double
the already-significant COM-call cost of CopyTextFormatting for a
much rarer case. Covers Line (outline), StrikeThrough/
DoubleStrikeThrough, Glow, Reflection, Shadow, SoftEdge, and Bevel/
3-D (TextFrame2.ThreeD) - every text-level DrawingML effect
confirmed via reflection on Font2/TextFrame2. This is still a
manually-enumerated list, not a generic "copy everything" - Font2
has no bulk clone method - but CopyViaPickUpApply (see above)
provides the generic broad pass at the whole-SHAPE level; this
method's job is specifically the effects a shape-level format
painter pass can't guarantee at true per-run granularity.

## `CopyTextEffects` - text bevel/3-D

Text "bevel"/3-D (real-user-confirmed gap, 2026-09-23):
TextFrame2.ThreeD (confirmed via reflection to be the same
ThreeDFormat type as Shape.ThreeD) - a whole-text-frame
property, not per-character, so it's set here rather than
per-run below.

## `CopyTextEffects` - Reflection bug

Real-user-confirmed bug (2026-09-23): reflection came out
"always set" regardless of the source. Root cause:
ReflectionFormat also has no Visible gate - msoReflection-
TypeNone (confirmed via reflection to be the real "off"
value, =0) is the only off switch, but the previous version
wrote Type/Size/Transparency/Blur unconditionally every
time, which - like Glow above - risked materializing a
reflection effect on the destination merely by touching
these properties at all, independent of which Type value
was written. Now only writes anything when the source's
own Type is affirmatively NOT None. Offset is deliberately
left alone (its exact type wasn't independently verified).

## `CopyGradientFill`

Real-user-confirmed (2026-09-23, live testing): the original version
of this method - which bootstrapped gradient mode by calling
destFill.TwoColorGradient(srcFill.GradientStyle, srcFill.
GradientVariant), i.e. replaying the SOURCE's own style/variant -
produced a visibly wrong gradient. Root cause: GradientStyle/
GradientVariant aren't safe to round-trip for every source fill -
a fill set up via PresetGradient, a theme gradient, or anything
other than an explicit TwoColorGradient call can report a
style/variant combination TwoColorGradient itself rejects or
reinterprets differently (msoGradientMixed being the most obvious
case, but not the only one). This version instead ALWAYS bootstraps
with a fixed, safe style/variant just to enter gradient mode and
get a real GradientStops collection to write into - the actual
visible result comes entirely from the explicit per-stop
color/position/transparency copy below, which doesn't depend on
the bootstrap style at all. GradientAngle is copied separately
afterward for the linear styles (horizontal/vertical/diagonal),
since that's the one visually-significant property the bootstrap
style choice would otherwise have baked in.

## `CopyGradientFill` - stop-count reconciliation (second bug round)

Real-user-confirmed (2026-09-23, live testing, SECOND round):
still wrong after the bootstrap-style fix above. Root cause:
the bootstrap leaves exactly 2 default stops, but a real-world
gradient fill - especially PowerPoint's own Shape Style
gallery gradients - very often has 3+ stops (a subtle sheen
effect), not 2. The previous version kept the 2 defaults and
only INSERT()-ed extras beyond index 2, on an unverified
assumption about Insert()'s index semantics. This version
instead first makes the dest stop COUNT match the source's
(deleting extras or inserting - confirmed via reflection that
GradientStops.Delete(index) exists alongside Insert) using
fresh, re-read indices/counts at every step rather than
precomputed ones, then does one final explicit pass setting
every stop's Color/Position/Transparency by matching index -
so even if an insert's initial values didn't stick exactly
right, the final pass corrects them.

## `CopyElementStyle`

Format painter. Source is addressed via the standard
slideIndex+shapeIndex (ResolveTopLevelShape, same as every other
single-shape PowerPoint tool) rather than a "sourceShapeIndex"
field, so it gets the existing nested-shape rejection for free.
Never touches position/size. Cross-slide (user-requested,
2026-09-22): each target now carries its own slideIndex, not just a
shapeIndex resolved against the source's slide - a breaking schema
change from targetShapeIndexes:number[] to targets:{slideIndex,
shapeIndex}[], safe since this tool has never shipped on main.
