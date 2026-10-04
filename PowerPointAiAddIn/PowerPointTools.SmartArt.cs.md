# PowerPointTools.SmartArt.cs

## `ResolveSmartArtLayout` - name-matching vs. index lookup

PP-22 Task 2 Step 4: index-based lookup (SmartArtLayouts is
index-addressable, and the built-in gallery order is stable across
installs) was considered as a locale-independent alternative to
name-matching. Not switched - index stability across Office versions
is an assumption no better founded than the display-name assumption,
and names at least fail loudly with a diagnostic message below,
whereas a wrong index would silently insert the wrong diagram.

## `ResolveSmartArtLayout` - non-English install diagnosis

Distinct from the unknown-key case above: the key was valid, but
this Office install's gallery has no layout under that display
name - the one diagnosis that actually points at a non-English
install, which a silent fallback could never surface.

## `ListSmartArtShapesOnSlide`

Ported from WordTools.cs (2026-08-27) to bring PowerPoint's SmartArt
surface to parity with Word's, which already had edit/read.

The one real difference from Word's version: Word finds SmartArt in a
flat document (doc.InlineShapes + doc.Shapes), while PowerPoint's
shapes live per-slide. So a diagram is addressed by (slideIndex,
smartArtIndex-within-that-slide) rather than a single document-wide
index - slide-scoped indices match how every other PowerPoint tool
here addresses shapes.

## `AddSmartArt` - clearing default placeholder nodes

Post-hoc fix (2026-08-24, found via Word's identical port of
this code, PP-23): AddSmartArt seeds the diagram with the
layout's own default "[Text]" placeholder nodes. Without
clearing them first, the requested items were appended after
the placeholders instead of replacing them.
