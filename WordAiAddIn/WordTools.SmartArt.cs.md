# WordTools.SmartArt.cs

## `ResolveSmartArtGalleryItem`

Post-hoc addition (2026-08-24, user-reported: "smart art has no
change style/color for an existing element"). Unlike layouts,
SmartArt color schemes and quick styles are NOT a fixed enum in
this Office object model (confirmed via reflection against the
referenced Office 15 PIA - Microsoft.Office.Core has no
MsoSmartArtColorType/MsoSmartArtQuickStyleType at all) - they are
live COM collections (Application.SmartArtColors /
.SmartArtQuickStyles) of SmartArtColor/SmartArtQuickStyle objects,
each with a .Name populated at runtime by this install's own
gallery. Rather than guess a curated list of exact display-name
strings (unverifiable without a live session, and a wrong guess
would either fail or - worse - silently match nothing), this
resolves by case-insensitive SUBSTRING match against whatever
names this install actually has, and a miss lists the real
available names so the caller can retry correctly instead of
guessing blind a second time.

## `AddSmartArt` - clearing default placeholder nodes

Post-hoc fix (2026-08-24, user-reported): AddSmartArt seeds the
new diagram with the layout's own default placeholder nodes
(the same "[Text]" prompts the ribbon's SmartArt gallery shows) -
same bug shape as the chart-data fix above (pre-seeded content
never cleared before writing). Without clearing them first, the
requested items were APPENDED after the placeholders instead of
replacing them, leaving visible "[Text]" nodes above the real
ones. Delete every existing node before adding the real ones,
same idea as the chart fix's sheet.Cells.Clear().

## `ListSmartArtShapes`

PP-23 Task 5: SmartArt shapes are not chart shapes and are not
tables - a small, separate list-and-resolve helper, mirroring
ListChartShapes'/ResolveTable's shape but for shape.HasSmartArt
instead of shape.HasChart.

Post-hoc fix (2026-08-24, user-reported): HasSmartArt returns an
MsoTriState, not a real bool, exactly like HasChart elsewhere in
this file - a plain (bool) cast either throws (silently swallowed
by the try/catch below) or never matches, so no shape was ever
recognized as SmartArt and read_smartart/edit_smartart always
reported "no SmartArt diagrams" even right after add_smartart had
just created one. Fixed with the same (int)x == -1 comparison
ListChartShapes already uses for HasChart.

## `EditSmartArt` - `set_layout` case

Post-hoc addition (2026-08-24, user-reported: "smart art
cant change layout"). Reuses ResolveSmartArtLayout
verbatim (same curated 7-key map + gallery lookup
add_smartart already uses) - SmartArt.Layout is
settable (confirmed via the same reflection pass that
found .Color/.QuickStyle), so changing an EXISTING
diagram's layout is the same resolve-then-assign shape
as creating one, just against smartArt.Layout instead
of the AddSmartArt call.
