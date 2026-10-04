## SmartArtLayouts

```
/// Friendly SmartArt layout key to the layout's display name in Office's
/// built-in gallery. Shared by Word and PowerPoint, which previously held
/// byte-identical copies.
///
/// The values are English display names, matched case-insensitively
/// against `Application.SmartArtLayouts`. That is a real constraint worth
/// knowing: on a non-English Office install the gallery's display names
/// differ, and the lookup finds nothing. Both apps surface that as a
/// distinct error rather than silently falling back to some other layout,
/// because "no layout by that name in this install" is the one diagnosis
/// that actually points at the localisation problem.
///
/// Resolving a key to a live layout object stays app-side: it walks
/// `Globals.ThisAddIn.Application.SmartArtLayouts`, and `Globals` is
/// generated per add-in project, so it cannot move here.
```
