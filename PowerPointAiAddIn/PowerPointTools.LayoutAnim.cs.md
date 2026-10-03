# PowerPointTools.LayoutAnim.cs

## `ResolveCustomLayout`

Mirrors WordTools.cs's ResolveSmartArtGalleryItem (PP-23): resolves
by case-insensitive substring match against this deck's own live
theme layouts, since custom layout names are not a fixed enum -
a miss lists the real available names so the caller can retry
correctly instead of guessing blind. Shares its exact-match-
preference + ambiguous-match detection with PowerPointTools.
Master.cs's ResolveLayoutByName via the same ResolveLayoutByQuery
core (review finding: this used to silently take the first
substring match with no ambiguity check at all - unlike
ResolveLayoutByName's own multi-design search, this function's
scope is unchanged, always just this one slide's own Design).

## `SetSlideLayout` - empty layoutName guard

Review finding: unlike ResolveMasterTarget's explicit
!string.IsNullOrEmpty(query) check for the identical field
on add_master_element/read_master_elements/etc., this path
never guarded against an empty string - IndexOf("") matches
every layout, so it silently picked the first one (or threw
a confusing "matches more than one" error with >1 layout)
for input that named nothing.
