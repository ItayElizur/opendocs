# WordTools.Commands.Style.cs

## `BulletPresets`

PP-12 Task 2: fixed, explicit preset set - each implemented by
applying Word's own proven default bullet/number list (rather than
constructing a ListTemplate from a gallery index, which the plan
itself flags as unstable across Office versions/locales) and then,
where the preset needs more than the default, overriding the
resulting level's NumberStyle/NumberFormat explicitly. The two
Wingdings-glyph variants (diamond/checkbox) are the least certain
of the seven without an interactive Word session to verify against -
flagged in this plan's verification file; narrow the enum to drop
them if they don't render correctly (Step 7's sanctioned fallback).
