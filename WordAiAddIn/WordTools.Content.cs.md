# WordTools.Content.cs

## `ParagraphIndexResolver`

Resolves a Range's 0-based paragraph index (matching
ActiveDoc.Paragraphs' own indexing, which read_blocks/
apply_commands/find_text/get_headings all address by) without
Word's slow positional Paragraphs[i] lookup - indexing the
Paragraphs collection by position has to re-walk the document
from the start on EVERY single access, which turned a scan of N
positions into roughly O(N^2) internally (confirmed root cause of
a real freeze report). This instead marches forward once via the
cheap Paragraph.Next() chain, and only as far as needed since the
last call - callers must request positions in non-decreasing
document order (true for both find_text and get_headings, which
each only ever move forward through the document), so the total
marching work across a whole call is O(N), not O(N) per lookup.

## `FindText`

Read-only search - unlike apply_commands' find_replace, this never
touches the document. Added because there was previously no way to
locate text without either mutating it (find_replace) or reading
the whole document paragraph-by-paragraph via read_blocks.

Plain-substring queries use Word's own native Find engine - the
same one behind Ctrl+F - which does a single optimized traversal
and only costs work proportional to the number of MATCHES, not the
number of paragraphs in the document (the original implementation
scanned every paragraph via positional Paragraphs[i] indexing
regardless of match count, which is what caused a real reported
freeze on a large document). Word's Find has no regex mode (only
its own more limited wildcard syntax), so a regex:true query still
needs a per-paragraph scan - but via the cheap forward
Paragraph.Next() chain, not positional indexing.

## `GetHeadings`

Navigation-Pane-style outline: every Heading-styled paragraph with
its index and level, so the model can see document structure
without reading every paragraph via read_blocks.

Uses Word's own wdGoToHeading jump - the same internal heading
index that powers the Navigation Pane and "Browse by Heading" -
which lands directly on each heading without ever touching a
non-heading paragraph, instead of scanning every paragraph's style
name to find the ones that are headings.

## `TextReadBlocksMaxParagraphs`

Post-hoc addition (2026-08-27, user-reported): 'text' mode (the
default) previously had NO cap at all - only one Range.Text read
per paragraph, so it was never capped the way 'html' mode was, but
an unbounded range on a very large document still means an
unbounded amount of walking and an unbounded output string (a
separate, non-perf concern - context/token budget). Not
independently benchmarked against a measured time budget the way
read_formats' 200-cell cap or html mode's 100-paragraph cap were -
chosen conservatively; raise it if real usage shows it's too tight.
