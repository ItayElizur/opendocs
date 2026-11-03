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

`ReadHeadings` is the one implementation, shared by `get_headings` (no cap) and the outline in `get_document_context` (first 25;
`more` says there are further ones). Two corrections made when the outline was added, both measured in Word on a
3,009-paragraph document with 14 headings, a table (with a heading inside a cell), floating text boxes and footnotes:

- Numbering. A heading's index used to come from `ParagraphIndexResolver`, a forward `Paragraph.Next()` march from the top of the
  document: 16,257 ms for the 13 headings it found. It is now the previous heading's index plus
  `doc.Range(previousStart, start).Paragraphs.Count` (and `Range(0, start)` for the first jump): 467 ms for all 14, with every
  index and offset identical to Word's own positional lookup (0 mismatches).
- The first paragraph. `wdGoToHeading` only finds headings after the cursor, so a heading at offset 0 was never returned. It is now
  added separately when the first paragraph's `OutlineLevel` is not body text (1-9 means a heading, 10 body text).

The level is the paragraph's `OutlineLevel`, falling back to the digits in the style name. NOT VERIFIED AGAINST LIVE WORD beyond
those measurements.

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

## `read_blocks` object annotations and `get_document_context` counts

`read_blocks` and `get_document_context` only ever saw `doc.Content` and `doc.Paragraphs`, so a
document built from text boxes, shapes or canvases looked empty (a floating shape is a single
anchor character in the body). They now report non-body objects without changing what a
paragraph index means:

- `get_document_context` appends one counts-only "Objects: ..." line (tables, charts, SmartArt,
  shapes with how many hold text, inline pictures, content controls, footnotes, endnotes,
  header/footer parts, TOC). It is omitted when the document has none, so plain documents'
  output is unchanged. No per-object list: the model has to call `read_object` anyway.
- `read_blocks` appends "  {table#0, shape "TextBox 3" textbox hasText, ...}" after each
  paragraph that holds or anchors an object (capped at 8 labels per block). Only the objects
  anchored inside the requested paragraphs are collected (`CollectObjectAnchorsInRange`, using
  Word's range-scoped collections) and merged during the existing forward `Paragraph.Next()` walk
  by `ObjectAnchors.LabelsFor` (OfficeAi.Shared, unit tested). The end of the range is one
  positional `Paragraphs[endIndex + 1]` lookup (about 4 ms on 3,000 paragraphs); the walk stays
  forward-only. The per-paragraph `Range.Start/End` reads only happen when the range holds an object.
  Floating shapes are tagged by name, not index: see WordTools.Objects.cs.md for why.
- The suffix is not paragraph text. It is appended after the closing tag in `html` mode too.

NOT VERIFIED AGAINST LIVE WORD: builds clean, but the anchor offsets (`Shape.Anchor.Start`,
`Footnote.Reference.Start`, `Table.Range.Start`, `ContentControl.Range.Start`) were not
checked against a real document, nor was timing on a large one.
