# WordTools.Objects.cs

## `read_object`

One read tool for every non-body Word object, replacing `read_table`, `read_chart` and
`read_smartart`. `kind` is table | chart | smartart | shape | content_control | header_footer |
footnote | endnote | toc. Each kind has its own 0-based index space, so `read_blocks` keeps
meaning "paragraph index" and never has to double as a shape index. Omitting `index` lists the
kind (one summary line each, with the paragraph it is anchored in); `index` reads one in full.
The old tools defaulted a missing index to the first object (tables, charts) or read every
diagram (SmartArt) - list mode is the uniform replacement. The old tool names are gone from the
C# switch and `entry.ts`; there are no hidden aliases.

Index orders match what the edit tools already use: tables in `doc.Tables` order; charts and
SmartArt via `ListChartShapes`/`ListSmartArtShapes` (inline first, then floating).

### Shapes

`shape` covers `doc.Shapes` minus charts and SmartArt (excluded with the same `HasChart` /
`HasSmartArt` test the chart/SmartArt lists use, so no shape is in two kinds). The has-text flag
is `Shape.TextFrame.HasText` (an MsoTriState; true is -1, tested as != 0), guarded by try/catch
because `TextFrame` throws on pictures and lines. Groups and canvases report hasText when any
child does (`GroupItems`/`CanvasItems`, depth-capped at 3 because canvases can nest groups) and
are read recursively. Only anchors in the main story (`Anchor.StoryType == 1`) are mapped to a
paragraph; shapes anchored in headers/footers are listed but have no paragraph.

`doc.Shapes` is enumerated with `foreach`, never `Shapes[i]` (see the positional-indexing
incident in CLAUDE.md).

### Header/footer parts

Only parts that apply are listed: first page / even pages only when the section enables them,
parts linked to the previous section are skipped (they repeat its text), empty parts are skipped.
The index is into that filtered list, so it is for reading only; the write tool addresses a
part by section/which/type instead.

NOT VERIFIED AGAINST LIVE WORD: compiles clean in Debug and Release, but none of the new kinds
were exercised against a running Word (shape text, canvas/group recursion, content controls,
header/footer filtering, footnote/TOC reads, paragraph mapping). The enum values used (MsoShapeType
1/5/6/9/13/17/20/24, WdContentControlType 0-9, WdInlineShapeType 3/4, WdStoryType main = 1) and the
members used (Shape.Anchor/TextFrame/GroupItems/CanvasItems/AlternativeText, TextFrame.HasText/TextRange,
HeaderFooter.LinkToPrevious, TableOfContents.UpperHeadingLevel) were confirmed by reflection against
the referenced PIAs.

## Cost of collecting objects, and long text (2026-11)

Measured with a hidden Word instance driven over COM (late-bound, so comparable to this file's `dynamic` calls), on a
generated document: 300 anchored text boxes, 3,261 paragraphs, 20 tables, 30 footnotes, 20 content controls.
Best of 2-3 runs:

- Classifying shapes the way `ListShapes` does (7 COM calls each): ~3,150 ms. Only `Anchor.Start` per shape: ~1,730 ms,
  so the `Anchor` lookup (~5 ms each) is most of the cost and trimming the other calls barely helps.
- Tables 182 ms (59 ms for `Range.Start` alone), content controls 100 ms (50), footnotes 143 ms (86), header/footer
  parts 90 ms. The plain `.Count` properties together: 19 ms; `doc.Shapes.Count`: 3 ms.
- Range-scoped alternative for a 5-paragraph range: `Range.ShapeRange.Count` ~4 ms, footnotes + controls + tables +
  inline shapes counts ~11 ms, shape name + anchor through the range ~10 ms.

So `read_blocks` as written (it collects every object in the document on every call) costs ~3.7 s on that document,
~85% of it shapes. Not yet changed: asking Word for only the objects in the requested range would make it scale with
the range instead of the document. The `Objects:` counts line, which is called once per conversation, now stops
classifying shapes above `ShapeScanCap` (100, about 1 s) and reports just the total.

Long text: `read_object` returns at most `ObjectTextCap` (4000) characters of a text box, note, header/footer or TOC,
chosen by `offset`, and the reply names the offset for the next window (`OfficeAi.Shared/TextWindow.cs`, unit-tested).
The old behaviour was a silent cut at 2000 characters with no way to read the rest. NOT VERIFIED AGAINST LIVE WORD
beyond a clean build.
