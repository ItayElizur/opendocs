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
