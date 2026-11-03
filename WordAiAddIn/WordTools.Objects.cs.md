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

## read_blocks collects only the requested range (2026-11)

The measurements above showed the problem: `read_blocks` collected every object in the document on every call,
and every shape fetch costs ~3.6 ms whichever property is read (Name, Type, ID, HasChart: all ~1.1 s per 300 shapes;
`Anchor.Start` 1.4 s; `Shapes.Item(i)` 2.4 s), so trimming calls per shape could never fix it. It now uses Word's
range-scoped collections for just the paragraphs being read (`CollectObjectAnchorsInRange`):

- Same document: the old document-wide collection measured **5,756 ms** per call (chart pass + SmartArt pass + shape
  classification + tables + notes). The scoped version measured: no objects in range 24 ms; 5 paragraphs 47 ms;
  50 paragraphs 85 ms; 200 paragraphs 339 ms; 1,000 paragraphs (the maximum, ~130 objects) 1,016 ms. It now scales
  with what is in the range, not with the document.
- The range end is one positional `Paragraphs[endIndex + 1]` lookup: 4 ms on 3,260 paragraphs (so it is not the
  O(n^2) trap, which is a lookup per loop iteration). `Range.Paragraphs.Count` over 1,000 paragraphs: 15 ms.
- Tables: index = `doc.Range(0, table.Start).Tables.Count`. Verified in Word, nested table included
  (`doc.Tables` is top-level only).
- Content controls: the same formula **minus 1**. `Range(0, start)` also overlaps the control's own opening tag, so it
  counts itself; measured +1 for every control, nested ones too. Footnotes/endnotes: `Footnote.Index - 1`.
- Floating shapes, charts and SmartArt are tagged by **name** (`shape "Text Box 3" textbox hasText`), not by index.
  A shape's index is its position in `doc.Shapes` (charts/SmartArt excluded), and finding it means fetching every earlier
  shape: ~3.6 ms each, i.e. the very cost being removed. `read_object` therefore accepts `name` for kinds shape,
  chart and smartart (it resolves the name to the index from the same collection it already builds). Inline charts and
  SmartArt keep `chart#N` / `smartart#N` (a scan of the inline shapes, only when the range holds one). Names are
  Word's auto names ("Text Box 3") unless the user renamed shapes; two shapes with one name resolve to the first.
- Shapes anchored outside the main story (headers, footers) are skipped; they still appear in `read_object` listings.

The `Objects:` counts line, called once per conversation, stops classifying shapes above `ShapeScanCap` (100, ~1 s) and
reports just the total.

Test harness caveat: building the benchmark document with content controls made the hidden Word instance stall or crash
(RPC 0x800706BE) in three separate runs when they were added to a 3,000-paragraph document, so the timings above were
taken without them (their cost, measured earlier on a smaller document: 20 controls ~50-100 ms). The content-control
index rule was verified on a 200-paragraph document with a nested pair. NOT VERIFIED AGAINST LIVE WORD end to end:
the C# itself has only been built; the timings are of the equivalent COM calls driven from PowerShell.

Long text: `read_object` returns at most `ObjectTextCap` (4000) characters of a text box, note, header/footer or TOC,
chosen by `offset`, and the reply names the offset for the next window (`OfficeAi.Shared/TextWindow.cs`, unit-tested).
The old behaviour was a silent cut at 2000 characters with no way to read the rest. NOT VERIFIED AGAINST LIVE WORD
beyond a clean build.
