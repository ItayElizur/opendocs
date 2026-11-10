# `WordTools.Direction.cs`

## Why `ApplyAutoDirection` exists

Hebrew written through the Word object model (`Range.Text`, a table cell, a text
box) comes out left-to-right and left-aligned. Word's editor flips direction while a
person types, but a COM write does not. Confirmed by hand with the mock server:
Hebrew into a body paragraph, a table cell, a text box and a SmartArt node all
stayed left-to-right. The direction is therefore decided per write from the
paragraph's own text with `TextUtil.DetectDirection`, the same detector PowerPoint
uses.

## Why only `ReadingOrder`, never `Alignment`

Checked against an installed Word with a throwaway document (alignment read back
after switching `ParagraphFormat.ReadingOrder`; 0 left, 1 center, 2 right, 3 justify):

| Alignment before | LTR to RTL | RTL to LTR |
|---|---|---|
| left | right | right |
| right | left | left |
| center | center | center |
| justify | justify | justify |

Setting `ReadingOrder` mirrors a left or right alignment and keeps a centered or
justified one, and setting the order a paragraph already has changes nothing. So
the helper writes `ReadingOrder` only. It needs no "currently right-to-left" guard
as PowerPoint does, and a deliberate center or justify survives in both directions.
`Range.Information` (character position) read back unreliably in a hidden window and
was ignored; only the alignment values were used.

## Why per paragraph

A range holding paragraphs of different direction reads `ParagraphFormat.ReadingOrder`
as `9999999` (undefined), and `replace_blocks` or `insert_content` can write several
paragraphs at once, so the text of each paragraph decides its own direction.
Neutral text (digits, punctuation, empty) changes nothing.

## Not verified

Rendering of the result in a live document was not looked at by eye, and the add-in
build and every call site were not exercised against running Word.
