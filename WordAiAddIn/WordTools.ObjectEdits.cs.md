# WordTools.ObjectEdits.cs

## Why five tools and not one `add_object`/`edit_object`

`read_object` could be one tool because its params are uniform (`kind`, `index`). The write
payloads are not: a content control has a type, title, tag, choice list and checked state; a
header is addressed by section/which/type; a note needs an anchor paragraph and optionally anchor
text; a TOC has heading levels; a text box has geometry. One generic tool would need a large per-kind
`oneOf` that the model re-reads on every call, plus a big dispatcher in C#, for a small saving.
So each family gets one tool with a small `action` enum over params that every action shares:
`edit_content_control`, `edit_header_footer`, `edit_note`, `edit_toc`, `edit_textbox`. The existing
`add_table`/`edit_table`, `add_smartart`/`edit_smartart`, `edit_chart` and `add_image` are left alone
for the same reason.

All five are mutating, so they are advertised from Track Changes up with no tier-list entry (they are
in neither `readOnlyTools` nor `commentOnlyExtraTools`, and not in `AlwaysAllowedTools`). `Execute`
forces `TrackRevisions` on for them like any content edit, so under Track Changes these edits are
revision-tracked.

## Anchors

`RangeAfterBlock` collapses to the END of the paragraph range, past the paragraph mark, i.e. into the
next paragraph. That is right for inserting whole blocks but wrong for a footnote mark or a content
control, which belong at the end of the paragraph's TEXT. `RangeAtEndOfBlockText` trims the mark
(`MoveEnd(wdCharacter, -1)`) first. `edit_toc add` is the exception: it uses `RangeAfterBlock`, the same
convention as `insert_content`, because a TOC is block-level.

## Indices

Indices are `read_object`'s per-kind indices (content controls, footnotes, endnotes, TOCs in document
order; text boxes in the `shape` list). Add and delete shift later indices, and a TOC changes paragraph
indices wholesale; the tool results and descriptions tell the model to re-read.

## Details worth keeping

- Headers/footers: for section 2+ `LinkToPrevious` is cleared before writing, otherwise the write goes
  through to the previous section. First-page / even-page parts are only shown when the section enables
  them, so writing one turns that option on (`OddAndEvenPagesHeaderFooter` is effectively document-wide).
- `edit_textbox edit` refuses groups/canvases for text edits: their text lives in child shapes, which
  `read_object` can show but nothing here can address.
- Dropdown `entries` on edit replace the whole list (`DropdownListEntries.Clear()` first).

NOT VERIFIED AGAINST LIVE WORD: compiles clean in Debug and Release, and every member used
(ContentControls.Add, ContentControl.Delete/DropdownListEntries/Checked/ID, Footnotes/Endnotes.Add,
Footnote.Index, TablesOfContents.Add, TableOfContents.Update/UpdatePageNumbers/Delete, Shapes.AddTextbox,
PageSetup.DifferentFirstPageHeaderFooter/OddAndEvenPagesHeaderFooter, HeaderFooter.LinkToPrevious) was
confirmed by reflection against the referenced PIA, but none of it was run against a real document. The
calls go through `dynamic`, so a wrong argument shape surfaces only at runtime.

## Found in first manual testing (2026-11)

- **Content controls: placement.** `ContentControls.Add(type, range)` failed with "the current selection
  partially covers a plain text content control" on the second add. The default position (end of the last
  paragraph's text) is the edge of the control the previous call created. `OutsideContentControls` now steps the
  insertion point past the end tag of any enclosing control (`ParentContentControl`) before adding. The exact
  Word behaviour is a hypothesis from the error text, NOT VERIFIED AGAINST LIVE WORD after the change.
- **Notes: `afterText`.** Only the first match used to be reachable, which is the wrong anchor when the phrase
  repeats in the paragraph. `occurrence` (1-based, default 1) picks the Nth match; one past the number of
  matches is an error that reports how many exist, so the model can correct itself instead of mis-anchoring.
- **Headers/footers: fields.** `set`/`clear` assign `Range.Text`, which destroys fields (PAGE, NUMPAGES, DATE...).
  The result now reports how many were removed. Inserting fields is not supported yet; `read_object` shows a
  field's displayed text, not the field itself. A token syntax (e.g. `{PAGE}` in `text` becoming a PAGE field,
  and the same tokens in the read output) is the likely design if this is wanted.
