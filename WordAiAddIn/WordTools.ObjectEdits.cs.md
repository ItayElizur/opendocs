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

- **Content controls: placement.** `doc.ContentControls.Add(type, range)` failed with "the current selection
  partially covers a plain text content control" whenever the user's cursor was inside a content control, and
  worked once the cursor was moved out (even right next to it). So Word consulted the selection, not the range we
  passed. `AddContentControl` binds the add to the range (`range.ContentControls.Add(type)`), and if Word still
  refuses it moves the selection to the range for the call and restores it. (An earlier attempt assumed the
  trouble was the insertion point touching the previous control; the tester's result disproved that.)
  NOT VERIFIED AGAINST LIVE WORD after the change.
- **Notes: `afterText`.** Only the first match used to be reachable, which is the wrong anchor when the phrase
  repeats in the paragraph. Choosing the Nth match by number was rejected (the model cannot see the matches to
  count them, and it adds a parameter a weaker model can misuse). Instead a phrase that matches more than once is an
  error listing each match with ~20 characters of context (first 5), so the model retries with a longer phrase it
  can see is unique. NOT VERIFIED AGAINST LIVE WORD.
  A phrase that cannot be made unique with context (the first word of a paragraph, or identical repeats like
  "the the") is handled by the optional `match` (1-based) - the number printed in that error list. It is only useful
  after the model has seen the list, so it is not blind counting like the rejected `occurrence` parameter.
  Matching does not use Word's `Find`: `Range.Text` contains control characters that are not visible words - a
  footnote/endnote reference mark (0x02) between "the" and " mouse" made `Find` miss "the mouse" and showed up as a
  square in the error snippets. The paragraph is read with field codes and hidden text included (so text index =
  document offset, checked by length), cleaned by `OfficeAi.Shared/TextAnchors.cs` (unit-tested), searched
  case-insensitively there, and the mark is inserted at the mapped document position right after the last matched
  character (before a reference mark that already follows it).
- **Headers/footers: fields.** `set`/`clear` assign `Range.Text`, which destroys fields. The result reports how many
  were replaced. `{PAGE}`, `{NUMPAGES}`, `{DATE}`, `{TIME}`, `{FILENAME}`, `{AUTHOR}` and `{TITLE}` in `text` are
  inserted left to right (text, `Fields.Add` for each token, more text) after the old content is cleared, and `read_object`
  reads the part with field codes on and maps them back to the same tokens (`OfficeAi.Shared/FieldTokens.cs`,
  unit-tested). Deliberately small: no format switches, no other field types; an unmapped field reads back as its
  displayed result. NOT VERIFIED AGAINST LIVE WORD.

## Code review fixes (2026-11)

Checked in a separate Word instance (the COM calls, not through the add-in):

- **Header/footer fields under Track Changes.** The old content is cleared (a tracked deletion in the header's own story) and the new
  content is appended before the part's final paragraph mark, so nothing is searched for afterwards. Result after accepting the
  header's revisions: `Page 1 of 1 / 1` with 3 fields. (The earlier Find-based approach also came out right for the input tried,
  so the reviewer's double-`{PAGE}` failure was not reproduced; the new build simply does not depend on where Word puts tracked
  text.) Header and footer revisions live in `HeaderFooter.Range.Revisions`, not `Document.Revisions`.
- **TOC levels.** Assigning `UpperHeadingLevel` above the current lower level fails with "Command failed" AND leaves the document
  without the table of contents. `edit_toc` therefore checks `upper <= lower` first and writes the lower level first when the new
  upper level exceeds the current lower one (1-3 -> 4-6 worked in that order).
- **`add` that fails half way.** Setting `Checked` on a plain-text control throws; the new control used to stay behind. Content
  controls and text boxes now delete what they just created before rethrowing.
- **Text box position.** `AddTextbox` positions were relative to the anchor paragraph (2/2 = column/paragraph); the code now sets
  page-relative (1/1) before `Left`/`Top`, matching "from the page corner".
- **Text box index.** Found by the shape's `ID` instead of its name (names can repeat).

Not run in Word: the `afterText` changes (the displayed-text matching and the end-of-cell guard are unit-tested in
`OfficeAi.Shared.Tests`), and the `edit_textbox` type restriction.

## After live testing (2026-11)

- **Content control `add` with a field the type cannot hold** (`checked` on a plain-text control). Cleaning up after the
  failure left the control's text behind as a tracked deletion under Track Changes, so the call now rejects `checked` /
  `entries` for the wrong type before anything is created, and the cleanup delete runs with tracking off. `edit` runs the
  same check against the existing control's type.
- **`edit_toc` edit 1-3 -> 4-6 failed with "Command failed" in the live add-in and left the TOC as "Error! Not a valid
  heading level range."** Not reproduced: the same two writes (lower first, then upper) succeed in a separate Word instance
  with Track Changes on and off, inside an undo record, in a visible window, on a document with and without headings.
  The edit now writes a line to `OpenDocsDebug.log` after each step to show which write fails.
- **Text box `x`/`y`** are measured from the page corner, so `x=72, y=72` lands on the first line of a page with 1-inch margins.
