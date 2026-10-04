# Comparison with genoffice

"genoffice" is a sibling, from-scratch web-based Office clone suite that officeoffice's
tool design was originally ported from (`Word`/`Excel`/`PowerPoint` only — Outlook has
no genoffice counterpart at all, see [`docs/tools/outlook.md`](tools/outlook.md)). This
page summarizes the high-level gaps and advantages; see each app's own doc under
[`docs/tools/`](tools/) for the full tool-by-tool detail.

## Explicitly out of scope everywhere (per project scope, not gaps)

- `web_search`, `image_search`, `generate_image`, `analyze_media` — no `ai-search`
  equivalent; air-gapped deployment target. (`get_attachment` is a partial exception
  for Outlook — it reads local + OpenXML attachment text, nothing remote, no
  PDF/images; see [`docs/tools/outlook.md`](tools/outlook.md).)
- The PDF app and the Markdown app have no officeoffice counterpart (Markdown's
  scope is folded into Word).
- PowerPoint's `execute_slide_script` DSL and entire deck-generation pipeline (see
  [`docs/tools/powerpoint.md`](tools/powerpoint.md)'s "Missing entirely" list).

## Summary: what genoffice has that officeoffice doesn't

1. **Web-sourced content** — no search, no AI image generation, no media analysis,
   no chat-attachment reading, anywhere. Image tools are local-file-only by design.
2. **PowerPoint's scripting DSL and generation pipeline** — no `execute_slide_script`,
   no `generate_deck`/`regenerate_slide`, no automatic QC/audit pass, no `dataSource`
   provenance enforcement on charts (Excel or PowerPoint).
3. **Block-indexed document context** — genoffice's `get_document_context`/
   `get_deck_context` return structured per-block/per-element inventories; Word's and
   PowerPoint's officeoffice equivalents return flat text previews only.
4. **`add_shape`'s preset breadth** in Excel/PowerPoint (26 named presets each) is
   still narrower than genoffice's "full OOXML preset-geometry set", and
   `set_data_validation`'s `checkbox` kind is unsupported (no COM equivalent exists).

## Summary: what officeoffice has that genoffice doesn't

1. **Real Word comments** (`add_comment`) — anchored, native, available in every
   editing mode including Comment Only. genoffice's docs surface has no comment tool.
2. **A real native `TablesOfContents` TOC** on `insertToc`, vs. genoffice's hand-built
   TOC field-XML workaround (needed because genoffice's own renderer doesn't
   paginate).
3. **A native error-cell scan** (`find_cells` with `errors_only`, via
   `SpecialCells(xlErrors)`) — a categorical COM-vs-Office.js/web advantage called out
   explicitly in the source.
4. **Server-enforced editing modes** (Read Only / Comment Only / Track Changes / Full
   Autonomy, or Outlook's four relabeled equivalents) as a first-class, uniformly-applied
   gate across all four apps' tool dispatch — genoffice's docs app has Track-Changes-aware
   writes but no equivalently formal, uniform mode-gating system across its apps.
5. **Live selection-push into context** (Word/Excel/PowerPoint/Outlook) — a real Office
   selection event pushes the user's current selection into `buildContext()`
   automatically on every turn, rather than app-side selection-range plumbing.
6. **`add_pivot` with calculated fields** — Excel's pivot op supports
   `PivotTable.CalculatedFields().Add(name, formula)` for formula-derived pivot
   values, in addition to row/column/page/data fields.
7. **PowerPoint's Slide Master tooling** (`add_master_element`/`read_master_elements`/
   `set_master_element_transform`/`remove_master_element`/`set_headers_footers`/
   `list_layouts`) and its native-clipboard cross-slide shape copy/move
   (`copy_element`/`move_element`) — both areas with no genoffice equivalent at all.
