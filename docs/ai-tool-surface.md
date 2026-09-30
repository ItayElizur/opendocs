# AI Tool Surface Reference — officeoffice

This document catalogs every mechanism by which the LLM (the AI assistant embedded in
the Word/Excel/PowerPoint/Outlook VSTO add-ins) can read or mutate a real Microsoft
Office document — or, for Outlook, the mailbox and calendar — in this repo. The
Word/Excel/PowerPoint sections compare against the equivalent surface documented for
`genoffice` (`C:\dev\genoffice\docs\ai-tool-surface.md`, a from-scratch web-based
Office clone suite that this project ports its tool design from); the **Outlook**
section (added 2026-08-27) has no genoffice counterpart and mirrors
`C:\dev\mcp-outlook` instead.

> **Note on `docs/tool-surface-todo.md`**: that checklist has been retired (see
> PP-8, `docs/superpowers/plans/2026-08-23-pp08-retire-stale-todo.md`, 2026-08-24)
> and now just points here. It was written against an earlier snapshot and marked
> many command/operation kinds as `[ ]` unimplemented that are, as of the current
> `main` (commit `1a478c5`), fully implemented. This document was produced by
> reading the actual current `WordTools.cs`, `ExcelTools.cs`, and
> `PowerPointTools.cs` end-to-end, not by trusting that file.

> **Update 2026-08-26 (search/replace + heading outline):** a user-reported gap -
> Word had no read-only search tool at all (only `apply_commands`' mutating
> `find_replace`, which also under-reported its replacement count - always 0 or 1
> regardless of how many occurrences it actually replaced, now fixed to loop and
> count accurately), Excel's `find_cells` had no write-side counterpart, and
> PowerPoint had neither search nor find/replace of any kind. Added: Word's
> `find_text` (read-only search) and `get_headings` (Navigation-Pane-style
> heading outline, `[index] H<level>: text`); Excel's `find_replace` op under
> `propose_operations` (text-value cells only, never formulas); PowerPoint's
> `find_text` (read-only, across shape text + notes) and `replace_text` (every
> text-frame shape + notes, NOT table cells or SmartArt node text - those keep
> using `edit_table_cell`/`edit_smartart`). Not re-verified against every section
> below - see each app's tool table for the pre-existing state this fixes.
>
> **Update 2026-08-27 (Word find_text/get_headings performance fix):** both
> tools' first cut scanned every paragraph via positional `Paragraphs[i]`
> indexing - `Paragraphs` is not a real array in Word's COM object model, so
> each indexed access re-walks the document from the start, turning a full
> scan into roughly O(n²) internally. A user reported this as Word visibly
> freezing on a large document (the automation call runs synchronously on
> Word's own UI thread, so it can't pump its message loop until the call
> returns). Fixed: `find_text`'s plain-substring path now uses Word's native
> `Range.Find` engine (the one behind Ctrl+F) via a `wdFindStop`-wrapped
> `Execute()` loop, costing work proportional to match count, not document
> size; `get_headings` now uses `Range.GoTo(wdGoToHeading, wdGoToNext)` - the
> same internal heading index that powers the Navigation Pane / "Browse by
> Heading" - so it never visits a non-heading paragraph at all. Both still
> need to report a 0-based paragraph index matching `read_blocks`/
> `apply_commands`' convention; resolving that no longer uses `Paragraphs[i]`
> either - a shared `ParagraphIndexResolver` marches forward via the cheap
> `Paragraph.Next()` chaining method, visiting each paragraph at most once
> across a whole call. `find_text`'s `regex:true` path still needs a
> per-paragraph scan (Word's Find has no regex mode, only its own more
> limited wildcard syntax), but via that same cheap forward chain now, not
> positional indexing.
>
> **Update 2026-08-27 (Excel find_cells/find_replace performance + scope):**
> a related but milder issue than Word's - `find_cells`' plain-substring path
> (and `find_replace`, added earlier the same day) read `.Text`/`.Formula` on
> every single cell in a sheet's `UsedRange` via a `foreach` loop. Not the
> same O(n²) bug (`foreach` over `Range.Cells` is a real enumerator, not
> positional re-indexing), but still one COM round-trip per cell regardless
> of match count, which adds up on a sheet with a large `UsedRange`. Fixed:
> the plain-substring path now uses Excel's own native `Range.Find`/
> `FindNext` (the engine behind Ctrl+F/Ctrl+H) - `xlValues`/`xlFormulas` are
> separate native passes (Excel's `Find` only searches one `LookIn` mode per
> call), de-duped by address for `look_in:"both"`. `regex:true` still needs
> the per-cell scan (same reason as Word - no regex mode, only wildcards).
> `find_replace` mirrors this: native `Find`/`FindNext` locates candidate
> cells fast, then only the matched cell's literal text `Value2` is read/
> replaced directly (same safety scope as before - never a formula, since a
> numeric/date/formula cell that merely *displays* a match via its formatted
> text has a non-string `Value2` and is skipped). **Also, per user request, a
> scope default change**: both tools previously searched the whole workbook
> when `sheetId` was omitted; they now default to the **active sheet only**
> (matching Ctrl+F/Ctrl+H's default "Within: Sheet"), with a new `allSheets`
> boolean to opt into workbook-wide search ("Within: Workbook") - `sheetId`
> still names one specific sheet directly, active or not.
>
> **Update 2026-08-27 (Word: the same positional-indexing bug found in 3 more
> places, plus a read_blocks cap):** a broader sweep for the same
> `Paragraphs[i]`-indexing anti-pattern (per user request) turned up three
> more hotspots, all now fixed the same way (`Paragraph.Next()` forward
> chaining instead of positional indexing):
> - **`ResolveTargetParagraphs`** - the shared Target matcher behind
>   `apply_commands`' `updateTextStyle`, `updateParagraphStyle`,
>   `deleteBlocks`, `createParagraphBullets`, and `deleteParagraphBullets` -
>   scanned every paragraph in the document via positional indexing
>   regardless of how narrow the Target filter was. Its return type also
>   changed from `List<int>` to `List<(int Index, Word.Paragraph Paragraph)>`
>   - every caller used to re-look-up `paragraphs[i + 1]` per match after
>   this function already had the paragraph in hand during its scan; now it
>   just hands the object back, removing that second round of positional
>   lookups too. All 5 call sites updated to match.
> - **`read_blocks`'s plain-`"text"` format** (the default) had no upper cap
>   at all - only `format:"html"` was capped, at 100 paragraphs. Added a
>   1000-paragraph cap for text mode (not independently benchmarked the way
>   `read_formats`' 200-cell cap or html mode's 100-paragraph cap were -
>   chosen conservatively, documented as such in the code) and switched its
>   indexing to the same forward-chaining walk.
> - `find_text`'s tool description and Word's system prompt now explicitly
>   tell the model that `find_text`'s returned `[index]` is the exact same
>   0-based paragraph index `read_blocks`/`replace_blocks`/
>   `apply_commands`' `Target.blockIndexes` use (no translation needed), and
>   to prefer `find_text` over reading a large range blindly.
>
> **Update 2026-08-27 (Phase 0 complete - shared pure-logic seam, plus two
> tool-facing bug fixes):** implemented
> `docs/superpowers/plans/2026-08-27-phase0-test-seam.md` in full (6 tasks,
> 6 commits, `dotnet test` 23 -> 90 passed). No tool *schema* changed. Two
> deliberate behavior changes did land, each its own commit:
> - **Hex-color validation** (`propose_operations`/`apply_commands`/
>   `set_element_fill` etc., any op taking a hex color): a malformed color
>   used to reach the model as a raw .NET exception - `"#abc"` threw
>   `ArgumentOutOfRangeException`, `"#GGGGGG"` threw `FormatException`,
>   neither naming the bad value or the expected format. Now: 3-digit CSS
>   shorthand (`"#abc"` == `"#aabbcc"`) is accepted, and anything else throws
>   a clean `ArgumentException` naming the offending value and the expected
>   `#RRGGBB` form.
> - **Word's `apply_commands`**: `set_bold`/`set_italic`'s `"value": null`
>   and `set_heading`'s `"level": null` used to reach `GetBoolean()`/
>   `GetInt32()` and throw an opaque `InvalidOperationException` where the
>   existing "missing required field" message belongs. Now rejected with
>   that same clean message. **Deliberately not a global rule** - Excel's
>   `set_cell`/`set_range` legitimately use `"value": null` to clear a cell,
>   so Excel's `RequiredFields` validation is untouched; rejection is
>   opt-in per field via a new `NonNullFields` table, Word-only so far.
>
> Also: a `Dictionary<string, MsoAutoShapeType>` cannot cross an assembly
> boundary from `OfficeAi.Shared` into the VSTO app projects - confirmed via
> `CS1769` ("embedded interop type" as a generic type argument). Excel's and
> PowerPoint's shape-type maps (previously separately duplicated, one with
> two extra aliases) are now one union table in `OfficeAi.Shared.ShapeTypes`,
> carried as `Dictionary<string, int>`, each app casting to
> `MsoAutoShapeType` at its own call site - the same split `ColorUtil`
> already used for color. Side effect: both apps' "unknown shapeType" error
> now lists the full union of valid names, not just each app's own subset.

> **Update 2026-08-27 (Word gains `barStacked`; chart types unified):** pulled
> forward out of Phase 2 at user request, in two commits.
> - **`feat(word)`:** Word's `edit_chart` now accepts **`barStacked`**
>   (`xlBarStacked`, 58). Word's chart-type map had 7 entries to Excel's and
>   PowerPoint's 8, so Word could not draw a stacked bar chart while the other
>   two could. Word's `entry.ts` enum honestly advertised only 7, so schema and
>   handler agreed — a genuine **capability gap**, not a mismatch. Its
>   `chartType` enum gains the value. Secondary fix: `read_chart`'s reverse
>   type-code lookup used the same map, so reading an existing stacked bar
>   chart reported *"unrecognized chart type code 58"* instead of `barStacked`.
> - **`refactor(shared)`:** the three per-app maps are now one
>   `OfficeAi.Shared.ChartTypes.ByName`. Sequencing the Word fix first was
>   deliberate — it made all three byte-identical, so the extraction itself
>   changed no behavior at all. Tests assert the exact `xlChartType` codes
>   (unlike `ShapeTypes`, where raw ints would be brittle): a wrong code here
>   is a **silent** wrong result, and that bug has shipped here before —
>   PowerPoint's copy mapped `"bar"` to `51`/`xlColumnClustered` instead of
>   `57`, so `chartType:'bar'` drew a column chart *and reported success*.
>   Tests also guard drift against every app's `entry.ts` enum in both
>   directions. `dotnet test` 90 → 102.
>
> **Not verified against live Office** — `barStacked` follows the identical
> code path as the seven chart types already supported (a `Dictionary` lookup
> handing an int to the same COM property), and all three apps build clean,
> but no live Word instance was reachable to confirm a stacked bar chart
> actually renders. Worth one manual check.

> **Update 2026-08-27 (Phase 2 duplication cleared; PowerPoint SmartArt parity;
> set_bullet):** three changes, at user request.
>
> **1 - Remaining cross-app duplication united into `OfficeAi.Shared`:**
> `ComRetry` (was `TransientComHResults`+`RetryTransientCom`, duplicated
> Word/PowerPoint - the copies differed only in whether they took a `label`)
> and `SmartArtLayouts` (was `SmartArtLayoutNames`, byte-identical in both).
> Resolving a layout key to a live object stays app-side: it walks
> `Globals.ThisAddIn.Application`, and `Globals` is generated per project.
> With the chart-type maps done earlier the same day, **the Phase 2
> duplication inventory is now empty.** `dotnet test` 102 -> 114.
>
> **2 - PowerPoint gained `read_smartart` and `edit_smartart`**, reaching
> parity with Word (which already had all three). One real adaptation, not a
> copy: Word finds SmartArt in a flat document (`InlineShapes`+`Shapes`),
> while PowerPoint's shapes are per-slide - so a diagram is addressed by
> `(slideIndex, smartArtIndex-within-that-slide)`, matching how every other
> PowerPoint tool addresses shapes. `read_smartart` is read-only and is
> registered in **both** the client `READER_TOOLS` list and the C#
> `AlwaysAllowedTools` gate; registering only one would have the client offer
> a tool the server then refuses.
> **NOT VERIFIED AGAINST LIVE POWERPOINT** - compiles clean in Debug and
> Release, but the SmartArt COM paths (`Nodes.Add`/`Delete`,
> `.Color`/`.QuickStyle`/`.Layout` assignment) were never exercised against a
> running instance. `set_style` and `set_layout` resolve against the local
> install's gallery **by English display name**, so a non-English Office will
> fail them with the "not in gallery" error. Needs a manual pass.
>
> **3 - `apply_commands` accepts `set_bullet`.** A model sent that kind and
> got back only "unknown command kind" - a dead end. The guess was reasonable:
> its neighbours are `set_bold`/`set_italic`/`set_heading`, while the real
> commands are camelCase `createParagraphBullets`/`deleteParagraphBullets`.
> `set_bullet` is now an accepted alias (same `target`, plus
> `value:true|false` picking which underlying command runs) and is in the
> schema so it is discoverable, not merely forgiven. Separately, **both
> Word's `apply_commands` and Excel's `propose_operations` now list every
> valid kind when rejecting an unknown one**, so any wrong guess is
> recoverable on retry - the same defect class as the malformed-hex-colour
> error fixed in Phase 0.

> **Update 2026-08-27 (Phases 1+3 - tool files split into partial classes):**
> the three `*Tools.cs` files were split by tool area into `partial class`
> file sets - Word 2,531 lines -> 10 files (2,676 total), Excel 2,183 -> 10
> (2,310), PowerPoint 1,820 -> 10 (1,956). Largest file is `WordTools.Content.cs`
> at 452 lines - 2 over the ~450 target; left as-is rather than re-splitting
> for 2 lines. **Structure only: no method body, tool schema, `entry.ts`, or
> system prompt changed**, and `dotnet test` stayed at 114. Each split was
> verified by an order-independent member-set diff (sorted declaration list
> identical before and after), which is what makes an otherwise unreviewable
> move-diff trustworthy. All three apps rebuilt clean in both Debug and
> Release afterward - Release matters because `deploy/package.ps1` only signs
> manifests in that configuration.
>
> Two things worth knowing before adding code to these projects:
> - The `.csproj` files are **classic format with explicit `<Compile Include>`
>   items** - they do not glob `*.cs`. A new file must be added by hand, or it
>   is silently not compiled and the error points at a call site elsewhere.
>   `tools/split-partial.py` did this split and can do further ones; it
>   refuses to run unless every member is assigned to exactly one destination.
> - Word's `ListChartShapes`/`ListSmartArtShapes` stayed `internal` (not
>   `private`) because `WordAiAddIn/TaskPaneHost.cs` calls them from outside
>   the class - the partial-class split keeps that legal since it's still one
>   class, one assembly.
>
> The cross-app duplication this phase might otherwise have surfaced
> (`ComRetry`, `SmartArtLayouts`) was already resolved into `OfficeAi.Shared`
> the same day, before this split ran - see the Phase 2 update above.

> **Update 2026-09-17 (Outlook gains color-tag/category tools):** three new
> tools, `OutlookAiAddIn/OutlookTools.Categories.cs`. `list_color_categories`
> (read-only) lists the profile's master `Namespace.Categories` list — name +
> friendly color name (the 26 `OlCategoryColor` values, mapped in that file
> since `OfficeAi.Shared` has no Outlook PIA reference, same split as
> `ColorUtil` uses for RGB). `set_event_categories` sets/clears an
> `AppointmentItem`'s `Categories` string (the colored block shown on a
> calendar event) — Full autonomy only. `set_category_color` creates a new
> tag or recolors an existing one via `Categories.Add`/`Category.Color` —
> also Full autonomy only. Verified against a live Outlook client (category
> enumeration, tag creation, and event `Categories` round-trip) before
> wiring into the tool switch; `dotnet test` unaffected (no new pure logic —
> the color-name map is a small dictionary, not extracted for unit testing).

> **Update 2026-09-27 (Outlook gains reschedule tools):** two new tools in
> `OutlookAiAddIn/OutlookTools.Calendar.cs`, filling the "no way to move an
> existing event" gap (previously only `draft_event`/`create_event` could
> place a *new* event at a time). **Fully replaced 2026-09-28 by
> `edit_event`/`draft_edit_event`** (see that dated update below) — everything
> described in this block carried over unchanged onto the new tool names.
> `draft_reschedule_event` (Draft tier) sets `Start`/`End` on the resolved
> `AppointmentItem` and `Display(false)`s it unsaved, same shape as
> `draft_event` — the user reviews the moved time and decides whether to
> save/send. `reschedule_event` (Full autonomy only, `SendTierTools`) sets
> `Start`/`End` then `.Save()` (no attendees) or `.Send()` (attendees present)
> to dispatch the reschedule notice, mirroring `create_event`'s
> attendee-present branch. Both required `event_id`, `start`, and `end`, with
> a clean `IsError` naming whichever of `start`/`end` is missing rather than
> defaulting or throwing a raw COM exception.
>
> **Organizer-only, by design.** Both tools first check
> `AppointmentItem.MeetingStatus`. On `olMeetingReceived` (the user is only an
> *attendee* on someone else's meeting, not the organizer) they refuse up
> front with an `IsError` explaining the user has no authority to move it and
> pointing at Outlook's own "Propose New Time" UI — see "Excluded / deferred"
> below for why no tool attempts that path itself. `olNonMeeting`/`olMeeting`
> (the user's own appointment or meeting) go through unchanged.
>
> **Verification status: unproven at runtime**, same caveat as the rest of
> this section (see "Unproven at runtime" below) — no live Outlook was
> available to exercise either tool's COM calls. One piece *was* checked
> concretely rather than assumed: whether the interop PIA exposes a "Propose
> New Time" member at all. .NET reflection against the actually-referenced
> `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA (same verification
> discipline `CreateEvent`'s comments use) confirmed neither
> `AppointmentItem`/`_AppointmentItem` nor `MeetingItem`/`_MeetingItem` expose
> anything with "Propose" (or "Counter"/"NewTime") in its name —
> `_AppointmentItem.Respond` only takes an `OlMeetingResponse`
> (Accept/Decline/Tentative), no counter-proposal overload. "Propose New
> Time" is real Outlook-client functionality, but it isn't reachable through
> the classic COM object model this add-in automates, which is why the
> `olMeetingReceived` case is a hard refusal rather than a half-working
> attempt.
>
> **Update 2026-09-28: fixed a gap in the organizer-authority check, wired
> into undo/redo.** The original check only excluded the exact
> `olMeetingReceived` value; `OlMeetingStatus` actually has 5 values
> (`olNonMeeting`, `olMeeting`, `olMeetingReceived`, `olMeetingCanceled`,
> `olMeetingReceivedAndCanceled`), so an attendee's copy of a meeting the
> organizer has since canceled (`olMeetingReceivedAndCanceled`) — and the
> organizer's own canceled copy (`olMeetingCanceled`) — both fell through to
> the silent `.Save()` path, reporting a `Mutated: true` "Rescheduled" success
> on a canceled or not-actually-yours meeting. Both tools now check for
> either canceled status first (a clean `IsError`: "has been canceled, so
> there's nothing to reschedule"), then the (now also two-valued)
> received-meeting check. Also merged with PR #23 (undo/redo) and #20
> (`set_event_availability`), which this branch predates: `reschedule_event`
> now records a `Start`/`End` snapshot (undo-able) on its `.Save()` branch and
> a barrier (not undo-able) on its `.Send()` branch, mirroring `create_event`'s
> two branches. `draft_reschedule_event` needs no wiring — like `draft_event`,
> it never saves/sends anything itself. (Carried over unchanged onto
> `edit_event`/`draft_edit_event` later the same day — see the dated update
> below.)

> **Update 2026-09-28 (Outlook gains cancel tools):** two more tools in
> `OutlookAiAddIn/OutlookTools.Calendar.cs`, right beside `edit_event`/
> `draft_edit_event` and reusing their `IsCanceledMeeting`/
> `IsReceivedMeeting` organizer-authority checks rather than duplicating them.
> `draft_cancel_event` (Draft tier) opens the item unchanged either way — for
> a meeting the user organizes, they cancel it themselves via Outlook's own
> Cancel Meeting/Send Cancellation buttons; for a plain appointment, via
> Delete. Refuses (clean `IsError`) on an already-canceled event (points at
> `cancel_event` instead — see below) or a meeting the user only attends
> (points at `decline_meeting`, not "Propose New Time" — canceling isn't
> something an attendee can do to someone else's meeting at all, unlike
> rescheduling where Outlook at least has a UI-level counter-proposal feature
> this add-in can't reach). (Originally set `MeetingStatus = olMeetingCanceled`
> unsaved before `Display(false)`, mirroring `draft_edit_event`'s (then
> `draft_reschedule_event`'s) unsaved `Start`/`End` — **removed the same day** after live testing showed
> Outlook persists that change when the Inspector closes even without the
> user clicking "Send Cancellation," silently canceling the meeting locally
> with attendees never notified. Confirmed live, fixed same-day by never
> mutating the item in `draft_cancel_event` at all.)
>
> `cancel_event` (Full autonomy only, `SendTierTools`) does it immediately:
> sets `olMeetingCanceled` and `.Send()`s the cancellation notice for an
> organized meeting, or skips straight to removal for a plain appointment;
> either way the item is then moved to Deleted Items (recoverable there),
> same as `delete_email`'s non-permanent path, not permanently deleted.
> Still refuses on a meeting the user only attends and isn't canceled yet
> (points at `decline_meeting`) — but **an already-canceled event (either the
> organizer's own `olMeetingCanceled` copy or an attendee's stale
> `olMeetingReceivedAndCanceled` one) is treated as cleanup, not refused**:
> there's nothing new to notify anyone of, so it just moves straight to
> Deleted Items, same as a plain appointment. Added the same day as the
> tools themselves, once live testing surfaced there was otherwise no way to
> dismiss a canceled event at all (not even the one this branch's own
> `draft_cancel_event` bug, above, could accidentally create).
>
> **Undo/redo:** `cancel_event` on a plain appointment, or on an
> already-canceled event, records a move (undo moves it back out of Deleted
> Items), matching `delete_email`'s pattern. `cancel_event` on a still-active
> organized meeting is a barrier, not a move snapshot — the cancellation
> notice already went out, so restoring the calendar entry would only
> half-undo the action and misleadingly imply it was fully reversed.
> `draft_cancel_event` needs no wiring — it never saves/sends.
>
> **Verification status: unproven at runtime**, same caveat as `edit_event`
> above — no live Outlook was available. One specific sequence is new and
> untested even relative to that PR: `cancel_event`'s organized-meeting branch
> calls `.Move()` right after `.Send()` on the same item, which no existing
> tool in this add-in does (every other `.Send()` call — `create_event`'s
> invite branch, `edit_event`'s notify branch — only reads properties
> off the item afterward, never mutates its folder). Whether Outlook still
> allows relocating a just-canceled-and-sent appointment is unconfirmed; the
> code fails soft (catches the exception, reports the notice went out
> regardless and tells the user to delete manually if the move didn't take) —
> see the manual test steps for this specifically.

> **Update 2026-09-27 (Outlook gains `set_event_availability`):** one new
> tool, added to `OutlookAiAddIn/OutlookTools.Categories.cs` right beside
> `set_event_categories`/`set_category_color` since it's the same shape (a
> single-property `AppointmentItem` mutation with a friendly-name map for an
> Outlook-PIA-only enum). Sets `AppointmentItem.BusyStatus` — the calendar's
> "Show As" dropdown — to one of the 5 `OlBusyStatus` values (`olFree`,
> `olTentative`, `olBusy`, `olOutOfOffice`, `olWorkingElsewhere`; names
> confirmed via .NET reflection against the referenced
> `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA, not assumed) via a
> case-insensitive friendly-name lookup (`free`/`tentative`/`busy`/`out of
> office`/`working elsewhere`); an unrecognized string returns a clean
> `IsError` result listing the valid names rather than throwing, unlike
> `set_category_color`'s `ParseColor` (which throws and relies on the outer
> `ExecuteAsync` catch). Unrelated to `Categories`/color tags — a distinct
> `AppointmentItem` property entirely. Placed in **Draft tier**
> (`DraftTierTools` + `entry.ts`'s `commentOnlyExtraTools`), not Full
> autonomy: it's purely local (`.BusyStatus` + `.Save()`, no `.Send()`),
> the same risk profile that put `set_event_categories`/`set_category_color`
> in Draft tier despite the "Mutating tools" table heading below still
> reading "Full autonomy only" (that heading predates the 2026-09-19
> four-tier gate for those two tools and was already stale before this
> change — not fixed here to keep this update focused). **Update
> 2026-09-28: verified against a live Outlook client** — all 5
> `OlBusyStatus` values round-trip correctly via `.BusyStatus` + `.Save()`,
> case/spacing-tolerant parsing confirmed, invalid values return a clean
> `IsError` instead of throwing, and saving a `BusyStatus` change on a
> meeting the user organizes (has attendees) does not prompt or notify
> attendees. Also wired into the undo/redo stack the same day (see that
> section below) after merging with PR #23.

> **Update 2026-09-27 (delete_email's permanent path split to its own tier):**
> another instance of the same class of gap closed by `aeae77c`/`b3fc5d2`
> (category tools left one gate too low) — here `delete_email` as a whole sat
> in `DraftTierTools`, but the tool bundles two risk classes under one name:
> `permanent: false` (default) just moves the message to Deleted Items, fully
> reversible, correctly Draft-tier; `permanent: true` additionally calls
> `.Delete()` from there, irreversible from within Outlook ("may still be
> server-recoverable" per its own result text — not a claim it can be undone
> here), which belongs at Full autonomy alongside `send_email` and friends,
> not one gate below it. `OutlookTools.cs`'s `ExecuteAsync` tier check is
> otherwise purely name-based (`AlwaysAllowedTools`/`DraftTierTools`/
> `ApprovalTierTools`/`SendTierTools`, checked before `input` is inspected at
> all) — added one narrow, input-aware special case immediately after the
> name-based check: if `name == "delete_email"` and `input.permanent == true`
> and the caller's mode is below Full autonomy, block with a message naming
> Full autonomy specifically (not delete_email in general — the tool is still
> fine at Draft only for the non-permanent path). Deliberately not
> generalized into a per-argument gating system for every tool; `delete_email`
> stays in `DraftTierTools` and in `entry.ts`'s `commentOnlyExtraTools`
> unchanged, since the tool overall is still reachable from Draft only.
> `entry.ts`'s `delete_email` description now says plainly that
> `permanent: true` requires Full autonomy, so the model doesn't attempt it
> needlessly at a lower tier and get a confusing runtime block. No existing
> automated test covers Outlook's tier-gate logic (it lives in the VSTO
> project, not `OfficeAi.Shared`); verified by code review plus
> `OutlookAiAddIn` MSBuild.

> **Update 2026-09-27 (undo/redo tooling — Word, PowerPoint; a custom
> undo/redo stack for Outlook; Excel's removed 2026-09-28):** added
> `undo_last_action`/`redo_last_action` to Word and (reversing this
> document's own earlier "no undo/redo tool" conclusion for it) PowerPoint,
> plus a custom stack of the assistant's own actions for Outlook. **None of this is runtime-verified
> against a live Office/Outlook client** — no live instance was reachable;
> every mechanism below was instead confirmed by .NET reflection directly
> against the exact PIA versions these projects reference (15.0.0.0), the
> same standard of evidence this document already applies elsewhere (e.g.
> the chart-type-map and SmartArt-parity updates above), and by a clean
> Debug build of all four touched projects plus their esbuild bundles.
>
> - **Word** (`WordAiAddIn/WordTools.History.cs`): the initial plan for this
>   feature assumed `Application.Undo()`/`Application.Redo()` exist — they do
>   **not** (reflection against the referenced PIA finds no such method on
>   `_Application`/`ApplicationClass` at all, only an unrelated `UndoRecord`
>   property for grouping automation edits into one user-visible undo entry).
>   The real, callable methods are one level down, on `Document`:
>   `Document.Undo(ref object Times)` / `Document.Redo(ref object Times)`,
>   both returning `bool`. The `Times` parameter is COM-optional (confirmed
>   via `ParameterInfo.IsOptional`), so C# calls them with no arguments -
>   `ActiveDoc.Undo()` / `ActiveDoc.Redo()` - exactly like real-world Word
>   automation code. The `bool` return gives an honest "Undid the last
>   action." vs. "Nothing to undo." (same for redo) instead of always
>   claiming success. Gated the same as every other Word mutating tool:
>   blocked in Read Only/Comment Only, available from Track Changes upward.
> - **Excel — removed (2026-09-28).** The first version dispatched
>   `Application.CommandBars.ExecuteMso("Undo"/"Redo")` (Excel has no
>   `Application.Redo()`, and `Application.Undo()` returns `void`). That
>   can't work for the assistant's edits: every write this add-in makes goes
>   through the object model (`propose_operations`, find/replace, and so on),
>   and Excel clears its whole undo stack on any object-model change. So
>   right after an AI edit there is nothing to undo, and later the tool would
>   undo the *user's* next manual edit instead. `Application.OnUndo` only
>   takes a VBA macro name, so it isn't usable from a C# add-in. A snapshot-
>   based custom undo was considered and not pursued; the tools were removed.
> - **PowerPoint** (`PowerPointAiAddIn/PowerPointTools.History.cs`) —
>   **reversing this document's earlier conclusion that PowerPoint should be
>   excluded entirely.** That conclusion was correct that no
>   `Application.Undo()`/`Redo()` method exists anywhere in the PowerPoint
>   interop surface (still true, reconfirmed here), but missed a distinct
>   mechanism: `Application.CommandBars.ExecuteMso(string)` /
>   `GetEnabledMso(string)` — the generic Office 2007+ ribbon-command
>   dispatch API, exposed via `Microsoft.Office.Core.CommandBars` (the
>   "Office" PIA reference already in every one of these four projects),
>   which every main Office host's `Application.CommandBars` returns,
>   PowerPoint included. Confirmed real and callable by reflection: the
>   `CommandBars` property on PowerPoint's `_Application` returns
>   `Microsoft.Office.Core.CommandBars`, and that type's `_CommandBars`
>   interface declares both `Void ExecuteMso(String)` and `Boolean
>   GetEnabledMso(String)`. `ExecuteMso("Undo"/"Redo")` dispatches by the
>   same ribbon-command ID the real Undo/Redo buttons and Ctrl+Z/Ctrl+Y use
>   internally, so it works despite there being no direct method to call.
>   `GetEnabledMso` is checked first (mirroring Excel above) since
>   `ExecuteMso` itself returns nothing, so without that check every call
>   would have to claim success unconditionally. See
>   `PowerPointTools.Master.cs`'s `RemoveMasterElement` (2026-09-22 incident,
>   in the PowerPoint section below) for why this repo cares: a Slide Master
>   placeholder deletion had no code path at all to undo at the time, one
>   direct consequence of `Application.Undo()` genuinely not existing —
>   `ExecuteMso` is the mechanism that generally closes that "no way back"
>   gap going forward (it does not retroactively change that method's own
>   still-valid placeholder refusal, which stands for an unrelated reason —
>   see that method's comment). Gated the same as every other PowerPoint
>   mutating tool (Track Changes upward).
> - **Outlook** (`OutlookAiAddIn/OutlookTools.Undo.cs`, stack logic in
>   `OfficeAi.Shared/ActionHistory.cs`) — a **custom** undo/redo stack of
>   the assistant's own actions, not a native-undo wrapper. Native Undo via
>   `Explorer.CommandBars.ExecuteMso("Undo")` was tried and tested by hand
>   on 2026-09-28. It's a single slot tied to the window that toggles
>   undo/redo, it doesn't see object-model changes (a `move_email` followed by
>   the ribbon Undo did nothing), and after one use both our tool and the
>   ribbon button failed with "The operation cannot be performed because the
>   message has changed." Outlook has no API to put an object-model change
>   onto that slot.
>
>   How it works: an in-memory stack per mailbox chat (keyed like
>   `ModeByMailbox`, capped at 50, lost on restart). Each mutating handler
>   records one entry after its change succeeds:
>   - **Property snapshots** (before/after values, restored and `Save()`d):
>     `mark_email_read/unread` (`UnRead`), `flag_email_important` (full
>     `Importance`, so Low is preserved), `set_event_categories`,
>     `set_event_availability` (`BusyStatus`), `edit_event` (added 2026-09-28,
>     replacing `reschedule_event`) whenever the call doesn't send/barrier
>     (see below) — `Start`/`End`, `Subject`, `Body`, `Location`, whichever
>     subset was actually touched that call, via a dynamically-built props
>     list (replacing the old fixed `{Start,End}` `RescheduleProps` constant) —
>     including an occurrence-level edit via `occurrence_date`:
>     `RecurrencePattern.GetOccurrence()` returns a real item with its own
>     `EntryID` once saved, so this same mechanism covers it with no new entry
>     type — `set_reminder`, and `update_task` (task fields as a group, or the
>     flagged-mail fields).
>   - **`set_email_reminder`**: if the message wasn't flagged before, undo
>     calls `ClearTaskFlag()`, and redo calls `MarkAsTask` again.
>   - **Moves**: `move_email`, non-permanent `delete_email`, and `cancel_event`
>     on a plain appointment or an already-canceled event (no one to notify
>     either way, so canceling it is just a move to Deleted Items). The
>     folder is resolved by its own EntryID/StoreID via
>     `Namespace.GetFolderFromID`, and the new EntryID after each move is
>     rewritten into every entry for that item.
>   - **Created items**: `create_task` and `create_event` without
>     attendees. Undo moves the item to Deleted Items (recoverable); redo
>     moves it back.
>   - **`set_category_color`**: undo restores the old color, or removes a tag
>     the assistant created.
>   - **Barriers**: `send_*`, `create_event` with attendees, `edit_event`
>     and `cancel_event` on a still-active meeting the user organizes (both
>     send a notice to attendees), `accept/decline/tentative_meeting`, `delete_email
>     permanent:true`, — regardless of plain-appointment vs. meeting — any
>     occurrence-level cancellation via `cancel_event`'s `occurrence_date`
>     (added 2026-09-28: `RecurrencePattern.Exceptions`/`Exception` are
>     read-only via COM, confirmed by reflection, so a deleted occurrence has
>     no API to reverse it), and — again regardless of plain-appointment vs.
>     meeting — `edit_event` on the **whole series** whenever it includes a
>     time change and `event_id` resolves to a genuinely recurring master
>     (confirmed live 2026-09-28: Outlook rejects setting `Start`/`End`
>     directly on a recurring master at all, so this branch writes to
>     `RecurrencePattern` fields instead, which the undo mechanism can't yet
>     read/write). This is the one case where an otherwise-undo-able-looking
>     `edit_event` call (no attendees, no `occurrence_date`) is still a
>     barrier — only true non-recurring events, or a single occurrence, stay
>     undo-able via the dynamically-built props snapshot. `edit_event` is also
>     a barrier any time it results in a `.Send()` — including converting a
>     plain event into a meeting by adding `required_attendees`/
>     `optional_attendees` for the first time (added 2026-09-28, unverified at
>     runtime — see the dated update below). Undo stops at a barrier instead
>     of reaching past it.
>   - **Known gap — `set_event_categories`**: the snapshot only covers the
>     appointment's own `Categories` string. If the assigned name wasn't
>     already in the mailbox's master category list, Outlook auto-adds it
>     on `Save()` with an arbitrary color (see the mutating-tools table
>     below); undo restores the appointment but does not remove that
>     auto-created master category entry, which is a permanent side effect
>     undo can't see or reverse.
>
>   Before reversing, each entry checks that the item still holds what the
>   assistant left there. If it was changed since (by the user or anything
>   else), undo refuses rather than overwrite it. A failed or refused entry
>   is dropped, never retried. `redo_last_action` re-applies undone entries
>   until the next new action. Both tools are gated at Draft tier. **Not
>   verified against a live Outlook client.**

> **Update 2026-09-28 (Outlook gains recurring-series support: occurrence
> targeting, recurrence creation):** two additions to
> `OutlookAiAddIn/OutlookTools.Calendar.cs` and one to
> `OutlookAiAddIn/OutlookTools.Compose.cs`, closing the "recurring occurrences
> share one EntryID" gap flagged in "Structural fragility" below (see that
> section's own correction, same date) and the "no way to create a repeating
> series" gap.
>
> **1 — Occurrence targeting.** `edit_event`/`draft_edit_event`/
> `cancel_event`/`draft_cancel_event` all gained an optional `occurrence_date`
> string param. A new `ResolveOccurrenceTarget` helper resolves it via
> `RecurrencePattern.GetOccurrence(DateTime)` — confirmed present and callable
> via .NET reflection against the referenced PIA, the same evidence standard
> this document uses elsewhere — into one specific occurrence, instead of
> always acting on the recurring master. Omitting `occurrence_date` still acts
> on the whole series (or a non-recurring event), unchanged from before this
> feature.
>
> **2 — Recurrence creation.** `create_event`/`draft_event` gained an optional
> `recurrence` object supporting all 6 `OlRecurrenceType` values (`daily`/
> `weekly`/`monthly`/`monthlyNth`/`yearly`/`yearlyNth`), validated by a new
> pure, unit-tested `OfficeAi.Shared.RecurrenceValidator` (23 tests) before any
> COM call — every validation failure names the specific field and the fix
> (e.g. `"recurrence.days_of_week is required for type \"weekly\"..."`, not a
> generic "invalid recurrence"). Applied via `AppointmentItem.GetRecurrencePattern()`
> before `.Save()`/`.Send()`, after attendee/`MeetingStatus` setup — **that
> ordering is a design choice, not yet verified against live Outlook** (see
> "Unproven at runtime" below). Series creation's own undo/redo is unchanged
> from `create_event`'s existing non-recurring behavior (a `RecordCreated`
> move-to-Deleted-Items entry on the master) — deleting a recurring master
> deletes the whole series, so no new undo mechanism was needed here either.
>
> **3 — Undo/redo asymmetry (the most important nuance of this feature).**
> Occurrence-level **reschedule** stays undo-able, and needed no new
> undo-entry type: `RecurrencePattern.GetOccurrence()`'s returned item becomes
> a real, independently-resolvable item with its own `EntryID` once saved, so
> the existing generic `SnapshotEntry`/`RecordSnapshot`/`ItemEntryIdOf`
> mechanism already works on it exactly like any other item. Occurrence-level
> **cancellation is always a barrier, never undo-able** — plain appointment or
> meeting occurrence, no exception either way. This is a deliberate asymmetry
> from whole-event cancellation (`cancel_event` without `occurrence_date` on a
> plain appointment stays undo-able, per the existing "Undo/redo tools"
> section). The reason: `RecurrencePattern.Exceptions`/`Exception` — the COM
> objects Outlook uses to track per-occurrence deletions — are entirely
> read-only. Confirmed via .NET reflection against the referenced PIA:
> `Exception` exposes only getters (`AppointmentItem`, `Deleted`,
> `OriginalDate`, `ItemProperties`) and no `Delete`/`Remove`/`Add` method
> anywhere, so there is no API to reverse a deleted occurrence once it's gone.
>
> **4 — A suspected bug fixed by reasoning, not by a live-confirmed observation
> (organizer-authority checks against the wrong item).** The four
> occurrence-aware methods originally checked `appt.MeetingStatus`/
> `IsCanceledMeeting(appt)`/`IsReceivedMeeting(appt)`, where `appt` is the
> resolved occurrence or master. A live reschedule attempt against what was
> believed at the time to be a recurring meeting occurrence failed with
> `COMException "Cannot save this item."` Reasoning about `GetOccurrence()`'s
> known-unreliable behavior for this property — rather than a live-confirmed
> observation on a verified meeting occurrence — identified a likely bug:
> `GetOccurrence()`'s returned occurrence item does **not** reliably report
> `MeetingStatus == olMeeting` even when it genuinely belongs to a recurring
> meeting with attendees. If so, this would cause `edit_event` on a
> meeting occurrence to wrongly take the `.Save()` branch — crashing with
> that same `COMException` — instead of `.Send()`, and would cause
> `cancel_event` on a meeting occurrence to silently delete it without ever
> notifying attendees. Fixed by checking
> `master.MeetingStatus`/`IsCanceledMeeting(master)`/`IsReceivedMeeting(master)`
> instead (the master is always reliable) — the actual mutation still targets
> `appt` (the occurrence or master, whichever `ResolveOccurrenceTarget`
> resolved). After applying this fix and testing again, the user confirmed
> the test event had actually been a plain (non-meeting) appointment all
> along — so the original failure had a different, unrelated cause (see
> point 5 below). **Caveat: this fix has never actually been exercised against
> a genuine meeting occurrence with attendees** — every live test run against
> this feature used a plain (non-meeting) recurring series. The code and its
> review are sound, but the specific scenario it was written to fix is
> unverified.
>
> **5 — A second real bug found and fixed (false-negative `.Save()`/
> `.Delete()` failures).** Outlook can throw `COMException "Cannot save this
> item."` from an occurrence's `.Save()` (in `EditEvent`, formerly `RescheduleEvent`) or `.Delete()`
> (in `CancelEvent`, both the plain-appointment and meeting-occurrence
> branches) **even when the mutation already persisted** — confirmed live: a
> reschedule call that reported this exception had, per a follow-up
> `list_events` call, actually moved the occurrence. Property setters on
> `AppointmentItem` write immediately via RPC, independent of `.Save()`/
> `.Delete()`'s own finalize step, which can fail on its own. Fixed by
> wrapping these calls in try/catch and re-verifying the real outcome (via
> `ResolveOccurrenceTarget`) before deciding what to report, instead of
> trusting the exception at face value.
>
> **6 — A third real, confirmed-live constraint: Outlook rejects reordering
> occurrences relative to each other.** Confirmed by reproducing the exact
> error manually in Outlook's own UI: *"Cannot reschedule an occurrence of the
> recurring appointment ... if it skips over a later occurrence of the same
> appointment"* — and separately confirmed that two occurrences of the same
> series also can't share the same calendar day. Both surface from automation
> only as the same generic `COMException "Cannot save this item."` — confirmed
> no `InnerException` carries the specific reason, checked via debug log
> across every occurrence of the failure. A new proactive check,
> `CheckOccurrenceReorderCollision`, now runs before `.Save()` is ever
> attempted (in both `EditEvent` and `DraftEditEvent`, formerly `RescheduleEvent`/
> `DraftRescheduleEvent`, occurrence-level only): it queries the Calendar folder for other occurrences
> of the same series between the occurrence's original date and the target
> date, and if any exist, returns an exact error naming the valid range
> instead of relying on Outlook's generic exception. This is Outlook-level
> behavior, not a limitation of this add-in's own code.
>
> **Tier placement unchanged.** `create_event`/`edit_event`/
> `cancel_event` stay Full-autonomy-only (`SendTierTools`); `draft_*` stay
> Draft-tier. `occurrence_date`/`recurrence` are additional scope on existing
> tools, not new risk categories — no edits to `DraftTierTools`/`SendTierTools`.

> **Update 2026-09-28 (`reschedule_event`/`draft_reschedule_event` fully
> replaced by `edit_event`/`draft_edit_event`):** a **full replacement, not an
> alias** — the retired tools are removed entirely from `entry.ts` and the C#
> switch in `OutlookAiAddIn/OutlookTools.cs`; `edit_event`/`draft_edit_event`
> are the only way to change an event's time (or anything else about it) going
> forward. Same tier placement as before (`edit_event` in `SendTierTools`,
> `draft_edit_event` in `DraftTierTools`/`entry.ts`'s `commentOnlyExtraTools`).
> Generalizes reschedule into a full edit: beyond `start`/`end` (still required
> together if either is given), a single call can now also change `subject`,
> `body`, `location`, and/or replace `required_attendees`/`optional_attendees`
> wholesale (not a diff/merge — the caller supplies the full new list each
> time, reading the current one first via `get_event` if it needs to preserve
> someone) — at least one of the seven optional fields must be given, else a
> clean `IsError` naming all seven. Attendee edits are whole-series/
> non-recurring only (`occurrence_date` + an attendee field together is a
> clean `IsError`); `subject`/`body`/`location`/`start`/`end` remain editable
> at the occurrence level, same as reschedule always was. Every mechanism
> `reschedule_event`/`draft_reschedule_event` had — `ResolveOccurrenceTarget`
> occurrence targeting, the whole-series `RecurrencePattern.PatternStartDate`/
> `StartTime`/`EndTime` fix, the false-negative `.Save()` retry via
> `ResolveOccurrenceTarget` re-verification, `CheckOccurrenceReorderCollision`,
> and the `IsCanceledMeeting`/`IsReceivedMeeting` organizer-authority checks —
> is reused unchanged by `edit_event`/`draft_edit_event`, not duplicated (see
> the "Occurrence targeting" / "whole-series `RecurrencePattern`" / "false-
> negative `.Save()`" / "reordering collision" points in the dated update
> above, now describing `edit_event`). New pieces specific to this change:
> a `ReplaceAttendees` helper — per-category, not a blanket clear: it removes
> existing `olRequired`-type `Recipients` entries only if `required_attendees`
> is given, and existing `olOptional`-type entries only if `optional_attendees`
> is given (omitting one leaves that category's existing attendees untouched;
> the organizer recipient is never touched either way), then re-adds via the
> existing `AddAttendees` helper `create_event`/`draft_event` already use, then
> `Recipients.ResolveAll()`; a call flips `MeetingStatus`
> to `olMeeting` and takes the `.Send()` branch (instead of `.Save()`) whenever
> the event is or becomes a meeting, setting `ForceUpdateToAllAttendees = false`
> explicitly first — confirmed via .NET reflection to be Outlook's own default
> for notifying only added/removed attendees, not everyone, but set explicitly
> here so correctness doesn't depend on that default never changing; and the
> fixed `RescheduleProps` (`{"Start","End"}`) undo/redo constant is gone,
> replaced by a dynamically-built props list per call, snapshotting only the
> fields the call actually touched (`DescribeEditEventChanges` builds the
> matching human-readable summary). A whole-series time change stays a barrier
> exactly as before (`RecurrencePattern` fields aren't reachable by
> `SnapshotEntry`); if other fields are *also* changed in that same call, the
> entire call stays a barrier too — no partial-undo of a mixed pattern-field +
> item-property change.
>
> **Unverified at runtime, ranked by consequence of being wrong** (per the
> design spec's risk-ordering, `docs/superpowers/specs/2026-09-28-outlook-edit-event-design.md`):
> 1. **`ReplaceAttendees`'s recipient-clearing** — highest risk, since a wrong
>    organizer-exclusion could drop the user's own organizer entry or fail to
>    actually clear old attendees, feeding directly into a `.Send()` with real
>    people.
> 2. **Converting a plain event into a meeting via `edit_event`** — adding
>    attendees to an *existing*, previously-saved item (unlike `create_event`,
>    which only ever adds attendees to a brand-new one) is an untested
>    combination.
> 3. **`subject`/`body`/`location` edits on a recurring master** — likely fine
>    (none of these are part of `RecurrencePattern`), but unverified.
>
> None of this has been exercised against a live Outlook client yet — same
> caveat as every other addition in this document's "Unproven at runtime"
> section below, which this update also folds into.

> **Update 2026-09-29 (Outlook gains `tentative_meeting` + draft meeting-response
> tools, plus an optional comment on all three):** `OutlookAiAddIn/OutlookTools.Calendar.cs`'s
> `RespondMeeting` — previously hardcoded to a `bool accept` — was generalized to take
> the actual `OlMeetingResponse` value, so one shared helper now backs `accept_meeting`,
> `decline_meeting`, and the new `tentative_meeting` (same "Automate approvals" tier as
> the other two — it already calls `resp.Send()` to notify the organizer, same rationale
> as the original two). All three also gained an optional `message` parameter: when
> given, it's set as `resp.Body` before `.Send()`, a short comment attached to the
> accept/decline/tentative response. **Unverified live** whether the organizer actually
> sees this text on the delivered response — needs a real received invite to test
> against, not a self-organized item (same category of gap as this document's other
> "Unproven at runtime" entries below). A new `DraftRespondMeeting` helper mirrors
> `RespondMeeting` for Draft tier: same `appt.Respond(response, true, false)` call, but
> ends in `resp.Display(false)` for the user to review and send themselves instead of
> `.Send()`-ing directly — `draft_accept_meeting`/`draft_decline_meeting`/
> `draft_tentative_meeting`, added to `entry.ts`'s `commentOnlyExtraTools` alongside the
> existing draft tools, never record undo/redo (same contract as `draft_event`/
> `draft_edit_event`/`draft_cancel_event` — nothing is saved or sent until the user acts
> on the opened window). `message` pre-fills the same `resp.Body` there, subject to the
> same unverified-live caveat. A second, separate unverified-live risk applies to the
> draft tools specifically: it is **unconfirmed whether `AppointmentItem.Respond()`
> itself commits local calendar changes at call time** — independent of whether
> `.Send()` or `.Display()` is subsequently called — such as replacing the appointment
> with a new EntryID on accept/tentative, or removing/moving the original appointment
> on decline. If Outlook's COM implementation does this, the draft tools' "nothing
> persists until the user acts" contract would not actually hold, even though they
> correctly never call `.Send()`/`RecordIrreversible`/`RecordSnapshot`. Needs a live
> test: open each draft response, close the window without sending, then confirm the
> appointment still exists under the same EntryID, `ResponseStatus` is unchanged, and
> the item is still in the inbox/calendar as before — `draft_decline_meeting` most
> carefully, since decline is the destructive direction.

> **Update 2026-09-29 (PR #29 code review: `draft_respond_meeting` redesign,
> plus smaller `RespondMeeting`/wording fixes):** a code review of the
> tentative_meeting/draft-meeting-response work above (the immediately
> preceding "Update 2026-09-29" block) raised the unverified `Respond()`
> side-effect risk it flagged from a caveat to a Critical finding, and this
> update is the response to it — not just a footnote on the same design, an
> actual redesign.
>
> **The `Respond()` side-effect risk is now the reason a design changed, not
> just a caveat.** The review treated `AppointmentItem.Respond()`'s
> documented/known behavior — committing a real calendar change at call time
> (a new EntryID on accept/tentative, a move to Deleted Items on decline),
> independent of whether the resulting response is ever sent — as
> confirmed-plausible enough that shipping the three draft tools
> (`draft_accept_meeting`/`draft_decline_meeting`/`draft_tentative_meeting`)
> unchanged would mean a "draft" tool could silently alter or destroy a real
> calendar item the instant it's invoked, before the user takes any action at
> all. That's the exact shape of bug `draft_cancel_event` itself hit and fixed
> the same way months earlier (see the "Update 2026-09-28" `draft_cancel_event`
> note above: an unsaved `MeetingStatus` change that Outlook still persisted on
> window-close). The fix follows that same precedent: `draft_respond_meeting`
> (replacing all three retired tools) never calls `Respond()` at all. It just
> opens the original, completely unmodified item via `Display(false)` and lets
> the user pick Accept/Tentative/Decline themselves from Outlook's own native
> ribbon buttons. Since the redesigned tool no longer calls `Respond()` with a
> specific response type, there is no longer a technical reason for three
> separate draft tools — one unified `draft_respond_meeting` replaces them,
> gated the same as they were (`DraftTierTools`/`entry.ts`'s
> `commentOnlyExtraTools`). `message` can no longer be pre-filled into a
> response body (that would require calling `Respond()` to get the
> `MeetingItem`, the exact call this redesign avoids) — it's returned in the
> tool's output text instead, for the user to paste in themselves via
> Outlook's own "Edit response before sending" option. See the
> `draft_respond_meeting` table row below for the mechanics.
>
> **Smaller fixes to `RespondMeeting`/`DraftRespondMeeting` from the same
> review, all in `OutlookAiAddIn/OutlookTools.Calendar.cs`:** a new shared
> `ResolveMeetingAppointment` helper replaces the duplicated
> item-to-`AppointmentItem` resolution logic both methods had. Both methods
> now refuse (`IsError`) up front on a meeting that's already been canceled by
> the organizer, or on an item that isn't a meeting the user was actually
> invited to (an organizer's own `olMeeting` copy, or a plain `olNonMeeting`
> appointment) — via two new shared helpers, `AlreadyCanceledRespondError`/
> `NotInvitedError` — instead of calling `Respond()` unconditionally and
> letting Outlook's own behavior in those cases go unchecked.
> `RespondMeeting`'s subject is now captured *before* `Respond()` runs, since
> `Respond()` replacing the item on accept/tentative (per the risk described
> above) could otherwise leave `appt.Subject` reading a stale reference
> afterward. `RespondMeeting`'s result text now says explicitly when
> `.Send()` fails ("...but the response could not be sent to the organizer")
> instead of unconditionally reporting success — the local `Respond()` had
> already gone through, but the organizer was never notified and any comment
> was lost. Two stale "accept/decline" tool-capability descriptions predating
> `tentative_meeting` (in `entry.ts`'s `trackChangesExtraTools` comment, and
> `OutlookTools.cs`'s comment on the `CommentOnly`/`TrackChanges` tier meaning
> just above `DraftTierTools`) now say "accept/decline/tentatively-respond".
> `draft_cancel_event`'s description, the system prompt's cancel-tools
> sentence, and `ReceivedMeetingCancelError`'s message all pointed an attendee
> who can't cancel someone else's meeting at `decline_meeting` — one tier
> above `draft_cancel_event` itself (Draft only vs. Automate approvals), so a
> Draft-only caller couldn't actually reach that suggestion; all three now
> also mention `draft_respond_meeting` as the Draft-tier-reachable
> alternative.
>
> **Still unverified live, pending the project owner's own test:** whether
> the organizer actually sees the `message`/comment text on the delivered
> response for `accept_meeting`/`decline_meeting`/`tentative_meeting` — this
> redesign didn't touch that mechanism (`resp.Body` set before `.Send()`) or
> add any new evidence toward confirming it; it remains exactly as
> unverified as the immediately preceding "Update 2026-09-29" block already
> described it.

## Architecture

officeoffice drives the **real desktop Office applications** via VSTO + COM interop
(`Microsoft.Office.Interop.{Word,Excel,PowerPoint,Outlook}`), unlike genoffice's
from-scratch web renderers. The chat UI runs in a WebView2 page inside a CustomTaskPane;
tool calls cross a `chrome.webview.postMessage` ⇄ `CoreWebView2.PostWebMessageAsJson`
JSON bridge (`OfficeAi.Shared/ToolProtocol.cs`) into C# handlers that call the COM
object model directly — there is no Electron/IPC hop. Word/Excel/PowerPoint attach one
pane per document window (keyed by `Hwnd`); Outlook attaches one pane per `Explorer`
window (keyed by COM identity, since `Explorer` has no `Hwnd`) and nothing to
Inspectors — see the Outlook section.

`packages/agent-core` and `packages/ai-provider` from genoffice are copied verbatim
into `shared/web-src/{agent-core,ai-provider}` (same `AgentLoop`, same `AgentSkill`
contract, same `maxTurns` default of 8, same multi-provider types including
`genspark`/`anthropic`/`gemini`/`deepseek`/`openai`/`custom`). **However, none of the
three add-ins actually wire up provider selection yet** — each `entry.ts` hardcodes
`streamOpenAiCompatible` against a local test endpoint (`http://127.0.0.1:9000/v1`),
and `onSettingsSave` is a stub ("Not yet wired to the transport/provider config —
deferred"). So the provider-abstraction layer exists but is not live.

> **Stale (PP-0, 2026-08):** the paragraph above predates the shared-app-shell
> refactor. Provider/model/key selection, the settings screen, and the transport
> now live once in `shared/web-src/app-shell/` (`getSettings` / `makeTransport` /
> `onSettingsSave`, persisted in WebView `localStorage`), shared by all four
> add-ins including Outlook — not re-audited tool-by-tool here.

There is **no `packages/ai-search` equivalent** in officeoffice: no `web_search`,
`image_search`, `generate_image`, or `analyze_media` anywhere in the repo. This is a
deliberate scope decision (the deployment target is air-gapped — see
`add_image`/`replace_image`'s explicit rejection of remote URLs below), not an
oversight. **`read_attachment` is a partial exception:** Outlook's `get_attachment`
(added 2026-08-27) saves an email attachment to a local file and extracts text from
text-family and OpenXML (`.docx/.xlsx/.pptx`) types via `OfficeAi.Shared/AttachmentText/`
— but never fetches anything remote, never handles PDF or images, and cannot feed a
binary to the model. See the Outlook section.

Each add-in has a governance layer genoffice's docs surface doesn't have in this
form: a shared `EditingMode` enum (`ReadOnly | CommentOnly | TrackChanges |
FullAutonomy`), filtered client-side (which tools are advertised to the model) *and*
re-enforced server-side in each `Tools.Execute()` (mutating tools blocked outright in
Read Only / Comment Only mode, regardless of what the model requests).

---

## Word (`WordAiAddIn/WordTools.cs`)

### Top-level tools (7, pre-2026-08-26 snapshot — see notes below)

> **Stale count, not rewritten in place** (same convention as the PowerPoint
> section below): this table predates `find_text`/`get_headings` (2026-08-26)
> and `add_image` (2026-08-27ish, contradicting the "No image-insertion tool"
> line right after the table), plus now **`undo_last_action`/
> `redo_last_action`** (2026-09-27 — see the dated Update block near the top
> of this document for what they do and how they're implemented). Actual
> current top-level tool count is 9 (7 below + `undo_last_action` +
> `redo_last_action`); `find_text`/`get_headings`/`add_image` bring the real
> total higher still — see each app's own "Update" block above for detail
> rather than this table.

| Tool | Implemented | Notes vs. genoffice |
|---|---|---|
| `get_document_context` | Yes | Much thinner: paragraph/word count + a flat 300-char text preview. No block-indexed list (index\|type\|preview) like genoffice's version. |
| `read_blocks` | Yes, but plain-text only | Paragraph-indexed range read matches genoffice's index shape, but returns plain text (`[i] text` lines) rather than genoffice's restricted-HTML serialization, so returned content loses formatting/structure markup. |
| `insert_content` | Yes, but narrow | Takes plain `text` only — **always appends at the end of the document**. No HTML, no `afterBlockIndex` positioning, no rich content (images, charts, lists) via this tool. |
| `replace_blocks` | Yes, but plain-text only | Same start/end-index-range shape as genoffice (empty text deletes the range), but the replacement is a plain `text` string set via `.Text` — no HTML/rich formatting on the replacement content, unlike genoffice's HTML-parsing version. |
| `apply_commands` | Yes — gateway, see below | |
| `edit_chart` | Yes, but narrow | Combines genoffice's separate `insert_chart`+`edit_chart` into one create-or-edit call, but only sets a title and a **single series'** numeric values — no categories, no chart-type selection, no multi-series support. |
| `add_comment` | Yes | **Not in genoffice's docs surface at all.** Anchors a real Word comment to the first match of given text. Available in every editing mode, including Comment Only. |
| `undo_last_action` / `redo_last_action` | Yes (2026-09-27) | `Document.Undo()`/`Document.Redo()` (not `Application` — see the dated Update block above), returning `bool` for an honest result message. Available from Track Changes mode upward, same gate as every other content-mutating Word tool. |

No image-insertion tool exists for Word at all (genoffice's docs has `insert_image`) — **stale, see the note above the table: `add_image` exists.**

### `apply_commands` command kinds (12 of 12 genoffice kinds + 4 officeoffice-only aliases — all genuinely implemented)

| Command | Implemented | Notes |
|---|---|---|
| `updateTextStyle` | Yes, but 9/10 fields | bold/italic/underline/strike/sizeHalfPoints/font/color/baselineOffset/link implemented. **`highlight` (text highlight color) is missing** — no handling in `UpdateTextStyle` (WordTools.cs:379-416) and no `highlight` property in the `entry.ts` schema. |
| `updateParagraphStyle` | Yes | align/lineSpacing/indentLeft/indentRight/indentFirstLine/spaceBefore/spaceAfter/pageBreakBefore/shadingFill/borders — full parity. |
| `deleteBlocks` | Yes | Same `Target` matcher (nodeType/headingLevel/containsText/blockIndexes/scope). Deleting every paragraph clears content instead, leaving one empty paragraph (explicitly mirrors genoffice's own guard). |
| `moveBlocks` | Yes | Captures moved paragraphs as OOXML snapshots before deleting, reinserts via `InsertXML` — preserves formatting through the move. |
| `copyBlocks` | Yes | **New — no genoffice equivalent.** Duplicates paragraphs matched by the same `Target` matcher `deleteBlocks` uses, leaving originals in place. Reuses `moveBlocks`' `WordOpenXML`/`InsertXML` capture technique minus the delete step. `afterBlockIndex` may reference one of the copied paragraphs' own indices (allowed, unlike `moveBlocks`). |
| `copyFormat` | Yes | **New — no genoffice equivalent.** Format-painter semantics: copies Font (bold/italic/underline/strike/size/name/color/superscript/subscript/highlight) and ParagraphFormat/Shading/4-side Borders (align/lineSpacing/indents/spacing/pageBreakBefore/shadingFill/border LineStyle+Color per side) from one source paragraph onto one or more target paragraphs, atomically. Hyperlinks are never copied (matches real Word Format Painter). Whole-paragraph granularity only — no sub-paragraph text-run targeting. |
| `createParagraphBullets` / `deleteParagraphBullets` | Yes | Heading paragraphs matched-but-skipped, non-list matched-but-skipped — explicitly mirrors genoffice. |
| `updateImageProperties` | Yes | width/height (proportional scale from current size)/align, targets `doc.InlineShapes` by index. |
| `insertToc` | Yes | Uses Word's **native** `TablesOfContents.Add(UseHeadingStyles: true)` — real, auto-paginating, more direct than genoffice's hand-built TOC field-XML (which has to work around its web renderer not paginating). |
| `set_bold` / `set_italic` | Yes | officeoffice-only convenience aliases for `updateTextStyle`'s bold/italic fields, addressed by paragraph-index range rather than a `Target`. |
| `set_heading` | Yes | officeoffice-only alias, ≈ genoffice's `setHeadingLevel`. |
| `find_replace` | Yes | officeoffice-only alias, ≈ genoffice's `replaceAllText`. |

Every genoffice `apply_commands` kind has a real implementation here — the todo
checklist's claim that these are unimplemented is wrong for the current source.

---

## Excel (`ExcelAiAddIn/ExcelTools.cs`)

### Top-level tools (10)

All 9 read/query tools plus `propose_operations` are implemented and genuinely
functional — full 1:1 parity with genoffice's naming and shape (`get_workbook_context`,
`read_range`, `read_cells`, `select_range`, `read_formats`, `read_sheet_features`,
`find_cells`, `trace_precedents`, `trace_dependents`). `load_guide` has no equivalent
(deliberately out of scope — genoffice's is an internal prompt-budget mechanism for
managing its larger op count in context, not needed at officeoffice's current scale).

No `undo_last_action`/`redo_last_action` for Excel: they were added on 2026-09-27 and
removed on 2026-09-28, because Excel clears its undo stack on any object-model write
(see the dated Update block near the top of this document).

Notable native-COM advantage: `find_cells`'s `errors_only` mode uses
`Range.SpecialCells(xlCellTypeFormulas, xlErrors)` — a genuinely native error-cell
scan the code comments call out as the categorical VSTO/COM advantage over
Office.js's wildcard-only `Range.find`.

### `propose_operations` operation kinds — **all named kinds from genoffice's list are implemented**

Every operation kind genoffice documents (writing, formatting, layout, structure,
charts, table, pivot, data — 51 distinct named kinds across those groups) has a real
handler in the `ProposeOperations` switch. This directly contradicts the now-retired
`tool-surface-todo.md`'s "9 of 65 implemented" claim — that count was from an earlier
snapshot; the current source is functionally complete against genoffice's named op
list. The switch now has **53** kinds total: the 51 above, plus `copy_range`/
`move_range` (arbitrary-block relocation via `Range.Copy`/`Range.Cut`), a superset
addition beyond genoffice's own list — genoffice has no equivalent named kind for
duplicating or relocating an arbitrary rectangular range.

Where officeoffice's version is **narrower** than genoffice's:

| Op | Gap |
|---|---|
| `format_range` | Only `bold`/`italic`/`numberFormat`/`fillColor` (4 properties). Missing: font family, size, font color, underline, strikethrough, align, wrap, rotation, indent, borders — all present in genoffice's version. |
| `add_chart` | Basic path only supports `column`/`line`/`pie` (chartType silently falls back to column for anything else). The richer chart vocabulary (bar/area/doughnut, legend, dataLabels, seriesColors, per-series renaming) only exists on `edit_chart`, not on creation — genoffice supports the richer set on both. |
| `add_image` | **Local file paths only** — remote URLs throw `NotSupportedException` ("air-gapped deployment"). genoffice downloads from `image_search`/`generate_image` results. This is a deliberate scope boundary, not a bug. |
| `add_shape` | 26 named preset types + textbox — narrower than genoffice's "full OOXML preset-geometry set" but still substantial (rect/roundRect/ellipse/triangle/parallelogram/trapezoid/diamond/pentagon/hexagon/octagon/pie/chord/donut/foldedCorner/heart/lightningBolt/sun/moon/cloud/arc/star5/4 arrow directions). |
| `set_data_validation` | `checkbox` kind explicitly rejected — Excel's Data Validation COM API (verified via reflection against the referenced PIA) has no boolean-checkbox validation type; only 8 kinds exist total, none map to it. |

No `dataSource`/provenance-enforcement mechanism exists anywhere (genoffice's slides
app gates chart data-source claims; nothing analogous exists in officeoffice's Excel
or PowerPoint chart tools).

---

## PowerPoint (`PowerPointAiAddIn/PowerPointTools.cs`)

> **Stale as of PP-24 (2026-08-24):** this section predates PP-19 through PP-24 and undercounts the tool list (now 35: the table below plus `delete_slide`/`move_slide`/`duplicate_slide` from PP-19, `set_slide_layout`/`set_slide_transition`/`add_animation`/`read_animations`/`edit_animation` from PP-24, and `duplicate_element`/`copy_element`/`move_element`/`copy_element_style` — same-slide shape duplication, cross-slide shape copy/move via PowerPoint's own native Copy/Paste (uses the real Windows clipboard, an explicit exception to this codebase's usual rule against it - every shape kind is supported, since it reproduces the exact underlying OOXML the same way Ctrl+C/Ctrl+V does), and a shape-level format painter). Not rewritten line-by-line here; see `docs/superpowers/plans/2026-08-24-pp24-powerpoint-layout-transitions-animations.md`, this plan's own doc, and `docs/superpowers/plans/STATUS.md` for the current, accurate state.

> **Update 2026-09-27 (`undo_last_action`/`redo_last_action` added — 35 → 37):**
> this document previously concluded PowerPoint should be excluded from the
> Word/Excel/Outlook undo/redo tooling pass, on the grounds that no
> `Application.Undo()`/`Redo()` method exists anywhere in the PowerPoint
> interop surface. That premise is still correct (reconfirmed here by
> reflection, not just re-assumed), but it missed a distinct mechanism:
> `Application.CommandBars.ExecuteMso("Undo"/"Redo")` — the generic Office
> 2007+ ribbon-command dispatch API, confirmed real and callable on this
> project's referenced PIA (`PowerPointAiAddIn/PowerPointTools.History.cs`;
> full detail in the dated Update block near the top of this document).
> `RemoveMasterElement` right below (2026-09-22 incident) is the reason this
> mechanism matters here specifically: at the time of that incident there was
> no `Application`-level Undo at all, so a bad Slide Master placeholder
> deletion had no code path back — that method now refuses the deletion
> outright for its own, separate reason (see its comment), and
> `undo_last_action`/`redo_last_action` are what generally close the "no way
> back" gap for whatever the model does elsewhere in a session, via ribbon
> dispatch rather than a direct method call. Gated the same as every other
> PowerPoint mutating tool (Track Changes upward). Not runtime-verified — no
> live PowerPoint instance was reachable; both tools build clean and their
> mechanism is confirmed by reflection, not by an actual Ctrl+Z/Ctrl+Y round
> trip.

### Tools (23 total, all genuinely implemented) — pre-PP-19/PP-24 snapshot, see note above

| Tool | Implemented | Notes vs. genoffice |
|---|---|---|
| `get_deck_context` | Yes | Per-slide flat text preview (120 chars), no per-element type/id inventory like genoffice's version. |
| `read_slide` | Yes | Shape name + text per element, indexed. |
| `set_element_text` | Yes | Matches. |
| `set_element_style` | Yes (widened, PP-20) | bold/italic/fontSize/color/fontName/underline/shadow/alignment(left/center/right/justify)/baselineOffset(SUPERSCRIPT/SUBSCRIPT/NONE). Reports which properties actually applied instead of a flat "Style updated." **Strikethrough deliberately not implemented**: unlike Excel (whose `TextFrame2`/`TextRange2` newer text model has it), this PowerPoint PIA has no `TextFrame2` member on `Shape` at all (confirmed via a direct `CS0234` compile failure) — no field was added to the schema for it, so nothing silently no-ops. |
| `set_element_transform` | Yes | left/top/width/height/rotation — matches. |
| `add_text_box` | Yes | Matches. |
| `add_shape` | Yes (widened, PP-20) | Now shares Excel's 26-preset `ShapeTypeMap` (ported, not extracted to `OfficeAi.Shared` — see rationale below), plus `rectangle`/`oval` kept as aliases for `rect`/`ellipse` so existing calls keep working. Unrecognized name errors listing valid ones (previously silently became a rectangle). No fill/line params — chain `set_element_fill`/`set_element_stroke` for those. |
| `delete_element` | Yes | Matches. |
| `add_slide` | Yes | Duplicates a source slide's layout, optional text-clear. |
| `set_element_fill` | Yes | Solid fill or none. |
| `set_element_stroke` | Yes | Color/width or remove. |
| `set_slide_background` | Yes | Single color; `slideIndex = -1` applies to all slides — matches genoffice. |
| `ungroup_element` | Yes | Promotes children; explicitly tells the model to re-`read_slide` for fresh indices afterward. |
| `add_table` / `edit_table_cell` / `edit_table_structure` / `edit_table_style` | Yes | Native `Shapes.AddTable`; structure supports insert/delete row+col; style supports firstRow/bandRow/shading/borders. Reasonably close to genoffice's parity. |
| `add_chart` / `edit_chart` | Yes | Writes real data into the chart's embedded Excel workbook (`ChartData.Workbook`), explicitly closes/releases the COM object to avoid a leaked hidden Excel process. Supports bar/barStacked/line/area/pie/doughnut, legend, dataLabels, gridlines. |
| `add_smartart` | Yes | Maps 7 layout keys (list/process/cycle/hierarchy/pyramid/matrix/venn) to native SmartArt layouts by display name; flat item list only (matches genoffice's own flat-list scope, per source comment). |
| `crop_image` | Yes | Fractional crop against current on-slide size (documented imprecision under repeated crops — no reliable "natural size" once already resized in classic Interop). |
| `replace_image` | Yes, narrow | **Local file path only** (same air-gapped constraint as Excel's `add_image`) — no AI-generation pairing since `generate_image` doesn't exist here. |
| `set_picture_opacity` | Yes | Via `Fill.Transparency`. |
| `undo_last_action` / `redo_last_action` | Yes (2026-09-27, not in the "35" count above) | Via `Application.CommandBars.ExecuteMso("Undo"/"Redo")` + `GetEnabledMso` pre-check — no direct `Undo`/`Redo` method exists on this object model at all. See the Update block right above this table. |

### Missing entirely (confirmed absent from both the C# switch and the advertised tool list)

- `delete_slide` — no way for the AI to remove a slide.
- `execute_slide_script` — no scripting DSL; every multi-property/multi-element edit
  must go tool-by-tool (`set_element_transform` one shape at a time), rather than
  genoffice's atomic AST-interpreted batch script.
- The entire deck-generation pipeline: `ask_clarification`, `plan_deck`,
  `generate_deck`, `regenerate_slide`, `save_style_template`, `list_style_templates`.
- No automatic post-edit audit/QC pass — genoffice's `auditSlideLayout` (geometric
  overflow/overlap/bounds check after every script run) and `slide-qc.ts` (vision-based
  QC sub-loop after generated pages) have no counterpart; nothing here checks the
  result of an edit automatically.
- No `add_comment`-equivalent for PowerPoint (unlike Word), so Comment Only mode
  currently behaves identically to Read Only — a documented gap in the code comments.

### Structural fragility note

Shapes are addressed by **positional index** (`slideIndex`, `shapeIndex` into
`slide.Shapes`) rather than a stable id — indices shift whenever shapes are
added/removed/reordered, unlike genoffice's `sourceId`-based addressing. The model
is instructed to re-read the slide after structural changes (e.g. `ungroup_element`'s
result text says so explicitly), but there's no protection against acting on a stale
index otherwise.

---

## Outlook (`OutlookAiAddIn/OutlookTools.*.cs`)

> **Added 2026-08-27 — current as of this section.** genoffice has no mail or
> calendar app, so there is no "vs. genoffice" column here. The tool set mirrors
> `C:\dev\mcp-outlook` (a self-hostable EWS MCP server) as closely as the Outlook
> COM object model allows — but this add-in drives the **already-running, already-
> authenticated desktop Outlook client** via `Microsoft.Office.Interop.Outlook`, not
> EWS. There is no `Namespace`/credential setup; it acts as the signed-in user.
> **One exception:** `search_contacts` makes a single EWS call (see below) — the COM
> object model cannot do a multi-result directory search.

### Shape of the integration (differs from Word/Excel/PowerPoint)

- **Explorer-only.** The chat pane docks in the main Outlook window (`Explorer`).
  There is **no task pane in Inspector windows** (pop-out read/compose). The ribbon
  button is on the Explorer's Mail tab (`idMso="TabMail"`, via a new
  `RibbonBase.HomeTabIdMso` / `ProvidesRibbonFor` hook) and suppressed on every
  Inspector ribbon surface.
- **Per-mailbox chat.** `GetChatId()` returns `mbx-` + `SHA256(primary SMTP)[:16]` —
  one rolling conversation per mailbox. No file path, so none of Word/Excel/
  PowerPoint's provisional-id / `ChatStore.Migrate` / `DocSettingsStore` save
  lifecycle applies. Primary SMTP is resolved lazily (`ExchangeUser.PrimarySmtpAddress`
  → `PR_SMTP_ADDRESS` proptag → `AddressEntry.Address` → first account → display name).
- **Pane lifecycle.** `Explorer` has no `Hwnd`, so panes are keyed by the COM
  identity pointer (`Marshal.GetIUnknownForObject`). Creation is **lazy** — on an
  explorer's first `Activate` or the ribbon button, never in `ThisAddIn_Startup` —
  because Outlook auto-disables add-ins whose median startup exceeds ~1 s, and
  `install.ps1` also writes `Resiliency\DoNotDisableAddinList`.
- **Editing modes — four tiers, added 2026-09-19.** Outlook repurposes all four
  `EditingMode` slots with its own meaning, rather than dropping two (previously only
  `ReadOnly`/`FullAutonomy` were used): `ReadOnly` → **Read only** (unchanged);
  `CommentOnly` → **Draft only** (every mutating/drafting tool that never leaves the
  mailbox unreviewed — `mark_email_read`/`unread`, `flag_email_important`,
  `move_email`, `delete_email`, `create_task`, `update_task`, `set_reminder`,
  `set_email_reminder`, `draft_email`, `reply_email`, `reply_all_email`,
  `forward_email`, `draft_event`, `draft_edit_event`, `draft_cancel_event`,
  `draft_respond_meeting`); `TrackChanges` → **Automate approvals** (adds
  `accept_meeting`/`decline_meeting`/`tentative_meeting` — these already call `resp.Send()` to notify the
  organizer, so they get their own explicit tier rather than hiding in Draft only or
  Full autonomy); `FullAutonomy` → unchanged name, adds the seven auto-send tools (see
  below). Each tier is a strict superset of the one before it. `entry.ts` passes
  `availableModes: ['readOnly','commentOnly','trackChanges','fullAutonomy']` and a
  `modeOverrides` map so the mode menu shows Outlook's own labels/descriptions instead
  of Word/Excel/PowerPoint's generic "Comment only"/"Track changes" text — those two
  apps' own use of `CommentOnly`/`TrackChanges` for real editing is untouched (Outlook's
  relabeling is purely its own `entry.ts` config, not a shared-string change). Gate
  logic (`OutlookTools.cs`'s `ExecuteAsync`) is an ordinal check on the enum's own
  declared order (`ReadOnly < CommentOnly < TrackChanges < FullAutonomy`). A fresh
  session now defaults to **Draft only**, not Full autonomy (`defaultMode: 'commentOnly'`,
  resolved via `chat-ui.ts`'s `defaultModeFor()` — the same helper Word/Excel/PowerPoint
  use to default to their own `trackChanges`).
- **Selection context.** The Explorer's `SelectionChange` pushes the selected mail
  item(s) / conversation into per-turn context as a `mail` `SelectionContext` variant
  (`shared/web-src/app-shell/bootstrap.ts`) carrying each item's `EntryID`, so the
  model prefers the selection over searching. The scope pill shows the subject (one
  message) or a count (`scopeUnit: 'mailbox'`).
- **`message_id` / `event_id` / `task_id` = Outlook `EntryID`**, re-resolved via
  `Namespace.GetItemFromID` on every call. EntryID is stable within a folder but
  **changes on `Move` and across stores** — `move_email` / `delete_email` return the
  new id. Recurring calendar instances all share the master appointment's EntryID.
- **Native query APIs are mandatory** (see the `native-query-apis` memory / the slow
  Word-search incident): `list_*` use `Folder.GetTable` (an in-memory rowset, no
  per-item COM object); `search_emails` uses `Items.Sort` → `Items.Restrict("@SQL=" + DASL)`;
  a capped linear scan is a fallback only when `Restrict` rejects the filter. The
  DASL builder is pure and unit-tested (`OfficeAi.Shared/OutlookDasl.cs`). `search_contacts`
  is the deliberate, documented exception: server-side EWS ANR (`ResolveName`), not a COM
  scan — see the EWS carve-out below.
- **`search_contacts` and `find_meeting_slots`' work-week default are the two things that
  touch EWS.** Every other Outlook tool is pure `Microsoft.Office.Interop.Outlook` COM
  against the running client. All EWS-dependent code lives apart from the pure-COM tool
  files, in two layers: `OutlookEws.cs` (the raw EWS Managed API wire calls —
  `ResolveNamesAsync`, `GetWorkingHoursAsync`) and `OutlookTools.Ews.cs` (2026-09-19,
  split out of the file this section used to describe — the tool-facing orchestration on
  top: `SearchContactsAsync`, `ResolveWorkWeekAsync`, and the shared endpoint/account
  resolution both call, `ResolveEwsUrlAsync`/`FindExchangeAccountInfo`). Contact resolution
  calls **EWS `ResolveName` twice, `ContactsOnly` then `DirectoryOnly` (both
  `returnContactDetails: true`), merging both result sets** (EWS Managed API 2.2,
  `Microsoft.Exchange.WebServices` 2.2.0) — **fixed 2026-09-30**: the original single-call
  `ContactsThenDirectory` short-circuited on any Contacts-folder hit (whose ANR only matches
  `DisplayName`) and never reached the Directory/GAL phase (whose ANR does cover
  given name/surname), so a query could resolve only against display names. Querying both
  locations unconditionally and merging (`ContactSearchFormat.Format` already dedupes by
  email/name) fixes that at the cost of one extra EWS round trip. Contact resolution
  runs with
  `ExchangeService.UseDefaultCredentials` (Windows Integrated Auth as the signed-in user —
  the .NET equivalent of `mcp-outlook`'s `auth_type=sspi`; no stored credentials). Endpoint
  is parsed from the cached `Outlook.Account.AutoDiscoverXml` (`<EwsUrl>`/`<ASUrl>`, `EXCH`
  preferred over `EXPR`; pure parser `OfficeAi.Shared/EwsAutodiscoverXml.cs`), falling back
  to `ExchangeService.AutodiscoverUrl`, then cached in a process-static `Uri` shared by both
  EWS-dependent tools. The call runs off the UI thread (`await Task.Run`, `svc.Timeout` 15 s)
  so Outlook stays responsive. **On-prem Exchange only.** EWS unreachable / SSPI failure /
  endpoint not found / timeout → a clear `IsError` result for `search_contacts` (EWS isn't
  optional there); `find_meeting_slots` instead falls back to its old hardcoded Sun-Thu/9-18
  default, since EWS is an enhancement over an already-working COM-only path there, not the
  only way to do the job. This is also
  the pilot for async tool execution — the shared `ToolExecutor` delegate is now
  `Task<ToolResult>`-returning (`WebViewBridgeHost.OnWebMessageReceived` is `async`).
  `find_meeting_slots` is the second Outlook tool to go async (2026-09-19, for its own
  EWS work-week lookup — see its row below); every other tool in all four add-ins is
  still synchronous, wrapped in `Task.FromResult`.

### Read tools (11 — always allowed, never gated)

| Tool | Notes |
|---|---|
| `list_emails` | `Folder.GetTable`; columns EntryID/Subject/ReceivedTime/SenderName/UnRead + `PR_HASATTACH` proptag; `[UnRead] = true` restriction when `unread_only`; sorted newest-first; non-mail rows filtered by `MessageClass` not starting `IPM.Note`. Args: `folder`, `limit` (20), `unread_only`. |
| `search_emails` | `Items.Sort("[ReceivedTime]")` then `Restrict("@SQL=" + BuildSearchFilter(...))`. `LIKE '%q%'` on subject + body; UTC-ISO date range; sender by `fromemail =` OR `fromname LIKE` (Exchange senders carry `legacyExchangeDN`, not SMTP — display-name fuzzy match). `ci_phrasematch`/`ci_startswith` are **not** usable via `Restrict` (they throw). `recipient` is a client-side filter, capped 500. Fallback: capped linear scan on a malformed filter. |
| `apply_search` | View-only (`Mutated: false`) — pushes the same `BuildSearchDasl` filter `search_emails` computes into `Explorer.CurrentFolder` + `Explorer.Search("@SQL=" + dasl, olSearchScopeCurrentFolder)`, so the user's own Outlook window shows the results. `_Explorer.Search(string Query, OlSearchScope SearchScope)`'s signature was confirmed via .NET reflection against the referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA; the `@SQL=` DASL string itself (same one `Items.Restrict` already proves works) has **not yet been exercised against a live Outlook session** — first real use should confirm it's accepted as-is by `Explorer.Search`, or fall back to plain-text `query` only. |
| `get_email` | Full `Body` (≤ 40k), To/CC via `Recipient.Type`, `ConversationID`/`ConversationTopic`, importance, unread, and an `attachments` array — `{index (1-based), name, type (byValue/embeddedItem/ole/reference), size}` — feed the index to `get_attachment`. |
| `open_email` | `MailItem.Display(false)` on an existing item resolved via `ItemById` — opens the message in its own Outlook reading window, unmodified. `Mutated: false`. |
| `get_attachment` | `Attachment.SaveAsFile` into `%LOCALAPPDATA%\OutlookAiAddIn\Attachments\`; returns the path. `extracted_text` (≤ 40k) is populated **only** for text-family extensions (`.txt .csv .tsv .md .json .xml .log`, `.html` tag-stripped) and OpenXML (`.docx .xlsx .pptx`), via the swappable `OfficeAi.Shared/AttachmentText/` module (`DocumentFormat.OpenXml` 2.20.0). **No PDF, no images, no vision** — those return the path + type only. `olOLE` throws (rejected); `olByReference` has no data (rejected); `olEmbeddedItem` saves as `.msg`. |
| `list_folders` | Recursive walk of every store's `Folders`, mail folders only (`DefaultItemType == olMailItem`), with item + unread counts; capped ~800 / depth 8. |
| `search_contacts` | **EWS `ResolveName` over Contacts then GAL** (server-side ANR), run off the UI thread via EWS Managed API 2.2 with `UseDefaultCredentials`. Endpoint discovery: `Account.AutoDiscoverXml` → `AutodiscoverUrl` → process-static `Uri` cache. Each `NameResolution` mapped to `(name, email)` — GAL X500/`EX` addresses fall back to the resolved contact's own `EmailAddress1..3`; entries with no `@` address are dropped (mirrors `mcp-outlook`). Deduped by lowercased address (name fallback), capped at `limit` (default 10). Pure helpers `EwsAutodiscoverXml.ParseEwsUrl` + `ContactSearchFormat.Format` are unit-tested. **No `folder`/scope arg.** On-prem Exchange only; unreachable / auth failure / no endpoint / 15 s timeout → a specific `IsError` message, no COM fallback. (The pre-2026-09 recursive multi-store contact-folder crawl froze then crashed Outlook — removed.) |
| `list_events` | `Items.Sort("[Start]")` → `Items.IncludeRecurrences = true` → `Items.Restrict("[Start] <= end AND [End] >= start")` — **this order is load-bearing** and rules out `GetTable`. Recurring instances share the master `event_id`; each row carries its own `start` to disambiguate. Args: `start_date` (today), `end_date` (+7d), `mailbox` (optional — email address for a shared calendar; if given, resolves via `Ns.CreateRecipient(mailbox).Resolve()` and opens via `Ns.GetSharedDefaultFolder(recipient, olFolderCalendar)`; returns `IsError` if resolution or folder access fails), `limit` (50). When `mailbox` is given, each event's output gains a `calendar_owner: <name>` line (the **resolved recipient's own display name**, not the raw `mailbox` input string — fixed 2026-09-29 per PR #28 review, see dated update below) and a `store_id: <StoreID>` line — feed it to `get_event` or any of the calendar-editing tools' own `store_id` param (see dated update below) to act on that specific event on the shared calendar. Subject/location are always whatever Outlook itself resolves, private items included — this add-in does not redact or otherwise restrict shared-calendar content beyond what the caller's real Exchange permissions already govern (see dated update below for why an earlier private-item redaction was reverted). |
| `get_event` | `Body` (≤ 40k), `RequiredAttendees`/`OptionalAttendees`, organizer, response status, recurring flag. Optional `store_id` (added 2026-09-29 per PR #28 review) is passed through to `ItemById`/`GetItemFromID` so an `event_id` from someone else's shared calendar (via `list_events`' `mailbox` param) can actually be resolved — `ItemById` only searches the caller's own default store when no store hint is given, so without this a shared-calendar `event_id` from `list_events` likely couldn't be read back at all. Omit for the caller's own events, unchanged from before this parameter existed. Whether the resolved item can actually be read is governed entirely by the caller's real Outlook/Exchange permissions on that calendar — see dated update below. |
| `find_meeting_slots` | `Recipient.FreeBusy(anchor, 30, true)` — a per-30-min status string — for `Namespace.CurrentUser` + each resolved attendee; then `OfficeAi.Shared.MeetingSlots.Rank` (pure, unit-tested) slides a `duration_minutes` window in 30-min steps across each work day's `[start_hour, end_hour)` and scores each candidate by how many people are free (so a best partial match still comes back). **Work week/hours are read from the mailbox's own EWS `GetUserAvailability` → `AttendeeAvailability.WorkingHours` (added 2026-09-19; see `OutlookEws.GetWorkingHoursAsync`), not hardcoded** — falls back to Sun–Thu 09:00–18:00 only if that call fails (non-Exchange profile, EWS unreachable, etc.), cached per process like `OutlookEws.CachedUrl`. Default range is today through the end of the current contiguous work-day run (generalizes the old "today→Thursday, or next week if Fri/Sat" to any work-days shape), max 28 days. Times past the returned free/busy window are assumed free. Async (like `search_contacts`) only because of the EWS work-week lookup; the FreeBusy/ranking work itself is still synchronous COM. Args: `attendees` (req), `duration_minutes` (req), `start_date`, `end_date`, `start_hour`, `end_hour` (both default to the resolved work hours, or 9/18 as a last resort), `limit` (5). |
| `list_tasks` | `Folder.GetTable` over the default Tasks folder; open tasks only unless `include_completed`. Columns EntryID/Subject/Due/Start/Status/PercentComplete/Complete/ReminderTime. |
| `list_color_categories` | `Namespace.Categories` — the profile's master color-tag ("Category") list shared by mail/calendar/tasks, same list Outlook's Categorize picker shows. Each entry: `{name, color}`; color is one of the 26 `OlCategoryColor` values (None/Red/Orange/…/Dark Maroon), mapped to a friendly display name in `OutlookTools.Categories.cs` (not in `OfficeAi.Shared` — that project doesn't reference the Outlook PIA, same split as `ColorUtil`). |

> **Update 2026-09-29 (Outlook `list_events` gains `mailbox` parameter for shared calendars):**
> `list_events` now accepts an optional `mailbox` parameter (email address) to list
> events on a shared calendar instead of the default user's own calendar. When
> provided, `mailbox` is resolved via `Ns.CreateRecipient(mailbox).Resolve()` and
> the calendar folder is opened via `Ns.GetSharedDefaultFolder(recipient,
> olFolderCalendar)`. Resolution failures or folder access errors return an `IsError`
> naming the mailbox and the likely cause (not shared with the user, or needs
> adding via Outlook's own "Open Calendar" first). The identical `Sort`/
> `IncludeRecurrences`/`Restrict` query is reused unchanged for both the default
> calendar and shared calendars. When `mailbox` is given, each event's output gains
> a `calendar_owner: <mailbox>` line. The zero-results message becomes "No events on
> <mailbox>'s calendar between X and Y." when `mailbox` is given (vs. "No events
> between X and Y." when omitted).
>
> **Verification status: Exchange calendar-sharing permission tiers untested.** The
> code path `GetSharedDefaultFolder` should work identically across all permission
> levels (Full Access / Editor / Reviewer / etc.), but behavior has not yet been
> exercised live against a real second mailbox with different permission tiers
> configured. First live use against a shared calendar with limited permissions
> (free-busy only, titles+locations only, or no-access) should verify the exact
> error messages and whether partial-read tiers degrade gracefully or fail outright.

> **Update 2026-09-29 (PR #28 code review: cross-mailbox write refusal, `get_event`
> `store_id`, private-item redaction, resolved display name) — since reversed, see
> the next dated update below:**
> A review of the `list_events` `mailbox` feature above (PR #28) found that nothing
> stopped an `event_id` obtained from someone else's shared calendar from being fed
> into a write tool — `edit_event`, `draft_edit_event`, `cancel_event`,
> `draft_cancel_event`, and `accept_meeting`/`decline_meeting`/`tentative_meeting`'s
> shared `RespondMeeting`/`draft_respond_meeting` all resolved `event_id` via
> `ItemById` with no check on which mailbox store the resolved item actually lived
> in. This was fixed with a new shared helper, `RefuseIfNotOwnStore`, that refused
> any of those six tools on an item outside the caller's own default store, plus a
> redaction of `subject`/`location` to `"(private)"` for `Sensitivity == olPrivate`
> items on the shared-calendar path in `list_events`. **Both of these were reverted
> the same day — see the dated update immediately below for the project owner's
> explicit reasoning and the resulting design.** `get_event`'s `store_id` parameter
> (described in the `get_event` row above) was the one piece of this review kept
> as-is, and is now extended to the calendar-editing tools too.
>
> Separately, same review: `calendar_owner` and the zero-results message now use the
> resolved `Recipient`'s own `.Name` (falling back to the raw `mailbox` string only
> if that's empty) instead of echoing back the raw `mailbox` input verbatim — so
> `mailbox: "dana"` reports back who it actually resolved to, not the ambiguous
> string the caller typed. This part is unaffected by the reversal below.
>
> **Update 2026-09-29 (explicit project-owner reversal: `RefuseIfNotOwnStore` and
> private-item redaction removed; `store_id` added to all calendar-editing tools
> instead):** The project owner explicitly overrode the `RefuseIfNotOwnStore` check
> and the private-item redaction added by the PR #28 review immediately above, on
> this principle: this add-in should not layer its own authorization or
> content-redaction logic on top of Outlook/Exchange's own permission model.
> Whatever a caller's real Exchange sharing permissions would let them do through
> Outlook's own UI on a calendar they've been given access to — view an event
> (including a private one, if their access level exposes it), edit it, cancel it,
> or respond to it — these tools should allow too, with Outlook/Exchange itself
> (not this code) the only thing that can refuse. `RefuseIfNotOwnStore` and its six
> call sites (`EditEvent`, `DraftEditEvent`, `CancelEvent`, `DraftCancelEvent`,
> `RespondMeeting`, `DraftRespondMeeting`) were deleted outright from
> `OutlookAiAddIn/OutlookTools.Calendar.cs`; `QueryCalendarItems`' `isPrivate`
> redaction of `subject`/`location` was deleted too, so shared-calendar rows from
> `list_events` now show whatever Outlook itself resolves for a private item, same
> as any other.
>
> In their place, `edit_event`, `cancel_event`, `accept_meeting`, `decline_meeting`,
> `tentative_meeting`, `set_event_categories`, `set_event_availability`, and their
> `draft_` counterparts (8 tools total) all gained the same optional `store_id`
> parameter `get_event` already had, passed straight through to `ItemById`/
> `GetItemFromID` — purely a lookup aid for resolving an `event_id` outside the
> caller's own default store (as returned by `list_events`' `store_id` field for a
> shared-calendar event), with zero authorization logic attached. If the caller
> lacks real Exchange permission for the action, that now surfaces as whatever
> COMException `.Save()`/`.Send()`/etc. naturally throws, caught by each tool's
> existing generic exception handling (or the outer `ExecuteAsync` catch, for a
> path with no local try/catch) — not as a custom pre-check message.
>
> **Verification status: unverified.** This reversal has not been exercised live
> against a real second mailbox with restricted permissions — it is not yet known
> what error (if any) actually surfaces from Outlook when a caller genuinely lacks
> permission for a cross-mailbox write (e.g. attempting `edit_event` with a
> `store_id` from a calendar where the caller only has Reviewer access). It's
> assumed to come back as a COMException from `.Save()`/`.Send()`, per this design's
> own reasoning, but that assumption itself is untested, same as the underlying
> calendar-sharing permission-tier behavior called out in the block above.

> **Update 2026-09-30 (shared-calendar path rewritten onto EWS - was freezing Outlook):**
> The `GetSharedDefaultFolder` + `Items.Sort/IncludeRecurrences/Restrict/foreach` COM
> enumeration above froze Outlook - confirmed live by the project owner, even for a
> single-day range. Root cause: that folder is normally not cached offline the way the
> caller's own default calendar is, so per-property reads during enumeration could mean
> a live, blocking round trip to Exchange for every property of every event, all on
> Outlook's own UI thread (Outlook COM objects are STA-bound, unlike this add-in's EWS
> calls). `list_events`' shared-calendar path now queries EWS's `FindAppointments` +
> `CalendarView` instead (server-side date-range filtering and recurrence expansion,
> off the UI thread via `Task.Run` - the same pattern `search_contacts` already uses for
> the same reason). `GetSharedDefaultFolder` is still called exactly once per
> `list_events` call, but only to read `.Store.StoreID` for `get_event`'s `store_id`
> parameter - it never touches `.Items`. Output format is unchanged (`list_events`'
> shared-calendar path now produces the same per-event text shape via the new
> `SharedCalendarEventFormat.Format` that `QueryCalendarItems` already produces
> inline for the own-calendar path — the two paths intentionally use separate code,
> not a shared call, per the design doc's rationale). See
> `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md` for the
> full design, including the one assumption this fix rests on that still needs live
> confirmation (whether `GetSharedDefaultFolder` alone, independent of enumeration, was
> ever part of the freeze).

### Mutating tools (13 — Full autonomy only; `Mutated = true`)

| Tool | Notes |
|---|---|
| `mark_email_read` / `mark_email_unread` | `MailItem.UnRead` + `.Save()`. |
| `flag_email_important` | `Importance = olImportanceHigh/Normal` + `.Save()`. `important` defaults true. |
| `move_email` | `MailItem.Move(ResolveFolder(destination))`; returns `{message_id: <new EntryID>, old_message_id}`. |
| `delete_email` | Non-permanent (default) → `Move` to Deleted Items (returns new id) — Draft only, same risk class as `move_email`. `permanent: true` → then `.Delete()` from there (no single-call hard delete in the OM — documented as "may still be server-recoverable") — **Full autonomy required**, gated by an input-aware check in `ExecuteAsync` alongside the name-based tier tables (2026-09-27; see update below), since the tool name alone sits in `DraftTierTools`. |
| `accept_meeting` / `decline_meeting` / `tentative_meeting` (added 2026-09-29, refusal checks + wording fixed 2026-09-29 per PR #29 review) | Resolves to `AppointmentItem` via the shared `ResolveMeetingAppointment` helper (`MeetingItem.GetAssociatedAppointment(false)` when the id is a meeting request). Refuses (`IsError`) up front on an already-canceled meeting (`IsCanceledMeeting`) or an item that isn't a meeting the user was actually invited to (`!IsReceivedMeeting` — covers both an organizer's own `olMeeting` copy and a plain `olNonMeeting` appointment), via two new shared helpers (`AlreadyCanceledRespondError`/`NotInvitedError`) also reused by `draft_respond_meeting`. Otherwise `appt.Respond(olMeetingAccepted/Declined/Tentative, true, false)`, then `.Send()` on the response if non-null. One shared `RespondMeeting` helper backs all three (generalized from a `bool accept` parameter to the actual `OlMeetingResponse` value so `tentative_meeting` could reuse it). Optional `message` sets `resp.Body` before `.Send()` — a short comment attached to the response; **unverified live** whether the organizer actually sees this text on the delivered response, since that needs a real received invite to test against, not a self-organized item. If `.Send()` fails, the result now says so explicitly ("...but the response could not be sent to the organizer") instead of unconditionally claiming success — the local `Respond()` still went through, but the organizer was never notified. |
| `set_event_categories` | `AppointmentItem.Categories` (comma-separated tag names, the color shown on the event in the calendar grid) + `.Save()`; empty/omitted `categories` clears all tags. A name outside the master list is auto-added by Outlook on `Save` with an arbitrary color — call `set_category_color` first to control it. |
| `set_category_color` | `Namespace.Categories[name]` — updates `.Color` if the tag exists, else `Categories.Add(name, color)` creates it. Same master list `list_color_categories` reads. |
| `set_event_availability` | `AppointmentItem.BusyStatus` (the calendar's "Show As" dropdown — Free/Tentative/Busy/Out of Office/Working Elsewhere) + `.Save()`. Case-insensitive friendly-name lookup; an unrecognized value returns `IsError` listing the valid names instead of throwing. Draft tier, not Full autonomy — same rationale as `set_event_categories`/`set_category_color` above (purely local, never `.Send()`); the "Full autonomy only" in this table's own heading predates that tiering and is already stale for those two rows. Added 2026-09-27, **verified against a live Outlook client 2026-09-28** (including on an organized meeting with attendees — no notification triggered). |
| `create_task` | `Application.CreateItem(olTaskItem)` + `.Save()` — no window (a task doesn't send anything, so it follows the mutate-directly pattern, not draft-and-display). Args: `subject` (req), `body`, `due_date`, `start_date`, `reminder_time`, `importance`. |
| `update_task` | `(TaskItem)GetItemFromID`; only passed fields change; `mark_complete: true` → `Complete = true` + `PercentComplete = 100`. |
| `set_reminder` | `ReminderSet` / `ReminderTime` on an appointment **or** task, addressed by its `item_id` (EntryID); `clear: true` turns it off. |
| `set_email_reminder` | `MailItem.MarkAsTask(mapped interval)` + `TaskStartDate`/`TaskDueDate` + `ReminderSet`/`ReminderTime` + `.Save()` — the confirmed COM path for "flag an email for follow-up with a reminder". `MailItem` does expose `ReminderSet`/`ReminderTime`. |

> This table's "Full autonomy only" heading predates the four-tier rework
> below ("Editing modes — four tiers") — every row above is actually
> reachable from Draft only upward (`DraftTierTools`), not gated to Full
> autonomy. Not rewritten here; see that section for the real gate.

### Undo/redo tools (2 — Draft tier or higher; `Mutated` on success)

> **Added 2026-09-27, reworked 2026-09-28** into a custom undo/redo stack of
> the assistant's own actions (`OutlookTools.Undo.cs` +
> `OfficeAi.Shared/ActionHistory.cs`). Outlook's native Undo can't be used —
> see the dated Update block near the top of this document for what was
> tested and the full list of what each tool records.

| Tool | Notes |
|---|---|
| `undo_last_action` | Reverses the assistant's most recent recorded action; call it again to step further back. Covers `mark_email_read/unread`, `flag_email_important`, `move_email`, non-permanent `delete_email`, `create_task`, `update_task`, `set_reminder`, `set_email_reminder`, `set_event_categories`, `set_category_color`, `set_event_availability`, `edit_event` (added 2026-09-28, replaces `reschedule_event`) on a plain (non-meeting) event that is not becoming a meeting in that call - a non-recurring event, a single occurrence via `occurrence_date`, or a whole recurring series' `subject`/`body`/`location` only (not a time change) — same `SnapshotEntry` mechanism as before, now against a dynamically-built props list (`Subject`/`Body`/`Location` in addition to `Start`/`End`, whichever fields the call actually touched) instead of the old fixed `{Start,End}` pair, since a saved occurrence has its own real `EntryID` — `cancel_event` on a plain or already-canceled appointment, and `create_event` without attendees. Sends, invites, meeting responses, `edit_event`/`cancel_event` on a still-active organized meeting (attendees notified) — including `edit_event` converting a plain event into a meeting for the first time by adding `required_attendees`/`optional_attendees` (unverified at runtime, see "Unproven at runtime" below) — permanent deletes, any occurrence-level cancellation via `cancel_event`'s `occurrence_date` (plain appointment or meeting alike — `RecurrencePattern.Exceptions`/`Exception` are read-only via COM, so a deleted occurrence can't be reversed), and **`edit_event` on the whole series whenever it includes a time change and `event_id` is a genuinely recurring master** (Outlook rejects setting `Start`/`End` directly on a recurring master, so this path writes `RecurrencePattern` fields instead, which undo can't yet target; a non-recurring event's whole-item edit is unaffected and stays undo-able, and if other fields are changed in the same call as a whole-series time change, the entire call stays a barrier — no partial-undo of a mixed pattern-field + item-property change) are barriers: undo reports it can't go past them. It refuses if the item was changed since the assistant's action, and never touches the user's own manual changes. **Not verified against a live Outlook client**, except `set_event_availability`'s undo/redo round trip (verified 2026-09-28) and, for the recurring-series feature, occurrence-level reschedule/cancellation against a plain (non-meeting) recurring series (verified 2026-09-28 against the retired `reschedule_event`/`cancel_event`, mechanism carried over unchanged onto `edit_event` — see the dated update near the top of this document; the meeting-occurrence case, and the whole-series-recurring `RecurrencePattern` time-change fix, are both unverified live, as are all of `edit_event`'s new `subject`/`body`/`location`/attendee-editing surface — see the 2026-09-28 `edit_event` update below). |
| `redo_last_action` | Re-applies the most recently undone action, with the same "changed since" check. The redo list is cleared by any new recorded action or barrier. |

### Draft-and-display tools (8 — Draft only or higher; open a native Outlook window for the user to review and send; `Mutated = false`)

| Tool | Notes |
|---|---|
| `draft_email` | `CreateItem(olMailItem)` → set `To`/`Subject`/`Body` → `Display(false)`. `"— Created with OpenDocs"` appended to a non-empty body (mirrors mcp-outlook's signature seed). |
| `reply_email` | `orig.Reply()` (sender only); `body` HTML-encoded and prepended above the quoted original; `Display(false)`. |
| `reply_all_email` | `orig.ReplyAll()`. A **distinct tool**, not a `reply_all` boolean on `reply_email` — clearer for the model, and the user sees the full recipient list before sending. |
| `forward_email` | `orig.Forward()`, optional `To`, `body` prepended; `Display(false)`. |
| `draft_event` | `CreateItem(olAppointmentItem)`; with `required_attendees`/`optional_attendees` → `MeetingStatus = olMeeting`, `Recipients.Add(...).Type`, `Recipients.ResolveAll()`; `Display(false)`. Same `"— Created with OpenDocs"` signature appended to a non-empty body. Optional `recurrence` object (added 2026-09-28), same `RecurrenceValidator`-validated shape and `GetRecurrencePattern()` application as `create_event` — see the dated update above. |
| `draft_edit_event` (added 2026-09-28, replaces `draft_reschedule_event`) | Resolves `event_id`/`occurrence_date` via the shared `ResolveOccurrenceTarget`, refusing (`IsError`) up front on an already-canceled event or one the user only attends (`IsCanceledMeeting`/`IsReceivedMeeting`, checked against the master — same shared helpers `edit_event` uses). Requires at least one of `start`+`end` (together), `subject`, `body`, `location`, `required_attendees`, `optional_attendees` — a clean `IsError` naming all seven if none are given; `required_attendees`/`optional_attendees` reject `occurrence_date` together (attendee edits are whole-series/non-recurring only). Applies every change unsaved: `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` for a whole-series-recurring time change, else `appt.Start`/`.End` directly (occurrence or non-recurring — same proactive `CheckOccurrenceReorderCollision` check `edit_event` runs before an occurrence reorder); `Subject`/`Body`/`Location` set directly; attendees replaced via the new `ReplaceAttendees` helper (per-category, not a blanket clear — see `edit_event` below for the exact rule), flipping `MeetingStatus` to `olMeeting` if the event wasn't already a meeting. Ends in `appt.Display(false)` for the user to review — **never** touches `ForceUpdateToAllAttendees`, never calls `.Send()`/`.Save()`, and never records undo/redo, same contract as `draft_event`. |
| `draft_cancel_event` (added 2026-09-28) | Resolves `event_id`, `Display(false)`s it **unchanged** either way — never touches `MeetingStatus` or anything else (an earlier version set `MeetingStatus = olMeetingCanceled` unsaved first; removed same day after live testing showed the change persists on window-close without an explicit Send, silently canceling with attendees never notified). The user cancels it themselves via Outlook's own UI from the opened window. Refuses (`IsError`) on an already-canceled event (points at `cancel_event` for that) or a still-active meeting the user only attends (points at `draft_respond_meeting`, the Draft-tier-reachable alternative — or `decline_meeting`, one tier up — fixed 2026-09-29 per PR #29 review: it previously pointed only at `decline_meeting`, a tool a Draft-only caller couldn't actually reach) — see the "Update 2026-09-28" note above. Optional `occurrence_date` (same day) targets one specific occurrence via `ResolveOccurrenceTarget` before displaying it. |
| `draft_respond_meeting` (added 2026-09-29, replacing `draft_accept_meeting`/`draft_decline_meeting`/`draft_tentative_meeting` the same day per a PR #29 code review) | **Never calls `Respond()`.** Resolves `event_id` via the shared `ResolveMeetingAppointment` helper, refuses (`IsError`) on an already-canceled meeting or one the user wasn't actually invited to (same `AlreadyCanceledRespondError`/`NotInvitedError` helpers `RespondMeeting` uses), then just `appt.Display(false)`s the original item **completely unchanged** — the user picks Accept/Tentative/Decline themselves from Outlook's own native ribbon buttons. Optional `message` is returned as suggested text in the tool's own output ("paste it in if you use Outlook's 'Edit response before sending' option") rather than pre-filled into a response body, since pre-filling would require calling `Respond()` to obtain the `MeetingItem` — the exact call this redesign avoids. Replaces what were three separate tools because the redesigned version no longer picks a specific `OlMeetingResponse` up front, so one unified tool is simpler and matches the design exactly. Never records undo/redo — same contract as `draft_event`/`draft_edit_event`/`draft_cancel_event`, since nothing is saved or sent until the user acts on the opened window. **Why the redesign:** the prior three-tool version called `appt.Respond(response, true, false)` *before* ever displaying anything, exactly like `RespondMeeting`'s immediate-response path — and `Respond()` is documented/known to commit a real calendar change at call time (a new EntryID on accept/tentative, a move to Deleted Items on decline), independent of whether the resulting response is ever sent or the window ever closed with an action taken. A PR #29 code review flagged this as Critical: it broke this codebase's "draft tools persist nothing until the user acts" guarantee, the same shape of bug `draft_cancel_event` itself hit and fixed the same way (see the "Update 2026-09-28" `draft_cancel_event` note above) — never mutate the item at all. |

**These never call `.Send()` (mail) or save a calendar event.** The user sends from the
opened Outlook window. Available from Draft only mode upward (tier 2 — see "Editing
modes" above), not gated behind Full autonomy.

### Auto-send tools (7 — Full autonomy only; send/create immediately, no review window; `Mutated = true`)

> **Added 2026-09-19**, reversing part of the "no auto-send tool" decision below —
> narrowed to Full autonomy specifically, not removed generally. Every draft/compose
> tool above is unaffected and stays draft-and-display-only in every mode.

| Tool | Notes |
|---|---|
| `send_email` | Same construction as `draft_email`, but `m.Send()` instead of `m.Display(false)`. |
| `send_reply` | Same as `reply_email`, but `.Send()`. |
| `send_reply_all` | Same as `reply_all_email`, but `.Send()`. |
| `send_forward` | Same as `forward_email`, but `.Send()`; `to` is required (unlike `forward_email`, where it's optional). |
| `create_event` | Same construction as `draft_event`. No attendees → `a.Save()` (a plain calendar entry, nobody to notify). Attendees present → `MeetingStatus = olMeeting` then `a.Send()`, dispatching the invite. Both `AppointmentItem.Send()`/`.Save()` confirmed present via .NET reflection against the referenced PIA before writing this — not assumed from `draft_event`'s non-sending shape. Optional `recurrence` object (added 2026-09-28) builds a repeating series via `GetRecurrencePattern()`, applied after attendee/`MeetingStatus` setup and validated first by the pure `OfficeAi.Shared.RecurrenceValidator` — see the dated update above. |
| `edit_event` (added 2026-09-28, replaces `reschedule_event`) | General-purpose event edit, not just a reschedule — Full autonomy only. Resolves `event_id`/`occurrence_date` via `ResolveOccurrenceTarget`, with the same `IsCanceledMeeting`/`IsReceivedMeeting` organizer-authority refusals `reschedule_event` had. Requires at least one of `start`+`end` (together, a clean `IsError` naming all seven optional fields otherwise), `subject`, `body`, `location`, `required_attendees`, `optional_attendees`; attendee fields reject `occurrence_date` together (attendee edits are whole-series/non-recurring only, per the design). Applies `start`/`end` via `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` for a whole-series-recurring time change, else directly on `appt.Start`/`.End` (occurrence or non-recurring — same proactive `CheckOccurrenceReorderCollision` check as before); `subject`/`body`/`location` set directly; attendees replaced via the new `ReplaceAttendees` helper — per-category, not a blanket clear: it clears and replaces only the required-attendee category if `required_attendees` is given, and/or the optional-attendee category if `optional_attendees` is given (omitting one leaves that category's existing attendees untouched; the organizer recipient is never touched), re-adding via the existing `AddAttendees` helper then `Recipients.ResolveAll()` — flipping `MeetingStatus` to `olMeeting` if the event wasn't already a meeting — converting a plain event into a meeting this way is supported, per the design, and unverified at runtime (see "Unproven at runtime" below). `.Send()`s if the result is or becomes a meeting — setting `ForceUpdateToAllAttendees = false` explicitly first (confirmed to already be Outlook's own default) — else `.Save()`s. Barrier (`RecordIrreversible`) whenever it sends or changes a whole recurring series' time via `RecurrencePattern`; otherwise `RecordSnapshot`s a dynamically-built props list — only the fields actually touched that call — replacing the old fixed `{Start,End}` `RescheduleProps` constant. Same false-negative `.Save()` risk `reschedule_event` had, now in `edit_event`'s occurrence path (re-verifies via `ResolveOccurrenceTarget` on a `COMException` before reporting failure) and the same proactive occurrence-reorder collision check. |
| `cancel_event` (added 2026-09-28) | Resolves `event_id`. Still-active organized meeting → `MeetingStatus = olMeetingCanceled` then `.Send()` (dispatches the cancellation notice), then moves the item to Deleted Items (best-effort — see the "Update 2026-09-28" note above for the unverified `.Send()`-then-`.Move()` sequence). Plain appointment, or an **already-canceled event** (`olMeetingCanceled`/`olMeetingReceivedAndCanceled` — the only way to dismiss one) → moves straight to Deleted Items, nobody to notify. Refuses (`IsError`) only on a still-active `olMeetingReceived` (points at `decline_meeting`, not "Propose New Time" — canceling isn't something an attendee can request at all). Optional `occurrence_date` (added 2026-09-28) targets one specific occurrence the same way as `edit_event`; occurrence-level cancellation is **always** an undo barrier, never undo-able, regardless of plain-appointment vs. meeting — see the dated update above and "Undo/redo tools" below. |

Gated by `OutlookTools.cs`'s `SendTierTools` set, requiring `EditingMode.FullAutonomy`
exactly (the ordinal check's top tier) — not reachable from Draft only or Automate
approvals. `accept_meeting`/`decline_meeting`/`tentative_meeting` are **not** in this table; they live one
tier down, in Automate approvals (see "Editing modes" above), since they already
call `resp.Send()` themselves — accept/decline before this addition, `tentative_meeting`
(added 2026-09-29) grouped alongside them for the same reason.

### Excluded / deferred

- **`update_event` / `delete_event`** — present in mcp-outlook's `server.py` but not
  its README; not ported. Trivial parity adds if wanted.
- **Proposing a new time on a meeting the user only attends** (`MeetingStatus ==
  olMeetingReceived`) — real Outlook's "Propose New Time" feature, deliberately
  *not* implemented (added 2026-09-27 alongside the now-retired `reschedule_event`/
  `draft_reschedule_event`, carried forward unchanged into `edit_event`/
  `draft_edit_event` on 2026-09-28). Those tools refuse outright on a received meeting
  rather than attempt anything, because there's no clean way to honor the request:
  a plain `.Send()` here wouldn't be an authoritative reschedule Outlook actually
  honors (the user isn't the organizer), and .NET reflection against the
  actually-referenced `Microsoft.Office.Interop.Outlook` 15.0.0.0 PIA found no
  "Propose"/"Counter"/"NewTime"-named member on `AppointmentItem`/
  `_AppointmentItem` or `MeetingItem`/`_MeetingItem` to call instead —
  `_AppointmentItem.Respond` only accepts `OlMeetingResponse`
  (Accept/Decline/Tentative). "Propose New Time" appears to be a ribbon/UI-level
  feature (MAPI counter-proposal properties) not cleanly exposed through classic
  COM automation. A future implementation would likely need raw `PropertyAccessor`
  MAPI-property manipulation rather than a typed object-model call — untried here,
  and risky without a live mailbox to validate against.

(`find_meeting_slots`, deferred in the first cut, is now implemented — see the Read
tools table. `Recipient.FreeBusy` is all local-time, so no cross-timezone math was
needed; the ranking is a pure, unit-tested helper.)

### Object Model Guard

In-process VSTO add-ins that use the **VSTO-supplied `Application`** object are trusted
by default — reading `Body` / `Recipients` / `SenderEmailAddress` /
`AddressEntry.PrimarySmtpAddress`, calling `PropertyAccessor.GetProperty` or
`Attachment.SaveAsFile` do **not** raise the "a program is trying to access…" prompt on
default settings. Prompts appear only under Trust Center → Programmatic Access set to
"Always warn", or an Exchange public-folder security form. `MailItem.Send` — the
highest-risk call — is never used.

### Structural fragility

Everything is addressed by `EntryID`. It is stable while an item stays put but changes
on `Move` and is store-specific; the mutating tools that move items return the new id,
and every other tool re-resolves via `GetItemFromID` each call. All recurring
occurrences of a calendar series **share one `EntryID`**, so `get_event` /
`accept_meeting` / `decline_meeting` / `tentative_meeting` still cannot target a single occurrence
unambiguously — they necessarily act on the master series. `list_events` carries each
occurrence's `start` as the disambiguator.

**Corrected 2026-09-28:** this section previously claimed `reschedule_event` /
`draft_reschedule_event` / `cancel_event` / `draft_cancel_event` were bound by the same
limitation, documenting it as though it were a fundamental COM constraint. That was
wrong: those four tools (the first two since fully replaced, same day, by `edit_event`/
`draft_edit_event` — see the dated update near the top of this document) took an
optional `occurrence_date`, resolved via a new `ResolveOccurrenceTarget` helper against
`RecurrencePattern.GetOccurrence(DateTime)` — a real, callable member, confirmed via
.NET reflection against the referenced PIA rather than assumed absent — so they can
target one specific occurrence precisely. The shared-`EntryID` limitation still applies
to `edit_event`/`cancel_event`/`draft_edit_event`/`draft_cancel_event` only when
`occurrence_date` is omitted (they then act on the whole series, exactly as before this
feature). See the "Update 2026-09-28 (Outlook gains recurring-series support...)" block
near the top of this document for full detail, including a since-discovered undo/redo
asymmetry between occurrence-level edit (undo-able) and occurrence-level cancellation
(always a barrier).

**Known limitation, confirmed live 2026-09-28 (Gmail-connected calendar):** on a
mailbox connected via Google's Gmail/Google Workspace sync, moving a calendar item
to Deleted Items does not appear to be durable the way it is on Exchange — a
`cancel_event` call that reported success moving an item to Deleted Items was
followed immediately (no other action in between) by that item relocating itself
to a `Drafts` folder, with no code in this add-in touching it a second time. Most
likely Google Calendar's own sync reconciling the move shortly after, outside this
add-in's control. `undo_last_action`'s own "did the item change since?" conflict
check caught the mismatch and refused rather than guessing or overwriting — the
system's designed safety net worked correctly, and no data was lost — but this
means `RecordMove`'s core assumption (an item stays wherever the last recorded
move put it, until this add-in moves it again) does not reliably hold for
Gmail-connected calendars specifically. Not something to build a targeted
workaround for without more data — surfacing this as a known account-type-specific
risk rather than a code bug in `cancel_event`, `edit_event`, or the undo
stack.

### Unproven at runtime (as of 2026-08-28)

The project builds clean (MSBuild Debug + Release, 0 warnings; `dotnet test` 136 pass
including the DASL and meeting-slot helpers). But **most COM paths have not been
exercised in a real Outlook** — the Explorer pane lifecycle / IUnknown-identity
keying, WebView2 rendering inside an Explorer task pane, `Folder.GetTable` column
names, the `list_events` recurrence ordering, `Recipient.FreeBusy`'s string format /
month coverage (`find_meeting_slots`), and several enum-name / method-signature
assumptions (`olEmbeddeditem` casing, `AppointmentItem.Respond` argument types,
`MailItem.MarkAsTask`) compiled against the interop assembly but are not yet confirmed
live. Early manual testing via the mock server has exercised `list_emails`,
`search_emails`, `list_folders`, `list_events`, and `list_tasks` against a real
mailbox successfully.

`search_contacts`'s EWS path (added 2026-09) is likewise **compiled but not confirmed
against a live on-prem Exchange**: the `ResolveName` call and its `NameResolution`
mapping, `ExchangeService.UseDefaultCredentials` (Windows Integrated Auth to on-prem
CAS), `Account.AutoDiscoverXml` shape / `EwsAutodiscoverXml.ParseEwsUrl`, the
`AutodiscoverUrl` fallback, and the off-thread `Task.Run` actually keeping Outlook
responsive during the call. The async delegate refactor (`ToolExecutor` →
`Task<ToolResult>`, `async void OnWebMessageReceived`) builds clean for all four
add-ins; a runtime smoke of one tool per app confirms nothing regressed is still
pending.

`edit_event`/`draft_edit_event` (added 2026-09-28, replacing `reschedule_event`/
`draft_reschedule_event` added 2026-09-27) are likewise **compiled but not exercised
against a live Outlook**: no environment with a real mailbox was available while
writing either version. The `MeetingStatus` branching
(`olNonMeeting`/`olMeeting`/`olMeetingReceived`) and the `AppointmentItem.Send()`/
`.Save()` calls reuse `CreateEvent`'s already-referenced members, so the original
risk was concentrated in the `Start`/`End` assignment on an item resolved via
`GetItemFromID` (not freshly created, unlike `draft_event`/`create_event`) and in
whether Outlook accepts a plain `.Send()` on a modified organizer-owned meeting as
a real update notice (vs., say, needing `ClearRecipients`/re-resolution first for
edge cases like an all-day flag or a recurrence-pattern change). `edit_event`
carries that same unverified `Start`/`End`/`.Send()` risk forward, plus three new
ones introduced by its wider field surface — see the dated `edit_event` update
below for the full, ranked list: `ReplaceAttendees`'s recipient-clearing (highest
risk, since it feeds directly into a `.Send()` with real people), converting a
plain event into a meeting via `edit_event`, and `subject`/`body`/`location` edits
on a recurring master. The one thing that *was* checked concretely, not assumed:
the "no Propose New Time member" claim behind the `olMeetingReceived` refusal,
confirmed via .NET reflection against the referenced PIA (see the "Update
2026-09-27" note above).

**Open question raised by the `draft_cancel_event` finding below, not yet checked:**
whether `draft_edit_event`'s unsaved `Start`/`End` assignment (same shape —
mutate an existing item, `Display(false)`, never `.Save()`/`.Send()`) has the same
"persists on window-close without an explicit action" risk that `MeetingStatus`
turned out to have. If so, closing a `draft_edit_event` window on a meeting
the user organizes without clicking Send could silently move it in the organizer's
own calendar while attendees still see the old time — the reschedule equivalent of
the bug fixed below. Not confirmed either way; flagging for whoever next has a live
Outlook session, since this risk predates `edit_event` (it was already open for
`draft_reschedule_event`, unchanged by the 2026-09-28 rename/generalization).

`cancel_event`/`draft_cancel_event` (added 2026-09-28) are likewise **compiled but
not exercised against a live Outlook**, and carry one risk beyond what
`edit_event` already flags: `cancel_event`'s organized-meeting branch calls
`.Move()` immediately after `.Send()` on the same item, a sequence no existing tool
here performs (every prior `.Send()` call only reads properties off the item
afterward). Whether Outlook still permits relocating a just-canceled-and-sent
appointment is unconfirmed; the code catches a failure there and reports the
(genuinely-sent) cancellation succeeded regardless, telling the user to delete the
stray calendar entry manually if the move didn't take — see the "Update 2026-09-28"
note above.

The recurring-series feature (occurrence targeting + recurrence creation, added
2026-09-28 — see the dated update near the top of this document) is only **partly**
live-verified. What live testing did confirm, against a plain (non-meeting) recurring
series: occurrence resolution via `ResolveOccurrenceTarget`/`RecurrencePattern.GetOccurrence`,
occurrence-level reschedule (including its `SnapshotEntry`-based undo), occurrence-level
cancellation (as an always-barrier action), the false-negative `.Save()`/`.Delete()`
exception-vs-actual-outcome fix, and the two specific occurrence-reordering rejections
`CheckOccurrenceReorderCollision` guards against (reordering past a later occurrence;
two occurrences sharing a calendar day — both first reproduced manually in Outlook's
own UI). What remains unconfirmed:

- **The master-vs-occurrence organizer-authority fix (point 4 in the dated update) has
  never been exercised against a genuine recurring *meeting* occurrence with
  attendees** — every live test run used a plain, non-meeting series. Checking
  `master.MeetingStatus`/`IsCanceledMeeting(master)`/`IsReceivedMeeting(master)` instead
  of the occurrence's own value is a reasoned fix based on `GetOccurrence()`'s
  known-unreliable behavior for this property — prompted by a live failure whose
  actual cause later turned out to be unrelated (see point 5 in the dated update) —
  not a fix confirmed by live-observing the bug on a genuine meeting occurrence, so
  the fix itself remains unverified against the scenario it targets.
- **PR #21's Finding #2 (whole-series `Start`/`End` reassignment on a recurring
  master) is resolved — confirmed live 2026-09-28, and it was worse than
  "unverified": setting `AppointmentItem.Start`/`.End` directly on a genuinely
  recurring master throws `COMException 0xAF620009 "The object does not support
  this method."` from `set_Start`, unconditionally. Outlook simply does not allow
  it. Fixed in `EditEvent` (formerly `RescheduleEvent`) by writing to
  `RecurrencePattern.PatternStartDate`/`StartTime`/`EndTime` instead (the same
  mechanism `create_event`'s recurrence support already uses for a brand-new series
  — see `ApplyRecurrence` in `OutlookTools.Compose.cs`) whenever the whole series
  (not a single occurrence) is being time-shifted. This is undocumented-but-native
  Outlook behavior, not a workaround. Consequence: a whole-series time change on a
  genuinely recurring master is now a **barrier**, not undo-able (for both the
  meeting and non-meeting case) — the existing `SnapshotEntry` undo mechanism only
  reads/writes plain item properties (exactly what just failed), and there's no
  `RecurrencePattern`-aware undo entry type yet. This is a deliberate asymmetry
  from every other `edit_event` branch, which stay undo-able where the underlying
  mutation actually works. Occurrence-level time changes and non-recurring plain
  events are unaffected.
  **Still unverified live** (this specific fix, not the underlying limitation it
  works around): whether `PatternStartDate`/`StartTime`/`EndTime` interact
  correctly with the series' existing `DayOfWeekMask`/other pattern fields when
  only the time (not the day) changes, whether calling `.Send()`/`.Save()` on
  `appt` (which equals `master` here) after mutating the pattern behaves
  identically to mutating `appt` directly, and what redo/undo-refusal messaging
  actually looks like for this barrier in practice.
- **`CheckOccurrenceReorderCollision` may not cover every reason Outlook can reject an
  occurrence reorder.** It's confirmed live for the two specific rejection reasons
  reproduced by hand in Outlook's own UI (see above), but every such rejection surfaces
  from automation as the same generic `COMException "Cannot save this item."` with no
  `InnerException` detail — so a rejection reason this check doesn't anticipate would
  still reach the model as that same unhelpful message rather than the check's
  exact-range error.
- **`create_event`/`draft_event`'s `recurrence` param**: `RecurrenceValidator`'s 23 unit
  tests cover only the pure validation logic. The COM application path —
  `AppointmentItem.GetRecurrencePattern()` called after attendee/`MeetingStatus` setup,
  before `.Save()`/`.Send()` — is unverified against live Outlook; whether that ordering
  (rather than calling `GetRecurrencePattern()` before attendee/`MeetingStatus` setup)
  actually matters is unconfirmed either way.
- **`accept_meeting`/`decline_meeting`/`tentative_meeting`'s `message` parameter
  (added 2026-09-29)** is likewise **compiled but not exercised against a live
  Outlook**: `resp.Body = message` is set before `.Send()`, but whether the
  organizer actually sees that text on the delivered response — as opposed to it being
  silently dropped or overwritten by Outlook's own response-body template — has not
  been confirmed against a real received invite (a self-organized item can't exercise
  this path meaningfully). Still open as of the 2026-09-29 PR #29 redesign below — that
  update didn't touch this mechanism or add any new evidence toward confirming it.
  (`draft_respond_meeting`'s `message` is not subject to this gap at all: since it's
  returned as plain suggested text in the tool's own output rather than set as
  `resp.Body`, there's nothing Outlook-side to verify.)
- **RESOLVED 2026-09-29 (was: `draft_accept_meeting`/`draft_decline_meeting`/
  `draft_tentative_meeting`'s "nothing persists until the user acts" claim rested on
  an unverified assumption about `AppointmentItem.Respond()` itself).** A PR #29 code
  review treated this gap as confirmed-plausible enough to be a Critical finding
  rather than leave it as an open question — `Respond()` is documented/known to
  commit a real calendar change at call time (a new EntryID on accept/tentative, a
  move to Deleted Items on decline) independent of whether `.Send()`/`.Display()` is
  subsequently called, which would have silently violated the draft contract. Fixed
  by redesign, not by verification: the replacement `draft_respond_meeting` never
  calls `Respond()` at all, so this class of risk no longer applies to the draft
  tool — see the "Update 2026-09-29 (PR #29 code review...)" block above. The
  underlying question (does `Respond()` itself commit changes at call time) remains
  formally unconfirmed, but it's now moot for the draft path specifically, since
  nothing in `draft_respond_meeting` ever calls it.

---

## Explicitly out of scope everywhere (per project scope, not gaps)

> This scope boundary originated in the project's original feasibility report and
> the toolset-port plan's Global Constraints. It was previously stated only in
> `docs/tool-surface-todo.md`'s header before that file was retired (see PP-8,
> `docs/superpowers/plans/2026-08-23-pp08-retire-stale-todo.md`); this is now its
> canonical location.

- `web_search`, `image_search`, `generate_image`, `analyze_media` — no `ai-search`
  equivalent; air-gapped deployment target. (`read_attachment` is a partial exception
  for Outlook — `get_attachment` reads local + OpenXML attachment text, nothing
  remote, no PDF/images; see the Outlook section.)
- The PDF app and the Markdown app have no officeoffice counterpart (Markdown's
  scope is folded into Word).
- PowerPoint's `execute_slide_script` DSL and entire deck-generation pipeline (see
  above: `ask_clarification`, `plan_deck`, `generate_deck`, `regenerate_slide`,
  `save_style_template`, `list_style_templates`).

**Inconsistency flagged, not resolved here:** the retired `tool-surface-todo.md`
bundled `delete_slide` into this same out-of-scope list, alongside the DSL and
generation pipeline. This document's own "Missing entirely" section above already
treats `delete_slide` separately, and
`docs/superpowers/plans/2026-08-23-pp19-powerpoint-scope-and-delete-slide.md` argues
explicitly that `delete_slide` is a small, clearly-in-scope fix with no dependency on
the DSL (`add_slide` already exists; deleting a slide needs no scripting language).
That plan's Task 1 ships `delete_slide` independently of its Task 2 decision gate on
the larger DSL/generation question. Treat `delete_slide` as in-scope going forward;
this out-of-scope list covers only the DSL, the generation pipeline, and the QC pass.

---

## Summary: what genoffice has that officeoffice doesn't

1. **Live multi-provider selection** — the abstraction is copied in, but no add-in
   actually wires user-selected provider/model/API key to the transport yet; all
   three hardcode a local OpenAI-compatible test endpoint.
2. **Web-sourced content** — no search, no AI image generation, no media analysis,
   no chat-attachment reading, anywhere. Image tools are local-file-only by design.
3. **PowerPoint's scripting DSL and generation pipeline** — no `execute_slide_script`,
   no `generate_deck`/`regenerate_slide`, no automatic QC/audit pass, no `delete_slide`.
4. **Richer per-op parameter sets** in several places (status as of the items below —
   several since fixed by their own PP item, noted inline; this list otherwise
   reflects the original audit and is not re-verified wholesale here): Excel's
   `format_range` (missing ~7 of ~11 style properties genoffice supports), Excel's
   `add_chart` on creation (missing chart-type breadth `edit_chart` has),
   ~~PowerPoint's `set_element_style` (missing underline/align/family)~~ and
   ~~`add_shape` (3 types vs a full preset-geometry set)~~ — **both fixed, PP-20**
   (see the PowerPoint tool table above), Word's `insert_content`
   (plain-text-append-only, no positioning or rich content), `read_blocks`/
   `replace_blocks` (plain text only, no HTML), `updateTextStyle` (missing
   `highlight`), and `edit_chart` (single-series, no categories).
5. **Word has no image-insertion tool at all** (genoffice's docs app does).
6. **`dataSource`/provenance enforcement** on chart and data-bearing content — genoffice
   gates this at the tool layer for slides; officeoffice has no equivalent anywhere.
7. **Block-indexed document context** — genoffice's `get_document_context`/
   `get_deck_context` return structured per-block/per-element inventories; Word's and
   PowerPoint's officeoffice equivalents return flat text previews only.

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
   Autonomy) as a first-class, uniformly-applied gate across all three apps' tool
   dispatch — genoffice's docs app has Track-Changes-aware writes but no equivalently
   formal, uniform mode-gating system across its apps.
5. **Live selection-push into context** (Word) — `WindowSelectionChange` pushes the
   user's current selection text into `buildContext()` automatically on every turn,
   driven by a real Office event rather than app-side selection-range plumbing.
6. **`add_pivot` with calculated fields** — Excel's pivot op supports
   `PivotTable.CalculatedFields().Add(name, formula)` for formula-derived pivot
   values, in addition to row/column/page/data fields.

---

## Schema-vs-implementation audit

> **Update 2026-08-24 (PP-5 landed):** the structural root cause this whole section
> points at — both gateway tools' `commands`/`operations` items being a bare
> `{type:'object'}` with the entire per-kind contract living only in prose — is fixed.
> `apply_commands` and `propose_operations` now carry real per-kind JSON Schema
> (`WORD_COMMAND_SCHEMAS` in `WordAiAddIn/web-src/entry.ts`; `EXCEL_OPS` +
> `opSchemas`/`opsDescription` in `ExcelAiAddIn/web-src/entry.ts`), and `kind` parsing
> in both `ApplyCommands`/`ProposeOperations` moved inside the per-command try/catch
> with a required-field precheck (`WordTools.cs`'s and `ExcelTools.cs`'s
> `RequiredFields`/`ValidateRequired`) — so a malformed command now fails only itself,
> with a specific error naming the missing field, instead of aborting the whole batch
> (Word finding #1 above) or silently reaching a COM handler.
>
> **Cross-checked exhaustively, not sampled:** every `case` in `WordTools.cs`'s
> `ApplyCommands` switch (12) has exactly one matching entry in `WORD_COMMAND_SCHEMAS`,
> and vice versa. Same for `ExcelTools.cs`'s `ProposeOperations` switch (51) against
> `EXCEL_OPS`. Both diffs are empty. Excel's schema uses a **grouped variant**
> (`DETAILED_KINDS` in `entry.ts`) rather than full `oneOf` detail on all 51 kinds — a
> 51-branch schema measured ~4,870 added tokens (cl100k_base, before/after
> `JSON.stringify(ALL_TOOLS)`), over the ~4k budget PP-5 set as the threshold; only the
> 7 highest-ambiguity kinds (`format_range`, `add_conditional_format`, `add_chart`,
> `edit_chart`, `add_shape`, `set_data_validation`, `add_pivot`) get full structural
> detail, the rest collapse to a `kind` enum + generated prose (still complete — the
> cut is schema-size only, not documentation). The grouped version measured ~2,120
> added tokens.
>
> **What PP-5 did NOT fix, and is not meant to:** every specific finding below —
> `highlight` unimplemented, `bulletPreset` collapsing to two effective values,
> conditional-format's silent-fallback operators/kinds, `add_shape`'s undocumented
> preset names, chart-type gaps, etc. — is a **capability** or **silent-fallback**
> defect in the handler itself, not a schema-structure problem, and is owned by its own
> PP item (PP-9, PP-12, PP-13, PP-14, PP-15, PP-16, PP-21, PP-22 respectively — see
> `docs/superpowers/plans/2026-08-23-pp-index.md`). The tables below are left exactly as
> written at the time of the original audit (a dated record), not updated in place, so
> whoever implements those items has the original evidence rather than a paraphrase.
> The schema **now correctly states the narrow truth** for each of these (e.g.
> `bulletPreset`'s schema enum is `['BULLET','NUMBERED']` with a note that anything else
> collapses to BULLET — matching the handler exactly, not overselling it) — the
> resulting behavior is unchanged until the owning PP item widens it.

Triggered by a real bug: Word's `edit_chart` lets the model set a chart title and one
series' numeric values, but has no `categories` parameter at all — the model can never
label a chart's axis categories or name its series, even though the *schema and its
description* don't oversell this (they only ever mention title+values). That specific
case is a narrow-but-honest tool, not a schema/handler mismatch. The question this
section answers is broader: **across every tool/op in all three add-ins, does the JSON
schema advertised to the LLM (`entry.ts`) actually match what the C# handler
(`*Tools.cs`) reads and does?** Verified by reading every schema definition against its
handler body directly, not by inference.

Two bug classes turned up, ranked by how badly they mislead the model:

- **Silent no-op with false success** — the tool accepts a parameter, does nothing
  useful with an out-of-range value, and still reports success. Worse than an error,
  because the model has no signal to retry or ask the user.
- **Undocumented schema** — the wire-level JSON Schema for a parameter is just
  `{type: 'object'}` or `{type: 'string'}` with no enum/field list; the *real* contract
  lives only in a free-text description (or nowhere). The model can only guess valid
  values, and any guess outside the handler's recognized set falls into one of the
  silent-no-op cases above.

### Word (`WordTools.cs` / `entry.ts`)

| # | Tool / command | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `apply_commands` (whole tool) | The wire schema for `commands` items is just `{type:'object'}` — no `kind` field, no per-kind shape; everything is prose-only in the description. Consequence: `cmd.GetProperty("kind")` sits **outside** the per-command try/catch, so one malformed command (missing `kind`) throws and aborts the **entire remaining batch** with a generic error — while any commands already applied earlier in the same batch stay applied (no rollback). The tool's `IsError:true` result can under-report what actually changed. | Undocumented schema → robustness gap | entry.ts (commands: array<object>) / WordTools.cs:210-268 |
| 2 | `updateTextStyle` | Schema's `style` field is `{type:'object'}` with zero enumerated keys — valid keys only exist in prose. Concretely: **`highlight` is not implemented** (9 of genoffice's 10 fields), and because nothing enumerates valid keys, requesting `highlight` (or any hallucinated key) is silently ignored — `fields.Contains(x)` never matches, nothing throws, and the tool still returns `"updateTextStyle: ok"`. | Silent no-op + false success | entry.ts (style: object) / WordTools.cs:236-238, 379-416 |
| 3 | `createParagraphBullets` (`bulletPreset`) | Schema implies distinct named presets (mirroring genoffice's `BULLET_DISC_CIRCLE_SQUARE`-style names). Handler only ever checks `bulletPreset.StartsWith("NUMBERED")` — every other value, including a correctly-formed preset name, collapses to the same generic `ApplyBulletDefault()`. The model can ask for a specific bullet style and be silently downgraded. | Silent no-op | entry.ts / WordTools.cs:248-250, 544-561 |
| 4 | `edit_chart` | No mismatch — schema and description accurately describe the narrow capability (title + one series' values, no categories, no chart-type choice). Listed here for completeness since it's the finding that prompted this audit. | Honest but narrow (not a bug) | entry.ts / WordTools.cs:132-169 |

`updateParagraphStyle` has the same "no enumerated keys in the wire schema" structural
issue as `updateTextStyle`, but has full 10/10 field parity with genoffice underneath,
so it currently works only because the model happens to send valid keys — a latent
version of the same risk, not an active bug today.

### Excel (`ExcelTools.cs` / `entry.ts`)

`propose_operations`'s formal schema is `{operations: {type:'array', items:
{type:'object'}}}` — like Word's `apply_commands`, **no per-operation-kind schema
exists at all**; every op's real parameter shape lives only in one long free-text
description block (`entry.ts:198-224`). That description is the de facto schema for
everything below.

| # | Operation | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `add_conditional_format` | **Worst gap in the file.** The description says `rule: {kind, ...}` and never lists what fields any of its 8 `kind` values need. Actual per-kind requirements: `number`→`operator,value,value2`; `text`→`text`; `top10`→`rank,percent,bottom,format.{bold,fontColor,fillColor}`; `formula`→`formula`; `colorScale`→`minColor,midColor,maxColor`; `dataBar`→`color` — none named in the schema. The model must guess field names by convention. | Undocumented schema | entry.ts:222 / ExcelTools.cs:400-476 |
| 2 | `add_conditional_format` (`kind:"number"`, `operator`) | Any `operator` string other than `greaterThan`/`lessThan`/`equal`/`between` silently becomes `equal` — no error. | Silent no-op | ExcelTools.cs:388-398 |
| 3 | `add_pivot` (`values[].agg`) | Any `agg` value other than `count`/`average`/`max`/`min` (e.g. a typo like `"avg"`) silently becomes `sum`, unstated in the description. | Silent no-op | entry.ts:218 / ExcelTools.cs:1202-1211 |
| 4 | `add_shape` (`shapeType`) | 26 valid preset names + `"textbox"` exist in the handler; **none are listed in the schema or description**, and an unrecognized name silently becomes a plain rectangle. | Undocumented schema + silent no-op | entry.ts:213 / ExcelTools.cs:27-56, 746-773 |
| 5 | `edit_chart` (`chartType`) | The handler supports 6 chart types (column/bar/line/area/pie/doughnut, `ExcelChartTypeMap`), but only `add_chart`'s narrower 3-type enum (column/line/pie) is ever documented — the model has no way to discover `edit_chart` can do bar/area/doughnut at all. | Undocumented schema (capability hidden, not broken) | entry.ts:210-211 / ExcelTools.cs:58-66, 665-726 |
| 6 | `edit_chart` (data rebinding) | No parameter exists to repoint an existing chart at a different range — the direct Excel analog to Word's `edit_chart` category gap. Less severe than Word's because `add_chart` binds live to a sheet range via `SetSourceData`, so categories/series update automatically when the model edits the underlying cells with `set_cell`/`set_range` — only *rebinding to a new range* is actually missing. | Capability gap (not schema/handler mismatch) | ExcelTools.cs:643-663 |
| 7 | `set_page_setup` (`scale` + `fitToWidth`/`fitToHeight` together) | If both are set in one op, the `fitToWidth` branch runs after `scale` and unconditionally sets `Zoom = false`, silently discarding the `scale` value — Excel's own UI treats these as mutually exclusive, but nothing in the tool description warns the model. | Silent no-op | ExcelTools.cs:857-865 |
| 8 | `add_defined_name` / `delete_defined_name` | Only ever touch workbook-scoped names (no `sheet?` param) — yet `read_sheet_features` reports sheet-scoped defined names as something that can exist. The model can discover a sheet-scoped name but has no op to create one. Schema and handler agree with each other, so not a mismatch, but a real read/write asymmetry. | Capability gap | entry.ts:171 / ExcelTools.cs:1124-1135 |
| 9 | `delete_table` | Calls `.Unlist()` — converts the table back to a plain range **and keeps all the data**. The name/description could easily be read as "remove the table and its data." | Misleading description | entry.ts:217 / ExcelTools.cs:1085-1088 |
| 10 | `add_sparkline` (`targetCell` omitted) | Defaults to the *same* cells as `dataRange` — the sparkline draws inside the very cells holding its own source data. Not mentioned in the description. | Undocumented default | entry.ts:212 / ExcelTools.cs:728-744 |
| 11 | `add_conditional_format` (`kind:"text"`) | Always uses "contains" (`xlContains`) — no "starts with"/"ends with"/"not contains", despite Excel natively offering those and the schema giving no indication only one mode exists. | Capability gap, undocumented | ExcelTools.cs:418-423 |
| 12 | `add_conditional_format` (`kind:"duplicate"`) | Hardcoded to highlight duplicates only (`xlDuplicate`) — no way to flip to "highlight uniques," and nothing documents the restriction. | Capability gap, undocumented | ExcelTools.cs:427-430 |

### PowerPoint (`PowerPointTools.cs` / `entry.ts`)

Best-behaved of the three: all 23 tools have exact 1:1 name correspondence between
schema and handler, and **`add_chart` here does not have Word's bug** — `categories`
and `series[].name`/`series[].values` are all genuinely read and written into the
chart's embedded workbook. The gaps that do exist are narrower and more contained:

| # | Tool | Issue | Class | File:line (schema / handler) |
|---|---|---|---|---|
| 1 | `edit_chart` (`chartType`) | Handler only applies the change if the value is found in `PptChartTypeMap` — an unrecognized/typo'd value is silently ignored, and the tool still returns `"Chart updated."` (success). | Silent no-op + false success | PowerPointTools.cs:510-513 |
| 2 | `edit_chart` (`legendPos`) | Schema is bare `{type:'string'}` with no enum or valid-value guidance. Handler expects exactly `"none"`/`"r"`/`"t"`/`"l"` — any natural-language guess a model would plausibly send (`"right"`, `"bottom"`, `"top"`, `"left"`) silently falls into the bottom-position branch. | Undocumented schema + silent no-op | entry.ts:356 / PowerPointTools.cs:519-531 |
| 3 | `add_chart` (`kind`) | Unrecognized value silently defaults to `"bar"`, no enum declared in the schema (unlike `add_shape.shapeType`, which does declare one). | Undocumented schema + silent fallback | PowerPointTools.cs:442 |
| 4 | `add_smartart` (`layout`) | Unrecognized value silently defaults to `"list"` ("Basic Block List"), no enum declared. | Undocumented schema + silent fallback | PowerPointTools.cs:564 |
| 5 | `edit_table_structure` (`kind`) / `edit_table_style` (`borderPreset`) | Documented only in free-text description, no JSON-schema `enum`, inconsistent with `add_shape`'s stricter pattern. | Undocumented schema | entry.ts (both) |

`set_element_style`'s missing `underline`/`align`/`fontFamily` is **not** a
schema/handler mismatch — both sides consistently omit them, so it's a smaller feature
set rather than a case of the model being told it can do something it can't.

### Pattern across all three

The two gateway tools (`apply_commands` in Word, `propose_operations` in Excel) both
use a completely untyped `items: {type:'object'}` schema for their batched sub-commands
— the entire per-kind parameter contract lives in prose. This is *architecturally*
consistent with genoffice's own equivalents (which have the same untyped-envelope
design), but genoffice's command/op TypeScript interfaces are internally documented and
tightly matched to their handlers (verified in the original genoffice audit); here, the
prose-only contract combined with several silent-fallback branches in the handlers is
what turns "vague schema" into "the model can silently fail and be told it succeeded."
The single highest-value fix across all three add-ins would be replacing free-text
parameter descriptions with real per-kind JSON-schema shapes (at least `enum` arrays for
every closed set of string values) — that alone would have caught 8 of the 17 issues
above before they could reach the handler.
