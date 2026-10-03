# PowerPointTools.Master.cs

## `ResolveSlideMaster`

Confirmed via .NET reflection against the real referenced PIA
(Microsoft.Office.Interop.PowerPoint.dll, GAC 15.0.0.0), same
discipline as PP-24's enum maps: _Presentation.SlideMaster returns
Master directly - no fallback needed. Known limitation, same class
as ActivePresentation itself (PP-1): a deck combining more than one
theme has more than one Slide Master, and only this default
(first) one is targeted.

## `MasterTarget`

User-requested (2026-09-22, live testing): add_master_element/
read_master_elements/remove_master_element/set_master_element_
transform previously only ever targeted the single top-level
Slide Master - there was no way to address one specific layout
(Title Slide, Title and Content, etc.), even though each layout
has its own independent shape collection (the exact thing that
caused the footer-text-not-inheriting bug). optional layoutName
resolves by case-insensitive substring against this presentation's
own layout names - same resolve-by-substring pattern
ResolveCustomLayout (PowerPointTools.LayoutAnim.cs) already uses
for set_slide_layout - and returns that layout's own Shapes
collection instead of the master's when given.

## `ResolveLayoutByQuery`

Review finding: neither of this struct's two callers used to
detect an ambiguous match - both took the FIRST layout whose name
merely CONTAINED the query substring, in whatever order
CustomLayouts happened to enumerate, with no signal to the caller
that a second (or third) candidate existed at all. Two layouts can
collide on name (nothing in the object model enforces uniqueness -
ListLayouts's own count-by-index logic already works around this
same risk) or simply both happen to contain the same substring
(e.g. query "Title" against both "Title Slide" and "Title and
Content"). Exact (case-insensitive) name matches are checked as
their own tier, ahead of substring matches, so typing the literal
full name is never "ambiguous" just because some other layout's
name happens to contain it.

## `ResolveLayoutByName`

Review finding: this used to only ever search
ActivePresentation.SlideMaster (the presentation's default/first
design), because there was no slide to derive scope from - unlike
ResolveCustomLayout (PowerPointTools.LayoutAnim.cs), which is
handed a specific slide and correctly searches THAT slide's own
Design. A deck combining more than one theme/design (e.g. slides
pasted in with "Preserve Source Formatting") could never reach a
non-default design's layout by name through add_master_element/
read_master_elements/remove_master_element/set_master_element_
transform's layoutName. Now searches every design in the
presentation (confirmed via reflection: Presentation.Designs is a
real, enumerable collection, each with its own SlideMaster).

## `ListLayouts`

User-requested (2026-09-22): before this, the only way to
discover a deck's real layout names was to deliberately pass a
bad layoutName and read the error message's "Available: ..."
list. Read-only, lists every layout in every design/theme in this
presentation (not just the default one - same widening as
ResolveLayoutByName above, for the same reason), plus how many
slides currently use each one.

## `ListLayouts` - keying by (design, layout)

Review finding: CustomLayout.Index is only unique WITHIN its
own design's SlideMaster - once more than one design is in
play, a bare Dictionary<int,int> keyed on Index alone would
silently conflate two different designs' layouts that happen
to share an Index (e.g. both designs' first layout is Index 1).
Key on the (design, layout) pair instead. CustomLayout.Design
and Design.Index are both confirmed via reflection.

## `DateAutoFormats`

Curated subset of the real PpDateTimeFormat enum (confirmed via
.NET reflection against the referenced PIA, Microsoft.Office.
Interop.PowerPoint.dll GAC 15.0.0.0) for set_headers_footers's
dateFormat. Deliberately excludes: the four time-only members
(Hmm/Hmmss/hmmAMPM/hmmssAMPM - a "date format" picker showing no
date at all doesn't fit this field); the two combined date+time
members (out of scope here - dateMode:"fixed" + dateText already
covers a time-bearing footer if ever needed); the seven locale-only
UAQ1-7 members (undocumented/regional); ppDateTimeFigureOut (a
sentinel PowerPoint uses internally, not a settable format); and
ppDateTimeFormatMixed (what PowerPoint reports when a range's
formats differ, never a value you set).

## `ApplyHeadersFooters`

Slide.HeadersFooters.DateAndTime/.Footer/.SlideNumber are all the
same HeaderFooter type. DisplayOnTitleSlide is NOT among them here -
real-Office-confirmed (2026-09-21, live error): setting it via a
SLIDE's HeadersFooters throws "HeadersFooters (unknown member) :
Invalid request. This property must be set using slide or title
master." even though reflection shows the member on the shared
HeadersFooters type with no hint of that restriction - a runtime
COM restriction reflection can't see, only member existence/type.
See ApplySkipTitleSlide below, which sets it on the MASTER instead.

Takes a HeadersFooters directly (not a Slide) so the SAME logic
applies to a slide's OWN HeadersFooters and to the Slide/Title
Master's (Master.HeadersFooters is the identical COM type) -
real-user-confirmed (2026-09-22): a slide added AFTER
set_headers_footers ran did not inherit the deck-wide toggle,
because only existing slides were ever touched. Now
SetHeadersFooters also applies to the master(s) on a deck-wide
call, on the hypothesis that a brand-new slide's initial
HeadersFooters state is seeded from its master at creation time -
unverified without a live retest, but harmless either way (the
master's own header/footer placeholders, if the theme has them,
are otherwise never addressed by this tool at all).

## `TryApplyHeadersFooters`

Review finding: SetHeadersFooters used to hand-repeat the same
try { ApplyHeadersFooters(...) } catch { DebugLog...; counter++ }
block 6 times (per-slide loop, SlideMaster, per-layout loop,
TitleMaster, per-title-layout loop, single-slide branch), each with
its own hand-typed DebugLog label - a future 7th target or a
signature change had 6 near-identical sites to update by hand.
Returns null on success, or the caught exception's message on
failure (never throws) - callers that only need a pass/fail signal
just check for null; the single-slide branch also needs the actual
message for its error response.

## `ApplySkipTitleSlide`

DisplayOnTitleSlide (the "Don't show on title slide" checkbox) is
deck-wide by nature, same as PageSetup.FirstSlideNumber - it lives
on the Slide Master's (and, if the deck has one, the Title
Master's - a deck can carry a second, separate master used only by
title-layout slides) HeadersFooters, not any individual slide's.

Review finding: the SlideMaster write below used to be unguarded,
while the TitleMaster write right after it was individually
wrapped - inconsistent with this file's own discipline of guarding
every master-level HeadersFooters write independently. If the
SlideMaster write threw (the same "Master (unknown member)" COM
restriction this file documents extensively for sibling writes),
the TitleMaster branch was never even attempted. Both writes are
now independently guarded; the TitleMaster write's own failure
stays best-effort/silent as before (unaffected by this fix), and
only the primary SlideMaster write's failure is surfaced to the
caller (returns its message, or null on success).

## `ApplySkipTitleSlide` - reading TitleMaster can also throw

Real-user-confirmed (2026-09-22): even READING pres.TitleMaster
(not just writing to it) can throw "Master (unknown member) :
Invalid request" in some deck states despite HasTitleMaster
reporting true first - best-effort, never let this abort the
rest of set_headers_footers.

## `SetHeadersFooters` - inferring dateMode

Review finding: dateText alone (no dateMode) used to be
silently dropped - never applied, never counted in `applied`,
and the call could even return the "no recognized fields" error
despite dateText being a valid, given field. Infer dateMode the
same way footerText infers footerVisible:true just above.
User-directed: "auto" (and therefore dateFormat, which only
applies to "auto") is opt-in only and never inferred - dateMode
defaults to "fixed" and stays "fixed" unless the caller states
dateMode:"auto" explicitly, regardless of which dateFormat is
given.

## `SetHeadersFooters` - per-slide loop guarding (fourth round)

Real-user-confirmed (2026-09-22, fourth round): this per-SLIDE loop
was the one place left unguarded - a dateVisible:false +
dateMode:"auto" combination (never exercised before; every prior
test used dateVisible:true) threw "HeaderFooter (unknown member)"
here, aborting the whole call before even reaching the master/layout
seeding below, so none of THEIR failure-counting ever got a chance
to run. Same per-item best-effort treatment as the master/layout
loops now.

## `SetHeadersFooters` - seeding masters and layouts

Also seed the master(s) AND every custom layout, so a
slide added AFTER this call starts with the same
state, instead of only ever touching slides that
existed at call time. Real-user-confirmed (2026-09-22):
seeding only the Master was NOT enough - slideNumber/
date still showed on a new slide (this theme's layouts
already defaulted those on), but footerText did not,
because CustomLayout.HeadersFooters is confirmed (via
reflection) to be a SEPARATE object from the Master's,
not something a layout reads through to the master for -
a new slide inherits from whichever layout it's based
on, not from the top-level Master directly.

Real-user-confirmed (2026-09-22, second round): SOME
layouts throw "HeaderFooter (unknown member) : Invalid
request" on at least one of these property writes -
exact cause unconfirmed (a "Blank"-style layout with no
footer/date/number placeholder at all is the leading
guess, mirroring DisplayOnTitleSlide's own confirmed
must-be-set-on-master-or-slide restriction), but the
whole deck-wide call must not abort partway through
over one uncooperative layout. Best-effort per layout;
failures are counted, not silently hidden, and don't
stop the rest.

Every one of the calls below is wrapped individually -
real-user-confirmed (2026-09-22, third round): even
pres.SlideMaster.HeadersFooters / pres.TitleMaster
itself (not just a per-layout write) can throw
"Master (unknown member) : Invalid request" - exact
trigger still unconfirmed, so nothing here is trusted
to succeed unguarded any more.

## `SetHeadersFooters` - single-slide branch keeps going on failure

Review finding: this branch used to return immediately
on failure, before ever reaching the startNumber/
skipTitleSlide handling below - since those two fields
are documented as ALWAYS deck-wide regardless of
slideIndex, a rejected per-slide field combination on
one slide had the side effect of silently dropping an
unrelated deck-wide field given in the same call.
Record the failure and keep going instead.

## `SetHeadersFooters` - startNumber/skipTitleSlide guarding

Both deck-wide regardless of slideIndex, same as
set_slide_background's slideIndex:-1 convention but stronger:
FirstSlideNumber (PageSetup) and DisplayOnTitleSlide (the
Slide/Title Master's HeadersFooters) are simply not per-slide
properties in PowerPoint's object model at all. Applied even
when the per-slide write above failed (see singleSlideError).

Review finding: these two writes used to be unguarded, unlike
every other write in this function - if a deck-wide call had
already mutated every slide's footer/date/number and THEN
PageSetup.FirstSlideNumber or ApplySkipTitleSlide's unguarded
DisplayOnTitleSlide write threw (the same class of COM
restriction this file documents extensively elsewhere), the
exception would escape to Execute's outer catch and report a
bare error with no Mutated:true, hiding the changes that had
already landed. Guard each independently so one failing
doesn't mask the other succeeding or the per-slide work above.

## `SetHeadersFooters` - accurate scope description

Review finding: this message used to unconditionally describe
"every slide, plus every layout and the Slide/Title Master" for
a deck-wide call even when anyPerSlideField was false (e.g. only
skipTitleSlide/startNumber given) - overstating what actually
changed, since the per-slide/layout/master loop above never ran
in that case.

## `AddMasterElement` - left/top must be given together

Review finding: left/top used to only count as an explicit
position when BOTH were given; giving just one alongside/without
corner silently fell through to corner-based placement (or the
margin default) with no error - the supplied coordinate was
dropped without a trace.

## `AddMasterElement` - corner validated only when used

Review finding: corner's validity used to be checked before
hasExplicitPosition was known, so valid left+top alongside a
stray/garbage corner value threw even though corner would
never actually be used once an explicit position is given.
Only require/validate corner when it will actually be used.

## `AddMasterElement` - resolving color/font size up front

Review finding: for kind:"text", ColorUtil.HexToOle used to run
AFTER the textbox was already created on the master/layout - an
invalid hex string (e.g. "gray") threw with the shape already
inserted but never positioned or named, orphaning it at (0,0)
with no shapeIndex returned for cleanup. Resolve/validate
color and font size up front, before creating anything.

## `AddMasterElement` - ApplyOptionalName dedup fix

Review finding: ApplyOptionalName's dedup-against-siblings
used to resolve the shape's parent via `shape.Parent as
PowerPoint.Slide`, which is always null for a master/layout
shape - two add_master_element calls with the same name
both got the literal name, unlike slide shapes which
auto-suffix. Pass target.Shapes explicitly so dedup works
here too.

## `SetMasterElementTransform`

User-reported gap (2026-09-22, live testing): add_master_element
takes a starting position/size, but there was no way to adjust an
EXISTING master element afterward short of remove + re-add.
Mirrors PowerPointTools.Elements.cs's SetElementTransform exactly,
just addressed by masterShapeIndex instead of slideIndex+shapeIndex.

## `RemoveMasterElement` - the placeholder-deletion incident

Real-user-confirmed (2026-09-22, live testing): the Slide Master's
Title/Text/Footer/DateAndTime/SlideNumber theme placeholders live
in master.Shapes right alongside anything add_master_element adds
- masterShapeIndex 0 is the Title Placeholder on a typical theme,
NOT whatever was added last. Deleting one of these did NOT remove
its appearance from actual slides even after save/close/reopen -
each slide's own CustomLayout carries an independent placeholder
object (inheriting the master's formatting, not sharing the same
shape reference), so master.Shapes only ever reflects the Slide
Master's OWN copy. Net effect of the naive index-0-by-default
testing: two theme placeholders (Title, Text) were deleted from
the Slide Master with no visible effect and no way to undo via
this tool surface. Refusing to delete a placeholder here closes
that hole - a custom element add_master_element actually created
(msoTextBox/msoPicture, never msoPlaceholder) is unaffected.
