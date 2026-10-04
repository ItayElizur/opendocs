# TaskPaneHost.cs

## `TaskPaneHost` constructor

Deliberately does NOT dereference _document here (no .Path/.FullName
read). Accessing Application.ActiveDocument-equivalent state eagerly,
at the exact moment this constructor runs inside CustomTaskPanes.Add
(called from Application_WindowActivate), hits a COM timing issue in
Word's own startup/activation sequence and silently kills the whole
add-in connection (VSTO never connects it - no exception, no
resiliency-disabled entry, just Connect=False forever). Confirmed by
direct repro. GetChatId() computes lazily on first actual use, by
which point the pane is visible and the user has triggered a
message - so _document's state is guaranteed settled.

## `GetChatId`

A saved id is final - never re-checked again. An "unsaved-" id
is provisional: re-check the document's Path on every call, so
the first use after the user saves migrates chat history and
doc settings onto the real per-file id (FT-1 Task 7b). The Path
read is one cheap COM property on operations (load-history,
append-message, etc.) that are already doing file I/O.

## `GetChatId` - unsaved document id

An unsaved document has no on-disk Path; Document.FullName
falls back to its temp Name (e.g. "Document1") in that case,
which is not a stable key across sessions - and with
multiple panes now possible in one process, "unsaved-<pid>"
alone would collide across two different unsaved documents,
so the window handle is folded in too.

## `GetChatId` - saved-id stickiness

Once _chatId is set to a saved (non-"unsaved-") id, the guard at
the top of this method returns it immediately on every later
call - so a subsequent Save As (which changes _document.FullName
again) does NOT re-key. The conversation and guidelines stay
with the id first saved to, not the copy. Whether they should
follow the copy instead has no obviously right answer; changing
this silently would be worse than this documented quirk.

## `OnSelectionChanged` - selection.Text null for shapes

ROOT CAUSE FOUND (2026-08-24, via DebugLog from a real repro):
a shape selection (clicking a chart/SmartArt) DOES set
Start != End (hasSelection=true) as normal, but
selection.Text returns NULL rather than "" for that selection
type - fullText.Length below threw NullReferenceException on
every single chart/SmartArt click, aborting OnSelectionChanged
before any of the object-detection code even ran. This was a
pre-existing gap (predates this session's object-detection
work) that only became visible once the previously-silent
outer catch in ThisAddIn.cs's Application_WindowSelectionChange
started logging instead of swallowing.

## `OnSelectionChanged` - start/end block indices

Post-hoc fix (2026-08-24, user-reported): the selection payload
previously carried only text, with no addressability at all -
the one selection kind FT-2 left un-addressable (Excel's range
gets an A1 address, PowerPoint's gets slideIndex/shapeIndex).
Without a paragraph index, the model has no way to call
replace_blocks on exactly what the user selected, and fell back
to insert_content instead - "translate this" appended a new
paragraph rather than replacing the selected one. 0-based,
matching every other paragraph index in this tool surface.

## `OnSelectionChanged` - object kind/index detection

Post-hoc addition (2026-08-24, user-reported: selecting a
table/chart/SmartArt "doesn't appear under selection" -
selection.Text for a shape selection is empty or a placeholder
character, not a useful pointer. Detect these three object
kinds and report the SAME 0-based index the read_table/
read_chart/read_smartart tools use, so the UI and the model
both get an actionable pointer instead of nothing.

## `OnSelectionChanged` - Selection.Type diagnostic log

Diagnostic addition (2026-08-24): logs Word's own raw
Selection.Type - lets us tell definitively whether Word
itself considers a shape "selected" for a given click
(wdSelectionShape/wdSelectionInlineShape) versus the click
just moving the text cursor near the shape without
selecting it (which would still report Start==End,
withinTable=false, and an empty ShapeRange - all "normal"
Word behavior, not a bug in this detection code).

## `OnSelectionChanged` - object-detection catch

Post-hoc diagnostic addition (2026-08-24): was a silent
catch-all before - now logged, since "selection doesn't
show a pointer" could mean this is throwing every time
rather than just finding nothing.

## `OnSelectionChanged` - effectiveHasSelection

ROOT CAUSE FOUND (2026-08-24, via DebugLog from a real repro):
objectKind resolved correctly to "table" (confirmed in the
log) even though hasSelection (Start != End) was FALSE - a
plain cursor click/placement inside a table cell, or a click
directly on a chart/SmartArt shape, does NOT set Start != End
in Word's object model the way dragging across text does.
bootstrap.ts's toSelectionScopeUpdate/defaultDescribeSelection
both bail out immediately on hasSelection:false (line 1,
matching every other app's same-shaped payload), so a
correctly-detected object was being discarded before it ever
reached the UI or the model - this is what made table/chart/
SmartArt selection never show a pointer, regardless of how
correct the detection itself was. The EFFECTIVE hasSelection
sent downstream now also counts "an object was detected" as a
selection, since that is exactly as actionable as a text
range or an Excel/PowerPoint selection.
