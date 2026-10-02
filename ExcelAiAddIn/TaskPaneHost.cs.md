# TaskPaneHost.cs

## `GetChatId` - saved-id stickiness

A saved id is final - never re-checked again. An "unsaved-" id
is provisional: re-check the workbook's Path on every call, so
the first use after the user saves migrates chat history and
doc settings onto the real per-file id (FT-1 Task 7b). The Path
read is one cheap COM property on operations (load-history,
append-message, etc.) that are already doing file I/O.

## `GetChatId` - unsaved workbook id

An unsaved workbook has no on-disk Path; Workbook.FullName
falls back to its temp Name (e.g. "Book1") in that case,
which is not a stable key across sessions - and with
multiple panes now possible in one process, "unsaved-<pid>"
alone would collide across two different unsaved workbooks,
so the window handle is folded in too.

## `OnSelectionChanged` - effective extent

Task 2b: report the effective (UsedRange-intersected) extent
alongside the literal one for whole-column/row selections and
any large selection - a bare "B1:B1048576" is both useless to
show the user and something the model would try to read in
full, hitting read_range's 2000-cell cap. Only pay for
UsedRange when it can matter; an ordinary drag-selection never
needs it.
