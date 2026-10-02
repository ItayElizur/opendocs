# TaskPaneHost.cs

## `TaskPaneHost` constructor

Deliberately does NOT dereference _presentation here (no .Path/.FullName
read). Accessing the presentation's state eagerly, at the exact moment
this constructor runs during pane setup (EnsurePaneFor, called from
Application_WindowActivate / the startup call in ThisAddIn.cs), hits a
COM timing issue in PowerPoint's own startup/activation sequence and
silently kills the whole add-in connection (VSTO never connects it - no
exception, no resiliency-disabled entry, just Connect=False forever).
Confirmed by direct repro. GetChatId() computes lazily on first actual
use, by which point the pane is visible and the user has triggered a
message - so _presentation's state is guaranteed settled.

Kept in sync with WordAiAddIn/TaskPaneHost.cs.md's identical note (same
repro class; Word's Document vs. PowerPoint's Presentation).

## `GetChatId`

A saved id is final - never re-checked again. An "unsaved-" id
is provisional: re-check the presentation's Path on every call, so
the first use after the user saves migrates chat history and
doc settings onto the real per-file id (FT-1 Task 7b). The Path
read is one cheap COM property on operations (load-history,
append-message, etc.) that are already doing file I/O.

## `GetChatId` - saved-id stickiness

Once _chatId is set to a saved (non-"unsaved-") id, the guard at
the top of this method returns it immediately on every later
call - so a subsequent Save As (which changes _presentation.FullName
again) does NOT re-key. The conversation and guidelines stay
with the id first saved to, not the copy. Whether they should
follow the copy instead has no obviously right answer; changing
this silently would be worse than this documented quirk.

Kept in sync with WordAiAddIn/TaskPaneHost.cs.md's identical note.
