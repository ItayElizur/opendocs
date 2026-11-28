# ItemPaneHost.cs

The chat pane inside one Outlook item window: a mail draft, or an
appointment or meeting. `ThisAddIn.ToggleItemPane` creates it on the
first ribbon press in that window and disposes it on `Inspector.Close`.

## Whose chat, mode and undo

- **Chat.** The item's, not the window's. `ItemChatId.Resolve` gives a
  saved item `item-` + a hash of its store id and EntryID, so a draft
  reopened from Drafts finds its chat again. A new item has no EntryID
  until its first save (measured on a live Outlook: empty before
  `Save()`, 140 characters after), so it starts on a provisional `new-`
  id. The first `GetChatId()` after that save moves the chat and the
  per-chat settings onto the saved id, the same step Word, Excel and
  PowerPoint take for a new document. Outlook's AutoSave saves a draft
  every few minutes, so the move usually happens while the window is open.
  The saved id is then kept even if the EntryID changes later (the item
  is moved or sent).
- **Who the user is, and the work week.** The same lines in the prompt as the main window's pane
  (`TaskPaneHost.MailboxPromptContextAsync`).
- **Editing mode.** Per chat, like the main window's (`OutlookTools.SetMode`).
  `OutlookTools.MoveChat` carries it to the saved id at the first save, and
  `OutlookTools.ForgetChat` drops it when the window closes.

## The gate

`ExecuteTool` is the second copy of `entry.ts`'s item-window lists
(`itemReadOnlyTools`, `itemEditExtraTools`), checked against them by
`TierListsAgreeTests`. The item tools (`get_current_item`, `edit_current_body`,
`set_current_fields`, `read_main_chat`) run here against this pane's item,
through `OutlookTools.Draft.cs`. The other read tools pass on to
`OutlookTools.ExecuteAsync`; `find_meeting_slots` gets the meeting's attendees
when it is called without any. Every other name is refused as "done from the
main Outlook window": moving, deleting, flagging, drafting other mail,
answering meetings, undo of mailbox actions and every send stay there, so a
draft's chat, whose body holds the incoming email's text, cannot act on the
mailbox.

The edit tools need the CommentOnly slot ("Edit this draft" or "Edit this
meeting"). The pane keeps the mode the page last set in `_mode`; `PaneHostBase`
runs no tool before the page's first `set-mode`, so its `ReadOnly` start value
is never what decides a call.

The pane passes its kind (`"mail"` or `"appointment"`) to `PaneHostBase`, and
the page reads it as `window.__opendocsPaneKind` to offer the two item-window
modes and tools.

## The item as selection context

`PostItem` sends the item as the pane's selection (`itemKind` "mail" or
"appointment"), so the page tells the model what is in this window.
`PostSelection` drops a payload identical to the last one. It runs when
the page starts (`OnPageStarted`, because anything posted before the page
exists is dropped) and on every `PropertyChange`. Typing in the body does
not raise `PropertyChange`, so the body preview (the first 500
characters) can lag until another field changes.

An attendee's view of a meeting (`olMeetingReceived`) is marked
`isOrganizer: false`, and the page tells the model the user cannot change
it.
