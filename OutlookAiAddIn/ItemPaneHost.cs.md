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
- **Editing mode.** Per chat, like every pane (`OutlookTools.SetMode`).
- **Undo history.** The mailbox's (`TaskPaneHost.MailboxKey()`), so an
  action the model takes from here can be undone from the main window.

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
