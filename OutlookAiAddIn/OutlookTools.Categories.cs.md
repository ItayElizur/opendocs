# OutlookTools.Categories.cs

## `SetEventCategories` / `SetEventAvailability` - `store_id` parameter

Only needed for an event_id from someone else's shared
calendar (returned by list_events' mailbox parameter) -
ItemById/GetItemFromID can't find an item outside the
caller's own default store without it. Omit for your own
events, same as before this parameter existed. Whether the
caller actually has permission to act on the resulting item
is entirely up to Outlook/Exchange - this add-in doesn't add
its own authorization check on top of that.
