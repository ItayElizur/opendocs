using System;

namespace OfficeAi.Shared
{
    // The chat id for an Outlook item open in its own window (a mail draft, an appointment or meeting). A new item has
    // no EntryID until its first save, so its chat starts on a provisional id and moves to the saved one at that save.
    public static class ItemChatId
    {
        private const string ProvisionalPrefix = "new-";

        // current: the pane's id so far (null before the first call). Returns the id to use now; migrateFrom is the
        // provisional id whose chat must be moved onto it, or null. A saved id is sticky: an EntryID that changes later
        // (the item is moved or sent) keeps the chat where it is.
        public static string Resolve(string current, string storeId, string entryId, out string migrateFrom)
        {
            migrateFrom = null;
            bool provisional = current == null || current.StartsWith(ProvisionalPrefix, StringComparison.Ordinal);
            if (!provisional) return current;
            if (string.IsNullOrEmpty(entryId)) return current ?? ProvisionalPrefix + Guid.NewGuid().ToString("N");
            migrateFrom = current;
            return "item-" + ChatStore.ChatIdForKey((storeId ?? "") + "|" + entryId);
        }
    }
}
