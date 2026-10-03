using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // undo_last_action / redo_last_action - an undo/redo stack of the
    // assistant's OWN reversible actions, per mailbox chat, in memory only.
    // Never touches the user's manual actions, and never Outlook's native
    // Undo - see OutlookTools.Undo.cs.md for why that was tried and ruled
    // out. Every mutating handler records one entry AFTER its change
    // succeeds; irreversible actions push a barrier instead. Before
    // reversing, each entry checks the item still holds what the assistant
    // left there and refuses rather than overwrite an intervening change.
    public static partial class OutlookTools
    {
        private static readonly Dictionary<string, ActionHistory<UndoEntry>> HistoryByMailbox =
            new Dictionary<string, ActionHistory<UndoEntry>>();

        private static ActionHistory<UndoEntry> HistoryFor(string mbxKey)
        {
            ActionHistory<UndoEntry> h;
            if (!HistoryByMailbox.TryGetValue(mbxKey, out h))
            {
                h = new ActionHistory<UndoEntry>();
                HistoryByMailbox[mbxKey] = h;
            }
            return h;
        }

        private sealed class UndoConflictException : Exception
        {
            public UndoConflictException(string message) : base(message) { }
        }

        private abstract class UndoEntry
        {
            public string ToolName;
            public string Subject;

            // Each returns the user/model-facing result line, or throws.
            public abstract string Undo(ActionHistory<UndoEntry> history);
            public abstract string Redo(ActionHistory<UndoEntry> history);

            // An Outlook move assigns the item a new EntryID - every entry
            // pointing at the old one must follow it.
            public virtual void RemapItemId(string oldEntryId, string newEntryId, string newStoreId) { }
        }

        // ---- recording (called by the mutating handlers) ----

        internal static object[] ReadProps(object item, string[] props)
        {
            var values = new object[props.Length];
            for (int i = 0; i < props.Length; i++) values[i] = GetProp(item, props[i]);
            return values;
        }

        // Call after the handler's own Save(): reads the "after" values back
        // from the item so the conflict check compares against what Outlook
        // actually stored.
        internal static void RecordSnapshot(string mbxKey, string toolName, object item, string subject, string[] props, object[] before)
        {
            HistoryFor(mbxKey).Push(new SnapshotEntry
            {
                ToolName = toolName,
                Subject = subject,
                ItemEntryId = ItemEntryIdOf(item),
                ItemStoreId = ItemStoreIdOf(item),
                Props = props,
                Before = before,
                After = ReadProps(item, props),
            });
        }

        internal static void RecordEmailFlag(string mbxKey, Outlook.MailItem mail, bool wasMarkedAsTask, Outlook.OlMarkInterval interval, object[] before)
        {
            HistoryFor(mbxKey).Push(new EmailFlagEntry
            {
                ToolName = "set_email_reminder",
                Subject = mail.Subject ?? "",
                ItemEntryId = mail.EntryID,
                ItemStoreId = ItemStoreIdOf(mail),
                Props = EmailFlagEntry.FlagProps,
                Before = before,
                After = ReadProps(mail, EmailFlagEntry.FlagProps),
                WasMarkedAsTask = wasMarkedAsTask,
                Interval = interval,
            });
        }

        // The action moved the item from -> to (move_email, soft delete_email).
        // oldEntryId must be the item's EntryID as read before the move, the
        // same way the other handlers record it - earlier entries for this
        // item are rewritten to the new id so undo can still find it.
        internal static void RecordMove(string mbxKey, string toolName, string idLabel, string subject, string oldEntryId, string newEntryId, Outlook.Folder from, Outlook.Folder to)
        {
            ActionHistory<UndoEntry> history = HistoryFor(mbxKey);
            string newStoreId = to.StoreID;
            history.ForEachEntry(e => e.RemapItemId(oldEntryId, newEntryId, newStoreId));
            history.Push(new MoveEntry
            {
                ToolName = toolName,
                IdLabel = idLabel,
                Subject = subject,
                ItemEntryId = newEntryId,
                ItemStoreId = newStoreId,
                From = FolderRef.Of(from),
                To = FolderRef.Of(to),
            });
        }

        // A newly created item: undo moves it to Deleted Items (recoverable),
        // redo moves it back - i.e. a move from Deleted Items to its folder.
        internal static void RecordCreated(string mbxKey, string toolName, string idLabel, object item, string subject)
        {
            dynamic d = item;
            Outlook.Folder home = (Outlook.Folder)d.Parent;
            Outlook.Folder deleted = home.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems) as Outlook.Folder;
            HistoryFor(mbxKey).Push(new MoveEntry
            {
                ToolName = toolName,
                IdLabel = idLabel,
                Subject = subject,
                ItemEntryId = ItemEntryIdOf(item),
                ItemStoreId = home.StoreID,
                From = FolderRef.Of(deleted),
                To = FolderRef.Of(home),
            });
        }

        internal static void RecordCategoryColor(string mbxKey, string name, bool existed, Outlook.OlCategoryColor prior, Outlook.OlCategoryColor now)
        {
            HistoryFor(mbxKey).Push(new CategoryColorEntry
            {
                ToolName = "set_category_color",
                Subject = name,
                Existed = existed,
                Prior = prior,
                Now = now,
            });
        }

        internal static void RecordIrreversible(string mbxKey, string description)
        {
            HistoryFor(mbxKey).PushBarrier(description);
        }

        // ---- the tools ----

        private static ToolResult UndoLastAction(string mbxKey)
        {
            ActionHistory<UndoEntry> history = HistoryFor(mbxKey);
            UndoEntry entry;
            string barrier;
            if (!history.TryTakeUndo(out entry, out barrier))
            {
                return new ToolResult
                {
                    Output = barrier != null
                        ? "Can't undo: the last action (" + barrier + ") can't be reversed, so nothing before it can be undone either. Nothing was changed."
                        : "Nothing to undo - no reversible action by the assistant in this session. (Undo only covers the assistant's own actions, not changes made directly in Outlook.)",
                    Summary = "undo_last_action",
                };
            }

            try
            {
                string output = entry.Undo(history);
                history.CompleteUndo(entry);
                return new ToolResult { Output = output, Mutated = true, Summary = "undo_last_action" };
            }
            catch (Exception ex)
            {
                return FailedReversal("undo", entry, ex, "undo_last_action");
            }
        }

        private static ToolResult RedoLastAction(string mbxKey)
        {
            ActionHistory<UndoEntry> history = HistoryFor(mbxKey);
            UndoEntry entry;
            if (!history.TryTakeRedo(out entry))
                return new ToolResult { Output = "Nothing to redo.", Summary = "redo_last_action" };

            try
            {
                string output = entry.Redo(history);
                history.CompleteRedo(entry);
                return new ToolResult { Output = output, Mutated = true, Summary = "redo_last_action" };
            }
            catch (Exception ex)
            {
                return FailedReversal("redo", entry, ex, "redo_last_action");
            }
        }

        private static ToolResult FailedReversal(string verb, UndoEntry entry, Exception ex, string summary)
        {
            DebugLog.WriteException(summary, ex);
            string reason = ex is UndoConflictException
                ? ex.Message
                : "Outlook reported: " + ex.Message;
            return new ToolResult
            {
                Output = "Could not " + verb + " " + entry.ToolName + " on \"" + entry.Subject + "\": " + reason +
                         " This step was dropped from the history; earlier steps can still be undone.",
                IsError = true,
                Summary = summary,
            };
        }

        // ---- entry kinds ----

        private class SnapshotEntry : UndoEntry
        {
            public string ItemEntryId;
            public string ItemStoreId;
            public string[] Props;
            public object[] Before;
            public object[] After;

            public override string Undo(ActionHistory<UndoEntry> history) { return Apply(Before, After, "Undid"); }
            public override string Redo(ActionHistory<UndoEntry> history) { return Apply(After, Before, "Redid"); }

            protected string Apply(object[] target, object[] expected, string verb)
            {
                object item = ItemById(ItemEntryId, ItemStoreId);
                EnsureUnchanged(item, Props, expected);
                List<string> skipped = WriteProps(item, Props, target);
                ((dynamic)item).Save();
                return Describe(verb, target, expected, skipped);
            }

            protected string Describe(string verb, object[] target, object[] from, List<string> skipped)
            {
                var changes = new List<string>();
                for (int i = 0; i < Props.Length; i++)
                    if (!SameValue(from[i], target[i]) && !skipped.Contains(Props[i]))
                        changes.Add(Props[i] + ": " + FormatValue(from[i]) + " -> " + FormatValue(target[i]));
                string text = verb + " " + ToolName + " on \"" + Subject + "\"" +
                              (changes.Count > 0 ? " (" + string.Join(", ", changes) + ")" : "") + ".";
                if (skipped.Count > 0) text += " Outlook would not restore: " + string.Join(", ", skipped) + ".";
                return text;
            }

            public override void RemapItemId(string oldEntryId, string newEntryId, string newStoreId)
            {
                if (SameEntryId(ItemEntryId, oldEntryId))
                {
                    ItemEntryId = newEntryId;
                    ItemStoreId = newStoreId;
                }
            }
        }

        // set_email_reminder: MarkAsTask can't be reversed by writing
        // properties back - a message that wasn't flagged before needs
        // ClearTaskFlag(), and redoing it needs MarkAsTask again.
        private sealed class EmailFlagEntry : SnapshotEntry
        {
            public static readonly string[] FlagProps =
            {
                "FlagRequest", "FlagStatus", "TaskStartDate", "TaskDueDate", "ReminderTime", "ReminderSet",
            };

            public bool WasMarkedAsTask;
            public Outlook.OlMarkInterval Interval;

            public override string Undo(ActionHistory<UndoEntry> history)
            {
                if (WasMarkedAsTask) return Apply(Before, After, "Undid");

                Outlook.MailItem mail = (Outlook.MailItem)ItemById(ItemEntryId, ItemStoreId);
                EnsureUnchanged(mail, Props, After);
                mail.ClearTaskFlag();
                mail.ReminderSet = false;
                mail.Save();
                return "Undid set_email_reminder on \"" + Subject + "\" (follow-up flag and reminder cleared).";
            }

            public override string Redo(ActionHistory<UndoEntry> history)
            {
                if (WasMarkedAsTask) return Apply(After, Before, "Redid");

                Outlook.MailItem mail = (Outlook.MailItem)ItemById(ItemEntryId, ItemStoreId);
                EnsureUnchanged(mail, Props, Before);
                mail.MarkAsTask(Interval);
                List<string> skipped = WriteProps(mail, Props, After);
                mail.Save();
                return "Redid set_email_reminder on \"" + Subject + "\" (follow-up flag set again" +
                       (skipped.Count > 0 ? "; Outlook would not restore: " + string.Join(", ", skipped) : "") + ").";
            }
        }

        private sealed class FolderRef
        {
            public string EntryId;
            public string StoreId;
            public string Name;

            public static FolderRef Of(Outlook.Folder f)
            {
                return new FolderRef { EntryId = f.EntryID, StoreId = f.StoreID, Name = f.Name };
            }
        }

        // The action moved the item From -> To. Undo moves it To -> From,
        // redo From -> To, each only if the item is still where the previous
        // step left it.
        private sealed class MoveEntry : UndoEntry
        {
            public string IdLabel;
            public string ItemEntryId;
            public string ItemStoreId;
            public FolderRef From;
            public FolderRef To;

            public override string Undo(ActionHistory<UndoEntry> history) { return MoveBetween(history, To, From, "Undid"); }
            public override string Redo(ActionHistory<UndoEntry> history) { return MoveBetween(history, From, To, "Redid"); }

            private string MoveBetween(ActionHistory<UndoEntry> history, FolderRef expectedNow, FolderRef dest, string verb)
            {
                Outlook.Folder destFolder = (Outlook.Folder)Ns.GetFolderFromID(dest.EntryId, dest.StoreId);
                dynamic d = ItemById(ItemEntryId, ItemStoreId);
                Outlook.Folder parent = (Outlook.Folder)d.Parent;
                if (!Ns.CompareEntryIDs(parent.EntryID, expectedNow.EntryId))
                    throw new UndoConflictException("it is no longer in " + expectedNow.Name + " (now in " + parent.Name + ").");

                dynamic moved = d.Move(destFolder);
                string oldId = ItemEntryId;
                string newId = moved.EntryID;
                string newStore = destFolder.StoreID;
                RemapItemId(oldId, newId, newStore); // this entry is off both stacks right now
                history.ForEachEntry(e => e.RemapItemId(oldId, newId, newStore));

                return verb + " " + ToolName + ": moved \"" + Subject + "\" to " + dest.Name + ".\n" + IdLabel + ": " + newId;
            }

            public override void RemapItemId(string oldEntryId, string newEntryId, string newStoreId)
            {
                if (SameEntryId(ItemEntryId, oldEntryId))
                {
                    ItemEntryId = newEntryId;
                    ItemStoreId = newStoreId;
                }
            }
        }

        private sealed class CategoryColorEntry : UndoEntry
        {
            public bool Existed;
            public Outlook.OlCategoryColor Prior;
            public Outlook.OlCategoryColor Now;

            public override string Undo(ActionHistory<UndoEntry> history)
            {
                Outlook.Categories cats = Ns.Categories;
                Outlook.Category cat = cats[Subject];
                if (cat == null || cat.Color != Now)
                    throw new UndoConflictException("the color tag has been changed or removed since.");
                if (Existed)
                {
                    cat.Color = Prior;
                    return "Undid set_category_color: \"" + Subject + "\" is " + ColorName(Prior) + " again.";
                }
                cats.Remove(Subject);
                return "Undid set_category_color: removed the color tag \"" + Subject + "\" it had created.";
            }

            public override string Redo(ActionHistory<UndoEntry> history)
            {
                Outlook.Categories cats = Ns.Categories;
                Outlook.Category cat = cats[Subject];
                if (Existed)
                {
                    if (cat == null || cat.Color != Prior)
                        throw new UndoConflictException("the color tag has been changed or removed since.");
                    cat.Color = Now;
                }
                else
                {
                    if (cat != null)
                        throw new UndoConflictException("a color tag with that name exists again.");
                    cats.Add(Subject, Now);
                }
                return "Redid set_category_color: \"" + Subject + "\" is " + ColorName(Now) + ".";
            }
        }

        // ---- helpers ----

        private static object GetProp(object item, string name)
        {
            return item.GetType().InvokeMember(name, BindingFlags.GetProperty, null, item, null);
        }

        private static void SetProp(object item, string name, object value)
        {
            item.GetType().InvokeMember(name, BindingFlags.SetProperty, null, item, new[] { value });
        }

        // Writes every property it can; returns the names Outlook refused
        // (reported to the caller rather than failing the whole reversal).
        private static List<string> WriteProps(object item, string[] props, object[] values)
        {
            var skipped = new List<string>();
            for (int i = 0; i < props.Length; i++)
            {
                if (SameValue(GetProp(item, props[i]), values[i])) continue;
                try { SetProp(item, props[i], values[i]); }
                catch (Exception ex)
                {
                    DebugLog.WriteException("undo WriteProps " + props[i], ex);
                    skipped.Add(props[i]);
                }
            }
            return skipped;
        }

        private static void EnsureUnchanged(object item, string[] props, object[] expected)
        {
            for (int i = 0; i < props.Length; i++)
            {
                if (!SameValue(GetProp(item, props[i]), expected[i]))
                    throw new UndoConflictException("the item was changed since (" + props[i] + " is no longer " +
                                                    FormatValue(expected[i]) + "), so undoing would overwrite that change.");
            }
        }

        private static bool SameValue(object a, object b)
        {
            if (a is string || b is string)
                return string.Equals(Convert.ToString(a) ?? "", Convert.ToString(b) ?? "", StringComparison.Ordinal);
            if (a == null || b == null) return a == null && b == null;
            if (a is DateTime && b is DateTime)
                return Math.Abs(((DateTime)a - (DateTime)b).TotalSeconds) < 1;
            return a.Equals(b) ||
                   string.Equals(Convert.ToString(a, CultureInfo.InvariantCulture), Convert.ToString(b, CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        private static string FormatValue(object v)
        {
            if (v is DateTime)
            {
                DateTime d = (DateTime)v;
                return d.Year < 1900 || d.Year > 4000 ? "(none)" : Iso(d);
            }
            if (v is string) return "\"" + v + "\"";
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "(none)";
        }

        // The same item can be addressed by differently formatted EntryIDs
        // (e.g. an Exchange short-term "EF00..." id vs. the long-term one), so
        // fall back to MAPI's own comparison when the strings differ.
        private static bool SameEntryId(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
            try { return Ns.CompareEntryIDs(a, b); }
            catch { return false; }
        }

        private static string ItemEntryIdOf(object item)
        {
            return ((dynamic)item).EntryID;
        }

        private static string ItemStoreIdOf(object item)
        {
            try { return ((Outlook.Folder)((dynamic)item).Parent).StoreID; }
            catch { return null; }
        }
    }
}
