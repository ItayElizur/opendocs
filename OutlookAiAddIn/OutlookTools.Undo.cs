using System;
using System.Collections.Generic;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // undo_last_action - a small, deliberately narrow inverse-op mechanism,
    // NOT a general mutation log/undo framework. Outlook has no native undo
    // hook for item-level mutations (nothing like Word's Application/Document
    // Undo or Excel's/PowerPoint's ribbon-command dispatch - CommandBars.
    // ExecuteMso("Undo") has no meaning for a mailbox action), so this
    // records just enough about the single most recent qualifying action to
    // manually reverse it.
    //
    // Scoped to exactly the four cheaply-reversible mutations in
    // docs/ai-tool-surface.md's Outlook "Mutating tools" table:
    // move_email (reverse by moving back to the original folder),
    // mark_email_read/mark_email_unread (flip UnRead back), and
    // flag_email_important (flip Importance back). Deliberately EXCLUDED:
    // delete_email (especially permanent:true), the five auto-send tools
    // (send_email/send_reply/send_reply_all/send_forward/create_event), and
    // accept_meeting/decline_meeting (already notify the organizer -
    // irreversible in effect). There is no redo_last_action - once a single
    // one-shot record has been reversed there is nothing sensible left to
    // redo.
    public static partial class OutlookTools
    {
        private sealed class ReversibleAction
        {
            public string ToolName;
            public string Subject;

            // The item's EntryID/StoreID as of right after the recorded
            // action - re-resolved via ItemById on undo, same as every other
            // Outlook tool.
            public string ItemEntryId;
            public string ItemStoreId;

            // move_email only: the folder the item was moved OUT of, by the
            // FOLDER's own EntryID/StoreID (distinct from the item's) -
            // Namespace.GetFolderFromID resolves this precisely, unlike
            // ResolveFolder's name-based lookup which can't distinguish two
            // same-named folders in different parents/stores.
            public string SourceFolderEntryId;
            public string SourceFolderStoreId;
            public string SourceFolderName;

            // mark_email_read / mark_email_unread / flag_email_important
            // only: the value the flipped boolean held BEFORE this action
            // (UnRead for the mark tools, "was High importance" for the flag
            // tool) - restoring it is the entire undo.
            public bool PriorBool;
        }

        // Keyed by mailbox (GetChatId()'s mbx-<hash> key), same as
        // ModeByMailbox above - one pending reversible action per mailbox
        // chat, not a history. Recording a new action overwrites whatever
        // was there before; undoing one clears the slot (one-shot, not
        // re-undoable).
        private static readonly Dictionary<string, ReversibleAction> LastReversibleActionByMailbox = new Dictionary<string, ReversibleAction>();

        internal static void RecordMarkOrFlag(string mbxKey, string toolName, string itemEntryId, string itemStoreId, string subject, bool priorBool)
        {
            LastReversibleActionByMailbox[mbxKey] = new ReversibleAction
            {
                ToolName = toolName,
                ItemEntryId = itemEntryId,
                ItemStoreId = itemStoreId,
                Subject = subject,
                PriorBool = priorBool,
            };
        }

        internal static void RecordMove(string mbxKey, string itemEntryId, string itemStoreId, string subject, string sourceFolderEntryId, string sourceFolderStoreId, string sourceFolderName)
        {
            LastReversibleActionByMailbox[mbxKey] = new ReversibleAction
            {
                ToolName = "move_email",
                ItemEntryId = itemEntryId,
                ItemStoreId = itemStoreId,
                Subject = subject,
                SourceFolderEntryId = sourceFolderEntryId,
                SourceFolderStoreId = sourceFolderStoreId,
                SourceFolderName = sourceFolderName,
            };
        }

        private static ToolResult UndoLastAction(string mbxKey, JsonElement input)
        {
            ReversibleAction action;
            if (!LastReversibleActionByMailbox.TryGetValue(mbxKey, out action) || action == null)
            {
                return new ToolResult
                {
                    Output = "Nothing to undo - no reversible action (move/mark read-unread/flag importance) has been recorded yet in this session.",
                    Summary = "undo_last_action",
                };
            }

            // One-shot: clear immediately, before attempting the reversal, so
            // a failed or partial undo can never be retried against
            // already-changed state.
            LastReversibleActionByMailbox.Remove(mbxKey);

            switch (action.ToolName)
            {
                case "move_email":
                    return UndoMove(action);
                case "mark_email_read":
                case "mark_email_unread":
                    return UndoMark(action);
                case "flag_email_important":
                    return UndoFlag(action);
                default:
                    // Should be unreachable - only the four kinds above are
                    // ever recorded - but fail honestly rather than silently
                    // if that ever changes.
                    return new ToolResult
                    {
                        Output = "Nothing to undo - the last recorded action (\"" + action.ToolName + "\") is not one of the reversible kinds.",
                        Summary = "undo_last_action",
                    };
            }
        }

        private static ToolResult UndoMove(ReversibleAction action)
        {
            Outlook.Folder sourceFolder;
            try
            {
                sourceFolder = (Outlook.Folder)Ns.GetFolderFromID(action.SourceFolderEntryId, action.SourceFolderStoreId);
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Output = "Could not undo move_email: the original folder (\"" + action.SourceFolderName + "\") could not be resolved (" + ex.Message + ").",
                    IsError = true,
                    Summary = "undo_last_action",
                };
            }

            object item;
            try
            {
                item = ItemById(action.ItemEntryId, action.ItemStoreId);
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Output = "Could not undo move_email: the message no longer resolves (" + ex.Message + ").",
                    IsError = true,
                    Summary = "undo_last_action",
                };
            }

            dynamic d = item;
            dynamic movedBack = d.Move(sourceFolder);
            string newId = "";
            try { newId = movedBack.EntryID; } catch { }

            return new ToolResult
            {
                Output = "Undid move_email: moved \"" + action.Subject + "\" back to " + action.SourceFolderName + ".\nmessage_id: " + newId,
                Mutated = true,
                Summary = "undo_last_action",
            };
        }

        private static ToolResult UndoMark(ReversibleAction action)
        {
            Outlook.MailItem mail = ItemById(action.ItemEntryId, action.ItemStoreId) as Outlook.MailItem;
            if (mail == null)
            {
                return new ToolResult { Output = "Could not undo " + action.ToolName + ": the message no longer resolves.", IsError = true, Summary = "undo_last_action" };
            }
            mail.UnRead = action.PriorBool;
            mail.Save();
            return new ToolResult
            {
                Output = "Undid " + action.ToolName + ": restored \"" + action.Subject + "\" to " + (action.PriorBool ? "unread" : "read") + ".",
                Mutated = true,
                Summary = "undo_last_action",
            };
        }

        private static ToolResult UndoFlag(ReversibleAction action)
        {
            Outlook.MailItem mail = ItemById(action.ItemEntryId, action.ItemStoreId) as Outlook.MailItem;
            if (mail == null)
            {
                return new ToolResult { Output = "Could not undo flag_email_important: the message no longer resolves.", IsError = true, Summary = "undo_last_action" };
            }
            mail.Importance = action.PriorBool ? Outlook.OlImportance.olImportanceHigh : Outlook.OlImportance.olImportanceNormal;
            mail.Save();
            return new ToolResult
            {
                Output = "Undid flag_email_important: restored \"" + action.Subject + "\" to " + (action.PriorBool ? "high" : "normal") + " importance.",
                Mutated = true,
                Summary = "undo_last_action",
            };
        }
    }
}
