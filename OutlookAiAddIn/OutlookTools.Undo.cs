using System;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // undo_last_action - dispatches through Outlook's own ribbon Undo command
    // (Explorer.CommandBars.ExecuteMso("Undo")/GetEnabledMso("Undo")), the
    // same generic Office ribbon-command API Excel's/PowerPoint's
    // ExecuteRibbonUndoRedo use (Microsoft.Office.Core.CommandBars, from the
    // "Office" PIA reference already in this project - see
    // ExcelTools.History.cs/PowerPointTools.History.cs for the identical
    // pattern).
    //
    // CORRECTION (2026-09-28): this file originally implemented a small
    // custom one-shot inverse-op record (move_email/mark_email_read/
    // mark_email_unread/flag_email_important only), reasoned as necessary
    // because "Outlook has no undo hook for item-level mutations at all...
    // CommandBars.ExecuteMso('Undo') has no meaning for a mailbox action."
    // That claim was never reflection-verified the way this repo's other COM
    // surface claims are (see WordTools.History.cs/ExcelTools.History.cs's
    // own "confirmed by .NET reflection" comments) - .NET reflection against
    // the referenced Microsoft.Office.Interop.Outlook 15.0.0.0 PIA shows
    // Explorer (what Application.ActiveExplorer() already returns elsewhere
    // in this file, see OutlookTools.Search.cs's ApplySearch) DOES expose a
    // CommandBars property returning Microsoft.Office.Core.CommandBars - the
    // exact same type/mechanism Excel and PowerPoint use for this same
    // feature. Replaced the custom mechanism with this native one for
    // consistency with the other three apps.
    //
    // UNVERIFIED AGAINST LIVE OUTLOOK: the API surface exists and is
    // callable, but whether ExecuteMso("Undo") actually reverses a completed
    // item mutation (a move, a read/unread flip, an importance change) the
    // way it reverses a Word/Excel/PowerPoint document edit is NOT
    // confirmed - Outlook's native Ctrl+Z has a long-standing reputation for
    // being much more limited than Word's/Excel's/PowerPoint's own undo,
    // often only covering very recent UI-level actions (e.g. typing in the
    // reading pane or a compose window) rather than completed item
    // mutations made via automation. This is best-effort, same framing as
    // Excel's own caveat - a "Nothing to undo" result can be correct here
    // even right after a mutating tool call, not necessarily a bug.
    //
    // No redo_last_action for Outlook, same as before this change - not
    // added here since ExecuteMso("Redo") would carry the identical
    // unverified-scope caveat as Undo above, and this correction's scope is
    // limited to fixing the incorrect "no mechanism exists" premise, not
    // expanding the feature.
    public static partial class OutlookTools
    {
        private static ToolResult UndoLastAction()
        {
            try
            {
                Outlook.Explorer explorer = App.ActiveExplorer();
                if (explorer == null)
                    return new ToolResult { Output = "Nothing to undo - no active Outlook window.", Summary = "undo_last_action" };

                Microsoft.Office.Core.CommandBars bars = explorer.CommandBars;
                if (!bars.GetEnabledMso("Undo"))
                {
                    return new ToolResult
                    {
                        Output = "Nothing to undo (Outlook's own Undo may not cover this action - this is best-effort, same caveat as Excel).",
                        Summary = "undo_last_action",
                    };
                }
                bars.ExecuteMso("Undo");
                return new ToolResult { Output = "Undid the last action.", Mutated = true, Summary = "undo_last_action" };
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = "Could not undo: " + ex.Message, IsError = true, Summary = "undo_last_action" };
            }
        }
    }
}
