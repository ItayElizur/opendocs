using System;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        private static ToolResult ListFolders(JsonElement input)
        {
            var sb = new StringBuilder();
            int total = 0;
            try
            {
                WalkFolders(Ns.Folders, "", ref total, 800, 0, sb);
            }
            catch (Exception ex)
            {
                // A store/folder that's momentarily unreachable (shared mailbox,
                // archive, public folder) throws here and would otherwise wipe
                // out everything already collected from folders that were fine.
                DebugLog.WriteException("ListFolders", ex);
                if (total == 0) return new ToolResult { Output = "Could not list folders: " + ex.Message, IsError = true, Summary = "list_folders" };
                sb.AppendLine("! folder listing stopped early: " + ex.Message);
            }
            if (total == 0) return new ToolResult { Output = "No mail folders found.", Summary = "list_folders" };
            return new ToolResult { Output = sb.ToString(), Summary = "list_folders" };
        }

        private static void WalkFolders(Outlook.Folders folders, string path, ref int total, int cap, int depth, StringBuilder sb)
        {
            if (folders == null || depth > 8 || total >= cap) return;
            foreach (Outlook.Folder f in folders)
            {
                if (total >= cap) return;
                string name = "(unnamed folder)";
                try { name = f.Name; } catch { }
                string here = path.Length == 0 ? name : path + "\\" + name;
                bool isMail = false;
                try { isMail = f.DefaultItemType == Outlook.OlItemType.olMailItem; } catch { }
                if (isMail)
                {
                    total++;
                    int count = 0, unread = 0;
                    try { count = f.Items.Count; } catch { }
                    try { unread = f.UnReadItemCount; } catch { }
                    sb.AppendLine("- " + here + "  (items: " + count + ", unread: " + unread + ")");
                }
                try
                {
                    // ComRetry first: the exact RPC hiccups it retries for
                    // (RPC_S_CALL_FAILED etc.) are the "network problems
                    // connecting to Microsoft Exchange" class this whole fix
                    // is for, so most subfolder opens should just succeed on
                    // attempt 2/3 instead of being given up on immediately.
                    Outlook.Folders sub = null;
                    ComRetry.Run(() => { sub = f.Folders; }, "WalkFolders " + here);
                    WalkFolders(sub, here, ref total, cap, depth + 1, sb);
                }
                catch (Exception ex)
                {
                    // Whatever ComRetry couldn't resolve (exhausted retries,
                    // or a non-transient HResult it rethrows immediately)
                    // shouldn't abort siblings that are perfectly reachable.
                    DebugLog.WriteException("WalkFolders " + here, ex);
                    sb.AppendLine("  ! could not list subfolders of " + here + ": " + ex.Message);
                }
            }
        }
    }
}
