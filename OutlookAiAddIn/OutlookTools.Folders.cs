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
            WalkFolders(Ns.Folders, "", ref total, 800, 0, sb);
            if (total == 0) return new ToolResult { Output = "No mail folders found.", Summary = "list_folders" };
            return new ToolResult { Output = sb.ToString(), Summary = "list_folders" };
        }

        private static void WalkFolders(Outlook.Folders folders, string path, ref int total, int cap, int depth, StringBuilder sb)
        {
            if (folders == null || depth > 8 || total >= cap) return;
            foreach (Outlook.Folder f in folders)
            {
                if (total >= cap) return;
                string here = path.Length == 0 ? f.Name : path + "\\" + f.Name;
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
                WalkFolders(f.Folders, here, ref total, cap, depth + 1, sb);
            }
        }
    }
}
