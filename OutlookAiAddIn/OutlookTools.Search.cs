using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Native query first: Items.Sort then Items.Restrict("@SQL=" + DASL).
        // A capped linear scan is the fallback only when Restrict throws on a
        // filter it cannot parse - never the default path.
        private static ToolResult SearchEmails(JsonElement input)
        {
            string query = Str(input, "query", "");
            string folderName = Str(input, "folder", "inbox");
            int limit = Math.Max(1, Int(input, "limit", 20));
            bool unreadOnly = Bool(input, "unread_only", false);
            DateTime? start = DateArg(input, "start_date");
            DateTime? end = DateArg(input, "end_date");
            string sender = Str(input, "sender", null);
            string recipient = Str(input, "recipient", null);

            Outlook.Folder folder = ResolveFolder(folderName);
            Outlook.Items items = folder.Items;
            items.Sort("[ReceivedTime]", true);

            string dasl = BuildSearchDasl(query, start, end, sender);
            Outlook.Items filtered = null;
            bool usedRestrict = false;
            if (dasl.Length > 0)
            {
                try
                {
                    filtered = items.Restrict("@SQL=" + dasl);
                    usedRestrict = true;
                }
                catch (Exception ex)
                {
                    DebugLog.WriteException("SearchEmails Restrict", ex);
                }
            }

            var matches = new List<Outlook.MailItem>();
            int scanned = 0;
            const int scanCap = 500;
            Outlook.Items source = filtered ?? items;
            foreach (object obj in source)
            {
                Outlook.MailItem m = obj as Outlook.MailItem;
                if (m == null) continue;
                scanned++;

                if (!usedRestrict && dasl.Length > 0 && !MatchesClientSide(m, query, start, end, sender))
                {
                    if (scanned >= scanCap) break;
                    continue;
                }
                if (unreadOnly && !m.UnRead) { if (scanned >= scanCap) break; continue; }
                if (!string.IsNullOrEmpty(recipient) && !MatchesRecipient(m, recipient)) { if (scanned >= scanCap) break; continue; }

                matches.Add(m);
                if (matches.Count >= limit) break;
                if (scanned >= scanCap) break;
            }

            if (matches.Count == 0)
                return new ToolResult { Output = "No matching messages in " + folder.Name + ".", Summary = "search_emails" };

            var sb = new StringBuilder();
            sb.AppendLine((usedRestrict ? "Server-filtered" : "Scanned") + " search in " + folder.Name + ":");
            foreach (Outlook.MailItem m in matches)
            {
                sb.AppendLine("- message_id: " + m.EntryID);
                sb.AppendLine("  subject: " + (m.Subject ?? ""));
                sb.AppendLine("  from: " + (m.SenderName ?? ""));
                try { sb.AppendLine("  received: " + Iso(m.ReceivedTime) + "  unread: " + m.UnRead); } catch { }
            }
            return new ToolResult { Output = sb.ToString(), Summary = "search_emails" };
        }

        // Pushes an already-computed search filter into Outlook's own Explorer
        // window (the native Instant Search UI) instead of only returning
        // results to the model - the "show, don't just tell" counterpart to
        // SearchEmails. Reuses BuildSearchDasl verbatim; no new filter logic.
        //
        // _Explorer.Search(string Query, OlSearchScope SearchScope) signature
        // confirmed via .NET reflection against the actually-referenced
        // Microsoft.Office.Interop.Outlook 15.0.0.0 PIA in this repo (not
        // assumed). The "@SQL=" DASL prefix used here is the same one
        // SearchEmails already proves works with Items.Restrict, and is a
        // documented Outlook automation pattern for Explorer.Search's Query
        // parameter (Advanced Find generates exactly this string) - but this
        // specific call has NOT been exercised against a live Outlook session
        // in this environment (no mailbox profile available here). First real
        // use should confirm the DASL string is accepted as-is; if not, fall
        // back to passing `query` alone (plain AQS free-text).
        private static ToolResult ApplySearch(JsonElement input)
        {
            string query = Str(input, "query", "");
            string folderName = Str(input, "folder", "inbox");
            DateTime? start = DateArg(input, "start_date");
            DateTime? end = DateArg(input, "end_date");
            string sender = Str(input, "sender", null);

            Outlook.Folder folder = ResolveFolder(folderName);
            Outlook.Explorer explorer = App.ActiveExplorer();
            if (explorer == null)
                return new ToolResult { Output = "No active Outlook window to apply the search to.", IsError = true, Summary = "apply_search" };

            explorer.CurrentFolder = folder;

            string dasl = BuildSearchDasl(query, start, end, sender);
            if (dasl.Length == 0)
                return new ToolResult { Output = "Showed " + folder.Name + " (no filter criteria given).", Summary = "apply_search" };

            explorer.Search("@SQL=" + dasl, Outlook.OlSearchScope.olSearchScopeCurrentFolder);
            return new ToolResult { Output = "Applied the search to " + folder.Name + " in Outlook - the user can see the results now.", Summary = "apply_search" };
        }

        private static bool MatchesClientSide(Outlook.MailItem m, string query, DateTime? start, DateTime? end, string sender)
        {
            if (!string.IsNullOrEmpty(query))
            {
                string q = query.ToLowerInvariant();
                bool hit = (m.Subject ?? "").ToLowerInvariant().Contains(q) || (m.Body ?? "").ToLowerInvariant().Contains(q);
                if (!hit) return false;
            }
            try
            {
                if (start.HasValue && m.ReceivedTime < start.Value) return false;
                if (end.HasValue && m.ReceivedTime > end.Value) return false;
            }
            catch { }
            if (!string.IsNullOrEmpty(sender))
            {
                string s = sender.ToLowerInvariant();
                string name = (m.SenderName ?? "").ToLowerInvariant();
                string addr = (m.SenderEmailAddress ?? "").ToLowerInvariant();
                if (!name.Contains(s) && !addr.Contains(s)) return false;
            }
            return true;
        }

        private static bool MatchesRecipient(Outlook.MailItem m, string recipient)
        {
            string target = recipient.Trim().ToLowerInvariant();
            try
            {
                foreach (Outlook.Recipient r in m.Recipients)
                {
                    if ((r.Name ?? "").ToLowerInvariant().Contains(target)) return true;
                    string addr = SmtpOf(r.AddressEntry);
                    if (!string.IsNullOrEmpty(addr) && addr.ToLowerInvariant().Contains(target)) return true;
                }
            }
            catch { }
            return false;
        }
    }
}
