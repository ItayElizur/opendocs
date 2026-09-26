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
        // query/folder/start_date/end_date/sender are the common arg shape
        // shared by search_emails and apply_search (the same filter criteria,
        // applied via two different Outlook query engines - see
        // BuildSearchAqs's comment). Factored out so a future change to how
        // these are read/defaulted can't drift out of sync between the two.
        private struct SearchArgs
        {
            public string Query;
            public string FolderName;
            public DateTime? Start;
            public DateTime? End;
            public string Sender;
        }

        private static SearchArgs ReadSearchArgs(JsonElement input)
        {
            return new SearchArgs
            {
                Query = Str(input, "query", ""),
                FolderName = Str(input, "folder", "inbox"),
                Start = DateArg(input, "start_date"),
                End = DateArg(input, "end_date"),
                Sender = Str(input, "sender", null),
            };
        }

        // Native query first: Items.Sort then Items.Restrict("@SQL=" + DASL).
        // A capped linear scan is the fallback only when Restrict throws on a
        // filter it cannot parse - never the default path.
        private static ToolResult SearchEmails(JsonElement input)
        {
            SearchArgs a = ReadSearchArgs(input);
            string query = a.Query;
            string folderName = a.FolderName;
            DateTime? start = a.Start;
            DateTime? end = a.End;
            string sender = a.Sender;
            int limit = Math.Max(1, Int(input, "limit", 20));
            bool unreadOnly = Bool(input, "unread_only", false);
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
        // SearchEmails.
        //
        // _Explorer.Search(string Query, OlSearchScope SearchScope) signature
        // confirmed via .NET reflection against the actually-referenced
        // Microsoft.Office.Interop.Outlook 15.0.0.0 PIA in this repo. Does NOT
        // take a DASL/SQL filter, despite the "@SQL=" prefix this file
        // originally tried here (copied from SearchEmails' Items.Restrict
        // usage): live testing (2026-09-26) proved that string doesn't throw,
        // it just silently matches nothing, because Explorer.Search goes
        // through Windows Instant Search - a different query engine from
        // Items.Restrict's direct MAPI table access - and doesn't honor
        // urn:schemas:httpmail:* property URNs. It DOES take the same
        // Advanced Query Syntax (AQS) the Outlook search box itself accepts -
        // confirmed via https://support.microsoft.com/en-us/outlook/search-mail-and-people-in-outlook-com
        // and the user's own live "Advanced Search Options" list, which is
        // what BuildSearchAqs below builds: `From:value` for sender,
        // `Received:MM/DD/YYYY..MM/DD/YYYY` for a date range (the documented
        // two-dot syntax - deliberately NOT `>=`/`<=`, which that same page
        // does not confirm exists for this property). Plain free text with
        // no prefix (already live-tested and confirmed working) covers the
        // subject/body query term.
        private static string BuildSearchAqs(string query, DateTime? start, DateTime? end, string sender)
        {
            var terms = new List<string>();
            if (!string.IsNullOrEmpty(query)) terms.Add(query);
            if (!string.IsNullOrEmpty(sender))
                terms.Add("From:" + (sender.Contains(" ") ? "\"" + sender + "\"" : sender));
            if (start.HasValue || end.HasValue)
            {
                // Live testing (2026-09-26) proved hardcoding MM/dd/yyyy (US
                // order) was wrong: classic Outlook's search box parses typed
                // dates using the DEVICE's Windows regional format, not a
                // fixed order - on a day-first locale, "09/26/2036" reads as
                // day=09/month=26, an invalid month, so the whole Received:
                // clause silently failed to parse and the search returned 0
                // results (no error - same silent-failure shape as the
                // original "@SQL=" DASL bug this method replaced). Formatting
                // with CurrentCulture instead of InvariantCulture matches
                // whatever order this machine's own regional settings use,
                // which is exactly what Outlook's own parser reads back.
                string lo = (start ?? new DateTime(1900, 1, 1)).ToString("d", CultureInfo.CurrentCulture);
                string hi = (end ?? DateTime.Today.AddYears(10)).ToString("d", CultureInfo.CurrentCulture);
                terms.Add("Received:" + lo + ".." + hi);
            }
            return string.Join(" ", terms);
        }

        // OlSearchScope (confirmed via reflection): CurrentFolder=0,
        // AllFolders=1, AllOutlookItems=2, Subfolders=3, CurrentStore=4.
        // Default stays CurrentFolder - the one value actually exercised in
        // the live test that validated this AQS approach in the first place;
        // the wider scopes are unverified beyond being documented, valid
        // enum members. Unrecognized non-empty input (e.g. a folder name
        // passed here by mistake instead of in `folder`) throws rather than
        // silently falling back to CurrentFolder - a live test hit exactly
        // that mix-up and got a confusingly quiet "searched the wrong scope"
        // instead of a clear error.
        private static readonly string[] ValidScopes = { "current_folder", "subfolders", "mailbox", "current_mailbox", "all_mailboxes", "all_folders" };

        private static Outlook.OlSearchScope ParseScope(string s)
        {
            string v = (s ?? "").Trim().ToLowerInvariant();
            switch (v)
            {
                case "": return Outlook.OlSearchScope.olSearchScopeCurrentFolder;
                case "current_folder": return Outlook.OlSearchScope.olSearchScopeCurrentFolder;
                case "subfolders": return Outlook.OlSearchScope.olSearchScopeSubfolders;
                case "mailbox":
                case "current_mailbox": return Outlook.OlSearchScope.olSearchScopeCurrentStore;
                case "all_mailboxes": return Outlook.OlSearchScope.olSearchScopeAllOutlookItems;
                case "all_folders": return Outlook.OlSearchScope.olSearchScopeAllFolders;
                default:
                    throw new ArgumentException("Unknown scope \"" + s + "\". Valid values: " + string.Join(", ", ValidScopes) +
                                                 ". To search a specific folder, use the \"folder\" argument instead.");
            }
        }

        private static ToolResult ApplySearch(JsonElement input)
        {
            SearchArgs a = ReadSearchArgs(input);
            string query = a.Query;
            string folderName = a.FolderName;
            DateTime? start = a.Start;
            DateTime? end = a.End;
            string sender = a.Sender;
            Outlook.OlSearchScope scope = ParseScope(Str(input, "scope", "current_folder"));

            Outlook.Folder folder = ResolveFolder(folderName);
            Outlook.Explorer explorer = App.ActiveExplorer();
            if (explorer == null)
                return new ToolResult { Output = "No active Outlook window to apply the search to.", IsError = true, Summary = "apply_search" };

            explorer.CurrentFolder = folder;

            string aqs = BuildSearchAqs(query, start, end, sender);
            if (aqs.Length == 0)
                return new ToolResult { Output = "Showed " + folder.Name + " (no filter criteria given).", Summary = "apply_search" };

            explorer.Search(aqs, scope);
            return new ToolResult { Output = "Applied the search to " + folder.Name + " in Outlook (scope: " + scope + ") - the user can see the results now.", Summary = "apply_search" };
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
