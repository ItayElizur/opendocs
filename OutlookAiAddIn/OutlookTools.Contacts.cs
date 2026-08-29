using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        private static ToolResult SearchContacts(JsonElement input)
        {
            string query = ReqStr(input, "query");
            int limit = Math.Max(1, Int(input, "limit", 10));
            string folderName = Str(input, "folder", null);

            var results = new List<KeyValuePair<string, string>>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) Resolve the query against the GAL / configured address books
            //    (skipped when the caller scoped the search to one folder).
            if (folderName == null)
            {
                try
                {
                    Outlook.Recipient r = Ns.CreateRecipient(query);
                    r.Resolve();
                    if (r.Resolved)
                    {
                        string addr = SmtpOf(r.AddressEntry);
                        if (!string.IsNullOrEmpty(addr) && seen.Add(addr))
                            results.Add(new KeyValuePair<string, string>(r.Name ?? query, addr));
                    }
                }
                catch { }
            }

            // 2) Contact folders, each queried server-side via Restrict (DASL).
            //    Default: every contact folder across every store (custom
            //    folders, shared mailboxes, subfolders). `folder` narrows it
            //    to one named folder.
            string q = query.Replace("'", "''");
            string dasl = "@SQL=(\"urn:schemas:contacts:fileas\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:contacts:email1\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:contacts:email2\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:contacts:email3\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:contacts:givenName\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:contacts:sn\" LIKE '%" + q + "%')";

            var folders = new List<Outlook.Folder>();
            if (folderName != null)
            {
                folders.Add(ResolveContactFolder(folderName));
            }
            else
            {
                try { CollectContactFolders(Ns.Folders, folders, 0); } catch { }
                if (folders.Count == 0)
                    folders.Add((Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts));
            }

            foreach (Outlook.Folder folder in folders)
            {
                if (results.Count >= limit) break;
                try
                {
                    Outlook.Items found;
                    try { found = folder.Items.Restrict(dasl); }
                    catch { found = folder.Items; }

                    int scanned = 0;
                    foreach (object o in found)
                    {
                        if (results.Count >= limit || scanned++ > 500) break;
                        Outlook.ContactItem c = o as Outlook.ContactItem;
                        if (c == null) continue;
                        string email = c.Email1Address ?? c.Email2Address ?? c.Email3Address ?? "";
                        string name = c.FullName;
                        if (string.IsNullOrEmpty(name)) name = c.FileAs ?? "";
                        string key = !string.IsNullOrEmpty(email) ? email : name;
                        if (string.IsNullOrEmpty(key) || !seen.Add(key)) continue;
                        results.Add(new KeyValuePair<string, string>(name, email));
                    }
                }
                catch (Exception ex) { DebugLog.WriteException("SearchContacts folder " + folder.Name, ex); }
            }

            if (results.Count == 0)
                return new ToolResult { Output = "No contacts matched \"" + query + "\".", Summary = "search_contacts" };

            var sb = new StringBuilder();
            foreach (var kv in results)
                sb.AppendLine(string.IsNullOrEmpty(kv.Value) ? "- " + kv.Key : "- " + kv.Key + " <" + kv.Value + ">");
            return new ToolResult { Output = sb.ToString(), Summary = "search_contacts" };
        }

        private static void CollectContactFolders(Outlook.Folders folders, List<Outlook.Folder> into, int depth)
        {
            if (folders == null || depth > 8 || into.Count > 60) return;
            foreach (Outlook.Folder f in folders)
            {
                bool isContacts = false;
                try { isContacts = f.DefaultItemType == Outlook.OlItemType.olContactItem; } catch { }
                if (isContacts) into.Add(f);
                CollectContactFolders(f.Folders, into, depth + 1);
            }
        }

        private static Outlook.Folder ResolveContactFolder(string name)
        {
            if (WellKnownContacts(name)) return (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            var all = new List<Outlook.Folder>();
            CollectContactFolders(Ns.Folders, all, 0);
            foreach (Outlook.Folder f in all)
                if (string.Equals(f.Name, name.Trim(), StringComparison.OrdinalIgnoreCase)) return f;
            throw new ArgumentException("Contact folder not found: " + name);
        }

        private static bool WellKnownContacts(string name)
        {
            string n = (name ?? "").Trim().ToLowerInvariant();
            return n == "contacts" || n == "";
        }
    }
}
