using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // One search_contacts match. FullName and DisplayName are tracked
    // separately (both sourced from the same EWS NameResolution/Contact, see
    // OutlookEws.cs's FullNameFrom/DisplayNameFrom) because they can disagree:
    // a directory's DisplayName is often org-formatted (e.g. "Dept/Unit/Title")
    // and can look nothing like the person's actual name, while FullName
    // (GivenName + Surname) is what a query like "John Doe" will actually
    // match against. Confirmed live 2026-10-01: an agent saw only a
    // display-name-only result that didn't resemble the query and wrongly
    // concluded the search had failed, even though the match was correct.
    public struct ContactMatch
    {
        public string FullName;
        public string DisplayName;
        public string Email;
    }

    // Pure formatting + dedupe for search_contacts' output, factored out of the
    // Outlook add-in so it stays unit-testable without the Outlook PIA or a live
    // Exchange server.
    public static class ContactSearchFormat
    {
        public static string Format(IReadOnlyList<ContactMatch> matches, string query, int limit)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var kept = new List<ContactMatch>();

            foreach (ContactMatch m in matches)
            {
                if (kept.Count >= limit) break;

                string email = m.Email ?? "";
                // Dedupe by address (case/space-insensitive) when there is one,
                // else by whatever name is available - mirrors the original
                // accumulator's behavior.
                string key = email.Length > 0 ? email.Trim().ToLowerInvariant() : (m.FullName ?? m.DisplayName ?? "");
                if (key.Length == 0 || !seen.Add(key)) continue;

                kept.Add(m);
            }

            if (kept.Count == 0)
                return "No contacts matched \"" + query + "\".";

            var sb = new StringBuilder();
            foreach (ContactMatch m in kept)
                sb.AppendLine(FormatLine(m));
            return sb.ToString();
        }

        private static string FormatLine(ContactMatch m)
        {
            string full = m.FullName ?? "";
            string display = m.DisplayName ?? "";
            string email = m.Email ?? "";
            // FullName leads when available (it's what the caller's query
            // actually matches against); DisplayName is a parenthetical hint,
            // shown only when it adds information beyond the name already
            // shown. A match with no name at all (shouldn't happen in
            // practice - OutlookEws's DisplayNameFrom always falls back to
            // the email itself) still renders without throwing.
            string primary = full.Length > 0 ? full : (display.Length > 0 ? display : email);
            bool showDisplaySeparately = full.Length > 0 && display.Length > 0 &&
                !string.Equals(full, display, StringComparison.OrdinalIgnoreCase);

            string line = "- " + primary;
            if (showDisplaySeparately) line += " (" + display + ")";
            if (email.Length > 0 && primary != email) line += " <" + email + ">";
            return line;
        }
    }
}
