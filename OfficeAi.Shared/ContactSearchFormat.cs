using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // One search_contacts match. FullName and DisplayName are tracked
    // separately because they can disagree - a directory DisplayName is often
    // org-formatted (e.g. "Dept/Unit/Title") while FullName is what a name
    // query actually matches against. See ContactSearchFormat.cs.md.
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
                // else by whatever name is available.
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
            // FullName leads when available (what the query matches against);
            // DisplayName is a parenthetical hint shown only when it adds
            // information. See ContactSearchFormat.cs.md.
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
