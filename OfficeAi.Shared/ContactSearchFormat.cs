using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // Pure formatting + dedupe for search_contacts' output, factored out of the
    // Outlook add-in so it stays unit-testable without the Outlook PIA or a live
    // Exchange server. The output shape is unchanged from the pre-EWS
    // implementation: one "- Name <addr>" line per match (or "- Name" when the
    // match carries no address), or a single "No contacts matched ..." line.
    public static class ContactSearchFormat
    {
        public static string Format(IReadOnlyList<KeyValuePair<string, string>> matches, string query, int limit)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var kept = new List<KeyValuePair<string, string>>();

            foreach (KeyValuePair<string, string> m in matches)
            {
                if (kept.Count >= limit) break;

                string name = m.Key ?? "";
                string email = m.Value ?? "";
                // Dedupe by address (case/space-insensitive) when there is one,
                // else by display name - mirrors the original accumulator.
                string key = email.Length > 0 ? email.Trim().ToLowerInvariant() : name;
                if (key.Length == 0 || !seen.Add(key)) continue;

                kept.Add(new KeyValuePair<string, string>(name, email));
            }

            if (kept.Count == 0)
                return "No contacts matched \"" + query + "\".";

            var sb = new StringBuilder();
            foreach (KeyValuePair<string, string> kv in kept)
                sb.AppendLine(string.IsNullOrEmpty(kv.Value) ? "- " + kv.Key : "- " + kv.Key + " <" + kv.Value + ">");
            return sb.ToString();
        }
    }
}
