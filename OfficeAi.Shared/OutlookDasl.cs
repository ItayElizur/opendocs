using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeAi.Shared
{
    // Pure, COM-free DASL query construction for the Outlook add-in's
    // search_emails tool. Lives here (not in OutlookAiAddIn) so it can be unit
    // tested without the Outlook PIA. The add-in pushes the returned string into
    // Items.Restrict("@SQL=" + dasl) - a native server-side filter - rather than
    // iterating Items, per the "native query APIs over manual scans" rule.
    public static class OutlookDasl
    {
        public static string BuildSearchFilter(string query, DateTime? start, DateTime? end, string sender)
        {
            var parts = new List<string>();

            if (!string.IsNullOrEmpty(query))
            {
                string q = query.Replace("'", "''");
                parts.Add("(\"urn:schemas:httpmail:subject\" LIKE '%" + q + "%'" +
                          " OR \"urn:schemas:httpmail:textdescription\" LIKE '%" + q + "%')");
            }
            if (start.HasValue)
                parts.Add("\"urn:schemas:httpmail:datereceived\" >= '" +
                          start.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) + "'");
            if (end.HasValue)
                parts.Add("\"urn:schemas:httpmail:datereceived\" <= '" +
                          end.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) + "'");
            if (!string.IsNullOrEmpty(sender))
            {
                string s = sender.Replace("'", "''");
                parts.Add("(\"urn:schemas:httpmail:fromemail\" = '" + s + "'" +
                          " OR \"urn:schemas:httpmail:fromname\" LIKE '%" + s + "%')");
            }

            return string.Join(" AND ", parts);
        }
    }
}
