using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeAi.Shared
{
    // The line Outlook adds to the model's prompt about who the user is, so it knows their own address without reading
    // a message sent to them.
    public static class UserIdentityPrompt
    {
        // name: Outlook's display name for the user (an IMAP profile often shows the address itself). addresses: the
        // user's own email addresses, primary first. Returns null when there is nothing to say.
        public static string Describe(string name, IEnumerable<string> addresses)
        {
            List<string> own = (addresses ?? Enumerable.Empty<string>())
                .Where(a => !string.IsNullOrWhiteSpace(a) && a.Contains("@"))
                .Select(a => a.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            string shownName = string.IsNullOrWhiteSpace(name) || own.Contains(name.Trim(), StringComparer.OrdinalIgnoreCase) ? null : name.Trim();
            if (own.Count == 0) return shownName == null ? null : "The user's name is " + shownName + ".";

            string who = shownName == null ? "The user's" : "The user is " + shownName + ", and their";
            string list = own.Count == 1 ? "email address is " + own[0] : "email addresses are " + string.Join(", ", own) + " (the first is the main one)";
            return who + " " + list + ". Use it for \"me\", \"my address\" and \"send it to myself\", and to tell the " +
                   "user's own messages and invitations from other people's.";
        }
    }
}
