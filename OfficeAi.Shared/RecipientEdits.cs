using System;
using System.Collections.Generic;

namespace OfficeAi.Shared
{
    // One recipient of the item open in an Outlook window. Kind is "to", "cc", "bcc", or for a meeting "required",
    // "optional", "resource" or "organizer".
    public sealed class RecipientEntry
    {
        public string Name;
        public string Address;
        public string Kind;
    }

    // What set_current_fields does to an item's recipient list, worked out before any COM call: indexes are into the
    // current list (0-based). Apply ChangeKind, then Remove (highest index first), then Add.
    public sealed class RecipientEditPlan
    {
        public readonly List<int> Remove = new List<int>();
        public readonly List<KeyValuePair<int, string>> ChangeKind = new List<KeyValuePair<int, string>>();
        public readonly List<KeyValuePair<string, string>> Add = new List<KeyValuePair<string, string>>(); // text -> kind
        public readonly List<string> NotFound = new List<string>();
    }

    public static class RecipientEdits
    {
        // "Dana Levi <dana@x.com>" -> "dana@x.com"; anything else trimmed as given.
        public static string AddressText(string text)
        {
            string s = (text ?? "").Trim();
            int lt = s.LastIndexOf('<');
            if (lt >= 0 && s.EndsWith(">", StringComparison.Ordinal)) s = s.Substring(lt + 1, s.Length - lt - 2).Trim();
            return s;
        }

        private static bool Matches(RecipientEntry e, string query)
        {
            string q = AddressText(query);
            string whole = (query ?? "").Trim();
            return q.Length > 0 &&
                   (string.Equals(e.Address, q, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(e.Name, q, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(e.Name, whole, StringComparison.OrdinalIgnoreCase));
        }

        // Removals first, then additions. Adding someone who is already on the item under another kind moves them
        // ("make Dana optional"); adding someone already under that kind changes nothing. The organizer is never touched.
        public static RecipientEditPlan Plan(IList<RecipientEntry> current, IList<KeyValuePair<string, string>> add, IList<string> remove)
        {
            var plan = new RecipientEditPlan();
            var removed = new HashSet<int>();
            foreach (string q in remove ?? new string[0])
            {
                bool found = false;
                for (int i = 0; i < current.Count; i++)
                {
                    if (current[i].Kind == "organizer" || !Matches(current[i], q)) continue;
                    found = true;
                    if (removed.Add(i)) plan.Remove.Add(i);
                }
                if (!found) plan.NotFound.Add(q);
            }

            var kindOf = new Dictionary<int, string>();
            foreach (KeyValuePair<string, string> a in add ?? new KeyValuePair<string, string>[0])
            {
                string text = (a.Key ?? "").Trim();
                if (text.Length == 0) continue;
                bool onItem = false;
                for (int i = 0; i < current.Count; i++)
                {
                    if (removed.Contains(i) || current[i].Kind == "organizer" || !Matches(current[i], text)) continue;
                    onItem = true;
                    if (current[i].Kind != a.Value) kindOf[i] = a.Value;
                    else kindOf.Remove(i);
                }
                if (onItem) continue;
                string key = AddressText(text);
                int at = plan.Add.FindIndex(p => string.Equals(AddressText(p.Key), key, StringComparison.OrdinalIgnoreCase));
                if (at >= 0) plan.Add[at] = new KeyValuePair<string, string>(plan.Add[at].Key, a.Value);
                else plan.Add.Add(new KeyValuePair<string, string>(text, a.Value));
            }
            foreach (KeyValuePair<int, string> k in kindOf) plan.ChangeKind.Add(k);
            plan.Remove.Sort((x, y) => y.CompareTo(x)); // highest first, so earlier indexes stay valid while removing
            return plan;
        }
    }
}
