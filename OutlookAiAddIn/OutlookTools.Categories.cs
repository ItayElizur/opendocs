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
        // Friendly names <-> OlCategoryColor. Kept here rather than in
        // OfficeAi.Shared because that project doesn't reference the Outlook
        // PIA (see ColorUtil.cs's own split-by-app rationale for RGB colors).
        private static readonly Dictionary<string, Outlook.OlCategoryColor> ColorByName =
            new Dictionary<string, Outlook.OlCategoryColor>(StringComparer.OrdinalIgnoreCase)
            {
                { "none", Outlook.OlCategoryColor.olCategoryColorNone },
                { "red", Outlook.OlCategoryColor.olCategoryColorRed },
                { "orange", Outlook.OlCategoryColor.olCategoryColorOrange },
                { "peach", Outlook.OlCategoryColor.olCategoryColorPeach },
                { "yellow", Outlook.OlCategoryColor.olCategoryColorYellow },
                { "green", Outlook.OlCategoryColor.olCategoryColorGreen },
                { "teal", Outlook.OlCategoryColor.olCategoryColorTeal },
                { "olive", Outlook.OlCategoryColor.olCategoryColorOlive },
                { "blue", Outlook.OlCategoryColor.olCategoryColorBlue },
                { "purple", Outlook.OlCategoryColor.olCategoryColorPurple },
                { "maroon", Outlook.OlCategoryColor.olCategoryColorMaroon },
                { "steel", Outlook.OlCategoryColor.olCategoryColorSteel },
                { "darksteel", Outlook.OlCategoryColor.olCategoryColorDarkSteel },
                { "gray", Outlook.OlCategoryColor.olCategoryColorGray },
                { "grey", Outlook.OlCategoryColor.olCategoryColorGray },
                { "darkgray", Outlook.OlCategoryColor.olCategoryColorDarkGray },
                { "darkgrey", Outlook.OlCategoryColor.olCategoryColorDarkGray },
                { "black", Outlook.OlCategoryColor.olCategoryColorBlack },
                { "darkred", Outlook.OlCategoryColor.olCategoryColorDarkRed },
                { "darkorange", Outlook.OlCategoryColor.olCategoryColorDarkOrange },
                { "darkpeach", Outlook.OlCategoryColor.olCategoryColorDarkPeach },
                { "darkyellow", Outlook.OlCategoryColor.olCategoryColorDarkYellow },
                { "darkgreen", Outlook.OlCategoryColor.olCategoryColorDarkGreen },
                { "darkteal", Outlook.OlCategoryColor.olCategoryColorDarkTeal },
                { "darkolive", Outlook.OlCategoryColor.olCategoryColorDarkOlive },
                { "darkblue", Outlook.OlCategoryColor.olCategoryColorDarkBlue },
                { "darkpurple", Outlook.OlCategoryColor.olCategoryColorDarkPurple },
                { "darkmaroon", Outlook.OlCategoryColor.olCategoryColorDarkMaroon },
            };

        private static readonly Dictionary<Outlook.OlCategoryColor, string> ColorDisplayName =
            new Dictionary<Outlook.OlCategoryColor, string>
            {
                { Outlook.OlCategoryColor.olCategoryColorNone, "None" },
                { Outlook.OlCategoryColor.olCategoryColorRed, "Red" },
                { Outlook.OlCategoryColor.olCategoryColorOrange, "Orange" },
                { Outlook.OlCategoryColor.olCategoryColorPeach, "Peach" },
                { Outlook.OlCategoryColor.olCategoryColorYellow, "Yellow" },
                { Outlook.OlCategoryColor.olCategoryColorGreen, "Green" },
                { Outlook.OlCategoryColor.olCategoryColorTeal, "Teal" },
                { Outlook.OlCategoryColor.olCategoryColorOlive, "Olive" },
                { Outlook.OlCategoryColor.olCategoryColorBlue, "Blue" },
                { Outlook.OlCategoryColor.olCategoryColorPurple, "Purple" },
                { Outlook.OlCategoryColor.olCategoryColorMaroon, "Maroon" },
                { Outlook.OlCategoryColor.olCategoryColorSteel, "Steel" },
                { Outlook.OlCategoryColor.olCategoryColorDarkSteel, "Dark Steel" },
                { Outlook.OlCategoryColor.olCategoryColorGray, "Gray" },
                { Outlook.OlCategoryColor.olCategoryColorDarkGray, "Dark Gray" },
                { Outlook.OlCategoryColor.olCategoryColorBlack, "Black" },
                { Outlook.OlCategoryColor.olCategoryColorDarkRed, "Dark Red" },
                { Outlook.OlCategoryColor.olCategoryColorDarkOrange, "Dark Orange" },
                { Outlook.OlCategoryColor.olCategoryColorDarkPeach, "Dark Peach" },
                { Outlook.OlCategoryColor.olCategoryColorDarkYellow, "Dark Yellow" },
                { Outlook.OlCategoryColor.olCategoryColorDarkGreen, "Dark Green" },
                { Outlook.OlCategoryColor.olCategoryColorDarkTeal, "Dark Teal" },
                { Outlook.OlCategoryColor.olCategoryColorDarkOlive, "Dark Olive" },
                { Outlook.OlCategoryColor.olCategoryColorDarkBlue, "Dark Blue" },
                { Outlook.OlCategoryColor.olCategoryColorDarkPurple, "Dark Purple" },
                { Outlook.OlCategoryColor.olCategoryColorDarkMaroon, "Dark Maroon" },
            };

        private static string ColorName(Outlook.OlCategoryColor c)
        {
            string name;
            return ColorDisplayName.TryGetValue(c, out name) ? name : c.ToString();
        }

        private static Outlook.OlCategoryColor ParseColor(string s)
        {
            string key = (s ?? "").Trim().Replace(" ", "").Replace("_", "").Replace("-", "");
            Outlook.OlCategoryColor c;
            if (ColorByName.TryGetValue(key, out c)) return c;
            throw new ArgumentException("Unknown color \"" + s + "\". Valid colors: " + string.Join(", ", ColorDisplayName.Values) + ".");
        }

        // Ns.Categories is the profile's master color-tag list (what "Categorize"
        // shows in the Outlook UI) - shared across mail, calendar, and tasks.
        private static ToolResult ListColorCategories(JsonElement input)
        {
            Outlook.Categories cats = Ns.Categories;
            if (cats.Count == 0)
                return new ToolResult { Output = "No color tags are defined yet. Use set_category_color to create one.", Summary = "list_color_categories" };

            var sb = new StringBuilder();
            for (int i = 1; i <= cats.Count; i++)
            {
                Outlook.Category cat = cats[i];
                sb.AppendLine("- name: " + cat.Name + "  color: " + ColorName(cat.Color));
            }
            return new ToolResult { Output = sb.ToString(), Summary = "list_color_categories" };
        }

        // Sets (or, with an empty string, clears) the color tag(s) on a calendar
        // event. Category names are shown as a colored block on the event in the
        // calendar grid. A name outside the master list (list_color_categories)
        // is auto-added by Outlook on Save with an arbitrary color - pass an
        // existing name, or call set_category_color first to control the color.
        private static ToolResult SetEventCategories(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string categories = (Str(input, "categories", "") ?? "").Trim();

            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "set_event_categories" };

            appt.Categories = categories;
            appt.Save();

            string result = categories.Length == 0
                ? "Cleared color tags on: " + (appt.Subject ?? "")
                : "Tagged \"" + (appt.Subject ?? "") + "\" with: " + appt.Categories;
            return new ToolResult { Output = result, Mutated = true, Summary = "set_event_categories" };
        }

        // Creates a new color tag, or recolors an existing one, in the master
        // category list - the same list Categorize/list_color_categories use.
        private static ToolResult SetCategoryColor(JsonElement input)
        {
            string name = ReqStr(input, "name").Trim();
            Outlook.OlCategoryColor color = ParseColor(ReqStr(input, "color"));

            Outlook.Categories cats = Ns.Categories;
            Outlook.Category existing = cats[name];
            if (existing != null)
            {
                existing.Color = color;
                return new ToolResult { Output = "Updated \"" + existing.Name + "\" to " + ColorName(color) + ".", Mutated = true, Summary = "set_category_color" };
            }

            Outlook.Category created = cats.Add(name, color);
            return new ToolResult { Output = "Created color tag \"" + created.Name + "\" (" + ColorName(created.Color) + ").", Mutated = true, Summary = "set_category_color" };
        }
    }
}
