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
        private static ToolResult SetEventCategories(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string categories = (Str(input, "categories", "") ?? "").Trim();

            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "set_event_categories" };

            string[] props = { "Categories" };
            object[] before = ReadProps(appt, props);
            appt.Categories = categories;
            appt.Save();
            RecordSnapshot(mbxKey, "set_event_categories", appt, appt.Subject ?? "", props, before);

            string result = categories.Length == 0
                ? "Cleared color tags on: " + (appt.Subject ?? "")
                : "Tagged \"" + (appt.Subject ?? "") + "\" with: " + appt.Categories;
            return new ToolResult { Output = result, Mutated = true, Summary = "set_event_categories" };
        }

        // Creates a new color tag, or recolors an existing one, in the master
        // category list - the same list Categorize/list_color_categories use.
        private static ToolResult SetCategoryColor(string mbxKey, JsonElement input)
        {
            string name = ReqStr(input, "name").Trim();
            Outlook.OlCategoryColor color = ParseColor(ReqStr(input, "color"));

            Outlook.Categories cats = Ns.Categories;
            Outlook.Category existing = cats[name];
            if (existing != null)
            {
                Outlook.OlCategoryColor prior = existing.Color;
                existing.Color = color;
                RecordCategoryColor(mbxKey, existing.Name, true, prior, color);
                return new ToolResult { Output = "Updated \"" + existing.Name + "\" to " + ColorName(color) + ".", Mutated = true, Summary = "set_category_color" };
            }

            Outlook.Category created = cats.Add(name, color);
            RecordCategoryColor(mbxKey, created.Name, false, color, created.Color);
            return new ToolResult { Output = "Created color tag \"" + created.Name + "\" (" + ColorName(created.Color) + ").", Mutated = true, Summary = "set_category_color" };
        }

        // Friendly names <-> OlBusyStatus - exactly what the calendar's "Show
        // As" dropdown controls (Free/Tentative/Busy/Out of Office/Working
        // Elsewhere). Nothing to do with Categories/color tags above; kept
        // here for the same reason ColorByName is - a small friendly-name
        // map for an Outlook-PIA-only enum that OfficeAi.Shared can't see.
        // Member names confirmed via .NET reflection against the referenced
        // Microsoft.Office.Interop.Outlook 15.0.0.0 PIA.
        private static readonly Dictionary<string, Outlook.OlBusyStatus> BusyStatusByName =
            new Dictionary<string, Outlook.OlBusyStatus>(StringComparer.OrdinalIgnoreCase)
            {
                { "free", Outlook.OlBusyStatus.olFree },
                { "tentative", Outlook.OlBusyStatus.olTentative },
                { "busy", Outlook.OlBusyStatus.olBusy },
                { "outofoffice", Outlook.OlBusyStatus.olOutOfOffice },
                { "workingelsewhere", Outlook.OlBusyStatus.olWorkingElsewhere },
            };

        private static readonly Dictionary<Outlook.OlBusyStatus, string> BusyStatusDisplayName =
            new Dictionary<Outlook.OlBusyStatus, string>
            {
                { Outlook.OlBusyStatus.olFree, "Free" },
                { Outlook.OlBusyStatus.olTentative, "Tentative" },
                { Outlook.OlBusyStatus.olBusy, "Busy" },
                { Outlook.OlBusyStatus.olOutOfOffice, "Out of Office" },
                { Outlook.OlBusyStatus.olWorkingElsewhere, "Working Elsewhere" },
            };

        private static string BusyStatusName(Outlook.OlBusyStatus s)
        {
            string name;
            return BusyStatusDisplayName.TryGetValue(s, out name) ? name : s.ToString();
        }

        // Unlike ParseColor, returns bool rather than throwing - an
        // unrecognized availability string should come back as a clean
        // IsError result (with the valid list), not a raw cast/parse
        // exception surfaced through the outer ExecuteAsync catch.
        private static bool TryParseBusyStatus(string s, out Outlook.OlBusyStatus status)
        {
            string key = (s ?? "").Trim().Replace(" ", "").Replace("_", "").Replace("-", "");
            return BusyStatusByName.TryGetValue(key, out status);
        }

        // Sets an event's "Show As" availability - purely local
        // (AppointmentItem.BusyStatus + .Save()), never .Send(), same risk
        // profile as set_event_categories/set_category_color above.
        private static ToolResult SetEventAvailability(JsonElement input)
        {
            string id = ReqStr(input, "event_id");
            string raw = ReqStr(input, "availability");

            Outlook.AppointmentItem appt = ItemById(id, null) as Outlook.AppointmentItem;
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to an appointment.", IsError = true, Summary = "set_event_availability" };

            Outlook.OlBusyStatus status;
            if (!TryParseBusyStatus(raw, out status))
            {
                return new ToolResult
                {
                    Output = "Unknown availability \"" + raw + "\". Valid values: " + string.Join(", ", BusyStatusDisplayName.Values) + ".",
                    IsError = true,
                    Summary = "set_event_availability",
                };
            }

            string oldName = BusyStatusName(appt.BusyStatus);
            appt.BusyStatus = status;
            appt.Save();

            return new ToolResult
            {
                Output = "Set \"" + (appt.Subject ?? "") + "\" to " + BusyStatusName(status) + " (was " + oldName + ").",
                Mutated = true,
                Summary = "set_event_availability",
            };
        }
    }
}
