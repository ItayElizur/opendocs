using System.Collections.Generic;
using System.Text.Json;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Shared JSON-argument validation for apply_commands (Word) and
    /// propose_operations (Excel). The RequiredFields dictionaries themselves
    /// stay in each app's own file - they are app-specific data, each
    /// documented as mirroring that app's entry.ts schema, and that pairing
    /// must not be split up.
    /// </summary>
    public static class ToolArgs
    {
        /// <summary>
        /// Reads an optional boolean flag leniently: models sometimes send "true" or 1 instead of true. Only an
        /// explicit yes counts (true, "true"/"yes"/"1" in any case, or a non-zero number); anything else, including a
        /// missing property or a non-object input, is false.
        /// </summary>
        public static bool IsTrue(JsonElement input, string name)
        {
            if (input.ValueKind != JsonValueKind.Object || !input.TryGetProperty(name, out var e)) return false;
            switch (e.ValueKind)
            {
                case JsonValueKind.True: return true;
                case JsonValueKind.Number: return e.TryGetDouble(out var d) && d != 0;
                case JsonValueKind.String:
                    string s = (e.GetString() ?? "").Trim().ToLowerInvariant();
                    return s == "true" || s == "yes" || s == "1";
                default: return false;
            }
        }

        public static void ValidateRequired(string kind, JsonElement element, IReadOnlyDictionary<string, string[]> requiredFields, string noun,
            IReadOnlyDictionary<string, string[]> nonNullFields = null)
        {
            string[] required;
            if (!requiredFields.TryGetValue(kind, out required)) return;
            foreach (string f in required)
            {
                JsonElement value;
                if (!element.TryGetProperty(f, out value))
                    throw new System.ArgumentException(noun + " \"" + kind + "\" is missing required field \"" + f + "\".");

                // Null is a MEANINGFUL value for some fields (Excel's set_cell
                // uses it to clear a cell), so it is only rejected where the
                // owning app has opted in - never globally.
                if (value.ValueKind != JsonValueKind.Null || nonNullFields == null) continue;
                string[] nonNull;
                if (nonNullFields.TryGetValue(kind, out nonNull) && System.Array.IndexOf(nonNull, f) >= 0)
                    throw new System.ArgumentException(noun + " \"" + kind + "\" requires a non-null value for field \"" + f + "\".");
            }
        }

        public static void ValidateKnownFields(HashSet<string> fields, HashSet<string> known, string commandName)
        {
            foreach (string f in fields)
            {
                if (!known.Contains(f))
                    throw new System.ArgumentException(commandName + ": unknown style field '" + f + "'. Valid: " + string.Join(", ", known) + ".");
            }
        }
    }
}
