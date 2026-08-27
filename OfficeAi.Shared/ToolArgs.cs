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
