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
        public static void ValidateRequired(string kind, JsonElement element, IReadOnlyDictionary<string, string[]> requiredFields, string noun)
        {
            string[] required;
            if (!requiredFields.TryGetValue(kind, out required)) return;
            foreach (string f in required)
            {
                if (!element.TryGetProperty(f, out _))
                    throw new System.ArgumentException(noun + " \"" + kind + "\" is missing required field \"" + f + "\".");
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
