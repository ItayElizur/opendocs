using System.Text.Json;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure JSON conversion helpers shared by the Word/Excel/PowerPoint tool
    /// layers.
    /// </summary>
    public static class JsonUtil
    {
        // Any JsonValueKind other than String/Number/True/False (including
        // Null, Array, and Object) falls through to null - e.g. a nested
        // array passed as a cell value silently lands as an empty cell rather
        // than throwing. Pinned as-is; not a Phase 0 behavior change.
        public static object JsonValueToObject(JsonElement v)
        {
            switch (v.ValueKind)
            {
                case JsonValueKind.String: return v.GetString();
                case JsonValueKind.Number: return v.GetDouble();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                default: return null;
            }
        }
    }
}
