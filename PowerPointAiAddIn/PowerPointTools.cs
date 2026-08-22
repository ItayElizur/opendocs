using System.Text.Json;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    // Stub: no PowerPoint-specific tools yet (Task 18 adds readers,
    // Task 19 adds mutation tools + editing-mode gating).
    public static class PowerPointTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
        }
    }
}
