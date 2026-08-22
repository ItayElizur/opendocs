using System.Text.Json;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static class ExcelTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
        }
    }
}
