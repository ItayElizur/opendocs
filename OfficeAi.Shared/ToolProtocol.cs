using System;
using System.Text.Json;

namespace OfficeAi.Shared
{
    public struct ToolResult
    {
        public string Output;
        public bool IsError;
        public bool Mutated;
        public string Summary;
    }

    public delegate ToolResult ToolExecutor(string toolName, JsonElement input);
    public delegate void OtherMessageHandler(string kind, JsonElement root);

    public static class ToolProtocol
    {
        public static (string requestId, string toolName, JsonElement input) ParseToolCall(string json)
        {
            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                JsonElement root = doc.RootElement.Clone();
                string kind = root.GetProperty("kind").GetString();
                if (kind != "tool-call")
                {
                    throw new FormatException("Expected kind=tool-call, got: " + kind);
                }
                string requestId = root.GetProperty("requestId").GetString();
                string toolName = root.GetProperty("toolName").GetString();
                JsonElement input = root.GetProperty("input").Clone();
                return (requestId, toolName, input);
            }
        }

        public static string SerializeToolResult(string requestId, ToolResult result)
        {
            return JsonSerializer.Serialize(new
            {
                kind = "tool-result",
                requestId,
                output = result.Output,
                isError = result.IsError,
                mutated = result.Mutated,
                summary = result.Summary,
            });
        }
    }
}
