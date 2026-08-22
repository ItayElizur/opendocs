using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class ToolProtocolTests
{
    [Fact]
    public void ParseToolCall_ExtractsFields()
    {
        string json = "{\"kind\":\"tool-call\",\"requestId\":\"abc\",\"toolName\":\"insert_content\",\"input\":{\"text\":\"hi\"}}";
        var (requestId, toolName, input) = ToolProtocol.ParseToolCall(json);
        Assert.Equal("abc", requestId);
        Assert.Equal("insert_content", toolName);
        Assert.Equal("hi", input.GetProperty("text").GetString());
    }

    [Fact]
    public void ParseToolCall_ThrowsOnWrongKind()
    {
        string json = "{\"kind\":\"tool-result\",\"requestId\":\"abc\"}";
        Assert.Throws<System.FormatException>(() => ToolProtocol.ParseToolCall(json));
    }

    [Fact]
    public void SerializeToolResult_RoundTrips()
    {
        var result = new ToolResult { Output = "done", IsError = false, Mutated = true, Summary = "insert_content" };
        string json = ToolProtocol.SerializeToolResult("abc", result);
        using (JsonDocument doc = JsonDocument.Parse(json))
        {
            var root = doc.RootElement;
            Assert.Equal("tool-result", root.GetProperty("kind").GetString());
            Assert.Equal("abc", root.GetProperty("requestId").GetString());
            Assert.Equal("done", root.GetProperty("output").GetString());
            Assert.True(root.GetProperty("mutated").GetBoolean());
        }
    }
}
