using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class JsonUtilTests
{
    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement;

    [Fact]
    public void JsonValueToObject_String_ReturnsString()
    {
        Assert.Equal("hello", JsonUtil.JsonValueToObject(Parse("\"hello\"")));
    }

    [Fact]
    public void JsonValueToObject_Number_ReturnsDouble()
    {
        Assert.Equal(42.5, JsonUtil.JsonValueToObject(Parse("42.5")));
    }

    [Fact]
    public void JsonValueToObject_True_ReturnsBoolTrue()
    {
        Assert.Equal(true, JsonUtil.JsonValueToObject(Parse("true")));
    }

    [Fact]
    public void JsonValueToObject_False_ReturnsBoolFalse()
    {
        Assert.Equal(false, JsonUtil.JsonValueToObject(Parse("false")));
    }

    [Fact]
    public void JsonValueToObject_Null_ReturnsNull()
    {
        Assert.Null(JsonUtil.JsonValueToObject(Parse("null")));
    }

    [Fact]
    public void JsonValueToObject_Array_ReturnsNull()
    {
        // Pinned catch-all: a nested array silently lands as an empty cell
        // rather than throwing - not a Phase 0 behavior change.
        Assert.Null(JsonUtil.JsonValueToObject(Parse("[1,2,3]")));
    }

    [Fact]
    public void JsonValueToObject_Object_ReturnsNull()
    {
        Assert.Null(JsonUtil.JsonValueToObject(Parse("{\"a\":1}")));
    }
}
