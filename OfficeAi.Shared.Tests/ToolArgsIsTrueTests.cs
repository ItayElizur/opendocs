using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class ToolArgsIsTrueTests
{
    private static JsonElement Json(string s)
    {
        return JsonDocument.Parse(s).RootElement.Clone();
    }

    [Theory]
    [InlineData("{\"saveAs\":true}")]
    [InlineData("{\"saveAs\":\"true\"}")]
    [InlineData("{\"saveAs\":\" TRUE \"}")]
    [InlineData("{\"saveAs\":\"yes\"}")]
    [InlineData("{\"saveAs\":\"1\"}")]
    [InlineData("{\"saveAs\":1}")]
    [InlineData("{\"saveAs\":2.0}")]
    public void IsTrue_AcceptsExplicitYes(string json)
    {
        Assert.True(ToolArgs.IsTrue(Json(json), "saveAs"));
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"saveAs\":false}")]
    [InlineData("{\"saveAs\":\"false\"}")]
    [InlineData("{\"saveAs\":\"\"}")]
    [InlineData("{\"saveAs\":0}")]
    [InlineData("{\"saveAs\":null}")]
    [InlineData("{\"saveAs\":[true]}")]
    [InlineData("{\"other\":true}")]
    [InlineData("null")]
    [InlineData("[]")]
    public void IsTrue_EverythingElseIsFalse(string json)
    {
        Assert.False(ToolArgs.IsTrue(Json(json), "saveAs"));
    }
}
