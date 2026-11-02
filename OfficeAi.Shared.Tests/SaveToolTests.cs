using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class SaveToolTests
{
    private static JsonElement Json(string s)
    {
        return JsonDocument.Parse(s).RootElement.Clone();
    }

    private static ToolResult Run(string json, string path, out string calls)
    {
        string log = "";
        var result = SaveTool.Run(Json(json), path,
            () => { log += "[in-place]"; return @"C:\docs\a.docx"; },
            () => { log += "[dialog]"; return @"C:\docs\copy.docx"; });
        calls = log;
        return result;
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"saveAs\":false}")]
    [InlineData("{\"saveAs\":0}")]
    public void SavedFile_WithoutSaveAs_SavesInPlace(string json)
    {
        var r = Run(json, @"C:\docs", out string calls);
        Assert.Equal("[in-place]", calls);
        Assert.Equal(@"Saved C:\docs\a.docx.", r.Output);
        Assert.False(r.IsError);
        Assert.False(r.Mutated);
    }

    [Theory]
    [InlineData("{\"saveAs\":true}")]
    [InlineData("{\"saveAs\":1}")]
    public void SaveAsTrue_OpensTheDialog(string json)
    {
        var r = Run(json, @"C:\docs", out string calls);
        Assert.Equal("[dialog]", calls);
        Assert.Equal(@"Saved as C:\docs\copy.docx.", r.Output);
    }

    [Fact]
    public void NeverSavedFile_FallsThroughToTheDialog()
    {
        Run("{}", "", out string calls);
        Assert.Equal("[dialog]", calls);
    }

    [Fact]
    public void CancelledDialog_IsANonErrorCancelledResult()
    {
        var r = SaveTool.Run(Json("{\"saveAs\":true}"), @"C:\docs", () => "x", () => null);
        Assert.Equal(SaveTool.CancelledMessage, r.Output);
        Assert.False(r.IsError);
        Assert.False(r.Mutated);
    }

    [Theory]
    [InlineData("{\"saveAs\":\"true\"}")]
    [InlineData("{\"saveAs\":\"yes\"}")]
    [InlineData("{\"saveAs\":2}")]
    [InlineData("{\"saveAs\":null}")]
    [InlineData("{\"saveAs\":[true]}")]
    public void InvalidFlag_ThrowsAndTouchesNothing(string json)
    {
        string calls = "";
        var ex = Assert.Throws<System.ArgumentException>(() =>
            SaveTool.Run(Json(json), @"C:\docs", () => { calls += "i"; return "x"; }, () => { calls += "d"; return "x"; }));
        Assert.Contains("saveAs must be true or false", ex.Message);
        Assert.Equal("", calls);
    }
}
