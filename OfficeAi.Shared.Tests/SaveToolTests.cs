using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class SaveToolTests
{
    private const string Folder = @"C:\docs";
    private const string FullName = @"C:\docs\a.docx";

    private static JsonElement Json(string s)
    {
        return JsonDocument.Parse(s).RootElement.Clone();
    }

    private static ToolResult Run(string json, string path, bool dialogConfirms, out string calls)
    {
        string log = "";
        var result = SaveTool.Run(Json(json), path,
            () => { log += "[in-place]"; },
            () => { log += "[dialog]"; return dialogConfirms; },
            () => FullName);
        calls = log;
        return result;
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"saveAs\":false}")]
    [InlineData("{\"saveAs\":0}")]
    public void SavedFile_WithoutSaveAs_SavesInPlace(string json)
    {
        var r = Run(json, Folder, true, out string calls);
        Assert.Equal("[in-place]", calls);
        Assert.Equal("Saved " + FullName + ".", r.Output);
        Assert.False(r.IsError);
        Assert.False(r.Mutated);
    }

    [Theory]
    [InlineData("{\"saveAs\":true}")]
    [InlineData("{\"saveAs\":1}")]
    public void SaveAsTrue_OpensTheDialog(string json)
    {
        var r = Run(json, Folder, true, out string calls);
        Assert.Equal("[dialog]", calls);
        Assert.Equal("Saved as " + FullName + ".", r.Output);
    }

    [Fact]
    public void NeverSavedFile_FallsThroughToTheDialog()
    {
        Run("{}", "", true, out string calls);
        Assert.Equal("[dialog]", calls);
    }

    [Fact]
    public void CancelledDialog_IsANonErrorCancelledResult()
    {
        var r = Run("{\"saveAs\":true}", Folder, false, out string calls);
        Assert.Equal("[dialog]", calls);
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
            SaveTool.Run(Json(json), Folder, () => { calls += "i"; }, () => { calls += "d"; return true; }, () => FullName));
        Assert.Contains("saveAs must be true or false", ex.Message);
        Assert.Equal("", calls);
    }
}
