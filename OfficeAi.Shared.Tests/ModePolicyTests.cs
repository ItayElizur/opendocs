using System.IO;
using Xunit;
using OfficeAi.Shared;

public class ModePolicyTests
{
    private static readonly string[] All = { "readOnly", "commentOnly", "trackChanges", "fullAutonomy" };
    private static readonly string[] Apps = { "Word", "Excel", "PowerPoint", "Outlook" };

    [Fact]
    public void Parse_EachAppGetsOnlyItsOwnList()
    {
        var p = ModePolicy.Parse("{ \"disabledModes\": { \"outlook\": [\"fullAutonomy\"], \"Word\": [\"TRACKCHANGES\", \"fullAutonomy\"] } }");
        Assert.Equal(new[] { "fullAutonomy" }, p["Outlook"]);
        Assert.Equal(new[] { "trackChanges", "fullAutonomy" }, p["Word"]);
        Assert.Empty(p["Excel"]);
        Assert.Empty(p["PowerPoint"]);
    }

    [Theory]
    [InlineData("{ \"disabledModes\": {} }")]
    [InlineData("{ \"disabledModes\": { \"Word\": [] } }")]
    [InlineData("{ \"_comment\": \"x\" }")]
    public void Parse_NothingListed_DisablesNothing(string json)
    {
        foreach (var app in Apps) Assert.Empty(ModePolicy.Parse(json)[app]);
    }

    [Fact]
    public void Parse_AllowsCommentsAndTrailingCommas()
    {
        Assert.Equal(new[] { "fullAutonomy" }, ModePolicy.Parse("{ // admin note\n \"disabledModes\": { \"Excel\": [\"fullAutonomy\",], }, }")["Excel"]);
    }

    [Theory]
    [InlineData("{ \"disabledModes\": { \"Outlook\": [\"fullAutonmy\"] } }")] // mode typo
    [InlineData("{ \"disabledModes\": { \"Outlok\": [\"fullAutonomy\"] } }")] // app typo
    [InlineData("{ \"disabledModes\": { \"Word\": [1] } }")]
    [InlineData("{ \"disabledModes\": { \"Word\": \"fullAutonomy\" } }")]
    [InlineData("{ \"disabledModes\": [\"fullAutonomy\"] }")] // the old all-apps shape
    [InlineData("[\"fullAutonomy\"]")]
    [InlineData("{ not json")]
    public void Parse_BrokenPolicy_FailsClosedForEveryApp(string json)
    {
        var p = ModePolicy.Parse(json);
        foreach (var app in Apps) Assert.Equal(All, p[app]);
    }

    [Fact]
    public void Load_MissingFile_DisablesNothing()
    {
        var p = ModePolicy.Load(Path.Combine(Path.GetTempPath(), "no-such-dir-opendocs", "policy.json"));
        foreach (var app in Apps) Assert.Empty(p[app]);
    }

    // The copy the installer ships must parse, with nothing switched off - a typo there would lock every mode for everyone.
    [Fact]
    public void ShippedPolicy_DisablesNothing()
    {
        string dir = System.AppContext.BaseDirectory;
        while (dir != null && !File.Exists(Path.Combine(dir, "CLAUDE.md"))) dir = Path.GetDirectoryName(dir);
        Assert.True(dir != null, "repo root (CLAUDE.md) not found above the test folder");
        var p = ModePolicy.Load(Path.Combine(dir, "deploy", "policy.json"));
        foreach (var app in Apps) Assert.Empty(p[app]);
    }

    [Fact]
    public void Blocks_OnlyADisabledOrUnreportedModeWhenThePolicyDisablesSomething()
    {
        var none = new System.Collections.Generic.HashSet<string>();
        var noFull = new System.Collections.Generic.HashSet<string> { "fullAutonomy" };
        Assert.False(ModePolicy.Blocks(none, null));
        Assert.False(ModePolicy.Blocks(none, EditingMode.FullAutonomy));
        Assert.True(ModePolicy.Blocks(noFull, EditingMode.FullAutonomy));
        Assert.True(ModePolicy.Blocks(noFull, null));
        Assert.False(ModePolicy.Blocks(noFull, EditingMode.TrackChanges));
    }

    [Fact]
    public void WireName_MatchesEnumOrder()
    {
        Assert.Equal("readOnly", ModePolicy.WireName(EditingMode.ReadOnly));
        Assert.Equal("fullAutonomy", ModePolicy.WireName(EditingMode.FullAutonomy));
    }
}
