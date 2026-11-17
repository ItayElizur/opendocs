using System.Collections.Generic;
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
        Assert.Equal(new[] { "fullAutonomy" }, p.Disabled["Outlook"]);
        Assert.Equal(new[] { "trackChanges", "fullAutonomy" }, p.Disabled["Word"]);
        Assert.Empty(p.Disabled["Excel"]);
        Assert.Empty(p.Disabled["PowerPoint"]);
    }

    [Fact]
    public void Parse_Approvals_PerAppAndTool_WithWildcard()
    {
        var p = ModePolicy.Parse("{ \"requireApproval\": { \"excel\": { \"*\": [\"trackChanges\"], \"save\": [] }, \"Outlook\": { \"delete_email\": [\"commentOnly\", \"TrackChanges\"] } } }");
        Assert.Equal(new[] { "trackChanges" }, p.Approvals["Excel"]["*"]);
        Assert.Empty(p.Approvals["Excel"]["save"]);
        Assert.Equal(new[] { "commentOnly", "trackChanges" }, p.Approvals["Outlook"]["delete_email"]);
        Assert.Empty(p.Approvals["Word"]);
        foreach (var app in Apps) Assert.Empty(p.Disabled[app]);
    }

    [Theory]
    [InlineData("{ \"disabledModes\": {} }")]
    [InlineData("{ \"disabledModes\": { \"Word\": [] }, \"requireApproval\": { \"Word\": {} } }")]
    [InlineData("{ \"_comment\": \"x\" }")]
    public void Parse_NothingListed_NoRules(string json)
    {
        var p = ModePolicy.Parse(json);
        foreach (var app in Apps) { Assert.Empty(p.Disabled[app]); Assert.Empty(p.Approvals[app]); }
    }

    [Fact]
    public void Parse_AllowsCommentsAndTrailingCommas()
    {
        Assert.Equal(new[] { "fullAutonomy" }, ModePolicy.Parse("{ // admin note\n \"disabledModes\": { \"Excel\": [\"fullAutonomy\",], }, }").Disabled["Excel"]);
    }

    [Theory]
    [InlineData("{ \"disabledModes\": { \"Outlook\": [\"fullAutonmy\"] } }")] // mode typo
    [InlineData("{ \"disabledModes\": { \"Outlok\": [\"fullAutonomy\"] } }")] // app typo
    [InlineData("{ \"disabledModes\": { \"Word\": [1] } }")]
    [InlineData("{ \"disabledModes\": { \"Word\": \"fullAutonomy\" } }")]
    [InlineData("{ \"disabledModes\": [\"fullAutonomy\"] }")] // the old all-apps shape
    [InlineData("{ \"requireApproval\": { \"Excel\": { \"*\": [\"trackChanges\", \"fullAutonmy\"] } } }")] // mode typo
    [InlineData("{ \"requireApproval\": { \"Exel\": { \"*\": [\"trackChanges\"] } } }")] // app typo
    [InlineData("{ \"requireApproval\": { \"Excel\": { \"Write Cells\": [\"trackChanges\"] } } }")] // not a tool name
    [InlineData("{ \"requireApproval\": { \"Excel\": [\"trackChanges\"] } }")]
    [InlineData("[\"fullAutonomy\"]")]
    [InlineData("{ not json")]
    public void Parse_BrokenPolicy_FailsClosedForEveryApp(string json)
    {
        var p = ModePolicy.Parse(json);
        foreach (var app in Apps) Assert.Equal(All, p.Disabled[app]);
    }

    [Fact]
    public void Load_MissingFile_NoRules()
    {
        var p = ModePolicy.Load(Path.Combine(Path.GetTempPath(), "no-such-dir-opendocs", "policy.json"));
        foreach (var app in Apps) { Assert.Empty(p.Disabled[app]); Assert.Empty(p.Approvals[app]); }
    }

    // The copy the installer ships must parse with no mode switched off - a typo there would lock every mode for
    // everyone - and must hold the default approvals: Excel and PowerPoint ask before every change in Track changes,
    // Outlook before a delete below Full autonomy, and Full autonomy never asks.
    [Fact]
    public void ShippedPolicy_DisablesNothing_AndHasTheDefaultApprovals()
    {
        var p = ModePolicy.Load(Path.Combine(RepoRoot(), "deploy", "policy.json"));
        foreach (var app in Apps)
        {
            Assert.Empty(p.Disabled[app]);
            foreach (var modes in p.Approvals[app].Values) Assert.DoesNotContain("fullAutonomy", modes);
        }
        Assert.Empty(p.Approvals["Word"]);
        Assert.Equal(new[] { "trackChanges" }, p.Approvals["Excel"]["*"]);
        Assert.Equal(new[] { "trackChanges" }, p.Approvals["PowerPoint"]["*"]);
        Assert.Equal(new[] { "commentOnly", "trackChanges" }, p.Approvals["Outlook"]["delete_email"]);
    }

    [Fact]
    public void Blocks_OnlyADisabledOrUnreportedModeWhenThePolicyDisablesSomething()
    {
        var none = new HashSet<string>();
        var noFull = new HashSet<string> { "fullAutonomy" };
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

    internal static string RepoRoot()
    {
        string dir = System.AppContext.BaseDirectory;
        while (dir != null && !File.Exists(Path.Combine(dir, "CLAUDE.md"))) dir = Path.GetDirectoryName(dir);
        Assert.True(dir != null, "repo root (CLAUDE.md) not found above the test folder");
        return dir;
    }
}
