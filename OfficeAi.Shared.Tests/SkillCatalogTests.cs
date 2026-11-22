using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class SkillCatalogTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "SkillCatalogTests-" + Guid.NewGuid().ToString("N"));
    private readonly List<string> _junctions = new List<string>();

    public void Dispose()
    {
        // Remove junctions first so deleting the tree never follows one out of the test folder.
        foreach (var link in _junctions) if (Directory.Exists(link)) Directory.Delete(link);
        if (Directory.Exists(_root)) Directory.Delete(_root, true);
    }

    private string Write(string relPath, string text)
    {
        string full = Path.Combine(_root, relPath);
        Directory.CreateDirectory(Path.GetDirectoryName(full));
        File.WriteAllText(full, text);
        return full;
    }

    private static Skill ParseOk(string text, string dir = @"C:\skills\demo")
    {
        string error;
        Skill s = SkillCatalog.Parse(text, dir, new List<string>(), "user", out error);
        Assert.True(s != null, error);
        return s;
    }

    private static ToolResult Read(List<Skill> skills, string skill, string path)
    {
        return SkillCatalog.ReadSkillFile(skills, JsonDocument.Parse(JsonSerializer.Serialize(new { skill, path })).RootElement);
    }

    private void Junction(string link, string target)
    {
        _junctions.Add(link);
        var p = Process.Start(new ProcessStartInfo("cmd.exe", "/c mklink /J \"" + link + "\" \"" + target + "\"") { CreateNoWindow = true, UseShellExecute = false });
        p.WaitForExit();
        Assert.True(Directory.Exists(link), "junction created");
    }

    [Fact]
    public void Parse_ReadsFrontmatterAndBody()
    {
        var s = ParseOk("---\nname: weekly-report\ndescription: \"Builds the weekly report\"\nargument-hint: [week]\ndisable-model-invocation: true\n---\n\n# Steps\nDo it.\n");
        Assert.Equal("weekly-report", s.Name);
        Assert.Equal("Builds the weekly report", s.Description);
        Assert.Equal("[week]", s.ArgumentHint);
        Assert.Equal("# Steps\nDo it.", s.Body);
        Assert.False(s.ModelInvocable);
        Assert.True(s.UserInvocable);
    }

    [Fact]
    public void Parse_FoldedDescription_AndUserInvocableFalse()
    {
        var s = ParseOk("---\r\nname: a\r\ndescription: >\r\n  first line\r\n  second line\r\nuser-invocable: false\r\n---\r\nbody");
        Assert.Equal("first line second line", s.Description);
        Assert.False(s.UserInvocable);
        Assert.True(s.ModelInvocable);
    }

    [Fact]
    public void Parse_NameFallsBackToFolder()
    {
        Assert.Equal("my-skill", ParseOk("---\ndescription: d\n---\nx", @"C:\s\my-skill").Name);
    }

    [Theory]
    [InlineData("---\nname: Bad_Name\ndescription: d\n---\n", "name")]
    [InlineData("---\nname: ok\n---\n", "description")]
    [InlineData("---\nname: ok\ndescription: d\n", "closing")]
    [InlineData("no frontmatter at all", "description")]
    public void Parse_Invalid_ReturnsNullWithReason(string text, string reasonPart)
    {
        string error;
        Assert.Null(SkillCatalog.Parse(text, @"C:\s\x", null, "user", out error));
        Assert.Contains(reasonPart, error);
    }

    [Fact]
    public void Load_MachineWinsOverUser_AndListsFilesWithoutSkillMd()
    {
        Write(@"machine\report\SKILL.md", "---\nname: report\ndescription: machine\n---\nm");
        Write(@"user\report\SKILL.md", "---\nname: report\ndescription: user\n---\nu");
        Write(@"user\notes\SKILL.md", "---\nname: notes\ndescription: n\n---\nn");
        Write(@"user\notes\references\style.md", "style");
        Write(@"user\notes\.hidden", "x");
        Write(@"user\broken\SKILL.md", "---\nname: broken\n---\n");

        var skills = SkillCatalog.Load(Path.Combine(_root, "machine"), Path.Combine(_root, "user"));

        Assert.Equal(new[] { "report", "notes" }, skills.Select(s => s.Name));
        Assert.Equal("machine", skills[0].Description);
        Assert.Equal("machine", skills[0].Source);
        Assert.Equal(new[] { "references/style.md" }, skills[1].Files);
    }

    [Fact]
    public void Load_MissingFolders_IsEmpty()
    {
        Assert.Empty(SkillCatalog.Load(Path.Combine(_root, "nope1"), Path.Combine(_root, "nope2")));
    }

    [Fact]
    public void ReadSkillFile_ReadsListedFile_CaseAndSlashInsensitive()
    {
        Write(@"user\notes\SKILL.md", "---\nname: notes\ndescription: n\n---\nn");
        Write(@"user\notes\references\style.md", "use short sentences");
        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));

        var r = Read(skills, "notes", @"References\Style.md");
        Assert.False(r.IsError, r.Output);
        Assert.Equal("use short sentences", r.Output);
    }

    [Theory]
    [InlineData("../other/SKILL.md")]
    [InlineData("..\\secret.txt")]
    [InlineData("C:\\Windows\\win.ini")]
    [InlineData("\\\\server\\share\\x.md")]
    [InlineData("/etc/passwd")]
    [InlineData("SKILL.md")]
    [InlineData("")]
    public void ReadSkillFile_RefusesAnythingNotInTheScannedList(string path)
    {
        Write(@"user\notes\SKILL.md", "---\nname: notes\ndescription: n\n---\nn");
        Write(@"user\secret.txt", "secret");
        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));

        var r = Read(skills, "notes", path);
        Assert.True(r.IsError);
        Assert.DoesNotContain("secret", r.Output.Replace("secret.txt", ""));
    }

    [Fact]
    public void ReadSkillFile_UnknownSkill_IsError()
    {
        Assert.True(Read(new List<Skill>(), "nope", "a.md").IsError);
    }

    [Fact]
    public void ReadSkillFile_Binary_GivesFullPathInsteadOfBytes()
    {
        Write(@"user\brand\SKILL.md", "---\nname: brand\ndescription: b\n---\nb");
        string logo = Path.Combine(_root, @"user\brand\assets\logo.png");
        Directory.CreateDirectory(Path.GetDirectoryName(logo));
        File.WriteAllBytes(logo, new byte[] { 0x89, 0x50, 0x4E, 0x47, 0, 1, 2 });
        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));

        var r = Read(skills, "brand", "assets/logo.png");
        Assert.False(r.IsError);
        Assert.Contains("binary", r.Output);
        Assert.Contains(logo, r.Output);
    }

    [Fact]
    public void ReadSkillFile_TooLarge_IsError()
    {
        Write(@"user\big\SKILL.md", "---\nname: big\ndescription: b\n---\nb");
        Write(@"user\big\huge.md", new string('a', SkillCatalog.MaxFileBytes + 1));
        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));
        Assert.True(Read(skills, "big", "huge.md").IsError);
    }

    [Fact]
    public void Junctions_AreNeverScannedOrRead()
    {
        Write(@"outside\private.md", "private data");
        Write(@"user\notes\SKILL.md", "---\nname: notes\ndescription: n\n---\nn");
        Junction(Path.Combine(_root, @"user\notes\linked"), Path.Combine(_root, "outside"));
        Junction(Path.Combine(_root, @"user\linkedskill"), Path.Combine(_root, @"user\notes"));

        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));

        Assert.Equal(new[] { "notes" }, skills.Select(s => s.Name));
        Assert.Empty(skills[0].Files);
        Assert.True(Read(skills, "notes", "linked/private.md").IsError);
    }

    [Fact]
    public void ReadSkillFile_LinkAddedAfterScan_IsRefused()
    {
        Write(@"outside\refs\style.md", "private data");
        Write(@"user\notes\SKILL.md", "---\nname: notes\ndescription: n\n---\nn");
        Write(@"user\notes\refs\style.md", "fine");
        var skills = SkillCatalog.Load(Path.Combine(_root, "m"), Path.Combine(_root, "user"));

        Directory.Delete(Path.Combine(_root, @"user\notes\refs"), true);
        Junction(Path.Combine(_root, @"user\notes\refs"), Path.Combine(_root, @"outside\refs"));

        var r = Read(skills, "notes", "refs/style.md");
        Assert.True(r.IsError);
        Assert.DoesNotContain("private", r.Output);
    }

    [Fact]
    public void ToPageJson_UsesCamelCaseFields()
    {
        var json = SkillCatalog.ToPageJson(new List<Skill> { ParseOk("---\nname: a\ndescription: d\n---\nb") });
        var el = JsonDocument.Parse(json).RootElement[0];
        Assert.Equal("a", el.GetProperty("name").GetString());
        Assert.True(el.GetProperty("modelInvocable").GetBoolean());
        Assert.Equal("user", el.GetProperty("source").GetString());
    }
}
