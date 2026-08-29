using System;
using System.IO;
using Xunit;
using OfficeAi.Shared;

public class DocSettingsStoreTests : IDisposable
{
    private readonly string _testFolder = "OfficeAiTests_" + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _testFolder);
        if (Directory.Exists(root)) Directory.Delete(root, true);
    }

    [Fact]
    public void Load_ReturnsEmptyForMissingFile()
    {
        DocSettings result = DocSettingsStore.Load(_testFolder, "nochat");
        Assert.Equal("", result.SystemMessage);
    }

    [Fact]
    public void SaveThenLoad_RoundTrips()
    {
        DocSettingsStore.Save(_testFolder, "chat1", new DocSettings { SystemMessage = "Use formal tone." });
        DocSettings result = DocSettingsStore.Load(_testFolder, "chat1");
        Assert.Equal("Use formal tone.", result.SystemMessage);
    }

    [Fact]
    public void SaveThenLoad_HandlesHebrewAndSpecialCharacters()
    {
        string message = "כתוב בעברית, ותשתמש ב\"מרכאות\" ותווים מיוחדים: <>&";
        DocSettingsStore.Save(_testFolder, "chat2", new DocSettings { SystemMessage = message });
        DocSettings result = DocSettingsStore.Load(_testFolder, "chat2");
        Assert.Equal(message, result.SystemMessage);
    }

    [Fact]
    public void Save_TruncatesOverCapMessages()
    {
        string tooLong = new string('x', DocSettingsStore.MaxSystemMessageChars + 500);
        DocSettingsStore.Save(_testFolder, "chat3", new DocSettings { SystemMessage = tooLong });
        DocSettings result = DocSettingsStore.Load(_testFolder, "chat3");
        Assert.Equal(DocSettingsStore.MaxSystemMessageChars, result.SystemMessage.Length);
    }

    [Fact]
    public void Load_ReturnsEmptyForCorruptFile()
    {
        // Write garbage directly, bypassing Save(), to simulate a truncated/
        // corrupt file (e.g. from a crash mid-write, the exact case atomic
        // write via a temp file + move is meant to prevent going forward).
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _testFolder, "DocSettings");
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "chat4.json"), "{ not valid json ");

        DocSettings result = DocSettingsStore.Load(_testFolder, "chat4");
        Assert.Equal("", result.SystemMessage);
    }

    [Fact]
    public void Save_OverwritesPreviousValue()
    {
        DocSettingsStore.Save(_testFolder, "chat5", new DocSettings { SystemMessage = "first" });
        DocSettingsStore.Save(_testFolder, "chat5", new DocSettings { SystemMessage = "second" });
        DocSettings result = DocSettingsStore.Load(_testFolder, "chat5");
        Assert.Equal("second", result.SystemMessage);
    }

    // ---- FT-1 Task 7b: Migrate ----

    [Fact]
    public void Migrate_MovesProvisionalSettingsToFreshTarget()
    {
        DocSettingsStore.Save(_testFolder, "unsaved-123-456", new DocSettings { SystemMessage = "guidelines" });
        DocSettingsStore.Migrate(_testFolder, "unsaved-123-456", "abcd1234abcd1234");

        DocSettings migrated = DocSettingsStore.Load(_testFolder, "abcd1234abcd1234");
        Assert.Equal("guidelines", migrated.SystemMessage);

        DocSettings oldStillThere = DocSettingsStore.Load(_testFolder, "unsaved-123-456");
        Assert.Equal("", oldStillThere.SystemMessage); // source file deleted after migration
    }

    [Fact]
    public void Migrate_NonEmptyProvisionalWinsOverExistingTarget()
    {
        DocSettingsStore.Save(_testFolder, "realid1", new DocSettings { SystemMessage = "old real guidelines" });
        DocSettingsStore.Save(_testFolder, "unsaved-1", new DocSettings { SystemMessage = "new provisional guidelines" });

        DocSettingsStore.Migrate(_testFolder, "unsaved-1", "realid1");

        DocSettings result = DocSettingsStore.Load(_testFolder, "realid1");
        Assert.Equal("new provisional guidelines", result.SystemMessage);
    }

    [Fact]
    public void Migrate_EmptyProvisionalNeverOverwritesExistingTarget()
    {
        DocSettingsStore.Save(_testFolder, "realid2", new DocSettings { SystemMessage = "keep me" });
        DocSettingsStore.Save(_testFolder, "unsaved-2", new DocSettings { SystemMessage = "" });

        DocSettingsStore.Migrate(_testFolder, "unsaved-2", "realid2");

        DocSettings result = DocSettingsStore.Load(_testFolder, "realid2");
        Assert.Equal("keep me", result.SystemMessage);
    }

    [Fact]
    public void Migrate_MissingSourceIsANoOp()
    {
        DocSettingsStore.Migrate(_testFolder, "never-existed", "realid3");
        DocSettings result = DocSettingsStore.Load(_testFolder, "realid3");
        Assert.Equal("", result.SystemMessage);
    }

    [Fact]
    public void Migrate_RunTwiceIsANoOpTheSecondTime()
    {
        DocSettingsStore.Save(_testFolder, "unsaved-3", new DocSettings { SystemMessage = "once" });
        DocSettingsStore.Migrate(_testFolder, "unsaved-3", "realid4");
        // Second call: source is already gone, must not throw and must not
        // touch the already-migrated target.
        DocSettingsStore.Migrate(_testFolder, "unsaved-3", "realid4");

        DocSettings result = DocSettingsStore.Load(_testFolder, "realid4");
        Assert.Equal("once", result.SystemMessage);
    }
}
