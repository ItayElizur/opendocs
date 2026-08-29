using System;
using System.IO;
using Xunit;
using OfficeAi.Shared;

public class ChatStoreTests : IDisposable
{
    private readonly string _testFolder = "OfficeAiTests_" + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _testFolder);
        if (Directory.Exists(root)) Directory.Delete(root, true);
    }

    [Fact]
    public void ChatIdForFile_IsStableAndSixteenHexChars()
    {
        string id1 = ChatStore.ChatIdForFile(@"C:\docs\report.docx");
        string id2 = ChatStore.ChatIdForFile(@"C:\docs\report.docx");
        Assert.Equal(id1, id2);
        Assert.Equal(16, id1.Length);
    }

    [Fact]
    public void ChatIdForKey_IsStableSixteenHexAndDistinctPerKey()
    {
        string a1 = ChatStore.ChatIdForKey("alice@corp.local");
        string a2 = ChatStore.ChatIdForKey("alice@corp.local");
        string b = ChatStore.ChatIdForKey("bob@corp.local");

        Assert.Equal(a1, a2);
        Assert.NotEqual(a1, b);
        Assert.Equal(16, a1.Length);
        Assert.Matches("^[0-9a-f]{16}$", a1);
    }

    [Fact]
    public void ChatIdForKey_MatchesChatIdForFileForTheSameString()
    {
        Assert.Equal(ChatStore.ChatIdForFile(@"C:\x\y.docx"), ChatStore.ChatIdForKey(@"C:\x\y.docx"));
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEmptyForMissingFile()
    {
        var result = ChatStore.LoadSinceLastDivider(_testFolder, "nochat");
        Assert.Empty(result);
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEverythingWhenNoDividerYet()
    {
        string chatId = "chat1";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "hello");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "hi there");
        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);
        Assert.Equal(2, result.Count);
        Assert.Equal("hello", result[0].Text);
    }

    [Fact]
    public void LoadSinceLastDivider_OnlyReturnsRecordsAfterTheLastDivider()
    {
        string chatId = "chat2";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "first session question");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "first session answer");
        ChatStore.AppendDivider(_testFolder, chatId);
        ChatStore.AppendMessage(_testFolder, chatId, "user", "second session question");
        ChatStore.AppendMessage(_testFolder, chatId, "assistant", "second session answer");

        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);

        Assert.Equal(2, result.Count);
        Assert.Equal("second session question", result[0].Text);
        Assert.Equal("second session answer", result[1].Text);
    }

    [Fact]
    public void LoadSinceLastDivider_ReturnsEmptyImmediatelyAfterADividerWithNothingAfterIt()
    {
        string chatId = "chat3";
        ChatStore.AppendMessage(_testFolder, chatId, "user", "question");
        ChatStore.AppendDivider(_testFolder, chatId);

        var result = ChatStore.LoadSinceLastDivider(_testFolder, chatId);

        Assert.Empty(result);
    }

    // ---- FT-1 Task 7b: Migrate ----

    [Fact]
    public void Migrate_MovesProvisionalHistoryToFreshTarget()
    {
        ChatStore.AppendMessage(_testFolder, "unsaved-1-100", "user", "hello from provisional");
        ChatStore.Migrate(_testFolder, "unsaved-1-100", "realid1");

        var migrated = ChatStore.LoadSinceLastDivider(_testFolder, "realid1");
        Assert.Single(migrated);
        Assert.Equal("hello from provisional", migrated[0].Text);

        var oldStillThere = ChatStore.LoadSinceLastDivider(_testFolder, "unsaved-1-100");
        Assert.Empty(oldStillThere); // source file deleted after migration
    }

    [Fact]
    public void Migrate_AppendsOntoAnExistingTargetInChronologicalOrder()
    {
        ChatStore.AppendMessage(_testFolder, "realid2", "user", "existing question");
        ChatStore.AppendMessage(_testFolder, "realid2", "assistant", "existing answer");

        ChatStore.AppendMessage(_testFolder, "unsaved-2", "user", "provisional question");
        ChatStore.Migrate(_testFolder, "unsaved-2", "realid2");

        var result = ChatStore.LoadSinceLastDivider(_testFolder, "realid2");
        Assert.Equal(3, result.Count);
        Assert.Equal("existing question", result[0].Text);
        Assert.Equal("existing answer", result[1].Text);
        Assert.Equal("provisional question", result[2].Text);
    }

    [Fact]
    public void Migrate_MissingSourceIsANoOp()
    {
        ChatStore.Migrate(_testFolder, "never-existed", "realid3");
        var result = ChatStore.LoadSinceLastDivider(_testFolder, "realid3");
        Assert.Empty(result);
    }

    [Fact]
    public void Migrate_RunTwiceIsANoOpTheSecondTime()
    {
        ChatStore.AppendMessage(_testFolder, "unsaved-3", "user", "once");
        ChatStore.Migrate(_testFolder, "unsaved-3", "realid4");
        ChatStore.Migrate(_testFolder, "unsaved-3", "realid4");

        var result = ChatStore.LoadSinceLastDivider(_testFolder, "realid4");
        Assert.Single(result);
    }
}
