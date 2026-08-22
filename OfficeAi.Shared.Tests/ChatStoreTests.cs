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
}
