using Xunit;
using OfficeAi.Shared;

public class TextWindowTests
{
    [Fact]
    public void ShortText_IsReturnedUnchanged()
    {
        Assert.Equal("hello", TextWindow.Slice("hello", 0, 100));
        Assert.Equal("", TextWindow.Slice(null, 0, 100));
    }

    [Fact]
    public void LongText_IsCutAndSaysWhatOffsetContinues()
    {
        string text = new string('a', 10);
        Assert.Equal("aaaa\n... (characters 0-4 of 10; call read_object again with offset: 4 for the rest)", TextWindow.Slice(text, 0, 4));
    }

    [Fact]
    public void Offset_ReadsTheNextWindowAndTheEndIsMarked()
    {
        string text = "0123456789";
        Assert.Equal("4567\n... (characters 4-8 of 10; call read_object again with offset: 8 for the rest)", TextWindow.Slice(text, 4, 4));
        Assert.Equal("89\n(end of text: characters 8-10 of 10)", TextWindow.Slice(text, 8, 4));
    }

    [Fact]
    public void WindowsCoverTheWholeTextWithoutGapsOrOverlap()
    {
        string text = "abcdefghijklmnopqrstuvwxyz";
        string rebuilt = "";
        int offset = 0;
        while (offset < text.Length)
        {
            string w = TextWindow.Slice(text, offset, 7);
            string body = w.Contains("\n") ? w.Substring(0, w.IndexOf('\n')) : w;
            rebuilt += body;
            offset += body.Length;
        }
        Assert.Equal(text, rebuilt);
    }

    [Fact]
    public void OffsetPastTheEnd_SaysSoInsteadOfReturningNothing()
    {
        Assert.Equal("(offset 50 is past the end: the text has 10 characters)", TextWindow.Slice("0123456789", 50, 4));
    }

    [Fact]
    public void NegativeOffsetCountsAsZero()
    {
        Assert.Equal("0123", TextWindow.Slice("0123", -5, 10));
    }

    [Fact]
    public void DoesNotSplitASurrogatePair()
    {
        string text = "ab" + "😀" + "cd"; // a, b, U+1F600 (two chars), c, d
        string w = TextWindow.Slice(text, 0, 3);
        Assert.StartsWith("ab\n", w);
        Assert.Contains("offset: 2", w);
    }
}
