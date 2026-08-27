using System;
using Xunit;
using OfficeAi.Shared;

public class ColorUtilTests
{
    // ---- known-value round trips - would catch a BGR/RGB swap regression ----

    [Fact]
    public void HexToOle_Red_MatchesKnownOleValue()
    {
        Assert.Equal(0x0000FF, ColorUtil.HexToOle("#FF0000"));
    }

    [Fact]
    public void HexToOle_Blue_MatchesKnownOleValue()
    {
        Assert.Equal(0xFF0000, ColorUtil.HexToOle("#0000FF"));
    }

    [Fact]
    public void HexToOle_Black_IsZero()
    {
        Assert.Equal(0, ColorUtil.HexToOle("#000000"));
    }

    [Fact]
    public void HexToOle_White_IsAllBitsSet()
    {
        Assert.Equal(0xFFFFFF, ColorUtil.HexToOle("#FFFFFF"));
    }

    [Fact]
    public void HexToOle_WithoutLeadingHash_MatchesWithHash()
    {
        Assert.Equal(ColorUtil.HexToOle("#FF0000"), ColorUtil.HexToOle("FF0000"));
    }

    [Fact]
    public void HexToOle_Lowercase_MatchesUppercase()
    {
        Assert.Equal(ColorUtil.HexToOle("#FF0000"), ColorUtil.HexToOle("#ff0000"));
    }

    [Fact]
    public void HexToOle_SurroundingWhitespace_IsTrimmed()
    {
        Assert.Equal(ColorUtil.HexToOle("#FF0000"), ColorUtil.HexToOle(" #FF0000 "));
    }

    // ---- Task 2 Step 4: malformed input now gets an actionable message ----

    [Fact]
    public void HexToOle_ThreeDigitShorthand_ExpandsToSixDigit()
    {
        // "abc" is CSS shorthand for "aabbcc".
        Assert.Equal(ColorUtil.HexToOle("#aabbcc"), ColorUtil.HexToOle("#abc"));
    }

    [Fact]
    public void HexToOle_ThreeDigitShorthand_WithoutHash_AlsoExpands()
    {
        Assert.Equal(ColorUtil.HexToOle("#aabbcc"), ColorUtil.HexToOle("abc"));
    }

    [Fact]
    public void HexToOle_NonHexDigits_ThrowsWithOffendingValueInMessage()
    {
        var ex = Assert.Throws<ArgumentException>(() => ColorUtil.HexToOle("#GGGGGG"));
        Assert.Contains("#GGGGGG", ex.Message);
    }

    [Fact]
    public void HexToOle_WrongLength_Throws()
    {
        Assert.Throws<ArgumentException>(() => ColorUtil.HexToOle("#12345"));
    }

    [Fact]
    public void HexToOle_EmptyString_Throws()
    {
        Assert.Throws<ArgumentException>(() => ColorUtil.HexToOle(""));
    }

    [Fact]
    public void HexToOle_Null_Throws()
    {
        Assert.Throws<ArgumentException>(() => ColorUtil.HexToOle(null));
    }
}
