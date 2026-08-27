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

    // ---- current failure modes, as of the pure move (Task 2 Step 3) ----
    // Step 4 rewrites both of these into a clean ArgumentException.

    [Fact]
    public void HexToOle_ThreeDigitShorthand_ThrowsIndexError_BeforeTheFix()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ColorUtil.HexToOle("#abc"));
    }

    [Fact]
    public void HexToOle_NonHexDigits_ThrowsFormatException_BeforeTheFix()
    {
        Assert.Throws<FormatException>(() => ColorUtil.HexToOle("#GGGGGG"));
    }
}
