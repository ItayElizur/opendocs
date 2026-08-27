using System;
using Xunit;
using OfficeAi.Shared;

public class TextUtilTests
{
    // ---- ColumnLetter ----

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    public void ColumnLetter_ConvertsKnownColumnNumbers(int col, string expected)
    {
        Assert.Equal(expected, TextUtil.ColumnLetter(col));
    }

    [Fact]
    public void ColumnLetter_NonPositiveInput_ReturnsEmptyString()
    {
        // Current behavior, pinned rather than "fixed" - the while loop never
        // runs for col <= 0.
        Assert.Equal("", TextUtil.ColumnLetter(0));
        Assert.Equal("", TextUtil.ColumnLetter(-5));
    }

    // ---- ReplaceAllOccurrences ----

    [Fact]
    public void ReplaceAllOccurrences_Ordinal_IsCaseSensitive()
    {
        Assert.Equal("XbcXbc", TextUtil.ReplaceAllOccurrences("abcabc", "a", "X", StringComparison.Ordinal));
        Assert.Equal("abcabc", TextUtil.ReplaceAllOccurrences("abcabc", "A", "X", StringComparison.Ordinal));
    }

    [Fact]
    public void ReplaceAllOccurrences_OrdinalIgnoreCase_MatchesRegardlessOfCase()
    {
        Assert.Equal("XbcXbc", TextUtil.ReplaceAllOccurrences("abcAbc", "a", "X", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ReplaceAllOccurrences_ReplacementContainsSearchTerm_DoesNotLoopForever()
    {
        // The single most important case: pos advances by find.Length, not by
        // the replacement's, so replacing "a" with "aa" must not re-scan the
        // freshly-inserted text.
        Assert.Equal("aaaa", TextUtil.ReplaceAllOccurrences("aa", "a", "aa", StringComparison.Ordinal));
    }

    [Fact]
    public void ReplaceAllOccurrences_NoMatch_ReturnsInputUnchanged()
    {
        Assert.Equal("hello", TextUtil.ReplaceAllOccurrences("hello", "xyz", "X", StringComparison.Ordinal));
    }

    [Fact]
    public void ReplaceAllOccurrences_EmptyInput_ReturnsEmpty()
    {
        Assert.Equal("", TextUtil.ReplaceAllOccurrences("", "a", "b", StringComparison.Ordinal));
    }

    // ---- CountOccurrences ----

    [Fact]
    public void CountOccurrences_OverlappingCandidates_CountsNonOverlapping()
    {
        // "aaa" against "aa": pos skips past each match, so this is 1, not 2.
        // Pinning the current non-overlapping semantics deliberately.
        Assert.Equal(1, TextUtil.CountOccurrences("aaa", "aa", StringComparison.Ordinal));
    }

    [Fact]
    public void CountOccurrences_EmptyNeedle_ReturnsZero()
    {
        Assert.Equal(0, TextUtil.CountOccurrences("hello", "", StringComparison.Ordinal));
    }

    [Fact]
    public void CountOccurrences_CaseInsensitive_CountsAcrossCase()
    {
        // Every non-space character here is an "a"/"A" - 4 matches.
        Assert.Equal(4, TextUtil.CountOccurrences("Aa aA", "a", StringComparison.OrdinalIgnoreCase));
    }

    // ---- IsRtlMajority ----

    [Fact]
    public void IsRtlMajority_PureHebrew_IsTrue()
    {
        Assert.True(TextUtil.IsRtlMajority("שלום עולם"));
    }

    [Fact]
    public void IsRtlMajority_PureEnglish_IsFalse()
    {
        Assert.False(TextUtil.IsRtlMajority("Hello world"));
    }

    [Fact]
    public void IsRtlMajority_EmptyOrNull_IsFalse()
    {
        Assert.False(TextUtil.IsRtlMajority(""));
        Assert.False(TextUtil.IsRtlMajority(null));
    }

    [Fact]
    public void IsRtlMajority_DigitsAndPunctuationOnly_IsFalse()
    {
        // No letters at all, RTL or LTR - rtl stays 0, so the "rtl > 0" guard
        // keeps this false rather than vacuously true.
        Assert.False(TextUtil.IsRtlMajority("123 - 456 !"));
    }

    [Fact]
    public void IsRtlMajority_FiftyFiftyMix_TiesGoRtl()
    {
        // The rule is "rtl >= ltr", so an exact tie resolves to RTL - pinned
        // deliberately, not incidentally.
        Assert.True(TextUtil.IsRtlMajority("aב"));
    }

    [Fact]
    public void IsRtlMajority_HebrewWithLatinPunctuation_IsTrue()
    {
        // 4 Hebrew letters vs 2 Latin ("hi") - RTL majority, punctuation
        // counted as neither.
        Assert.True(TextUtil.IsRtlMajority("שלום, hi!"));
    }
}
