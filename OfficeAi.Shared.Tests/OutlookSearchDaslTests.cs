using System;
using Xunit;
using OfficeAi.Shared;

public class OutlookSearchDaslTests
{
    [Fact]
    public void EmptyInputs_ProduceEmptyFilter()
    {
        Assert.Equal("", OutlookDasl.BuildSearchFilter("", null, null, null));
    }

    [Fact]
    public void QueryOnly_MatchesSubjectAndBody()
    {
        string dasl = OutlookDasl.BuildSearchFilter("invoice", null, null, null);
        Assert.Contains("urn:schemas:httpmail:subject\" LIKE '%invoice%'", dasl);
        Assert.Contains("urn:schemas:httpmail:textdescription\" LIKE '%invoice%'", dasl);
        Assert.DoesNotContain(" AND ", dasl);
    }

    [Fact]
    public void SingleQuote_IsDoubledToEscape()
    {
        string dasl = OutlookDasl.BuildSearchFilter("O'Brien", null, null, "a'b@x.com");
        Assert.Contains("O''Brien", dasl);
        Assert.Contains("a''b@x.com", dasl);
    }

    [Fact]
    public void DatesUseUtcIsoFormat()
    {
        string dasl = OutlookDasl.BuildSearchFilter("", new DateTime(2026, 1, 5), new DateTime(2026, 2, 1), null);
        Assert.Contains(">= '2026-01-05T00:00:00Z'", dasl);
        Assert.Contains("<= '2026-02-01T00:00:00Z'", dasl);
        Assert.Contains(" AND ", dasl);
    }

    [Fact]
    public void AllFour_AreAndCombinedInOrder()
    {
        string dasl = OutlookDasl.BuildSearchFilter("q", new DateTime(2026, 1, 1), new DateTime(2026, 1, 31), "alice@corp.local");
        string[] segments = dasl.Split(new[] { " AND " }, StringSplitOptions.None);
        Assert.Equal(4, segments.Length);
        Assert.Contains("subject", segments[0]);
        Assert.Contains(">=", segments[1]);
        Assert.Contains("<=", segments[2]);
        Assert.Contains("fromemail", segments[3]);
    }
}
