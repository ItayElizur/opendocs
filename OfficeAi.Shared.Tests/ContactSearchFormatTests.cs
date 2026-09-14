using System.Collections.Generic;
using Xunit;
using OfficeAi.Shared;

public class ContactSearchFormatTests
{
    private static List<KeyValuePair<string, string>> M(params (string name, string email)[] pairs)
    {
        var list = new List<KeyValuePair<string, string>>();
        foreach (var p in pairs) list.Add(new KeyValuePair<string, string>(p.name, p.email));
        return list;
    }

    [Fact]
    public void Empty_ReturnsNoMatchMessage()
    {
        Assert.Equal("No contacts matched \"bob\".", ContactSearchFormat.Format(M(), "bob", 10));
    }

    [Fact]
    public void SingleWithEmail_FormatsWithAngleBrackets()
    {
        Assert.Equal("- Bob Smith <bob@x.com>\r\n", ContactSearchFormat.Format(M(("Bob Smith", "bob@x.com")), "bob", 10));
    }

    [Fact]
    public void EmailLessEntry_FormatsNameOnly()
    {
        Assert.Equal("- Bob Smith\r\n", ContactSearchFormat.Format(M(("Bob Smith", "")), "bob", 10));
    }

    [Fact]
    public void DedupeByLowercasedEmail_FirstWins()
    {
        string outp = ContactSearchFormat.Format(M(("Bob", "BOB@X.com"), ("Robert", "bob@x.com")), "bob", 10);
        Assert.Equal("- Bob <BOB@X.com>\r\n", outp);
    }

    [Fact]
    public void DedupeEmailLessByName()
    {
        string outp = ContactSearchFormat.Format(M(("Bob", ""), ("Bob", "")), "bob", 10);
        Assert.Equal("- Bob\r\n", outp);
    }

    [Fact]
    public void LimitAppliedAfterDedupe()
    {
        var matches = M(("A", "a@x.com"), ("B", "b@x.com"), ("C", "c@x.com"), ("D", "d@x.com"), ("E", "e@x.com"));
        string outp = ContactSearchFormat.Format(matches, "x", 2);
        Assert.Equal("- A <a@x.com>\r\n- B <b@x.com>\r\n", outp);
    }

    [Fact]
    public void LimitZero_ClampedToNoThrow()
    {
        // Format itself does not clamp; callers pass Math.Max(1, ...). A literal 0
        // must still not throw - it just yields the empty-result message.
        string outp = ContactSearchFormat.Format(M(("A", "a@x.com")), "x", 0);
        Assert.Equal("No contacts matched \"x\".", outp);
    }

    [Fact]
    public void OrderPreserved()
    {
        var matches = M(("Zoe", "zoe@x.com"), ("Amy", "amy@x.com"));
        string outp = ContactSearchFormat.Format(matches, "x", 10);
        Assert.Equal("- Zoe <zoe@x.com>\r\n- Amy <amy@x.com>\r\n", outp);
    }
}
