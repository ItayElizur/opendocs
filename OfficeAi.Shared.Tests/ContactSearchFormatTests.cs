using System.Collections.Generic;
using Xunit;
using OfficeAi.Shared;

public class ContactSearchFormatTests
{
    private static ContactMatch M(string fullName, string displayName, string email)
    {
        return new ContactMatch { FullName = fullName, DisplayName = displayName, Email = email };
    }

    private static List<ContactMatch> L(params ContactMatch[] matches) => new List<ContactMatch>(matches);

    [Fact]
    public void Empty_ReturnsNoMatchMessage()
    {
        Assert.Equal("No contacts matched \"bob\".", ContactSearchFormat.Format(L(), "bob", 10));
    }

    [Fact]
    public void FullNameDiffersFromDisplayName_ShowsBothWithDisplayNameParenthesized()
    {
        // The exact scenario that confused the agent: the directory's DisplayName is
        // an org-formatted string, not the person's name. FullName is what a query
        // like "John Doe" will actually match, so it leads; DisplayName becomes a
        // parenthetical hint rather than the only thing shown.
        string outp = ContactSearchFormat.Format(L(M("John Doe", "Company/Unit/CEO", "john@x.com")), "John Doe", 10);
        Assert.Equal("- John Doe (Company/Unit/CEO) <john@x.com>\r\n", outp);
    }

    [Fact]
    public void FullNameEqualsDisplayName_ShowsOnce()
    {
        string outp = ContactSearchFormat.Format(L(M("Bob Smith", "Bob Smith", "bob@x.com")), "bob", 10);
        Assert.Equal("- Bob Smith <bob@x.com>\r\n", outp);
    }

    [Fact]
    public void FullNameEqualsDisplayName_CaseInsensitiveStillShowsOnce()
    {
        string outp = ContactSearchFormat.Format(L(M("Bob Smith", "BOB SMITH", "bob@x.com")), "bob", 10);
        Assert.Equal("- Bob Smith <bob@x.com>\r\n", outp);
    }

    [Fact]
    public void NoFullName_FallsBackToDisplayNameOnly()
    {
        // A shared/role mailbox (e.g. "IT Helpdesk") has no GivenName/Surname.
        string outp = ContactSearchFormat.Format(L(M(null, "IT Helpdesk", "helpdesk@x.com")), "helpdesk", 10);
        Assert.Equal("- IT Helpdesk <helpdesk@x.com>\r\n", outp);
    }

    [Fact]
    public void NoFullName_EmptyStringAlsoFallsBackToDisplayName()
    {
        string outp = ContactSearchFormat.Format(L(M("", "IT Helpdesk", "helpdesk@x.com")), "helpdesk", 10);
        Assert.Equal("- IT Helpdesk <helpdesk@x.com>\r\n", outp);
    }

    [Fact]
    public void EmailLessEntry_FormatsNameOnly()
    {
        string outp = ContactSearchFormat.Format(L(M("Bob Smith", "Bob Smith", "")), "bob", 10);
        Assert.Equal("- Bob Smith\r\n", outp);
    }

    [Fact]
    public void DedupeByLowercasedEmail_FirstWins()
    {
        string outp = ContactSearchFormat.Format(L(M("Bob", "Bob", "BOB@X.com"), M("Robert", "Robert", "bob@x.com")), "bob", 10);
        Assert.Equal("- Bob <BOB@X.com>\r\n", outp);
    }

    [Fact]
    public void DedupeEmailLessByName()
    {
        string outp = ContactSearchFormat.Format(L(M("Bob", "Bob", ""), M("Bob", "Bob", "")), "bob", 10);
        Assert.Equal("- Bob\r\n", outp);
    }

    [Fact]
    public void LimitAppliedAfterDedupe()
    {
        var matches = L(
            M("A", "A", "a@x.com"), M("B", "B", "b@x.com"), M("C", "C", "c@x.com"),
            M("D", "D", "d@x.com"), M("E", "E", "e@x.com"));
        string outp = ContactSearchFormat.Format(matches, "x", 2);
        Assert.Equal("- A <a@x.com>\r\n- B <b@x.com>\r\n", outp);
    }

    [Fact]
    public void LimitZero_ClampedToNoThrow()
    {
        string outp = ContactSearchFormat.Format(L(M("A", "A", "a@x.com")), "x", 0);
        Assert.Equal("No contacts matched \"x\".", outp);
    }

    [Fact]
    public void OrderPreserved()
    {
        var matches = L(M("Zoe", "Zoe", "zoe@x.com"), M("Amy", "Amy", "amy@x.com"));
        string outp = ContactSearchFormat.Format(matches, "x", 10);
        Assert.Equal("- Zoe <zoe@x.com>\r\n- Amy <amy@x.com>\r\n", outp);
    }
}
