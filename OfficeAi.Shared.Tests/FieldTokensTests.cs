using System.Linq;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class FieldTokensTests
    {
        [Fact]
        public void Find_ReturnsKnownTokensInOrder()
        {
            Assert.Equal(new[] { "PAGE", "NUMPAGES", "PAGE" }, FieldTokens.Find("Page {PAGE} of {NUMPAGES} / {PAGE}").ToArray());
        }

        [Fact]
        public void Find_IgnoresUnknownAndLowercaseBraces()
        {
            Assert.Empty(FieldTokens.Find("{FOO} {page} {PAGE"));
            Assert.Empty(FieldTokens.Find(null));
        }

        [Fact]
        public void FromFieldCodes_ReplacesKnownFieldWithToken()
        {
            Assert.Equal("Page {PAGE} of {NUMPAGES}", FieldTokens.FromFieldCodes("Page \u0013 PAGE \u0014 3\u0015 of \u0013 NUMPAGES \u0014 9\u0015"));
        }

        [Fact]
        public void FromFieldCodes_TokenWithSwitchesAndMissingResult()
        {
            Assert.Equal("{DATE}|{TIME}", FieldTokens.FromFieldCodes("\u0013 DATE \\@ \"d/M/yyyy\" \u0014 1/1/2026\u0015|\u0013 time \u0015"));
        }

        [Fact]
        public void FromFieldCodes_UnknownFieldKeepsItsResult()
        {
            Assert.Equal("Ref 7", FieldTokens.FromFieldCodes("Ref \u0013 SEQ Figure \u00147\u0015"));
        }

        [Fact]
        public void FromFieldCodes_PlainTextUnchanged()
        {
            Assert.Equal("no fields", FieldTokens.FromFieldCodes("no fields"));
        }
    }
}
