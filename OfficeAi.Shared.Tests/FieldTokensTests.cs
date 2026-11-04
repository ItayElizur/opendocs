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

        [Fact]
        public void FromFieldCodes_NameIsTheLeadingLettersSoASwitchWithoutASpaceStillReadsAsDate()
        {
            Assert.Equal("{DATE}", FieldTokens.FromFieldCodes("\u0013 DATE\\@ \"d/M\" \u0014 5/11\u0015"));
            Assert.Equal("{PAGE}", FieldTokens.FromFieldCodes("\u0013PAGE\u00143\u0015"));
        }

        [Fact]
        public void FromFieldCodes_NestedFieldsLeaveNoControlCharacters()
        {
            string nested = "a\u0013 IF \u0013 PAGE \u0014 3\u0015 > 1 \u0014later\u0015b";
            string read = FieldTokens.FromFieldCodes(nested);
            Assert.Equal("alaterb", read);
            Assert.DoesNotContain("\u0013", read);
            Assert.DoesNotContain("\u0014", read);
            Assert.DoesNotContain("\u0015", read);
        }

        [Fact]
        public void Split_CutsPlainTextAndFieldsInOrder()
        {
            var parts = FieldTokens.Split("Page {PAGE} of {NUMPAGES} / {PAGE}");
            Assert.Equal(new[] { "Page ", "", " of ", "", " / ", "" }, parts.Select(p => p.Text).ToArray());
            Assert.Equal(new string[] { null, "PAGE", null, "NUMPAGES", null, "PAGE" }, parts.Select(p => p.Field).ToArray());
        }

        [Fact]
        public void Split_AdjacentTokensAndEdgeCases()
        {
            var parts = FieldTokens.Split("{AUTHOR}{DATE}");
            Assert.Equal(new string[] { "AUTHOR", "DATE" }, parts.Select(p => p.Field).ToArray());
            Assert.Single(FieldTokens.Split("no tokens"));
            Assert.Equal("no tokens", FieldTokens.Split("no tokens")[0].Text);
            Assert.Empty(FieldTokens.Split(""));
            Assert.Empty(FieldTokens.Split(null));
            // an unknown token stays plain text
            Assert.Equal("{FOO} x", FieldTokens.Split("{FOO} x").Single().Text);
        }

        [Fact]
        public void Split_ConcatenatingThePartsRebuildsTheText()
        {
            string text = "L {PAGE}\r{DATE} R";
            string rebuilt = string.Concat(FieldTokens.Split(text).Select(p => p.Field == null ? p.Text : "{" + p.Field + "}"));
            Assert.Equal(text, rebuilt);
        }
    }
}
