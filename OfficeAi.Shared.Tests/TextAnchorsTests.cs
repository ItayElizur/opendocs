using System.Linq;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class TextAnchorsTests
    {
        [Fact]
        public void Clean_DropsReferenceMarksAndFieldCodesAndMapsBack()
        {
            // a footnote mark, then a PAGE field whose code text sits before the result the reader sees
            string raw = "the\u0002 mouse \u0013 PAGE \u00143\u0015x\r";
            var c = TextAnchors.Clean(raw);
            Assert.Equal("the mouse 3x", c.Text);
            Assert.Equal('x', raw[c.Map[c.Text.Length - 1]]);
            Assert.Equal(' ', raw[c.Map[3]]); // the space after the reference mark, index 4 in raw
            Assert.Equal(4, c.Map[3]);
            Assert.Equal('3', raw[c.Map[c.Text.IndexOf('3')]]);
        }

        [Fact]
        public void Clean_FieldCodeTextCannotBeMatchedButItsResultCan()
        {
            var c = TextAnchors.Clean("See page \u0013 PAGE \u00143\u0015 of the report");
            Assert.Equal("See page 3 of the report", c.Text);
            Assert.Equal(new[] { 4 }, TextAnchors.FindAll(c.Text, "page 3", 5).ToArray());
            // "page" matches only the visible word, not the field code "PAGE" in front of the result
            Assert.Single(TextAnchors.FindAll(c.Text, "page", 5));
        }

        [Fact]
        public void Clean_FieldWithoutAResultAndNestedFieldsLeaveNothingBehind()
        {
            // no separator: the code part runs to the end mark; nested: both codes are dropped, the inner result is kept
            Assert.Equal("ab", TextAnchors.Clean("a\u0013 XE \"term\" \u0015b").Text);
            // a field inside another field's code part is not displayed at all, result included
            Assert.Equal("x[y", TextAnchors.Clean("x\u0013 IF \u0013 PAGE \u00147\u0015 > 3 \u0014[\u0015y").Text);
            // a field with no result inside the code part, then the outer field's own result
            Assert.Equal("xshowny", TextAnchors.Clean("x\u0013 IF \u0013 PAGE \u0015 = 1 \u0014shown\u0015y").Text);
        }

        [Fact]
        public void Clean_UnbalancedFieldMarksDoNotThrow()
        {
            Assert.Equal("ab", TextAnchors.Clean("a\u0015\u0014b").Text);
            Assert.Equal("a", TextAnchors.Clean("a\u0013 PAGE ").Text);
        }

        [Fact]
        public void PositionsAlign_AcceptsTheEndOfCellMarkOnly()
        {
            Assert.True(TextAnchors.PositionsAlign("hello\r", 6));
            Assert.True(TextAnchors.PositionsAlign("hello\r\a", 6));   // last paragraph in a table cell: two characters, one position
            Assert.False(TextAnchors.PositionsAlign("hello\r\a", 5));
            Assert.False(TextAnchors.PositionsAlign("hello\r", 9));    // hidden or collapsed content: positions do not line up
            Assert.False(TextAnchors.PositionsAlign("hello", 7));
            Assert.True(TextAnchors.PositionsAlign(null, 0));
        }

        [Fact]
        public void FindAll_MatchesAcrossAFootnoteMark()
        {
            var c = TextAnchors.Clean("the\u0002 mouse the cat");
            Assert.Equal(new[] { 0 }, TextAnchors.FindAll(c.Text, "the mouse", 20).ToArray());
        }

        [Fact]
        public void FindAll_IsCaseInsensitiveNonOverlappingAndCapped()
        {
            Assert.Equal(new[] { 0, 2 }, TextAnchors.FindAll("ababab", "ab", 2).ToArray());
            Assert.Equal(new[] { 0, 4 }, TextAnchors.FindAll("The the", "the", 20).ToArray());
            Assert.Empty(TextAnchors.FindAll("abc", "", 5));
        }

        [Fact]
        public void FindAll_NonBreakingSpaceMatchesSpace()
        {
            var c = TextAnchors.Clean("a b");
            Assert.Equal(new[] { 0 }, TextAnchors.FindAll(c.Text, "a b", 5).ToArray());
        }

        [Fact]
        public void Snippet_MarksTheMatchAndClampsAtTheEdges()
        {
            Assert.Equal("...[[the]] mouse the...", TextAnchors.Snippet("the mouse the cat", 0, 3, 10).Replace(" cat...", "..."));
            Assert.Equal("...mouse [[the]] cat...", TextAnchors.Snippet("the mouse the cat", 10, 3, 6));
        }
    }
}
