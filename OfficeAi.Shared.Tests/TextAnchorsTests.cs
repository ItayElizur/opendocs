using System.Linq;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class TextAnchorsTests
    {
        [Fact]
        public void Clean_DropsReferenceAndFieldMarksAndMapsBack()
        {
            string raw = "the\u0002 mouse \u0013 PAGE \u00143\u0015x\r";
            var c = TextAnchors.Clean(raw);
            Assert.Equal("the mouse  PAGE 3x", c.Text);
            Assert.Equal('x', raw[c.Map[c.Text.Length - 1]]);
            Assert.Equal(' ', raw[c.Map[3]]); // the space after the reference mark, index 4 in raw
            Assert.Equal(4, c.Map[3]);
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
