using System.Collections.Generic;
using System.Linq;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    // The pure parts of the Outlook item window's tools: set_current_fields' recipient step, read_main_chat's digest,
    // and the body position map edit_current_body edits through.
    public class ItemEditHelpersTests
    {
        private static RecipientEntry R(string name, string address, string kind) => new RecipientEntry { Name = name, Address = address, Kind = kind };

        private static KeyValuePair<string, string> A(string text, string kind) => new KeyValuePair<string, string>(text, kind);

        private static readonly List<RecipientEntry> Meeting = new List<RecipientEntry>
        {
            R("Me", "me@x.com", "organizer"),
            R("Dana Levi", "dana@x.com", "required"),
            R("Avi Cohen", "avi@x.com", "required"),
            R("Noa", "noa@x.com", "required"),
        };

        [Fact]
        public void Add_NewPersonIsAddedUnderTheirKind()
        {
            var p = RecipientEdits.Plan(Meeting, new[] { A("Yael <yael@x.com>", "optional") }, null);
            Assert.Equal(new[] { A("Yael <yael@x.com>", "optional") }, p.Add);
            Assert.Empty(p.Remove);
            Assert.Empty(p.ChangeKind);
        }

        [Fact]
        public void Add_SomeoneAlreadyThereUnderAnotherKind_MovesOnlyThem()
        {
            var p = RecipientEdits.Plan(Meeting, new[] { A("dana levi", "optional") }, null);
            Assert.Equal(new[] { new KeyValuePair<int, string>(1, "optional") }, p.ChangeKind);
            Assert.Empty(p.Add);
            Assert.Empty(p.Remove);
        }

        [Fact]
        public void Add_SomeoneAlreadyThereUnderTheSameKind_ChangesNothing()
        {
            var p = RecipientEdits.Plan(Meeting, new[] { A("avi@x.com", "required") }, null);
            Assert.Empty(p.Add);
            Assert.Empty(p.ChangeKind);
        }

        [Fact]
        public void Remove_ByNameOrAddress_HighestIndexFirst()
        {
            var p = RecipientEdits.Plan(Meeting, null, new[] { "dana@x.com", "Noa" });
            Assert.Equal(new[] { 3, 1 }, p.Remove);
            Assert.Empty(p.NotFound);
        }

        [Fact]
        public void Remove_SomeoneNotOnTheItem_IsReportedNotGuessed()
        {
            var p = RecipientEdits.Plan(Meeting, null, new[] { "Dan" });
            Assert.Empty(p.Remove);
            Assert.Equal(new[] { "Dan" }, p.NotFound);
        }

        [Fact]
        public void TheOrganizerIsNeverRemovedOrMoved()
        {
            var p = RecipientEdits.Plan(Meeting, new[] { A("me@x.com", "optional") }, new[] { "Me" });
            Assert.Empty(p.Remove);
            Assert.Empty(p.ChangeKind);
            Assert.Equal(new[] { "Me" }, p.NotFound);
            Assert.Equal(new[] { A("me@x.com", "optional") }, p.Add); // not on the list as an attendee, so it is an add
        }

        [Fact]
        public void Duplicates_InOneCall_AddOnce_LastKindWins()
        {
            var p = RecipientEdits.Plan(new List<RecipientEntry>(), new[] { A("a@x.com", "to"), A("A@X.com", "to"), A("Ann <a@x.com>", "cc") }, null);
            Assert.Equal(new[] { A("a@x.com", "cc") }, p.Add);
        }

        [Fact]
        public void RemovedThenAddedUnderAnotherKind_IsARemoveAndAnAdd()
        {
            var p = RecipientEdits.Plan(Meeting, new[] { A("noa@x.com", "optional") }, new[] { "Noa" });
            Assert.Equal(new[] { 3 }, p.Remove);
            Assert.Equal(new[] { A("noa@x.com", "optional") }, p.Add);
        }

        private static List<ChatRecord> Chat(int n, int chars = 5) =>
            Enumerable.Range(1, n).Select(i => new ChatRecord { Role = i % 2 == 1 ? "user" : "assistant", Text = chars > 5 ? new string('x', chars) : "m" + i + "." }).ToList();

        [Fact]
        public void Digest_Empty()
        {
            Assert.Contains("empty", MainChatDigest.Format(new List<ChatRecord>()));
        }

        [Fact]
        public void Digest_KeepsOnlyTheLast20()
        {
            string s = MainChatDigest.Format(Chat(25));
            Assert.StartsWith("The last 20 of 25 messages", s);
            Assert.Equal(20, s.Split('\n').Count(l => l.StartsWith("user: ") || l.StartsWith("assistant: ")));
            Assert.Contains("assistant: m6.", s); // message 6 is the first kept
            Assert.DoesNotContain("m5.", s);
        }

        [Fact]
        public void Digest_CutsEachMessageTo2000Characters()
        {
            string s = MainChatDigest.Format(Chat(1, 5000));
            Assert.Contains(new string('x', 2000) + " ...[cut]", s);
            Assert.DoesNotContain(new string('x', 2001), s);
        }

        [Fact]
        public void Positions_CellMarkIsOnePosition()
        {
            // "ab" in a cell, then "c" after the table
            Assert.Equal(new[] { 0, 1, 2, 2, 3 }, TextAnchors.DocumentPositions("ab\r\ac"));
        }

        [Fact]
        public void Clean_KeepBreaks_TurnsParagraphMarksIntoNewlines()
        {
            var c = TextAnchors.Clean("Hi Dana,\rThanks\vBest\r\a", true);
            Assert.Equal("Hi Dana,\nThanks\nBest\n", c.Text);
            Assert.Single(TextAnchors.FindAll(c.Text, "Dana,\nThanks", 5));
        }
    }
}
