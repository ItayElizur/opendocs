using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class ItemChatIdTests
    {
        [Fact]
        public void ANewItemGetsOneProvisionalIdUntilItIsSaved()
        {
            string first = ItemChatId.Resolve(null, "S", "", out string from1);
            string again = ItemChatId.Resolve(first, "S", null, out string from2);

            Assert.StartsWith("new-", first);
            Assert.Equal(first, again);
            Assert.Null(from1);
            Assert.Null(from2);
        }

        [Fact]
        public void TwoNewItemsGetDifferentIds()
        {
            Assert.NotEqual(ItemChatId.Resolve(null, "S", "", out _), ItemChatId.Resolve(null, "S", "", out _));
        }

        [Fact]
        public void TheFirstSaveMovesTheProvisionalChatOntoTheSavedId()
        {
            string provisional = ItemChatId.Resolve(null, "S", "", out _);
            string saved = ItemChatId.Resolve(provisional, "S", "ENTRY1", out string from);

            Assert.StartsWith("item-", saved);
            Assert.Equal(provisional, from);
        }

        [Fact]
        public void AnItemOpenedAlreadySavedGetsTheSameIdEveryTime_AndHasNothingToMove()
        {
            string a = ItemChatId.Resolve(null, "S", "ENTRY1", out string from);
            string b = ItemChatId.Resolve(null, "S", "ENTRY1", out _);

            Assert.Equal(a, b);
            Assert.Null(from);
            Assert.NotEqual(a, ItemChatId.Resolve(null, "OTHER-STORE", "ENTRY1", out _));
        }

        [Fact]
        public void ASavedIdStaysWhenTheEntryIdChangesLater()
        {
            string saved = ItemChatId.Resolve(null, "S", "ENTRY1", out _);
            Assert.Equal(saved, ItemChatId.Resolve(saved, "S", "ENTRY-AFTER-MOVE", out string from));
            Assert.Null(from);
        }
    }
}
