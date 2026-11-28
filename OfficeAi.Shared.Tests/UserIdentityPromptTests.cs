using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class UserIdentityPromptTests
    {
        [Fact]
        public void NamesTheUserAndTheirAddress()
        {
            string text = UserIdentityPrompt.Describe("Dana Levi", new[] { "dana@example.com" });
            Assert.StartsWith("The user is Dana Levi, and their email address is dana@example.com.", text);
            Assert.Contains("\"me\"", text);
        }

        [Fact]
        public void ListsEveryAccountOnce_PrimaryFirst()
        {
            string text = UserIdentityPrompt.Describe("Dana", new[] { "dana@work.com", "dana@gmail.com", "DANA@work.com" });
            Assert.Contains("email addresses are dana@work.com, dana@gmail.com (the first is the main one)", text);
        }

        [Fact]
        public void DoesNotRepeatAnAddressShownAsTheName_AndSkipsNonAddresses()
        {
            string text = UserIdentityPrompt.Describe("dana@gmail.com", new[] { "dana@gmail.com", "Dana Levi", "" });
            Assert.StartsWith("The user's email address is dana@gmail.com.", text);
        }

        [Fact]
        public void SaysNothingWhenNothingIsKnown()
        {
            Assert.Null(UserIdentityPrompt.Describe(null, new string[0]));
            Assert.Equal("The user's name is Dana.", UserIdentityPrompt.Describe("Dana", null));
        }
    }
}
