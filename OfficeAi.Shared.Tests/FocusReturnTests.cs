using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class FocusReturnTests
    {
        [Fact]
        public void ClickingTheDocumentDoesNotMoveFocusBack()
        {
            bool away = false;
            Assert.Equal(FocusReturnStep.Stop, FocusReturn.Next(true, ref away));
        }

        [Fact]
        public void WhileAnotherProgramIsInFrontItKeepsWaiting()
        {
            bool away = false;
            Assert.Equal(FocusReturnStep.Wait, FocusReturn.Next(false, ref away));
            Assert.Equal(FocusReturnStep.Wait, FocusReturn.Next(false, ref away));
            Assert.True(away);
        }

        [Fact]
        public void ComingBackAfterBeingAwayRestoresFocus()
        {
            bool away = false;
            FocusReturn.Next(false, ref away);
            Assert.Equal(FocusReturnStep.Restore, FocusReturn.Next(true, ref away));
        }
    }
}
