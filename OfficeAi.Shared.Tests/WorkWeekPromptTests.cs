using System;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class WorkWeekPromptTests
    {
        [Fact]
        public void Describe_ListsTheDaysInWeekOrder_AndAsksToConfirmTheHours()
        {
            string text = WorkWeekPrompt.Describe(
                new[] { DayOfWeek.Thursday, DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Wednesday, DayOfWeek.Tuesday }, 8, 17.5);

            Assert.Contains("Work days from the user's mailbox: Sunday, Monday, Tuesday, Wednesday, Thursday.", text);
            Assert.Contains("work hours are 08:00-17:30", text);
            Assert.Contains("ask the user to confirm their hours first", text);
            Assert.Contains("find_meeting_slots", text);
        }

        [Theory]
        [InlineData(124, new[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday })]
        [InlineData(248, new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday })]
        [InlineData(130, new[] { DayOfWeek.Sunday, DayOfWeek.Saturday })]
        [InlineData(0, new DayOfWeek[0])]
        public void DaysFromOutlookMask_ReadsOutlooksWorkDayBits(int mask, DayOfWeek[] expected)
        {
            Assert.Equal(expected, WorkWeekPrompt.DaysFromOutlookMask(mask));
        }

        [Fact]
        public void Unknown_TellsTheModelToAskInsteadOfAssumingMondayToFriday()
        {
            Assert.Contains("instead of assuming Monday to Friday", WorkWeekPrompt.Unknown());
        }

        [Theory]
        [InlineData(9, "09:00")]
        [InlineData(8.5, "08:30")]
        [InlineData(17.75, "17:45")]
        [InlineData(0, "00:00")]
        public void Hour_FormatsAsHoursAndMinutes(double hour, string expected)
        {
            Assert.Equal(expected, WorkWeekPrompt.Hour(hour));
        }
    }
}
