using System;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class MeetingArgsTests
    {
        [Fact]
        public void OccurrenceStart_DateAloneTakesTheSeriesTimeOfDay()
        {
            Assert.Equal(new DateTime(2026, 11, 24, 10, 0, 0), MeetingArgs.OccurrenceStart(new DateTime(2026, 11, 24), new TimeSpan(10, 0, 0)));
        }

        [Fact]
        public void OccurrenceStart_ADateWithATimeIsKept()
        {
            var given = new DateTime(2026, 11, 24, 14, 30, 0);
            Assert.Equal(given, MeetingArgs.OccurrenceStart(given, new TimeSpan(10, 0, 0)));
        }

        [Fact]
        public void TargetNeedsAction_AnIdWithNoActionOrNewIsRejected()
        {
            Assert.True(MeetingArgs.TargetNeedsAction("new", "abc", null));
            Assert.True(MeetingArgs.TargetNeedsAction(null, null, "abc"));
        }

        [Fact]
        public void TargetNeedsAction_EmptyIdsOrARealActionAreFine()
        {
            Assert.False(MeetingArgs.TargetNeedsAction("new", "", null));
            Assert.False(MeetingArgs.TargetNeedsAction("reply", "abc", null));
            Assert.False(MeetingArgs.TargetNeedsAction("edit", "abc"));
        }
    }
}
