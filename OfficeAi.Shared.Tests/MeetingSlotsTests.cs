using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OfficeAi.Shared;

public class MeetingSlotsTests
{
    // A day whose FreeBusy string is all-free ('0') for the whole day.
    private static string AllFree() => new string('0', 48); // 48 half-hours

    // Marks [fromHour, toHour) busy ('2') in an otherwise-free day string.
    private static string BusyBetween(int fromHour, int toHour)
    {
        char[] c = AllFree().ToCharArray();
        for (int i = fromHour * 2; i < toHour * 2 && i < c.Length; i++) c[i] = '2';
        return new string(c);
    }

    private static readonly DateTime Anchor = new DateTime(2026, 9, 7); // a Monday, midnight
    private static readonly List<DateTime> OneDay = new List<DateTime> { new DateTime(2026, 9, 7) };

    [Fact]
    public void EveryoneFree_FirstSlotIsAtStartHour_AllAvailable()
    {
        var fb = new Dictionary<string, string> { ["a"] = AllFree(), ["b"] = AllFree() };
        var slots = MeetingSlots.Rank(fb, Anchor, OneDay, 9, 17, 30, 30, 5);

        Assert.Equal(new DateTime(2026, 9, 7, 9, 0, 0), slots[0].Start);
        Assert.Equal(new DateTime(2026, 9, 7, 9, 30, 0), slots[0].End);
        Assert.All(slots, s => Assert.Equal(2, s.Available));
        Assert.All(slots, s => Assert.Empty(s.Missing));
    }

    [Fact]
    public void OneAttendeeBusy_ThatWindowRanksLower_AndListsWhoIsBusy()
    {
        var fb = new Dictionary<string, string>
        {
            ["free"] = AllFree(),
            ["busy9to11"] = BusyBetween(9, 11),
        };
        // 60-min meeting, 09:00-12:00 window => candidate starts 9:00, 9:30, 10:00, 10:30, 11:00
        var slots = MeetingSlots.Rank(fb, Anchor, OneDay, 9, 12, 60, 30, 10);

        var at9 = slots.Single(s => s.Start.Hour == 9 && s.Start.Minute == 0);
        Assert.Equal(1, at9.Available);
        Assert.Contains("busy9to11", at9.Missing);

        var at11 = slots.Single(s => s.Start.Hour == 11 && s.Start.Minute == 0);
        Assert.Equal(2, at11.Available);

        // Best (all-free) slot ranks first.
        Assert.Equal(2, slots[0].Available);
    }

    [Fact]
    public void NoFullyFreeSlot_StillReturnsBestPartialMatch()
    {
        var fb = new Dictionary<string, string>
        {
            ["x"] = BusyBetween(9, 18), // busy all day
            ["y"] = BusyBetween(9, 10), // free after 10
        };
        var slots = MeetingSlots.Rank(fb, Anchor, OneDay, 9, 18, 30, 30, 3);
        Assert.NotEmpty(slots);
        Assert.Equal(1, slots[0].Available); // y is free, x never is
        Assert.Contains("x", slots[0].Missing);
    }

    [Fact]
    public void MeetingMustFitEntirelyInsideTheDayWindow()
    {
        var fb = new Dictionary<string, string> { ["a"] = AllFree() };
        // 90-min meeting, 09:00-11:00 window => last valid start is 09:30 (ends 11:00)
        var slots = MeetingSlots.Rank(fb, Anchor, OneDay, 9, 11, 90, 30, 10);
        Assert.Equal(new DateTime(2026, 9, 7, 9, 30, 0), slots.Max(s => s.Start));
    }

    [Fact]
    public void MultiDay_RanksAcrossAllDays_ByAvailabilityThenTime()
    {
        var mon = new DateTime(2026, 9, 7);
        var tue = new DateTime(2026, 9, 8);
        var fb = new Dictionary<string, string>
        {
            // Monday fully busy for 'a', Tuesday free
            ["a"] = BusyBetween(0, 24) + AllFree(),
            ["b"] = AllFree() + AllFree(),
        };
        var slots = MeetingSlots.Rank(fb, mon, new List<DateTime> { mon, tue }, 9, 17, 30, 30, 3);
        Assert.Equal(tue.Date, slots[0].Start.Date); // Tuesday 09:00, both free
        Assert.Equal(2, slots[0].Available);
    }

    [Fact]
    public void PastTheKnownFreeBusyWindow_IsTreatedAsFree()
    {
        var fb = new Dictionary<string, string> { ["short"] = "000000" }; // only 3 hours of data
        var slots = MeetingSlots.Rank(fb, Anchor, OneDay, 14, 16, 30, 30, 2);
        Assert.All(slots, s => Assert.Equal(1, s.Available));
    }

    [Fact]
    public void EndHourNotAfterStartHour_Throws()
    {
        var fb = new Dictionary<string, string> { ["a"] = AllFree() };
        Assert.Throws<ArgumentException>(() => MeetingSlots.Rank(fb, Anchor, OneDay, 12, 12, 30, 30, 5));
    }
}
