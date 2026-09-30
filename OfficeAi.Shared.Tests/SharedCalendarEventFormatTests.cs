using System;
using System.Collections.Generic;
using Xunit;
using OfficeAi.Shared;

public class SharedCalendarEventFormatTests
{
    private static SharedCalendarEventRow Row(
        string entryId = "ID1", string subject = "Sync", string location = "Room 1",
        string organizer = "alice@x.com", bool allDay = false, bool recurring = false,
        string response = "Accept", string meetingStatus = "Meeting")
    {
        return new SharedCalendarEventRow
        {
            EntryId = entryId,
            Subject = subject,
            Start = new DateTime(2026, 10, 1, 9, 0, 0, DateTimeKind.Utc),
            End = new DateTime(2026, 10, 1, 10, 0, 0, DateTimeKind.Utc),
            Location = location,
            Organizer = organizer,
            AllDay = allDay,
            Recurring = recurring,
            ResponseStatus = response,
            MeetingStatus = meetingStatus,
        };
    }

    [Fact]
    public void SingleRow_NoCalendarOwner_MatchesOwnCalendarShape()
    {
        var rows = new List<SharedCalendarEventRow> { Row() };
        string outp = SharedCalendarEventFormat.Format(rows, 50, null, null);
        Assert.Equal(
            "- event_id: ID1\r\n" +
            "  subject: Sync\r\n" +
            "  start: 2026-10-01T09:00:00.0000000Z  end: 2026-10-01T10:00:00.0000000Z\r\n" +
            "  location: Room 1\r\n" +
            "  organizer: alice@x.com  all_day: False  recurring: False\r\n" +
            "  response: Accept  meeting_status: Meeting\r\n",
            outp);
    }

    [Fact]
    public void WithCalendarOwner_AppendsCalendarOwnerAndStoreIdLines()
    {
        var rows = new List<SharedCalendarEventRow> { Row() };
        string outp = SharedCalendarEventFormat.Format(rows, 50, "Dana Cohen", "STORE123");
        Assert.Equal(
            "- event_id: ID1\r\n" +
            "  subject: Sync\r\n" +
            "  start: 2026-10-01T09:00:00.0000000Z  end: 2026-10-01T10:00:00.0000000Z\r\n" +
            "  location: Room 1\r\n" +
            "  organizer: alice@x.com  all_day: False  recurring: False\r\n" +
            "  response: Accept  meeting_status: Meeting\r\n" +
            "  calendar_owner: Dana Cohen\r\n" +
            "  store_id: STORE123\r\n",
            outp);
    }

    [Fact]
    public void EmptyRows_ReturnsEmptyString()
    {
        Assert.Equal("", SharedCalendarEventFormat.Format(new List<SharedCalendarEventRow>(), 50, null, null));
    }

    [Fact]
    public void LimitTruncatesRows()
    {
        var rows = new List<SharedCalendarEventRow> { Row("A"), Row("B"), Row("C") };
        string outp = SharedCalendarEventFormat.Format(rows, 2, null, null);
        Assert.Contains("event_id: A", outp);
        Assert.Contains("event_id: B", outp);
        Assert.DoesNotContain("event_id: C", outp);
    }

    [Fact]
    public void NullSubjectLocationOrganizer_FormatsAsEmptyNotThrow()
    {
        var row = Row(subject: null, location: null, organizer: null);
        string outp = SharedCalendarEventFormat.Format(new List<SharedCalendarEventRow> { row }, 50, null, null);
        Assert.Contains("  subject: \r\n", outp);
        Assert.Contains("  location: \r\n", outp);
        Assert.Contains("organizer:   all_day:", outp);
    }
}
