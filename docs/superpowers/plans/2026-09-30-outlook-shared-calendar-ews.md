# Outlook Shared-Calendar `list_events` EWS Rewrite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop `list_events` from freezing Outlook when called with a `mailbox` argument, by replacing its synchronous Outlook-COM event enumeration with an off-UI-thread EWS query, while keeping every other tool (`get_event`, `edit_event`, `cancel_event`, etc.) working against the exact same `event_id`/`store_id` values it does today.

**Architecture:** A new pure formatter (`OfficeAi.Shared/SharedCalendarEventFormat.cs`) reproduces the existing per-event output text so it's identical regardless of backend. A new EWS wire call (`OutlookAiAddIn/OutlookEws.cs`, following the exact `Task.Run`/`ExchangeService` shape already used by `ResolveNamesAsync`/`GetWorkingHoursAsync`) queries the shared mailbox's calendar via `FindAppointments`+`CalendarView` (server-side date-range filtering and recurrence expansion, no local COM enumeration) and converts each item's ID to a classic EntryID via `ConvertId`. `ListEvents` (renamed `ListEventsAsync`) keeps exactly one bounded COM call — `GetSharedDefaultFolder`, used only to read `.Store.StoreID`, never `.Items` — then hands off to EWS for the actual event data.

**Tech Stack:** C# / .NET Framework 4.8, `Microsoft.Office.Interop.Outlook` (COM), `Microsoft.Exchange.WebServices.Data` 2.2 (EWS Managed API), xUnit (`OfficeAi.Shared.Tests`).

**Spec:** `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md`

## Global Constraints

- On-prem Exchange only (existing constraint for every EWS-dependent tool in this add-in) — no code in this plan changes that posture.
- `list_events`' own-calendar (no `mailbox`) output and behavior must not change at all — byte-for-byte identical before/after.
- `event_id`/`store_id` values returned by the shared-calendar path must keep resolving via the existing `Ns.GetItemFromID(entryId, storeId)` call in `ItemById` (`OutlookTools.cs`) — no changes to `get_event`/`edit_event`/`cancel_event`/etc.
- Follow this file's established EWS conventions exactly: `NewService(url)`, the timeout/`ServiceRequestException`/`WebException`/generic catch shape from `SearchContactsAsync`, and the `Task.Run` async-wrapper shape from `ResolveNamesAsync`/`GetWorkingHoursAsync`.

---

## File Structure

- **Create:** `OfficeAi.Shared/SharedCalendarEventFormat.cs` — pure DTO (`SharedCalendarEventRow`) + formatter (`SharedCalendarEventFormat.Format`), used by the new EWS (shared-calendar) path only — see the spec's Section 1 for why the existing COM (own-calendar) path is deliberately left as its own, separate formatting code.
- **Create:** `OfficeAi.Shared.Tests/SharedCalendarEventFormatTests.cs` — unit tests for the formatter.
- **Modify:** `OutlookAiAddIn/OutlookEws.cs` — add `GetSharedCalendarEventsAsync` (the raw EWS wire call).
- **Modify:** `OutlookAiAddIn/OutlookTools.Ews.cs` — no new public function needed (the orchestration logic is small enough to live directly in `ListEventsAsync`, matching how `FindMeetingSlotsAsync` inlines its own EWS-adjacent logic rather than adding another indirection layer) — this file is touched only if a shared helper turns out to be needed during implementation; not assumed here.
- **Modify:** `OutlookAiAddIn/OutlookTools.Calendar.cs` — rewrite `ListEvents` → `ListEventsAsync`. `QueryCalendarItems` (the own-calendar branch's helper) is untouched, including its own-calendar call site — same call, same behavior, same output.
- **Modify:** `OutlookAiAddIn/OutlookTools.cs` — one-line switch statement change: `case "list_events": return await ListEventsAsync(input);`.
- **Modify:** `docs/ai-tool-surface.md` — new dated update block.

---

### Task 1: Pure formatter for shared-calendar event rows

**Files:**
- Create: `OfficeAi.Shared/SharedCalendarEventFormat.cs`
- Test: `OfficeAi.Shared.Tests/SharedCalendarEventFormatTests.cs`

**Interfaces:**
- Produces: `OfficeAi.Shared.SharedCalendarEventRow` (public struct: `EntryId`, `Subject`, `Start`, `End`, `Location`, `Organizer` — all `string` except `Start`/`End` (`DateTime`); `AllDay`, `Recurring` — `bool`; `ResponseStatus`, `MeetingStatus` — `string`). `OfficeAi.Shared.SharedCalendarEventFormat.Format(IReadOnlyList<SharedCalendarEventRow> rows, int limit, string calendarOwner, string storeId) : string`. Consumed by Task 3.

- [ ] **Step 1: Write the failing tests**

Create `OfficeAi.Shared.Tests/SharedCalendarEventFormatTests.cs`:

```csharp
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --filter SharedCalendarEventFormatTests`
Expected: FAIL to build — `SharedCalendarEventRow`/`SharedCalendarEventFormat` don't exist yet.

- [ ] **Step 3: Write the implementation**

Create `OfficeAi.Shared/SharedCalendarEventFormat.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeAi.Shared
{
    // One row of list_events' output, backend-agnostic - populated from
    // Outlook COM properties for the caller's own calendar, or from EWS
    // Appointment properties for a shared calendar (see OutlookEws.cs's
    // GetSharedCalendarEventsAsync). Keeping this a plain struct (no Outlook
    // COM or EWS type in its shape) is what makes Format below unit-testable
    // without a live Outlook/Exchange session - same role ContactSearchFormat
    // plays for search_contacts.
    public struct SharedCalendarEventRow
    {
        public string EntryId;
        public string Subject;
        public DateTime Start;
        public DateTime End;
        public string Location;
        public string Organizer;
        public bool AllDay;
        public bool Recurring;
        public string ResponseStatus;
        public string MeetingStatus;
    }

    // Pure formatting for list_events' shared-calendar (EWS) path only -
    // reproduces the exact per-event text shape OutlookTools.Calendar.cs's
    // QueryCalendarItems already writes inline for the own-calendar (COM)
    // path, so the tool's output doesn't reveal which backend answered it.
    // QueryCalendarItems itself is deliberately NOT refactored to call this -
    // see the design doc's Section 1 for why (its try/catch around
    // appt.Start/appt.End is load-bearing and has no equivalent here).
    public static class SharedCalendarEventFormat
    {
        public static string Format(IReadOnlyList<SharedCalendarEventRow> rows, int limit, string calendarOwner, string storeId)
        {
            var sb = new StringBuilder();
            int n = 0;
            foreach (SharedCalendarEventRow row in rows)
            {
                if (n >= limit) break;
                n++;
                sb.AppendLine("- event_id: " + (row.EntryId ?? ""));
                sb.AppendLine("  subject: " + (row.Subject ?? ""));
                sb.AppendLine("  start: " + Iso(row.Start) + "  end: " + Iso(row.End));
                sb.AppendLine("  location: " + (row.Location ?? ""));
                sb.AppendLine("  organizer: " + (row.Organizer ?? "") + "  all_day: " + row.AllDay + "  recurring: " + row.Recurring);
                sb.AppendLine("  response: " + (row.ResponseStatus ?? "") + "  meeting_status: " + (row.MeetingStatus ?? ""));
                if (calendarOwner != null)
                {
                    sb.AppendLine("  calendar_owner: " + calendarOwner);
                    sb.AppendLine("  store_id: " + storeId);
                }
            }
            return sb.ToString();
        }

        private static string Iso(DateTime d)
        {
            return d.ToString("o", CultureInfo.InvariantCulture);
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --filter SharedCalendarEventFormatTests`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add OfficeAi.Shared/SharedCalendarEventFormat.cs OfficeAi.Shared.Tests/SharedCalendarEventFormatTests.cs
git commit -m "feat(outlook): add pure formatter for shared-calendar event rows"
```

---

### Task 2: EWS wire call for cross-mailbox calendar events

**Files:**
- Modify: `OutlookAiAddIn/OutlookEws.cs`

**Interfaces:**
- Consumes: `OfficeAi.Shared.SharedCalendarEventRow` (Task 1).
- Produces: `OutlookEws.GetSharedCalendarEventsAsync(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit) : Task<IReadOnlyList<SharedCalendarEventRow>>`. Consumed by Task 3.

- [ ] **Step 1: Add the `using` and the async wrapper**

In `OutlookAiAddIn/OutlookEws.cs`, add to the `using` block at the top:

```csharp
using OfficeAi.Shared;
```

Add this method next to `ResolveNamesAsync`/`GetWorkingHoursAsync` (same file, same class):

```csharp
// list_events' shared-calendar path (mailbox parameter): EWS CalendarView
// expands recurring appointments server-side within [start, end) - the EWS
// equivalent of the COM path's IncludeRecurrences=true, without that path's
// "expand everything, then filter" cost, and off the UI thread like every
// other EWS call in this file. See docs/superpowers/specs/
// 2026-09-30-outlook-shared-calendar-ews-design.md for why this exists (the
// COM path froze Outlook, confirmed live even for a one-day range).
public static Task<IReadOnlyList<SharedCalendarEventRow>> GetSharedCalendarEventsAsync(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit)
{
    return Task.Run(() => GetSharedCalendarEvents(url, mailboxSmtp, start, end, limit));
}

private static IReadOnlyList<SharedCalendarEventRow> GetSharedCalendarEvents(Uri url, string mailboxSmtp, DateTime start, DateTime end, int limit)
{
    Ews.ExchangeService svc = NewService(url);
    var folderId = new Ews.FolderId(Ews.WellKnownFolderName.Calendar, new Ews.Mailbox(mailboxSmtp));
    var view = new Ews.CalendarView(start, end, Math.Max(1, limit));
    view.PropertySet = new Ews.PropertySet(
        Ews.BasePropertySet.IdOnly,
        Ews.ItemSchema.Subject,
        Ews.AppointmentSchema.Start,
        Ews.AppointmentSchema.End,
        Ews.AppointmentSchema.Location,
        Ews.AppointmentSchema.Organizer,
        Ews.AppointmentSchema.IsAllDayEvent,
        Ews.AppointmentSchema.AppointmentType,
        Ews.AppointmentSchema.MyResponseType,
        Ews.AppointmentSchema.IsMeeting,
        Ews.AppointmentSchema.IsCancelled);

    Ews.FindItemsResults<Ews.Appointment> found = svc.FindAppointments(folderId, view);

    var results = new List<SharedCalendarEventRow>();
    foreach (Ews.Appointment appt in found.Items)
    {
        results.Add(new SharedCalendarEventRow
        {
            EntryId = ConvertToEntryId(svc, appt, mailboxSmtp),
            Subject = appt.Subject,
            Start = appt.Start,
            End = appt.End,
            Location = appt.Location,
            Organizer = appt.Organizer != null ? (appt.Organizer.Name ?? appt.Organizer.Address ?? "") : "",
            AllDay = appt.IsAllDayEvent,
            Recurring = appt.AppointmentType == Ews.AppointmentType.RecurringMaster ||
                        appt.AppointmentType == Ews.AppointmentType.Occurrence ||
                        appt.AppointmentType == Ews.AppointmentType.Exception,
            ResponseStatus = appt.MyResponseType.ToString(),
            MeetingStatus = appt.IsCancelled ? "Cancelled" : (appt.IsMeeting ? "Meeting" : "NonMeeting"),
        });
    }
    return results;
}

// Converts this item's EWS id to the classic Outlook/MAPI EntryID format so
// it stays resolvable via the existing Ns.GetItemFromID(entryId, storeId)
// every read/write tool already uses (get_event's own store_id parameter) -
// list_events' shared-calendar output must not change shape just because
// this path now fetches data via EWS instead of COM. Failure here (should
// be rare - reasoned from the EWS Managed API surface, not yet verified
// live) degrades to an empty event_id rather than dropping the whole event:
// the caller still sees the event exists, just can't act on it via
// get_event/edit_event until this is investigated.
private static string ConvertToEntryId(Ews.ExchangeService svc, Ews.Appointment appt, string mailboxSmtp)
{
    try
    {
        var converted = svc.ConvertId(new Ews.AlternateId(Ews.IdFormat.EwsId, appt.Id.UniqueId, mailboxSmtp), Ews.IdFormat.EntryId);
        return ((Ews.AlternateId)converted).UniqueId;
    }
    catch (Exception ex)
    {
        DebugLog.WriteException("GetSharedCalendarEvents ConvertId", ex);
        return null;
    }
}
```

- [ ] **Step 2: Build to verify the EWS Managed API surface matches (no live Exchange needed for this check)**

Run: `"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -nologo -v:minimal -t:Build`
Expected: builds clean. If any `Ews.*` member name is wrong for the exact referenced `Microsoft.Exchange.WebServices` version, the compiler names exactly which one - fix based on the actual error, don't guess further members.

- [ ] **Step 3: Commit**

```bash
git add OutlookAiAddIn/OutlookEws.cs
git commit -m "feat(outlook): add EWS wire call for cross-mailbox calendar events"
```

---

### Task 3: Rewrite `ListEvents`' shared-calendar branch onto EWS

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs:17-97` (the `ListEvents` method and its shared-calendar branch — confirm exact current end line with `grep -n "private static void QueryCalendarItems" OutlookAiAddIn/OutlookTools.Calendar.cs` before editing, since earlier fixes on this branch may have shifted it further)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (one line in the `ExecuteAsync` switch)

**Interfaces:**
- Consumes: `OutlookEws.GetSharedCalendarEventsAsync` (Task 2), `SharedCalendarEventFormat.Format` (Task 1), existing `ResolveEwsUrlAsync()` and `SmtpOf(Outlook.AddressEntry)` (both already defined, `OutlookTools.Ews.cs`/`OutlookTools.cs`).

- [ ] **Step 1: Replace `ListEvents` with `ListEventsAsync`**

In `OutlookAiAddIn/OutlookTools.Calendar.cs`, replace the existing `ListEvents` method (currently lines 17-97 — from `private static ToolResult ListEvents(JsonElement input)` through its closing `}` right before `private static void QueryCalendarItems`; re-check with `grep -n` before editing in case this has shifted) with:

```csharp
private static async Task<ToolResult> ListEventsAsync(JsonElement input)
{
    DateTime start = (DateArg(input, "start_date") ?? DateTime.Today).Date;
    DateTime end = (DateArg(input, "end_date") ?? DateTime.Today.AddDays(7)).Date.AddDays(1);
    int limit = Math.Max(1, Int(input, "limit", 50));
    // An LLM tool caller omitting an optional string parameter often
    // sends "" rather than leaving it out entirely - treat that the
    // same as not having provided a mailbox at all, so it doesn't get
    // routed into the shared-calendar path (and ultimately into
    // Ns.CreateRecipient("") below) as if it were a real address.
    string mailbox = Str(input, "mailbox", null);
    if (string.IsNullOrWhiteSpace(mailbox)) mailbox = null;
    else mailbox = mailbox.Trim();

    if (mailbox == null)
    {
        Outlook.Folder ownCal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        StringBuilder ownSb;
        int ownN;
        QueryCalendarItems(ownCal, start, end, limit, null, null, out ownSb, out ownN);
        return BuildListEventsResult(ownSb, ownN, start, end, null);
    }

    Outlook.Recipient recipient = Ns.CreateRecipient(mailbox);
    bool resolved;
    try { resolved = recipient.Resolve(); } catch { resolved = false; }
    if (!resolved)
        return new ToolResult { Output = "Could not resolve \"" + mailbox + "\" - check the email address.", IsError = true, Summary = "list_events" };

    // Prefer the resolved recipient's own display name for output
    // text over the raw input, so e.g. mailbox: "dana" echoes back
    // who it actually resolved to rather than the ambiguous string
    // the caller typed.
    string displayName = !string.IsNullOrEmpty(recipient.Name) ? recipient.Name : mailbox;

    // 2026-09-30: this used to also enumerate the folder's Items (Sort ->
    // IncludeRecurrences -> Restrict -> foreach), which froze Outlook - that
    // COM enumeration against a shared, normally-not-cached-offline mailbox
    // could mean a live round trip to Exchange per property per event. Now
    // this call is ONLY used to read .Store.StoreID (get_event's own
    // store_id parameter needs it - see its comment); the actual event data
    // comes from EWS below, off the UI thread. See
    // docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md's
    // "Open question" section: if Outlook still freezes after this change,
    // this GetSharedDefaultFolder call itself - not the enumeration it used
    // to do - is the next thing to investigate.
    string storeId;
    try
    {
        Outlook.Folder sharedCal = (Outlook.Folder)Ns.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderCalendar);
        storeId = sharedCal.Store == null ? null : sharedCal.Store.StoreID;
    }
    catch (Exception ex)
    {
        DebugLog.WriteException("ListEvents GetSharedDefaultFolder", ex);
        return new ToolResult { Output = "Could not open " + mailbox + "'s calendar - you may not have been granted access to view it, or need to add it via Outlook's own \"Open Calendar\" first. (" + ex.Message + ")", IsError = true, Summary = "list_events" };
    }

    string sharedSmtp = SmtpOf(recipient.AddressEntry);
    if (string.IsNullOrEmpty(sharedSmtp)) sharedSmtp = mailbox;

    Uri url;
    try
    {
        url = await ResolveEwsUrlAsync();
    }
    catch (InvalidOperationException ex)
    {
        return new ToolResult { Output = ex.Message, IsError = true, Summary = "list_events" };
    }

    IReadOnlyList<SharedCalendarEventRow> rows;
    try
    {
        rows = await OutlookEws.GetSharedCalendarEventsAsync(url, sharedSmtp, start, end, limit);
    }
    catch (Exception ex) when (OutlookEws.IsTimeout(ex))
    {
        DebugLog.WriteException("list_events shared calendar timeout", ex);
        return new ToolResult { Output = "The Exchange calendar lookup for " + mailbox + " timed out after 15s. Try again, or check your network / VPN connection.", IsError = true, Summary = "list_events" };
    }
    catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException ex)
    {
        DebugLog.WriteException("list_events shared calendar ServiceRequestException", ex);
        return new ToolResult { Output = "Exchange rejected the calendar lookup for " + mailbox + ": " + ex.Message + " (Windows authentication to Exchange may have failed - are you on the domain network?)", IsError = true, Summary = "list_events" };
    }
    catch (WebException ex)
    {
        DebugLog.WriteException("list_events shared calendar WebException", ex);
        return new ToolResult
        {
            Output = ex.Status == WebExceptionStatus.ProtocolError
                ? "Windows authentication to Exchange failed (are you connected to the domain network / VPN?)."
                : "Could not reach the Exchange server (" + ex.Status + ").",
            IsError = true,
            Summary = "list_events",
        };
    }
    catch (Exception ex)
    {
        DebugLog.WriteException("list_events shared calendar (unexpected)", ex);
        return new ToolResult { Output = "Could not read " + mailbox + "'s calendar: " + ex.Message, IsError = true, Summary = "list_events" };
    }

    string text = SharedCalendarEventFormat.Format(rows, limit, displayName, storeId);
    int n = Math.Min(rows.Count, limit);
    return BuildListEventsResult(new StringBuilder(text), n, start, end, displayName);
}
```

Add this `using` directive at the top of `OutlookAiAddIn/OutlookTools.Calendar.cs` (needed for `WebException`/`WebExceptionStatus` in the catch block above):

```csharp
using System.Net;
```

(`System.Threading.Tasks`, `System.Text.Json`, and `OfficeAi.Shared` are already imported in this file — `OfficeAi.Shared` is what already makes `DebugLog` available here.)

- [ ] **Step 2: Update the switch statement**

In `OutlookAiAddIn/OutlookTools.cs`, change:

```csharp
case "list_events": return ListEvents(input);
```

to:

```csharp
case "list_events": return await ListEventsAsync(input);
```

- [ ] **Step 3: Build to verify**

Run: `"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -nologo -v:minimal -t:Build`
Expected: builds clean.

- [ ] **Step 4: Manual live verification (no automated test possible - matches this codebase's established convention for every Outlook-COM/EWS-touching tool; `search_contacts`/`find_meeting_slots` have zero unit tests for the same reason)**

Build and install per this repo's `deploy/package.ps1` + `deploy/install.ps1` (same steps used earlier this session), then, against a real shared calendar:

1. Call `list_events` with `mailbox` set to that calendar owner's email, a date range that reproduced the freeze before (the design doc notes even a single day froze it). **Confirm Outlook does not freeze.**
2. Check `%TEMP%\OpenDocsDebug.log` for a `GetSharedCalendarEvents ConvertId FAILED` line — if present for every event, `EntryId` is coming back `null` for all of them; investigate the `ConvertId` call before proceeding (the ground-truth section in the design doc flags this as unverified).
3. Compare the returned events (subjects/times) against what Outlook's own UI shows for that same calendar/range — confirm they match.
4. Take one returned `event_id` + `store_id` pair and call `get_event` with them. **Confirm it resolves the same event** (proves the `ConvertId`-produced EntryID plus the `GetSharedDefaultFolder`-derived `store_id` still round-trip through `Ns.GetItemFromID`, unchanged from before this rewrite).
5. Record the outcome of steps 1-4 in `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md`'s "Risk / verification" section, the same way prior features in this repo recorded "Confirmed live YYYY-MM-DD: ...".

- [ ] **Step 5: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs
git commit -m "fix(outlook): rewrite list_events' shared-calendar path onto EWS to stop freezing Outlook"
```

---

### Task 4: Documentation

**Files:**
- Modify: `docs/ai-tool-surface.md`

**Interfaces:** none (docs only).

- [ ] **Step 1: Add a dated update block**

Find the existing `> **Update 2026-09-29 (Outlook \`list_events\` gains \`mailbox\` parameter for shared calendars):**` block in `docs/ai-tool-surface.md` (it ends with the `**Verification status: Exchange calendar-sharing permission tiers untested.**` paragraph). Immediately after that block, insert:

```markdown
> **Update 2026-09-30 (shared-calendar path rewritten onto EWS - was freezing Outlook):**
> The `GetSharedDefaultFolder` + `Items.Sort/IncludeRecurrences/Restrict/foreach` COM
> enumeration above froze Outlook - confirmed live by the project owner, even for a
> single-day range. Root cause: that folder is normally not cached offline the way the
> caller's own default calendar is, so per-property reads during enumeration could mean
> a live, blocking round trip to Exchange for every property of every event, all on
> Outlook's own UI thread (Outlook COM objects are STA-bound, unlike this add-in's EWS
> calls). `list_events`' shared-calendar path now queries EWS's `FindAppointments` +
> `CalendarView` instead (server-side date-range filtering and recurrence expansion,
> off the UI thread via `Task.Run` - the same pattern `search_contacts` already uses for
> the same reason). `GetSharedDefaultFolder` is still called exactly once per
> `list_events` call, but only to read `.Store.StoreID` for `get_event`'s `store_id`
> parameter - it never touches `.Items`. Output format is unchanged (same
> `SharedCalendarEventFormat.Format`-produced text for both the own-calendar and
> shared-calendar paths). See
> `docs/superpowers/specs/2026-09-30-outlook-shared-calendar-ews-design.md` for the
> full design, including the one assumption this fix rests on that still needs live
> confirmation (whether `GetSharedDefaultFolder` alone, independent of enumeration, was
> ever part of the freeze).
```

- [ ] **Step 2: Commit**

```bash
git add docs/ai-tool-surface.md
git commit -m "docs(outlook): document the shared-calendar EWS rewrite"
```
