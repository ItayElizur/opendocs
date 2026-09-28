# Outlook Shared Calendar Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let `list_events` view another user's calendar via an optional `mailbox` parameter, reusing the existing query logic unchanged.

**Architecture:** A single new resolution branch in `ListEvents` (`OutlookAiAddIn/OutlookTools.Calendar.cs`): resolve `mailbox` to a `Recipient`, open their calendar folder via `NameSpace.GetSharedDefaultFolder`, then fall through into the exact same `Sort`→`IncludeRecurrences`→`Restrict` query and per-event output loop already used for the caller's own calendar.

**Tech Stack:** C# / .NET Framework 4.8, `Microsoft.Office.Interop.Outlook` COM interop, TypeScript (`entry.ts` tool schema).

**Spec:** `docs/superpowers/specs/2026-09-29-outlook-shared-calendar-design.md`

## Global Constraints

- `mailbox` is optional; omitting it preserves `list_events`' exact current behavior (own default calendar) with zero output-format change for that path.
- Every other tool (`get_event`, `edit_event`, `cancel_event`, etc.) is unchanged — acting on someone else's calendar event is out of scope for this plan.
- The existing `Sort("[Start]")` → `IncludeRecurrences = true` → `Restrict(filter)` order is load-bearing (per `ListEvents`' own existing comment) and must not be reordered for either the own-calendar or shared-calendar path.
- A resolution failure (`mailbox` doesn't resolve) or an access failure (`GetSharedDefaultFolder` throws) must return a clear `IsError` naming the mailbox and the likely cause — never let either fail silently or produce a confusing generic exception message.
- The Exchange sharing-permission-tier behavior (what Outlook actually returns for redacted fields) is unverified pending the project owner's own live test — do not add speculative special-case handling for it; only pass through whatever Outlook resolves.

---

### Task 1: Add `mailbox` parameter to `ListEvents`

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (the `ListEvents` method)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (the `list_events` tool schema)

**Interfaces:**
- Consumes: `Ns` (existing `Outlook.NameSpace` property), `Str`/`DateArg`/`Int`/`Iso` (existing JSON/format helpers in `OutlookTools.cs`), `DebugLog.WriteException` (existing).
- Produces: no new shared helpers — this is a self-contained change to `ListEvents` alone.

- [ ] **Step 1: Modify `ListEvents`**

Find this exact method in `OutlookAiAddIn/OutlookTools.Calendar.cs`:

```csharp
        private static ToolResult ListEvents(JsonElement input)
        {
            DateTime start = (DateArg(input, "start_date") ?? DateTime.Today).Date;
            DateTime end = (DateArg(input, "end_date") ?? DateTime.Today.AddDays(7)).Date.AddDays(1);
            int limit = Math.Max(1, Int(input, "limit", 50));

            Outlook.Folder cal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = cal.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            string filter = "[Start] <= '" + end.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [End] >= '" + start.ToString("g", CultureInfo.CurrentCulture) + "'";
            Outlook.Items restricted = items.Restrict(filter);

            var sb = new StringBuilder();
            int n = 0;
            foreach (object o in restricted)
            {
                if (n >= limit) break;
                Outlook.AppointmentItem appt = o as Outlook.AppointmentItem;
                if (appt == null) continue;
                n++;
                sb.AppendLine("- event_id: " + appt.EntryID);
                sb.AppendLine("  subject: " + (appt.Subject ?? ""));
                try { sb.AppendLine("  start: " + Iso(appt.Start) + "  end: " + Iso(appt.End)); } catch { }
                sb.AppendLine("  location: " + (appt.Location ?? ""));
                sb.AppendLine("  organizer: " + (appt.Organizer ?? "") + "  all_day: " + appt.AllDayEvent + "  recurring: " + appt.IsRecurring);
                sb.AppendLine("  response: " + appt.ResponseStatus + "  meeting_status: " + appt.MeetingStatus);
            }

            if (n == 0)
                return new ToolResult { Output = "No events between " + start.ToShortDateString() + " and " + end.AddDays(-1).ToShortDateString() + ".", Summary = "list_events" };
            return new ToolResult { Output = sb + "\n(Recurring instances share the master event_id.)", Summary = "list_events" };
        }
```

Replace it with:

```csharp
        private static ToolResult ListEvents(JsonElement input)
        {
            DateTime start = (DateArg(input, "start_date") ?? DateTime.Today).Date;
            DateTime end = (DateArg(input, "end_date") ?? DateTime.Today.AddDays(7)).Date.AddDays(1);
            int limit = Math.Max(1, Int(input, "limit", 50));
            string mailbox = Str(input, "mailbox", null);

            Outlook.Folder cal;
            if (mailbox != null)
            {
                Outlook.Recipient recipient = Ns.CreateRecipient(mailbox);
                bool resolved;
                try { resolved = recipient.Resolve(); } catch { resolved = false; }
                if (!resolved)
                    return new ToolResult { Output = "Could not resolve \"" + mailbox + "\" - check the email address.", IsError = true, Summary = "list_events" };
                try
                {
                    cal = (Outlook.Folder)Ns.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderCalendar);
                }
                catch (Exception ex)
                {
                    DebugLog.WriteException("ListEvents GetSharedDefaultFolder", ex);
                    return new ToolResult { Output = "Could not open " + mailbox + "'s calendar - you may not have been granted access to view it, or need to add it via Outlook's own \"Open Calendar\" first. (" + ex.Message + ")", IsError = true, Summary = "list_events" };
                }
            }
            else
            {
                cal = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            }

            Outlook.Items items = cal.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            string filter = "[Start] <= '" + end.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [End] >= '" + start.ToString("g", CultureInfo.CurrentCulture) + "'";
            Outlook.Items restricted = items.Restrict(filter);

            var sb = new StringBuilder();
            int n = 0;
            foreach (object o in restricted)
            {
                if (n >= limit) break;
                Outlook.AppointmentItem appt = o as Outlook.AppointmentItem;
                if (appt == null) continue;
                n++;
                sb.AppendLine("- event_id: " + appt.EntryID);
                sb.AppendLine("  subject: " + (appt.Subject ?? ""));
                try { sb.AppendLine("  start: " + Iso(appt.Start) + "  end: " + Iso(appt.End)); } catch { }
                sb.AppendLine("  location: " + (appt.Location ?? ""));
                sb.AppendLine("  organizer: " + (appt.Organizer ?? "") + "  all_day: " + appt.AllDayEvent + "  recurring: " + appt.IsRecurring);
                sb.AppendLine("  response: " + appt.ResponseStatus + "  meeting_status: " + appt.MeetingStatus);
                if (mailbox != null) sb.AppendLine("  calendar_owner: " + mailbox);
            }

            string whoseCalendar = mailbox != null ? mailbox + "'s calendar " : "";
            if (n == 0)
                return new ToolResult { Output = "No events on " + whoseCalendar + "between " + start.ToShortDateString() + " and " + end.AddDays(-1).ToShortDateString() + ".", Summary = "list_events" };
            return new ToolResult { Output = sb + "\n(Recurring instances share the master event_id.)", Summary = "list_events" };
        }
```

Note: `whoseCalendar` deliberately ends with a trailing space (used only in the `mailbox != null` case, folded into one string so the `n == 0` sentence reads correctly either way - e.g. "No events on alice@example.com's calendar between ..." vs. "No events between ...").

- [ ] **Step 2: Add `mailbox` to the `list_events` schema in `entry.ts`**

Find this exact block in `OutlookAiAddIn/web-src/entry.ts`:

```typescript
    name: 'list_events',
    description:
      'Lists calendar events in a date range (expands recurring meetings). Returns event_id, subject, start/end, location, organizer, and your response status. Recurring instances share the master event_id.',
    inputSchema: {
      type: 'object',
      properties: {
        start_date: { type: 'string', description: 'Range start, inclusive (YYYY-MM-DD). Default today.' },
        end_date: { type: 'string', description: 'Range end, inclusive (YYYY-MM-DD). Default +7 days.' },
        limit: { type: 'number', description: 'Default 50.' },
      },
      required: [],
    },
  },
```

Replace it with:

```typescript
    name: 'list_events',
    description:
      'Lists calendar events in a date range (expands recurring meetings). Returns event_id, subject, start/end, location, organizer, and your response status. Recurring instances share the master event_id. Pass mailbox to view someone else\'s calendar instead of your own, if they\'ve granted you access to it in Exchange - visibility depends on what sharing level they set (full details, free/busy only, or none), and Outlook enforces that automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        start_date: { type: 'string', description: 'Range start, inclusive (YYYY-MM-DD). Default today.' },
        end_date: { type: 'string', description: 'Range end, inclusive (YYYY-MM-DD). Default +7 days.' },
        limit: { type: 'number', description: 'Default 50.' },
        mailbox: { type: 'string', description: 'Email address of the calendar owner to view. Omit to view your own calendar.' },
      },
      required: [],
    },
  },
```

- [ ] **Step 3: Build to confirm it compiles**

Run (from `C:\dev\opendocs`):
```
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Release -nologo -v:minimal
```
If it fails with a NuGet `RuntimeIdentifier` error, run the same command with `-t:Restore` first, then retry `-t:Build`.

Expected: 0 errors.

Run (from `OutlookAiAddIn/`):
```
./node_modules/.bin/esbuild.cmd web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap
```
Expected: success, no errors.

- [ ] **Step 4: Run the shared test suite (regression check)**

Run (from `C:\dev\opendocs`):
```
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo
```
Expected: 216/216 passing (this task touches no code the suite covers).

- [ ] **Step 5: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add mailbox parameter to list_events for viewing shared calendars"
```

---

### Task 2: Update `docs/ai-tool-surface.md`

**Files:**
- Modify: `docs/ai-tool-surface.md`

**Interfaces:**
- Consumes: nothing — documentation only, describing Task 1's already-committed code.
- Produces: nothing.

- [ ] **Step 1: Find and update `list_events`' table row**

Run: `grep -n "\`list_events\`" docs/ai-tool-surface.md` to find its current table row, then update it to describe the new `mailbox` parameter: resolves via `Ns.CreateRecipient(mailbox).Resolve()`, opens the folder via `Ns.GetSharedDefaultFolder`, reuses the identical `Sort`/`IncludeRecurrences`/`Restrict` query, adds a `calendar_owner` line per event when `mailbox` is given. Explicitly note the sharing-permission-tier behavior is unverified live as of this addition (matching this file's existing convention for flagging unverified behavior elsewhere).

- [ ] **Step 2: Add a dated update block**

Add a `**Update <today's date>:**` block (matching this file's existing convention for dated update blocks) summarizing the addition and explicitly naming the one unverified risk area (Exchange sharing-tier behavior) for a future reader.

- [ ] **Step 3: Commit**

```bash
git add docs/ai-tool-surface.md
git commit -m "docs(outlook): document list_events' mailbox parameter"
```

---

### Task 3: Build, test, package for the project owner's own live verification

**Files:** none (build/test/package only).

**Interfaces:** none — verification task.

- [ ] **Step 1: Run the full C# test suite**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo` (from `C:\dev\opendocs`).

Expected: 216/216 passing.

- [ ] **Step 2: Rebuild the add-in**

Run both commands from Task 1 Step 3 again from a clean state to confirm the final combined diff builds.

Expected: both succeed with 0 errors.

- [ ] **Step 3: Package**

```bash
cd deploy
powershell -ExecutionPolicy Bypass -File ./package.ps1 -App Outlook
```

Expected: package completes, produces `deploy/dist`.

- [ ] **Step 4: Hand off**

Tell the project owner the package is ready at `deploy/dist` and that they said they'd test this (specifically the Exchange sharing-permission-tier behavior — full details vs. free/busy-only vs. no access) against a real shared calendar elsewhere, since no second mailbox with configured sharing was available during development. Do not run `install.ps1` automatically on this machine unless asked — the owner may be deploying this package somewhere else specifically to get a real second-mailbox test environment.

No commit for this task — it's verification/packaging only.
