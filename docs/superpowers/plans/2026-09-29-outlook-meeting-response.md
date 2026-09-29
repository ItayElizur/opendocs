# Outlook Meeting Response Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `tentative_meeting` alongside `accept_meeting`/`decline_meeting`, an optional `message` comment on all three, and `draft_` counterparts for each.

**Architecture:** Generalize the existing `RespondMeeting` helper (`OutlookAiAddIn/OutlookTools.Calendar.cs`) from a `bool accept` parameter to a full `OlMeetingResponse` parameter plus an optional message, add a new `DraftRespondMeeting` helper with the same signature shape that opens the response for review instead of sending it, and wire six tool names (3 immediate + 3 draft) through both.

**Tech Stack:** C# / .NET Framework 4.8, `Microsoft.Office.Interop.Outlook` COM interop, TypeScript (`entry.ts` tool schemas).

**Spec:** `docs/superpowers/specs/2026-09-29-outlook-meeting-response-design.md`

## Global Constraints

- Three response tools (`accept_meeting`, `decline_meeting`, `tentative_meeting`), never one tool with a response-type parameter — matches this codebase's established split-tool convention.
- Each of the three, plus each of their `draft_` counterparts, takes `event_id` (required) and `message` (optional).
- Immediate tools (`accept_meeting`/`decline_meeting`/`tentative_meeting`) stay in `ApprovalTierTools`/"Automate approvals" (they already call `.Send()`) — do not move them to Full Autonomy.
- Draft tools (`draft_accept_meeting`/`draft_decline_meeting`/`draft_tentative_meeting`) go in `DraftTierTools`/"Draft only" — one tier below their immediate counterparts, matching every other draft/immediate pair in this codebase.
- `draft_*_meeting` tools never call `.Send()`, never call `RecordIrreversible`/`RecordSnapshot` — nothing persists until the user acts from the opened review window, matching `draft_event`'s existing contract.
- No forwarding tool of any kind — explicitly out of scope per the spec.
- The `message`/comment mechanism (`resp.Body = message` before `.Send()`) is unverified live pending the project owner's own test against a real received invite — implement it exactly as specified, but do not claim it's confirmed working in any commit message or doc update.

---

### Task 1: Generalize `RespondMeeting`, add `tentative_meeting`

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (the `RespondMeeting` method)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (switch cases, `ApprovalTierTools`)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (`accept_meeting`/`decline_meeting` schemas get `message`; new `tentative_meeting` schema; `OUTLOOK_TOOL_DISPLAY`; `trackChangesExtraTools`)

**Interfaces:**
- Consumes: `ReqStr`/`Str`/`ItemById` (existing helpers), `RecordIrreversible` (existing, `OutlookTools.Undo.cs`), `DebugLog.WriteException` (existing).
- Produces: `RespondMeeting(string mbxKey, JsonElement input, Outlook.OlMeetingResponse response, string toolName) -> ToolResult` (signature change from the current `bool accept` - Task 2's `DraftRespondMeeting` mirrors this shape).

- [ ] **Step 1: Replace `RespondMeeting`**

Find this exact method in `OutlookAiAddIn/OutlookTools.Calendar.cs`:

```csharp
        private static ToolResult RespondMeeting(string mbxKey, JsonElement input, bool accept)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);

            Outlook.AppointmentItem appt = item as Outlook.AppointmentItem;
            if (appt == null)
            {
                Outlook.MeetingItem mi = item as Outlook.MeetingItem;
                if (mi != null) appt = mi.GetAssociatedAppointment(false);
            }
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = accept ? "accept_meeting" : "decline_meeting" };

            Outlook.OlMeetingResponse response = accept
                ? Outlook.OlMeetingResponse.olMeetingAccepted
                : Outlook.OlMeetingResponse.olMeetingDeclined;
            object respObj = appt.Respond(response, true, false);
            Outlook.MeetingItem resp = respObj as Outlook.MeetingItem;
            if (resp != null)
            {
                try { resp.Send(); } catch (Exception ex) { DebugLog.WriteException("RespondMeeting Send", ex); }
            }
            RecordIrreversible(mbxKey, (accept ? "accept_meeting" : "decline_meeting") + " for \"" + (appt.Subject ?? "") + "\"");
            return new ToolResult
            {
                Output = (accept ? "Accepted: " : "Declined: ") + (appt.Subject ?? ""),
                Mutated = true,
                Summary = accept ? "accept_meeting" : "decline_meeting",
            };
        }
```

Replace it with:

```csharp
        // response is the actual OlMeetingResponse to send (not just a bool)
        // so this one helper covers all three: accept_meeting, decline_meeting,
        // tentative_meeting. message is an optional comment attached to the
        // response before it's sent - UNVERIFIED live whether the organizer
        // actually sees this text on the delivered response (needs a real
        // received invite to test, not a self-organized item).
        private static ToolResult RespondMeeting(string mbxKey, JsonElement input, Outlook.OlMeetingResponse response, string toolName)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);

            Outlook.AppointmentItem appt = item as Outlook.AppointmentItem;
            if (appt == null)
            {
                Outlook.MeetingItem mi = item as Outlook.MeetingItem;
                if (mi != null) appt = mi.GetAssociatedAppointment(false);
            }
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = toolName };

            string message = Str(input, "message", null);

            object respObj = appt.Respond(response, true, false);
            Outlook.MeetingItem resp = respObj as Outlook.MeetingItem;
            if (resp != null)
            {
                if (message != null) resp.Body = message;
                try { resp.Send(); } catch (Exception ex) { DebugLog.WriteException(toolName + " Send", ex); }
            }
            RecordIrreversible(mbxKey, toolName + " for \"" + (appt.Subject ?? "") + "\"" + (message != null ? " with a comment" : ""));
            string verb = response == Outlook.OlMeetingResponse.olMeetingAccepted ? "Accepted"
                        : response == Outlook.OlMeetingResponse.olMeetingTentative ? "Responded tentatively to"
                        : "Declined";
            return new ToolResult
            {
                Output = verb + ": " + (appt.Subject ?? "") + (message != null ? " (comment sent)" : ""),
                Mutated = true,
                Summary = toolName,
            };
        }
```

- [ ] **Step 2: Update the switch cases and add `tentative_meeting`**

Find this exact block in `OutlookAiAddIn/OutlookTools.cs`:

```csharp
                    case "accept_meeting": return RespondMeeting(mbxKey, input, true);
                    case "decline_meeting": return RespondMeeting(mbxKey, input, false);
```

Replace it with:

```csharp
                    case "accept_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingAccepted, "accept_meeting");
                    case "decline_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingDeclined, "decline_meeting");
                    case "tentative_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingTentative, "tentative_meeting");
```

- [ ] **Step 3: Add `tentative_meeting` to `ApprovalTierTools`**

Find this exact block in `OutlookAiAddIn/OutlookTools.cs`:

```csharp
        private static readonly HashSet<string> ApprovalTierTools = new HashSet<string>
        {
            "accept_meeting", "decline_meeting",
        };
```

Replace it with:

```csharp
        private static readonly HashSet<string> ApprovalTierTools = new HashSet<string>
        {
            "accept_meeting", "decline_meeting", "tentative_meeting",
        };
```

- [ ] **Step 4: Update `entry.ts` schemas for `accept_meeting`/`decline_meeting`, add `tentative_meeting`**

Find this exact block in `OutlookAiAddIn/web-src/entry.ts`:

```typescript
    name: 'accept_meeting',
    description: 'Accepts a meeting invitation and notifies the organizer.',
    inputSchema: { type: 'object', properties: { event_id: { type: 'string' } }, required: ['event_id'] },
  },
  {
    name: 'decline_meeting',
    description: 'Declines a meeting invitation and notifies the organizer.',
    inputSchema: { type: 'object', properties: { event_id: { type: 'string' } }, required: ['event_id'] },
  },
```

Replace it with:

```typescript
    name: 'accept_meeting',
    description: 'Accepts a meeting invitation and notifies the organizer. Pass message to add a comment to the response.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment sent to the organizer along with the response.' } },
      required: ['event_id'],
    },
  },
  {
    name: 'decline_meeting',
    description: 'Declines a meeting invitation and notifies the organizer. Pass message to add a comment to the response.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment sent to the organizer along with the response.' } },
      required: ['event_id'],
    },
  },
  {
    name: 'tentative_meeting',
    description: 'Responds tentatively to a meeting invitation and notifies the organizer. Pass message to add a comment to the response.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment sent to the organizer along with the response.' } },
      required: ['event_id'],
    },
  },
```

- [ ] **Step 5: Update `OUTLOOK_TOOL_DISPLAY` and `trackChangesExtraTools`**

Find this exact line in `OutlookAiAddIn/web-src/entry.ts`:

```typescript
  accept_meeting: d('Accept meeting', 'אישור פגישה', 'Accepts a meeting invitation.', 'מאשר הזמנה לפגישה.'),
  decline_meeting: d('Decline meeting', 'דחיית פגישה', 'Declines a meeting invitation.', 'דוחה הזמנה לפגישה.'),
```

Replace it with:

```typescript
  accept_meeting: d('Accept meeting', 'אישור פגישה', 'Accepts a meeting invitation.', 'מאשר הזמנה לפגישה.'),
  decline_meeting: d('Decline meeting', 'דחיית פגישה', 'Declines a meeting invitation.', 'דוחה הזמנה לפגישה.'),
  tentative_meeting: d('Tentative response', 'תגובה זמנית', 'Responds tentatively to a meeting invitation.', 'משיב תשובה זמנית להזמנה לפגישה.'),
```

Find this exact line in `OutlookAiAddIn/web-src/entry.ts`:

```typescript
  trackChangesExtraTools: ['accept_meeting', 'decline_meeting'],
```

Replace it with:

```typescript
  trackChangesExtraTools: ['accept_meeting', 'decline_meeting', 'tentative_meeting'],
```

- [ ] **Step 6: Build to confirm it compiles**

Run (from `C:\dev\opendocs`):
```
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" OutlookAiAddIn/OutlookAiAddIn.csproj -t:Build -p:Configuration=Release -nologo -v:minimal
```
If it fails with a NuGet `RuntimeIdentifier` error, run the same command with `-t:Restore` first, then retry `-t:Build`.

Run (from `OutlookAiAddIn/`):
```
./node_modules/.bin/esbuild.cmd web-src/entry.ts --bundle --outfile=web/bundle.js --alias:@genoffice/agent-core=../shared/web-src/agent-core/index.ts --alias:@genoffice/ai-provider=../shared/web-src/ai-provider/index.ts --alias:@officeai/chat-ui=../shared/chat-ui/chat-ui.ts --alias:@officeai/app-shell=../shared/web-src/app-shell/index.ts --target=chrome100 --format=iife --sourcemap
```

Expected: both succeed with 0 errors.

- [ ] **Step 7: Run the shared test suite (regression check)**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo` (from `C:\dev\opendocs`).

Expected: 216/216 passing.

- [ ] **Step 8: Commit**

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add tentative_meeting, optional comment on meeting responses"
```

---

### Task 2: Add `DraftRespondMeeting` and the three `draft_*_meeting` tools

**Files:**
- Modify: `OutlookAiAddIn/OutlookTools.Calendar.cs` (new `DraftRespondMeeting` method)
- Modify: `OutlookAiAddIn/OutlookTools.cs` (switch cases, `DraftTierTools`)
- Modify: `OutlookAiAddIn/web-src/entry.ts` (three new schemas, `OUTLOOK_TOOL_DISPLAY`, `commentOnlyExtraTools`)

**Interfaces:**
- Consumes: `ReqStr`/`Str`/`ItemById` (existing helpers) - does NOT consume `RecordIrreversible`/`RecordSnapshot` (draft tools never persist).
- Produces: `DraftRespondMeeting(JsonElement input, Outlook.OlMeetingResponse response, string toolName) -> ToolResult`.

- [ ] **Step 1: Add `DraftRespondMeeting`**

File: `OutlookAiAddIn/OutlookTools.Calendar.cs`. Insert immediately after the `RespondMeeting` method (Task 1's replacement):

```csharp

        // Draft-tier counterpart to RespondMeeting: same appt.Respond() call
        // (NoUIFlag=true, SendResponse=false - never sent by Outlook itself),
        // but opens the response for the user to review/edit and send
        // themselves instead of calling .Send() here. Never records
        // undo/redo - draft tools never persist anything, same contract as
        // draft_event.
        private static ToolResult DraftRespondMeeting(JsonElement input, Outlook.OlMeetingResponse response, string toolName)
        {
            string id = ReqStr(input, "event_id");
            object item = ItemById(id, null);

            Outlook.AppointmentItem appt = item as Outlook.AppointmentItem;
            if (appt == null)
            {
                Outlook.MeetingItem mi = item as Outlook.MeetingItem;
                if (mi != null) appt = mi.GetAssociatedAppointment(false);
            }
            if (appt == null)
                return new ToolResult { Output = "event_id does not resolve to a meeting.", IsError = true, Summary = toolName };

            string message = Str(input, "message", null);

            object respObj = appt.Respond(response, true, false);
            Outlook.MeetingItem resp = respObj as Outlook.MeetingItem;
            if (resp != null)
            {
                if (message != null) resp.Body = message;
                resp.Display(false);
            }

            string verbing = response == Outlook.OlMeetingResponse.olMeetingAccepted ? "an acceptance"
                            : response == Outlook.OlMeetingResponse.olMeetingTentative ? "a tentative response"
                            : "a decline";
            return new ToolResult
            {
                Output = "Opened " + verbing + " to \"" + (appt.Subject ?? "") + "\" in Outlook for the user to review and send." + (message != null ? " Comment pre-filled." : ""),
                Summary = toolName,
            };
        }
```

- [ ] **Step 2: Add the three switch cases**

Find this exact block in `OutlookAiAddIn/OutlookTools.cs` (Task 1's Step 2 replacement, now already in place):

```csharp
                    case "accept_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingAccepted, "accept_meeting");
                    case "decline_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingDeclined, "decline_meeting");
                    case "tentative_meeting": return RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingTentative, "tentative_meeting");
```

Insert immediately before it (draft-tier cases sit alongside this project's other `draft_*` cases earlier in the switch, but adding them here - right before their immediate counterparts - is equally correct and simpler to locate; group them with the other `draft_*` cases if you're already touching that block for clarity, implementer's choice):

```csharp
                    case "draft_accept_meeting": return DraftRespondMeeting(input, Outlook.OlMeetingResponse.olMeetingAccepted, "draft_accept_meeting");
                    case "draft_decline_meeting": return DraftRespondMeeting(input, Outlook.OlMeetingResponse.olMeetingDeclined, "draft_decline_meeting");
                    case "draft_tentative_meeting": return DraftRespondMeeting(input, Outlook.OlMeetingResponse.olMeetingTentative, "draft_tentative_meeting");
```

- [ ] **Step 3: Add the three draft tools to `DraftTierTools`**

Find this exact line in `OutlookAiAddIn/OutlookTools.cs`:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_cancel_event", "draft_edit_event",
```

Replace it with:

```csharp
            "set_event_categories", "set_category_color", "set_event_availability", "apply_search", "draft_cancel_event", "draft_edit_event",
            "draft_accept_meeting", "draft_decline_meeting", "draft_tentative_meeting",
```

- [ ] **Step 4: Add the three schemas to `entry.ts`**

File: `OutlookAiAddIn/web-src/entry.ts`. Find the `tentative_meeting` block added in Task 1 Step 4 (its closing `},`), and insert immediately after it:

```typescript
  {
    name: 'draft_accept_meeting',
    description: 'Opens an acceptance of a meeting invitation for the user to review and send themselves. Pass message to pre-fill a comment.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment pre-filled in the response.' } },
      required: ['event_id'],
    },
  },
  {
    name: 'draft_decline_meeting',
    description: 'Opens a decline of a meeting invitation for the user to review and send themselves. Pass message to pre-fill a comment.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment pre-filled in the response.' } },
      required: ['event_id'],
    },
  },
  {
    name: 'draft_tentative_meeting',
    description: 'Opens a tentative response to a meeting invitation for the user to review and send themselves. Pass message to pre-fill a comment.',
    inputSchema: {
      type: 'object',
      properties: { event_id: { type: 'string' }, message: { type: 'string', description: 'Optional comment pre-filled in the response.' } },
      required: ['event_id'],
    },
  },
```

- [ ] **Step 5: Add `OUTLOOK_TOOL_DISPLAY` entries and `commentOnlyExtraTools` entries**

Find this exact line in `OutlookAiAddIn/web-src/entry.ts` (added in Task 1 Step 5):

```typescript
  tentative_meeting: d('Tentative response', 'תגובה זמנית', 'Responds tentatively to a meeting invitation.', 'משיב תשובה זמנית להזמנה לפגישה.'),
```

Insert immediately after it:

```typescript
  draft_accept_meeting: d('Draft acceptance', 'טיוטת אישור', 'Opens a meeting acceptance to review and send.', 'פותח אישור פגישה לבדיקה ולשליחה.'),
  draft_decline_meeting: d('Draft decline', 'טיוטת דחייה', 'Opens a meeting decline to review and send.', 'פותח דחיית פגישה לבדיקה ולשליחה.'),
  draft_tentative_meeting: d('Draft tentative response', 'טיוטת תגובה זמנית', 'Opens a tentative response to review and send.', 'פותח תגובה זמנית לבדיקה ולשליחה.'),
```

Find this exact line in `OutlookAiAddIn/web-src/entry.ts`:

```typescript
    'draft_event',
    'draft_cancel_event',
    'draft_edit_event',
```

Replace it with:

```typescript
    'draft_event',
    'draft_cancel_event',
    'draft_edit_event',
    'draft_accept_meeting',
    'draft_decline_meeting',
    'draft_tentative_meeting',
```

- [ ] **Step 6: Build, test, commit**

Same build/test commands as Task 1 Steps 6-7. Expect 0 errors, 216/216.

```bash
git add OutlookAiAddIn/OutlookTools.Calendar.cs OutlookAiAddIn/OutlookTools.cs OutlookAiAddIn/web-src/entry.ts
git commit -m "feat(outlook): add draft_accept_meeting/draft_decline_meeting/draft_tentative_meeting"
```

---

### Task 3: Update system prompt text and `docs/ai-tool-surface.md`

**Files:**
- Modify: `OutlookAiAddIn/web-src/entry.ts` (system prompt string)
- Modify: `docs/ai-tool-surface.md`

**Interfaces:**
- Consumes: nothing — text/documentation only, describing Tasks 1-2's already-committed code.
- Produces: nothing.

- [ ] **Step 1: Update the system prompt's mode-description and undo_last_action sentences in `entry.ts`**

Run: `grep -n "accept/decline\|accept_meeting\|decline_meeting" OutlookAiAddIn/web-src/entry.ts` to find every remaining mention (the schema/display ones from Tasks 1-2 are already done; this step is the system prompt and `undo_last_action`'s own description). Update each to also name `tentative_meeting` and the three new draft tools alongside the existing `accept_meeting`/`decline_meeting` mentions - e.g. "accept/decline_meeting" becomes "accept/decline/tentative_meeting", and "Automate approvals (also auto-accept/decline meeting invitations...)" becomes "...auto-accept/decline/tentatively-respond to meeting invitations...". Keep the same sentence structure and tone as the surrounding text; this is a naming-completeness fix, not a rewrite.

- [ ] **Step 2: Update `docs/ai-tool-surface.md`**

Run: `grep -n "accept_meeting\|decline_meeting" docs/ai-tool-surface.md` to find its current table rows and prose mentions. Add `tentative_meeting`'s own row (same shape as `accept_meeting`/`decline_meeting`'s existing rows, describing the generalized `RespondMeeting` and the new `message` parameter) and the three `draft_*_meeting` rows in the draft-tools table section (same shape as `draft_cancel_event`'s row). Add a dated update block (`**Update <today's date>:**`, matching this file's existing convention) summarizing the addition and explicitly flagging the `message`/comment delivery as unverified live, same caveat style already used elsewhere in this file for other unverified behaviors.

- [ ] **Step 3: Commit**

```bash
git add OutlookAiAddIn/web-src/entry.ts docs/ai-tool-surface.md
git commit -m "docs(outlook): document tentative_meeting and draft_*_meeting tools"
```

---

### Task 4: Build, test, package, install for the project owner's live verification

**Files:** none (build/test/deploy only).

**Interfaces:** none — verification task.

- [ ] **Step 1: Run the full C# test suite**

Run: `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo` (from `C:\dev\opendocs`).

Expected: 216/216 passing.

- [ ] **Step 2: Rebuild the add-in**

Run both commands from Task 1 Step 6 again from a clean state to confirm the final combined diff builds.

Expected: both succeed with 0 errors.

- [ ] **Step 3: Check Outlook is closed, then package and install**

Run: `tasklist | grep -i outlook`

Expected: no Outlook process running. If one is running, ask the user to close Outlook before continuing.

```bash
cd deploy
powershell -ExecutionPolicy Bypass -File ./package.ps1 -App Outlook
cd dist
powershell -ExecutionPolicy Bypass -File ./install.ps1 -App Outlook
```

Expected: both scripts complete without error.

- [ ] **Step 4: Hand off for manual testing**

Tell the user the build is installed and ask them to restart Outlook. Provide `FORCE_TOOL` commands to test:
1. `accept_meeting`/`decline_meeting`/`tentative_meeting` with and without `message`, on a real received invite (not a self-organized item) - **critically check whether the organizer actually sees the comment text**, since this is the one load-bearing unverified assumption this whole feature depends on.
2. `draft_accept_meeting`/`draft_decline_meeting`/`draft_tentative_meeting` with `message` - confirm the opened window shows the pre-filled comment and nothing is sent until the user acts.

No commit for this task — it's verification only.
