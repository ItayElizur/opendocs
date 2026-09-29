# Design: meeting responses with a comment, plus Tentative

**Status:** approved design, not yet implemented. Branch: `feature/outlook-meeting-response` (fresh branch from `main`, post-PR#27). Forwarding a received meeting was considered and explicitly dropped from scope (see "Out of scope").

## Problem

`accept_meeting`/`decline_meeting` exist but only take `event_id` — no way to add a comment to the response (Outlook's own UI offers "Edit the response before sending" for exactly this), and there's no `tentative_meeting` at all, even though `OlMeetingResponse` has a third value for it and Outlook's own UI always offers all three (Accept/Tentative/Decline) side by side.

## Ground truth (verified this session)

- `AppointmentItem.Respond(OlMeetingResponse, object, object)` returns a `MeetingItem` (confirmed via reflection against `_AppointmentItem`) — this is the existing mechanism `RespondMeeting` already uses for accept/decline; `OlMeetingResponse.olMeetingTentative` is a real enum value, same enum as the two already used.
- **Not verified live**: whether setting `.Body` on the `MeetingItem` `Respond()` returns, before calling `.Send()` on it, actually results in the organizer seeing that text as a comment on the response (this is the core mechanism the whole "message" feature depends on). This needs a real received invite to test — one this add-in's own account is an attendee on, not one it organizes (a self-organized item may not even support `.Respond()` meaningfully). The user has agreed to verify this live once built.
- `AppointmentItem.Forward()` does **not** exist (confirmed via reflection — this was an earlier wrong assumption, corrected before writing this spec). The only forwarding-adjacent member is `ForwardAsVcal()`, which returns a `MailItem` with the meeting attached as a legacy `.vcs` file (confirmed live: subject becomes `"FW: <original>"`, no pre-filled recipient, one `.vcs` attachment) — not a real clickable invite, and of uncertain compatibility with non-Outlook clients. **Explicitly out of scope for this branch** per the project owner's decision.

## Scope decisions

1. **Three response tools, not one with an enum parameter** — `accept_meeting`, `decline_meeting`, and new `tentative_meeting`, matching the existing split-tool convention used everywhere else in this codebase (`draft_event`/`create_event`, never a single tool with a mode parameter). Mirrors Outlook's own three-button UI.
2. **Each of the three gets an optional `message` parameter** — a comment attached to the response before it's sent to the organizer.
3. **Each of the three gets a `draft_` counterpart** — `draft_accept_meeting`, `draft_decline_meeting`, `draft_tentative_meeting` — opening the response (with any given `message` pre-filled) in a review window instead of auto-sending, matching the draft/immediate pairing convention used everywhere else in this add-in.
4. **No forwarding** — dropped from scope per the project owner, given `ForwardAsVcal`'s weak fidelity (plain email + legacy attachment, not a real invite).
5. **Tier placement**: the three immediate tools stay at the tier `accept_meeting`/`decline_meeting` already occupy (`ApprovalTierTools`/"Automate approvals" — they already auto-notify the organizer via `.Send()`). The three new draft tools go one tier down (`DraftTierTools`/"Draft only"), matching how every other draft/immediate pair in this codebase is split by exactly one tier.

## Design

### 1. Schema

```
accept_meeting / decline_meeting / tentative_meeting:
  event_id (required)
  message (optional) - comment added to the response before it's sent to the organizer.

draft_accept_meeting / draft_decline_meeting / draft_tentative_meeting:
  event_id (required)
  message (optional) - comment pre-filled in the response for the user to review/edit before sending themselves.
```

### 2. `RespondMeeting` generalization (immediate tools)

Current `RespondMeeting(mbxKey, input, bool accept)` takes a bool; generalize to take the actual `OlMeetingResponse` value directly, since there are now three:

```
private static ToolResult RespondMeeting(string mbxKey, JsonElement input, Outlook.OlMeetingResponse response, string toolName)
{
    id = ReqStr(input, "event_id")
    appt = resolve same way as today (AppointmentItem or MeetingItem.GetAssociatedAppointment(false))
    if appt == null: IsError "event_id does not resolve to a meeting."

    message = Str(input, "message", null)

    respObj = appt.Respond(response, true, false)   // NoUIFlag=true, SendResponse=false - unchanged from today
    resp = respObj as MeetingItem
    if resp != null:
        if message != null: resp.Body = message     // NEW - only reachable if Respond() actually returned a real MeetingItem
        try { resp.Send() } catch (ex) { DebugLog.WriteException(toolName + " Send", ex) }   // unchanged

    RecordIrreversible(mbxKey, toolName + " for \"" + (appt.Subject ?? "") + "\"" + (message != null ? " with a comment" : ""))
    verb = response == olMeetingAccepted ? "Accepted" : response == olMeetingTentative ? "Responded tentatively to" : "Declined"
    return ToolResult {
        Output: verb + ": " + (appt.Subject ?? "") + (message != null ? " (comment sent)" : ""),
        Mutated: true,
        Summary: toolName,
    }
}
```

`accept_meeting`/`decline_meeting`/`tentative_meeting` each become a one-line call: `RespondMeeting(mbxKey, input, Outlook.OlMeetingResponse.olMeetingAccepted, "accept_meeting")` etc. This is a **signature change** to an existing shared helper — both existing call sites (`accept_meeting`, `decline_meeting` in the dispatch switch) need updating to pass the enum value instead of a bool, and the new `tentative_meeting` case added alongside them.

### 3. Draft tools

New method, one per response type or one shared helper parameterized by `OlMeetingResponse` (implementer's choice, whichever keeps the diff more readable — the plan should let this be decided during implementation, same as prior "keep as new methods or rename" calls this session):

```
private static ToolResult DraftRespondMeeting(JsonElement input, Outlook.OlMeetingResponse response, string toolName)
{
    id = ReqStr(input, "event_id")
    appt = resolve same way as RespondMeeting
    if appt == null: IsError "event_id does not resolve to a meeting."

    message = Str(input, "message", null)

    respObj = appt.Respond(response, true, false)   // NoUIFlag=true, SendResponse=false - same call, never sent
    resp = respObj as MeetingItem
    if resp != null:
        if message != null: resp.Body = message
        resp.Display(false)

    verbing = response == olMeetingAccepted ? "an acceptance" : response == olMeetingTentative ? "a tentative response" : "a decline"
    return ToolResult {
        Output: "Opened " + verbing + " to \"" + (appt.Subject ?? "") + "\" in Outlook for the user to review and send." + (message != null ? " Comment pre-filled." : ""),
        Summary: toolName,
    }
}
```

Never calls `.Send()`, never calls `RecordIrreversible`/`RecordSnapshot` — matches every other draft tool's contract in this codebase exactly (nothing persists until the user acts).

### 4. Organizer-authority / already-responded checks

None needed beyond what `Respond()` itself already enforces (unchanged from today's `accept_meeting`/`decline_meeting`) — Outlook's own `Respond()` call is the authority here; this design doesn't add new pre-checks beyond the existing null-resolution check, since the existing tools already work this way and no new failure mode has been identified that needs a proactive check.

### 5. Risk / verification

**Explicitly unverified, flagged for the user's own live test once built** (per their agreement to verify): does the organizer actually see the `message` text as a comment on the response notification? This is the single load-bearing assumption the whole feature depends on — if `resp.Body` doesn't survive to the delivered response the way expected, the comment may be silently dropped, similarly to (but distinct from) the attendee-resolution bug found and fixed earlier this session. Verify this specifically before considering the feature done, the same way that bug was only caught by checking real delivery, not by trusting a clean COM call with no exception.

## Out of scope

- Forwarding a received meeting to a third party (`ForwardAsVcal`'s weak fidelity, dropped per project owner decision — see Ground Truth above).
- "Propose New Time" — confirmed via reflection this session to have no COM automation API at all (no `Propose`/`Counter` member on either `AppointmentItem` or `MeetingItem`); not revisited here.
- Any change to `accept_meeting`/`decline_meeting`'s existing organizer-authority behavior beyond adding the `message` parameter and generalizing the shared helper's signature.
