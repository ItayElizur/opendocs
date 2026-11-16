# OutlookTools.cs

## `ModeFor` - why the fallback changed from FullAutonomy

Falls back to CommentOnly (Outlook's "Draft only" tier - see
AlwaysAllowedTools' comment below), matching the client's own new
default (chat-ui.ts's defaultModeFor()), for the brief window
before the client's first explicit "set-mode" bridge message
arrives (only sent on a user-initiated mode change, never on
mount). Previously fell back to FullAutonomy, which was safe only
because FullAutonomy was also the client's own default at the
time - no longer true.

## `DraftTierTools` - why set_event_categories/set_category_color/set_event_availability are here

set_event_categories/set_category_color/set_event_availability
belong here, not in SendTierTools - all three are purely local
(appt.Categories/.BusyStatus/.Save(), cats.Add()/.Color), never
call .Send(), and carry the same risk profile as
move_email/flag_email_important right next to them. An earlier
version of this fix put them in SendTierTools to match their old
(pre-four-tier) Full-Autonomy-only gate, but that was restoring
the OLD binary model rather than applying this PR's own tiering
logic - every other local-only mutation here was deliberately
downgraded from Full-Autonomy-only, and these two were simply
missed, not deliberately kept stricter.

## `DraftTierTools` - why apply_search is here

apply_search is here for a different reason than the mutating
tools above: it never mutates data, but unlike every other
AlwaysAllowedTools entry it has a real, visible side effect - it
hijacks the user's actual Outlook Explorer window (folder jump +
search overlay) with no consent step. "Read only" is supposed to
guarantee the assistant never touches the user's screen; leaving
it always-allowed broke that.

## `ExecuteAsync` - delete_email's permanent-flag gate

delete_email is a single tool spanning two risk classes:
permanent:false (default) just moves to Deleted Items -
fully reversible, same Draft-tier gate as move_email
above. permanent:true additionally calls .Delete() from
there, which is irreversible from within Outlook (see
DeleteEmail's own result text) - the same risk class as
SendTierTools, so it needs that gate too even though the
tool NAME sits in DraftTierTools. This is name-based
gating's one input-aware exception; keep it that way
rather than generalizing to a per-argument system.

DeleteEmail adds a second rule for permanent:true: the item must
already be in Deleted Items (or a subfolder of it), otherwise it
returns a Blocked result and touches nothing. A hard delete is
therefore always a second call after a recoverable move, so a
wrong id costs a move, not the message. Checked against a live
Outlook: IsInFolderTree is true for Deleted Items and for a
subfolder of it, and false for the Inbox, another subfolder and a
null folder. The whole DeleteEmail path was not run through the
add-in.

## `ExecuteAsync` - preview_action is handled before the tier gate

`preview_action` backs the panel's Approve/Deny card (`confirmActions` in `entry.ts`; see `bootstrap.ts.md`). It is read-only,
is not a model tool, and is in none of `AlwaysAllowedTools`, `DraftTierTools`, `ApprovalTierTools` or `SendTierTools`, so
`entry.ts` must not list it and `TierListsAgreeTests` does not see it. `ExecuteAsync` answers it before the gate, by name. It
takes `{ tool, input }` (the call about to run) and returns one JSON object: `blocked` (the exact refusal the call would get)
or `type`, `subject`, `from` and `date` of the item the call acts on, never the body. The permanent-delete refusals are shared
functions (`PermanentModeRefusal`, `PermanentLocationRefusal`), so the preview and the real call cannot say different things.
An id that does not resolve is reported as `blocked` so the card is skipped. Not run end to end against a live Outlook: only
the type, sender and date helpers were checked on real items.
