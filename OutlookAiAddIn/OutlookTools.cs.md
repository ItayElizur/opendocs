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
