# OutlookTools.Search.cs

## `BuildSearchAqs`

Pushes an already-computed search filter into Outlook's own Explorer
window (the native Instant Search UI) instead of only returning
results to the model - the "show, don't just tell" counterpart to
SearchEmails.

_Explorer.Search(string Query, OlSearchScope SearchScope) signature
confirmed via .NET reflection against the actually-referenced
Microsoft.Office.Interop.Outlook 15.0.0.0 PIA in this repo. Does NOT
take a DASL/SQL filter, despite the "@SQL=" prefix this file
originally tried here (copied from SearchEmails' Items.Restrict
usage): live testing (2026-09-26) proved that string doesn't throw,
it just silently matches nothing, because Explorer.Search goes
through Windows Instant Search - a different query engine from
Items.Restrict's direct MAPI table access - and doesn't honor
urn:schemas:httpmail:* property URNs. It DOES take the same
Advanced Query Syntax (AQS) the Outlook search box itself accepts -
confirmed via https://support.microsoft.com/en-us/outlook/search-mail-and-people-in-outlook-com
and the user's own live "Advanced Search Options" list, which is
what BuildSearchAqs below builds: `From:value` for sender,
`Received:MM/DD/YYYY..MM/DD/YYYY` for a date range (the documented
two-dot syntax - deliberately NOT `>=`/`<=`, which that same page
does not confirm exists for this property). Plain free text with
no prefix (already live-tested and confirmed working) covers the
subject/body query term.

## `BuildSearchAqs` - date range formatting

Live testing (2026-09-26) proved hardcoding MM/dd/yyyy (US
order) was wrong: classic Outlook's search box parses typed
dates using the DEVICE's Windows regional format, not a
fixed order - on a day-first locale, "09/26/2036" reads as
day=09/month=26, an invalid month, so the whole Received:
clause silently failed to parse and the search returned 0
results (no error - same silent-failure shape as the
original "@SQL=" DASL bug this method replaced). Formatting
with CurrentCulture instead of InvariantCulture matches
whatever order this machine's own regional settings use,
which is exactly what Outlook's own parser reads back.

## `ParseScope` / `ValidScopes`

OlSearchScope (confirmed via reflection): CurrentFolder=0,
AllFolders=1, AllOutlookItems=2, Subfolders=3, CurrentStore=4.
Default stays CurrentFolder - the one value actually exercised in
the live test that validated this AQS approach in the first place;
the wider scopes are unverified beyond being documented, valid
enum members. Unrecognized non-empty input (e.g. a folder name
passed here by mistake instead of in `folder`) throws rather than
silently falling back to CurrentFolder - a live test hit exactly
that mix-up and got a confusingly quiet "searched the wrong scope"
instead of a clear error.
