## ContactMatch

```
// One search_contacts match. FullName and DisplayName are tracked
// separately (both sourced from the same EWS NameResolution/Contact, see
// OutlookEws.cs's FullNameFrom/DisplayNameFrom) because they can disagree:
// a directory's DisplayName is often org-formatted (e.g. "Dept/Unit/Title")
// and can look nothing like the person's actual name, while FullName
// (GivenName + Surname) is what a query like "John Doe" will actually
// match against. Confirmed live 2026-10-01: an agent saw only a
// display-name-only result that didn't resemble the query and wrongly
// concluded the search had failed, even though the match was correct.
```

## ContactSearchFormat.FormatLine

```
// FullName leads when available (it's what the caller's query
// actually matches against); DisplayName is a parenthetical hint,
// shown only when it adds information beyond the name already
// shown. A match with no name at all (shouldn't happen in
// practice - OutlookEws's DisplayNameFrom always falls back to
// the email itself) still renders without throwing.
```
