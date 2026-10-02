## SharedCalendarEventRow

```
// One row of list_events' output, backend-agnostic - populated from
// Outlook COM properties for the caller's own calendar, or from EWS
// Appointment properties for a shared calendar (see OutlookEws.cs's
// GetSharedCalendarEventsAsync). Keeping this a plain struct (no Outlook
// COM or EWS type in its shape) is what makes Format below unit-testable
// without a live Outlook/Exchange session - same role ContactSearchFormat
// plays for search_contacts.
```

## SharedCalendarEventFormat

```
// Pure formatting for list_events' shared-calendar (EWS) path only -
// reproduces the exact per-event text shape OutlookTools.Calendar.cs's
// QueryCalendarItems already writes inline for the own-calendar (COM)
// path, so the tool's output doesn't reveal which backend answered it.
// QueryCalendarItems itself is deliberately NOT refactored to call this -
// see the design doc's Section 1 for why (its try/catch around
// appt.Start/appt.End is load-bearing and has no equivalent here).
```
