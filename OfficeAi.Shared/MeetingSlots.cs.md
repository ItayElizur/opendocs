## MeetingSlots - FreeBusy string format

```
// Each FreeBusy string has one character per `stepMinutes` minutes,
// starting at `rangeStartMidnight`. '0' = free; any other char (or an
// index past the end of the string, or before the range) is treated as
// free too only when past the end - a shorter-than-expected string means
// "no known busy info". Before the range is never reached in practice.
```
