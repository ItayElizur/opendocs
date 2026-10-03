## DebugLog

```
// Post-hoc diagnostic tool (2026-08-24): a plain-text, append-only log for
// manually reproducing bugs that have resisted blind fixing (the chart
// RPC failure, in particular, after two rounds of unverified guesses).
// Writes to a fixed, easy-to-find path so a user reproducing a bug can
// locate and share the file without hunting for it - %TEMP% is a well
// known Windows path nameable directly in Explorer's address bar.
// Deliberately NOT wired into every tool call, only where a specific bug
// needs step-by-step COM call tracing - this is a temporary debugging
// aid, not a permanent logging subsystem.
```
