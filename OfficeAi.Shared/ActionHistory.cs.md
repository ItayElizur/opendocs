## ActionHistory&lt;T&gt;

```
/// Undo/redo bookkeeping for a host that has no usable native undo for the
/// assistant's own actions (Outlook - see OutlookTools.Undo.cs). Pure data
/// structure: knows nothing about how an entry is reversed, only the order
/// entries come back in.
///
/// A barrier marks an irreversible action (a sent email, a meeting
/// response). Undo stops at it instead of silently reaching past it and
/// reversing something older than what the user just saw happen.
///
/// Protocol for callers: TryTakeUndo/TryTakeRedo remove the entry; on a
/// successful reversal hand it back with CompleteUndo/CompleteRedo. On
/// failure just drop it - never retry an entry against state that may
/// already have changed.
```
