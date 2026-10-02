## ChatStore.Migrate

```
// FT-1 Task 7b: called once a provisional ("unsaved-...") chat id has
// just resolved to a real, path-derived one. This store is append-only
// JSONL, so concatenating the provisional file's lines onto whatever
// the target already has (the user may have saved over a path they'd
// chatted about before) is trivially valid and chronologically
// correct - no merge logic needed beyond "append, then remove the
// source". A missing source is a silent no-op (nothing to migrate).
```
