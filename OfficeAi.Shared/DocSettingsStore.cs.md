## DocSettingsStore.Migrate

```
// FT-1 Task 7b: called once a provisional ("unsaved-...") chat id has
// just resolved to a real, path-derived one. Non-empty wins - a single
// JSON document cannot be concatenated the way ChatStore's JSONL can,
// so real guidelines already on the target must never be overwritten
// by an empty provisional value.
```
