# settings.ts — relocated comment history / rationale

## `PanelSettings.lang`

```
'en'/'he' are explicit overrides; 'default' follows Office's own UI
display language, read once at pane startup (see
OfficeAi.Shared/OfficeLanguage.cs) - not re-checked while the pane stays
open. Same save-gating as theme. Previously this preference was threaded
through chat-ui.ts's Save payload but never actually persisted here - a
pre-existing gap now closed alongside adding the 'default' option.
```

## `defaultsForThisRepo`

```
OpenDocs is an air-gapped/on-prem-oriented deployment; defaultAiSettings()
alone would default to Genspark (a hosted proxy that needs a login), which
would silently change out-of-box behavior for this repo's test/mock-server
flow. Override just the default provider + the custom slot's starting values, so a fresh
profile behaves exactly as it did before PP-6.
```

## `MAX_TOKENS`

```
Per-turn output budget. 1024 was too small for models that spend budget on
reasoning before emitting visible text: the provider returned
finish_reason=length with zero content and the run ended in an unexplained
empty reply (PP-4). 8192 is comfortably above a long tool-using turn's real
output while staying well under every supported provider's per-request cap
- checked against every model in shared/web-src/ai-provider/providers.ts
(Claude/GPT/Gemini/DeepSeek families all support output limits well above
8192 via their APIs). Not made user-configurable from Settings: it's a
footgun with no better default a user could pick, and Settings is already
growing in PP-6/FT-1.
```
