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
OpenDocs is an air-gapped/on-prem-oriented deployment. defaultAiSettings()
already defaults to the `custom` provider, but with empty base URL/key/model;
override just the custom slot's starting values (the test/mock-server
endpoint), so a fresh profile behaves exactly as it did before PP-6.
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

## API keys are not in localStorage (2026)

`persistSettings` blanks every provider's `apiKey` in the copy written to `localStorage` and sends the keys to the host
(`save-secrets`). The host keeps them encrypted with the Windows user's key (DPAPI) in `<app data folder>\secrets.bin`
(`OfficeAi.Shared/SecretStore.cs`) and injects them into the page before its scripts run as `window.__opendocsSecrets`
(`WebViewBridgeHost`), which `loadSettings` reads. A key still in `localStorage` from an older build is moved over the first time
the settings load. If the secrets file is lost, the keys are simply blank and the user enters them again.
