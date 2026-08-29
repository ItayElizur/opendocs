# PP-6: Wire Provider/Model Selection to the Transport — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-6 (P1). That item asked the implementer to "verify current wiring before scoping" — **that verification is done and recorded below**, and it narrows the scope.

> **Updated 2026-08-24, post-PP-0:** `2026-08-23-pp00-shared-app-shell.md` has landed since the verification below was performed. The `WordAiAddIn/web-src/entry.ts` line references in "Verified current state" are a historical record of that pre-PP-0 audit and are left as-written; the code they describe now lives once in `shared/web-src/app-shell/settings.ts` (`makeTransport`, `StoredSettings`, defaults) and `bootstrap.ts` (`onSettingsSave`, `initialSettings`), identically for all three apps. Task 2 below is rewritten as a single-file edit.

## Verified current state (do not re-derive)

- `makeTransport()` reads `currentSettings` **inside** `stream()`, at request time (`WordAiAddIn/web-src/entry.ts:133-156`). So `baseUrl`, `apiKey`, and `model` entered in Settings *do* reach the next request. That half of the item is **not** a live gap.
- The real gap is **provider protocol**: all three apps call `streamOpenAiCompatible` unconditionally (`WordAiAddIn/web-src/entry.ts:136`, Excel `:120`-ish, PowerPoint `:121`-ish). Anthropic's and Gemini's APIs are not OpenAI-compatible, so entering `https://api.anthropic.com` in Settings produces a failing request, not a working Claude connection.
- The router already exists and is already exported: `streamForProvider(provider, config, system, messages, tools, maxTokens, cb)` (`shared/web-src/ai-provider/stream.ts:847`) dispatches to `streamAnthropic` / `streamGemini` / `streamOpenAiCompatible`, including Genspark's three protocol-specific endpoints. `AI_PROVIDERS` (`shared/web-src/ai-provider/providers.ts:27-88`) carries labels, model lists, default models, key placeholders, and a `needsBaseUrl` flag. `defaultAiSettings()` and `resolveAiSettings()` (`:96-133`) already handle defaults and migration from the legacy single-endpoint shape.
- The Settings UI has no provider control at all. `ChatUIOptions.onSettingsSave` is typed `{ baseUrl, apiKey, model, skipTlsVerify, lang }` (`shared/chat-ui/chat-ui.ts:38`) and the panel renders three text fields (`STRINGS.settingsBaseUrl`/`settingsApiKey`/`settingsModel`, `:15-17`).
- Defaults point at a local test endpoint: `{ baseUrl: 'http://127.0.0.1:9000/v1', apiKey: 'test', model: 'test-model' }` (`WordAiAddIn/web-src/entry.ts:114`).

**Goal:** A user picks a provider (Claude / Gemini / OpenAI / DeepSeek / Genspark / Custom) and a model in Settings, enters a key, saves, and the next request uses that provider's real protocol — with each provider's key and model remembered separately so switching back and forth doesn't wipe credentials.

**Architecture:** Adopt `AiSettings` (the `{ provider, providers: Record<id, {apiKey, model, baseUrl?}> }` shape `providers.ts` already defines) as the panel's persisted settings model, replacing the flat `StoredSettings` in `shared/web-src/app-shell/settings.ts` (one place, post-PP-0, rather than three `entry.ts` files). Settings gains a provider `<select>` and a model control that repopulates from `AI_PROVIDERS`; the base-URL field shows only when `needsBaseUrl` (Custom) is chosen. `makeTransport` swaps its single `streamOpenAiCompatible` call for `streamForProvider`. `resolveAiSettings` handles migration so anyone with the old flat localStorage entry lands in the `custom` slot rather than losing their configuration.

**Tech Stack:** TypeScript — `shared/chat-ui/chat-ui.ts` + `.css`, `shared/web-src/app-shell/settings.ts` + `bootstrap.ts`, consuming the existing `shared/web-src/ai-provider` package unchanged.

## Global Constraints

- Do not modify `shared/web-src/ai-provider/` in this plan. Everything needed is already exported; changing it would widen blast radius across all three apps for no gain.
- The API key is a secret. It stays in the WebView2 profile's `localStorage` exactly as today (each app has its own user-data folder, so no cross-app leak), the input stays `type="password"`, and the key must never be written to `ChatStore`, a log, or a tool result. Note that `loop.ts` already redacts secret-looking tokens from outgoing user messages — that is a separate safeguard, not a substitute.
- Every new user-visible string goes through `chat-ui.ts`'s `STRINGS` table with `en` and `he` entries.
- Provider labels and model lists come from `AI_PROVIDERS` at runtime — never hardcode a model id in `chat-ui.ts` or the app shell.
- Keep `skipTlsVerify` and its `postTlsBypass` bridge call exactly as they are (`shared/web-src/app-shell/bridge.ts`/`bootstrap.ts`, post-PP-0); this plan does not touch TLS.
- Rebuild bundles + MSBuild for all three apps after any change (command in `2026-08-23-pp02-tool-steps-chronological-order.md`'s Global Constraints — note its 4-alias update from PP-0).

---

### Task 1: Provider + model controls in the Settings panel

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`

**Interfaces:**
- Produces: a widened `ChatUIOptions.onSettingsSave` payload — `{ provider, apiKey, model, baseUrl, skipTlsVerify, lang }` — and a widened `initialSettings`. Consumed by all three `entry.ts` in Task 2.

- [ ] **Step 1: Strings** — add `settingsProvider` (`en: 'Provider'`, `he: 'ספק'`) to `STRINGS`.

- [ ] **Step 2: Options type**

```ts
onSettingsSave: (settings: {
  provider: AiProviderId
  apiKey: string
  model: string
  baseUrl: string      // '' unless the provider needs one
  skipTlsVerify: boolean
  lang: Lang
}) => void

initialSettings?: {
  provider?: AiProviderId
  apiKey?: string
  model?: string
  baseUrl?: string
  skipTlsVerify?: boolean
  /** per-provider saved values, so switching providers restores that provider's own key/model */
  providers?: Record<string, { apiKey: string; model: string; baseUrl?: string }>
}
```

Import `AI_PROVIDERS` and the `AiProviderId` type from `@genoffice/ai-provider`. Confirm `chat-ui` can resolve that alias — the esbuild command aliases it for `entry.ts`; if `chat-ui.ts` is bundled through the same entry (it is, via `@officeai/chat-ui`), the alias applies. If a vitest run cannot resolve it, add the same alias to `shared/chat-ui/vitest.config.ts` rather than duplicating the provider table.

- [ ] **Step 3: Markup** — in the settings form, above the existing fields:
  - a `<select class="ai-settings-provider">` with one `<option>` per `AI_PROVIDERS` entry (`value` = `id`, text = `label`);
  - the model control becomes a `<select>` when the chosen provider has a non-empty `models` array, and stays a free-text `<input>` when it is empty (Custom). Rebuild it on provider change.
  - the base-URL field gets `hidden` unless the chosen provider's `needsBaseUrl` is true.
  - the API-key input's `placeholder` follows the chosen provider's `keyPlaceholder` (which is how Genspark communicates "not required — sign in").

- [ ] **Step 4: Provider-change behavior** — changing the provider swaps the form to that provider's remembered values from `initialSettings.providers`, falling back to `defaultModel` and an empty key. Do **not** fire `onSettingsSave` on change; saving stays explicit on the Save button, matching the existing tested behavior ("settings only call onSettingsSave when Save is clicked, not on field input", `shared/chat-ui/chat-ui.test.ts`).

- [ ] **Step 5: CSS** — style the two selects consistently with the existing settings inputs, using existing tokens.

**Verification:** `cd shared/chat-ui && npx vitest run` — existing settings tests still pass (adjust only for the widened payload).

---

### Task 2: `shared/web-src/app-shell/settings.ts` — persist `AiSettings` and route through `streamForProvider`

> **Updated 2026-08-24, post-PP-0:** `2026-08-23-pp00-shared-app-shell.md` has landed. `StoredSettings`, `loadSettings`, `makeTransport` now live once in `shared/web-src/app-shell/settings.ts` rather than in three `entry.ts` files — this task, including its old Step 6 ("do all three apps identically... trivial extraction later"), collapses to one edit.

**Files:**
- Modify: `shared/web-src/app-shell/settings.ts`
- Modify: `shared/web-src/app-shell/bootstrap.ts` (for `initialSettings`/`onSettingsSave`, Steps 4-5)

*(If PP-0 has not landed: modify `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts` identically, and do Task 2's original Step 6.)*

- [ ] **Step 1: Replace `StoredSettings` with `AiSettings` + the panel-local `skipTlsVerify`**

```ts
interface PanelSettings {
  ai: AiSettings           // from @genoffice/ai-provider
  skipTlsVerify: boolean
}

function loadSettings(): PanelSettings {
  const defaults = defaultAiSettings()
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY)
    const parsed = raw ? JSON.parse(raw) : {}
    return {
      // resolveAiSettings migrates the legacy flat {baseUrl, apiKey, model}
      // shape into the `custom` provider slot, so an existing user's
      // configuration survives this change instead of silently resetting.
      ai: resolveAiSettings(parsed.ai ?? parsed, defaults),
      skipTlsVerify: !!parsed.skipTlsVerify,
    }
  } catch {
    return { ai: defaults, skipTlsVerify: false }
  }
}
```

- [ ] **Step 2: Default provider for this repo**

`defaultAiSettings()` defaults to `provider: 'genspark'`. officeoffice is an air-gapped/on-prem-oriented deployment whose current default is a local endpoint — so override the default to `custom` with `baseUrl: 'http://127.0.0.1:9000/v1'`, `apiKey: 'test'`, `model: 'test-model'`, preserving today's out-of-box behavior and the mock-server testing flow in `docs/superpowers/plans/2026-08-22-mock-server-mode-testing.md`. Comment the override so it reads as deliberate.

- [ ] **Step 3: `makeTransport` routes by provider**

`settings.ts` as PP-0 left it reads its module-scope `currentSettings` variable directly inside `stream()`'s closure (not through the `getSettings()`/`setSettings()` accessors, which exist for *external* callers in `bootstrap.ts`) — keep that shape, just widen what `currentSettings` holds:

```ts
function makeTransport(): AgentTransport {
  return {
    stream(request, callbacks): AgentStreamHandle {
      const controller = new AbortController()
      // Read at request time (not at module load) so a Save takes effect on the
      // very next message without rebuilding the loop.
      const id = currentSettings.ai.provider
      const slot = currentSettings.ai.providers[id]
      const config: AiProviderConfig = { apiKey: slot.apiKey, model: slot.model, baseUrl: slot.baseUrl }
      streamForProvider(id, config, request.system, request.messages, request.tools, MAX_TOKENS, {
        onDelta: callbacks.onDelta,
        onToolCall: callbacks.onToolCall,
        onStopReason: callbacks.onStopReason,
        signal: controller.signal,
      })
        .then(() => callbacks.onDone())
        .catch((e: unknown) => callbacks.onError(e instanceof Error ? e.message : String(e)))
      return { cancel: () => controller.abort() }
    },
  }
}
```

Check `AiProviderConfig`'s actual field names in `shared/web-src/ai-provider/types.ts` before writing this — if `baseUrl` is not part of it, `streamForProvider`'s `custom` branch takes the base URL another way; follow whatever that branch expects rather than inventing a field.

- [ ] **Step 4: `onSettingsSave`** (in `bootstrap.ts`, via `settings.ts`'s `getSettings()`/`setSettings()`) writes into the per-provider slot, sets `ai.provider`, persists, and keeps the existing `postTlsBypass` call:

```ts
// bootstrap.ts
onSettingsSave: (s) => {
  const current = getSettings()
  const providers = {
    ...current.ai.providers,
    [s.provider]: {
      apiKey: s.apiKey || current.ai.providers[s.provider]?.apiKey || '',
      model: s.model || current.ai.providers[s.provider]?.model || '',
      baseUrl: s.baseUrl || current.ai.providers[s.provider]?.baseUrl,
    },
  }
  setSettings({ ai: { provider: s.provider, providers }, skipTlsVerify: s.skipTlsVerify })
  postTlsBypass(s.skipTlsVerify)
},
```

`setSettings` already persists to `localStorage` internally (PP-0), so there is no separate `saveSettings` call to make here. The `||` fallbacks preserve the existing "blank field keeps the old value" behavior, so a user who reopens Settings and saves without retyping their key doesn't wipe it.

- [ ] **Step 5: `initialSettings`** (in `bootstrap.ts`'s `mountChatUI` call) passes `provider`, the active slot's values, `skipTlsVerify`, and the whole `providers` map (Task 1 Step 4 needs it) — sourced from `getSettings()`.

**Verification:** all three bundles build (one source edit, three rebuilds — the shell is bundled separately into each app); all three projects MSBuild.

---

### Task 3: Connection test button

**Files:**
- Modify: `shared/chat-ui/chat-ui.ts`, `shared/chat-ui/chat-ui.css`, `shared/web-src/app-shell/bootstrap.ts`

**Rationale:** with six providers and a TLS-bypass toggle, "it just doesn't answer" becomes the dominant support case. A test button converts that into a specific message.

- [ ] **Step 1:** Add a "Test connection" button beside Save, plus `STRINGS` entries for it and its three states (testing / success / failure).
- [ ] **Step 2:** `ChatUIOptions` gains `onSettingsTest?: (settings: SameShapeAsSave) => Promise<string>` — resolving with a success message, rejecting with the error text.
- [ ] **Step 3:** In `bootstrap.ts` (once, for all three apps), implement it as a minimal `streamForProvider` call with a one-word prompt, no tools, `maxTokens: 16`, collecting the first delta. Report `err.message`, which already carries HTTP status detail via `shared/web-src/ai-provider/http-error.ts`.
- [ ] **Step 4:** Show the result inline in the settings panel, not as a chat message.

**Verification:** with a deliberately wrong key, the button reports an auth error rather than silently doing nothing; with the local mock server running, it reports success.

---

### Task 4: Manual verification across providers

- [ ] **Step 1:** Fresh profile → Settings shows Custom + the local endpoint; the mock-server flow from `docs/superpowers/plans/2026-08-22-mock-server-mode-testing.md` still works unchanged.
- [ ] **Step 2:** Existing profile with the old flat localStorage entry → after upgrade, Settings shows Custom with the previously saved base URL/key/model intact (migration path).
- [ ] **Step 3:** Switch to Claude, pick a model, enter a real key, Save, send a message → a real answer arrives and a tool call round-trips (Anthropic's tool protocol differs from OpenAI's; a tool-using prompt is the real test, not a plain chat).
- [ ] **Step 4:** Switch to Gemini, same test.
- [ ] **Step 5:** Switch back to Claude → the previously entered Claude key and model are still there (per-provider persistence).
- [ ] **Step 6:** Repeat Step 3 in Excel and PowerPoint.
- [ ] **Step 7:** Confirm no API key appears in `ChatStore`'s on-disk history files after any of the above.
