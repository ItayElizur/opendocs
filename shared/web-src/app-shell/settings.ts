import type { AgentStreamHandle, AgentTransport } from '@genoffice/agent-core'
import {
  defaultAiSettings,
  resolveAiSettings,
  streamForProvider,
  type AiProviderConfig,
  type AiSettings,
} from '@genoffice/ai-provider'

// Connection settings are user-editable via the panel's Settings dropdown
// (onSettingsSave, see bootstrap.ts) and persisted in this WebView2 profile's
// own localStorage - each app has its own separate WebView2 user-data folder
// (see WebViewBridgeHost's userDataFolder), so one shared key here never
// collides with or shares storage across Word/Excel/PowerPoint.
const SETTINGS_STORAGE_KEY = 'airchat-settings'

export interface PanelSettings {
  ai: AiSettings
  skipTlsVerify: boolean
}

// PP-0's flat { baseUrl, apiKey, model, skipTlsVerify } shape, kept only as
// the type loadSettings() migrates FROM - never written again.
interface LegacyStoredSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
  skipTlsVerify?: boolean
}

// officeoffice is an air-gapped/on-prem-oriented deployment; defaultAiSettings()
// alone would default to Genspark (a hosted proxy that needs a login), which
// would silently change out-of-box behavior for this repo's test/mock-server
// flow (docs/superpowers/plans/2026-08-22-mock-server-mode-testing.md). Override
// just the default provider + the custom slot's starting values, so a fresh
// profile behaves exactly as it did before PP-6.
function defaultsForThisRepo(): AiSettings {
  const defaults = defaultAiSettings()
  defaults.provider = 'custom'
  defaults.providers.custom = { baseUrl: 'http://127.0.0.1:9000/v1', apiKey: 'test', model: 'test-model' }
  return defaults
}

function loadSettings(): PanelSettings {
  const defaults = defaultsForThisRepo()
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY)
    if (!raw) return { ai: defaults, skipTlsVerify: false }
    const parsed = JSON.parse(raw) as Partial<PanelSettings> & LegacyStoredSettings
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

function persistSettings(settings: PanelSettings): void {
  localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings))
}

let currentSettings: PanelSettings = loadSettings()

// Per-turn output budget. 1024 was too small for models that spend budget on
// reasoning before emitting visible text: the provider returned
// finish_reason=length with zero content and the run ended in an unexplained
// empty reply (PP-4). 8192 is comfortably above a long tool-using turn's real
// output while staying well under every supported provider's per-request cap
// - checked against every model in shared/web-src/ai-provider/providers.ts
// (Claude/GPT/Gemini/DeepSeek families all support output limits well above
// 8192 via their APIs). Not made user-configurable from Settings: it's a
// footgun with no better default a user could pick, and Settings is already
// growing in PP-6/FT-1.
export const MAX_TOKENS = 8192

/**
 * Current connection settings, read through a function rather than a mutable
 * exported binding - every caller (in particular makeTransport()'s
 * request-time read, and bootstrap.ts's initialSettings/onSettingsSave) always
 * sees the latest saved value, never a snapshot captured at import time.
 */
export function getSettings(): PanelSettings {
  return currentSettings
}

export function setSettings(settings: PanelSettings): void {
  currentSettings = settings
  persistSettings(currentSettings)
}

export function makeTransport(): AgentTransport {
  return {
    stream(request, callbacks): AgentStreamHandle {
      const controller = new AbortController()
      // Read at request time (not at module load) so a Save takes effect on
      // the very next message without rebuilding the loop.
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
