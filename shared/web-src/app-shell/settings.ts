import type { AgentStreamHandle, AgentTransport } from '@genoffice/agent-core'
import {
  defaultAiSettings,
  resolveAiSettings,
  streamForProvider,
  type AiProviderConfig,
  type AiSettings,
  type ReasoningEffort,
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
  /**
   * 'light'/'dark' are explicit overrides; 'default' follows Office's own
   * theme, read once at pane startup (see OfficeAi.Shared/OfficeTheme.cs) -
   * not re-checked while the pane stays open. Save-gated: only written here
   * by bootstrap.ts's onSettingsSave handler, never live on click.
   */
  theme: 'light' | 'dark' | 'default'
  /** 'en'/'he' are explicit overrides; 'default' follows Office's own UI display
   * language, read once at pane startup. Same save-gating as theme (see `settings.ts.md`). */
  lang: 'en' | 'he' | 'default'
  /** Global, not per-provider - see chat-ui.ts's "More settings" reasoning section. 'default' omits the param entirely (today's behavior). */
  reasoningEffort: ReasoningEffort
}

// PP-0's flat { baseUrl, apiKey, model, skipTlsVerify } shape, kept only as
// the type loadSettings() migrates FROM - never written again.
interface LegacyStoredSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
  skipTlsVerify?: boolean
}

// Overrides defaultAiSettings()'s Genspark default with this repo's own
// test/mock-server provider + starting values (see `settings.ts.md`).
function defaultsForThisRepo(): AiSettings {
  const defaults = defaultAiSettings()
  defaults.provider = 'custom'
  defaults.providers.custom = { baseUrl: 'http://127.0.0.1:9000/v1', apiKey: 'test', model: 'test-model' }
  return defaults
}

const VALID_THEMES = ['light', 'dark', 'default'] as const
const VALID_LANGS = ['en', 'he', 'default'] as const
const VALID_REASONING_EFFORTS = ['default', 'off', 'low', 'medium', 'high', 'xhigh'] as const

function normalizeTheme(value: unknown): PanelSettings['theme'] {
  return (VALID_THEMES as readonly unknown[]).includes(value) ? (value as PanelSettings['theme']) : 'default'
}

function normalizeLang(value: unknown): PanelSettings['lang'] {
  return (VALID_LANGS as readonly unknown[]).includes(value) ? (value as PanelSettings['lang']) : 'default'
}

function normalizeReasoningEffort(value: unknown): ReasoningEffort {
  return (VALID_REASONING_EFFORTS as readonly unknown[]).includes(value) ? (value as ReasoningEffort) : 'default'
}

function loadSettings(): PanelSettings {
  const defaults = defaultsForThisRepo()
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY)
    if (!raw) return { ai: defaults, skipTlsVerify: false, theme: 'default', lang: 'default', reasoningEffort: 'default' }
    const parsed = JSON.parse(raw) as Partial<PanelSettings> & LegacyStoredSettings
    return {
      // resolveAiSettings migrates the legacy flat {baseUrl, apiKey, model}
      // shape into the `custom` provider slot, so an existing user's
      // configuration survives this change instead of silently resetting.
      ai: resolveAiSettings(parsed.ai ?? parsed, defaults),
      skipTlsVerify: !!parsed.skipTlsVerify,
      theme: normalizeTheme(parsed.theme),
      lang: normalizeLang(parsed.lang),
      reasoningEffort: normalizeReasoningEffort(parsed.reasoningEffort),
    }
  } catch {
    return { ai: defaults, skipTlsVerify: false, theme: 'default', lang: 'default', reasoningEffort: 'default' }
  }
}

function persistSettings(settings: PanelSettings): void {
  localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings))
}

let currentSettings: PanelSettings = loadSettings()

// Per-turn output budget (see `settings.ts.md` for why 8192 and not smaller/configurable).
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
      const config: AiProviderConfig = {
        apiKey: slot.apiKey,
        model: slot.model,
        baseUrl: slot.baseUrl,
        reasoningEffort: currentSettings.reasoningEffort,
      }
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
