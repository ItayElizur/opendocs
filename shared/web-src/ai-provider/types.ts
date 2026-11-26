export type AiProviderId = 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'custom'

/** 'default' sends no reasoning-related parameter; 'off' actively asks the
 * provider to suppress reasoning. Not every provider/model honors every tier -
 * unsupported combinations are silently ignored, not an error here (see `types.ts.md`). */
export type ReasoningEffort = 'default' | 'off' | 'low' | 'medium' | 'high' | 'xhigh'

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** only used by the custom (OpenAI-compatible) provider */
  baseUrl?: string | undefined
  reasoningEffort?: ReasoningEffort
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
}

/** pre-provider settings shape (single OpenAI-compatible endpoint); migrated into "custom" */
export interface LegacyAiSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
}
