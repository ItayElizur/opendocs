# types.ts — relocated comment history / rationale

## `ReasoningEffort`

```
'default' sends no reasoning-related parameter at all (today's behavior,
zero risk). 'off' actively asks the provider to suppress reasoning where
it can - distinct from 'default' for providers whose models reason by
default (e.g. a vLLM-served Qwen3 deployment), identical to 'default' for
providers that never reason unless asked (e.g. Anthropic). Not every
provider/model honors every tier - unsupported combinations are silently
ignored by the provider, not an error in this layer.
```
