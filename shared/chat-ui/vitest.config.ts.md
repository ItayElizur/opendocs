# vitest.config.ts — relocated comment history / rationale

## Path-alias resolve block

```
chat-ui.ts imports AI_PROVIDERS/AiProviderId from '@officeai/ai-provider'
(PP-6) - a path alias each app's esbuild/tsconfig resolves, not a real npm
package, so vitest needs the same alias explicitly or this import fails to
resolve when chat-ui.test.ts runs standalone (not bundled through an app).
'@officeai/ai-provider' (stream.ts) itself now imports randomId from
'@officeai/agent-core' - same deal, needs its own alias or the resolve
chain breaks one hop further in.
```
