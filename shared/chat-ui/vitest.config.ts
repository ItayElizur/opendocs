import { defineConfig } from 'vitest/config'
import { fileURLToPath, URL } from 'node:url'

// chat-ui.ts imports AI_PROVIDERS/AiProviderId from '@genoffice/ai-provider'
// (PP-6) - a path alias each app's esbuild/tsconfig resolves, not a real npm
// package, so vitest needs the same alias explicitly or this import fails to
// resolve when chat-ui.test.ts runs standalone (not bundled through an app).
// '@genoffice/ai-provider' (stream.ts) itself now imports randomId from
// '@genoffice/agent-core' - same deal, needs its own alias or the resolve
// chain breaks one hop further in.
export default defineConfig({
  test: { environment: 'jsdom' },
  resolve: {
    alias: [
      { find: '@genoffice/ai-provider', replacement: fileURLToPath(new URL('../web-src/ai-provider/index.ts', import.meta.url)) },
      { find: '@genoffice/agent-core', replacement: fileURLToPath(new URL('../web-src/agent-core/index.ts', import.meta.url)) },
    ],
  },
})
