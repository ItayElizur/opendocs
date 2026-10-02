import { defineConfig } from 'vitest/config'
import { fileURLToPath, URL } from 'node:url'

// chat-ui.ts/stream.ts import via path aliases that each app's esbuild/tsconfig
// resolves, not real npm packages - vitest needs the same aliases explicitly or
// these imports fail to resolve when chat-ui.test.ts runs standalone (see
// `vitest.config.ts.md`).
export default defineConfig({
  test: { environment: 'jsdom' },
  resolve: {
    alias: [
      { find: '@genoffice/ai-provider', replacement: fileURLToPath(new URL('../web-src/ai-provider/index.ts', import.meta.url)) },
      { find: '@genoffice/agent-core', replacement: fileURLToPath(new URL('../web-src/agent-core/index.ts', import.meta.url)) },
    ],
  },
})
