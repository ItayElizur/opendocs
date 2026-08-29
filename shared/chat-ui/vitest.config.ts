import { defineConfig } from 'vitest/config'

// chat-ui.ts imports AI_PROVIDERS/AiProviderId from '@genoffice/ai-provider'
// (PP-6) - a path alias each app's esbuild/tsconfig resolves, not a real npm
// package, so vitest needs the same alias explicitly or this import fails to
// resolve when chat-ui.test.ts runs standalone (not bundled through an app).
export default defineConfig({
  test: { environment: 'jsdom' },
  resolve: { alias: { '@genoffice/ai-provider': '../web-src/ai-provider/index.ts' } },
})
