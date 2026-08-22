// Ambient module declaration so tsc can resolve the CSS side-effect import
// inside @officeai/chat-ui (shared/chat-ui/chat-ui.ts imports './chat-ui.css').
// esbuild's default loader already handles bundling a raw .css import as
// text for a `--format=iife` bundle; this only satisfies the type checker.
declare module '*.css'
