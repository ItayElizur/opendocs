// crypto.randomUUID() is gated to secure contexts. This app's WebView2 host
// serves its own content from the insecure-scheme virtual host
// http://appassets.local (see OfficeAi.Shared/WebViewBridgeHost.cs) so that
// outbound fetches to a user-configured plain-HTTP LLM endpoint aren't
// blocked as mixed content - which means randomUUID is unavailable here.
// Falls back to a manual RFC4122 v4 UUID built from
// crypto.getRandomValues(), which is NOT secure-context-gated.
export function randomId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    try {
      return crypto.randomUUID()
    } catch {
      // not a secure context - fall through to the manual generator below
    }
  }
  const bytes = new Uint8Array(16)
  crypto.getRandomValues(bytes)
  bytes[6] = (bytes[6] & 0x0f) | 0x40
  bytes[8] = (bytes[8] & 0x3f) | 0x80
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('')
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`
}
