# id.ts — relocated comment history / rationale

## `randomId`

```
crypto.randomUUID() is gated to secure contexts. This app's WebView2 host
serves its own content from the insecure-scheme virtual host
http://appassets.local (see OfficeAi.Shared/WebViewBridgeHost.cs) so that
outbound fetches to a user-configured plain-HTTP LLM endpoint aren't
blocked as mixed content - which means randomUUID is unavailable here.
Falls back to a manual RFC4122 v4 UUID built from
crypto.getRandomValues(), which is NOT secure-context-gated.
```
