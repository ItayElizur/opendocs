# fetch.ts — relocated comment history / rationale

## Module-level rescue-fetch rationale

```
In Electron main processes AI requests run on Node's fetch (undici), which
connects directly instead of going through Chromium's network stack. Under
VPN/tun setups those direct connections can get reset (ECONNRESET) while
Chromium traffic — login, renderer fetches — works fine. Main processes
inject Electron's net.fetch here as a rescue path: when the primary fetch
fails at the network layer, the request is retried once over the Chromium
stack. Renderers never inject one (their fetch already is Chromium's).
```
