/**
 * Electron main processes can inject net.fetch here as a rescue path for when
 * the primary fetch fails at the network layer (see `fetch.ts.md`); renderers
 * never need one.
 */

type FetchLike = (url: string, init?: RequestInit) => Promise<Response>

let rescueFetch: FetchLike | null = null

export function setRescueFetch(fn: FetchLike | null): void {
  rescueFetch = fn
}

export async function aiFetch(url: string, init: RequestInit): Promise<Response> {
  try {
    return await fetch(url, init)
  } catch (primaryError) {
    const signal = init.signal as AbortSignal | null | undefined
    if (!rescueFetch || signal?.aborted) throw primaryError
    console.warn('[ai-provider] fetch failed, retrying via rescue fetch:', String(primaryError))
    try {
      return await rescueFetch(url, init)
    } catch {
      throw primaryError
    }
  }
}
