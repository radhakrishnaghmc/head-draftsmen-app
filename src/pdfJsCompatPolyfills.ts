// Marks this file as a module so `declare global` below is valid.
export {}

// Not yet in this project's configured TS lib version's typings.
declare global {
  interface Uint8Array {
    toHex(): string
  }
  interface Map<K, V> {
    getOrInsert(key: K, value: V): V
    getOrInsertComputed(key: K, callbackfn: (key: K) => V): V
  }
  interface WeakMap<K extends WeakKey, V> {
    getOrInsert(key: K, value: V): V
    getOrInsertComputed(key: K, callbackfn: (key: K) => V): V
  }
}

function patchGetOrInsert(proto: Map<unknown, unknown> | WeakMap<WeakKey, unknown>): void {
  const p = proto as unknown as Record<string, unknown>
  if (typeof p.getOrInsert !== 'function') {
    p.getOrInsert = function (this: Map<unknown, unknown>, key: unknown, value: unknown) {
      if (this.has(key)) return this.get(key)
      this.set(key, value)
      return value
    }
  }
  if (typeof p.getOrInsertComputed !== 'function') {
    p.getOrInsertComputed = function (
      this: Map<unknown, unknown>,
      key: unknown,
      callbackfn: (key: unknown) => unknown
    ) {
      if (this.has(key)) return this.get(key)
      const value = callbackfn(key)
      this.set(key, value)
      return value
    }
  }
}

/**
 * pdf.js (both on the main thread and inside its own dedicated Worker —
 * see pdfWorkerShim.ts) assumes an increasing number of brand-new JS engine
 * features exist unconditionally, with no fallback, one of which this app's
 * bundled Electron/Chromium version predates each time pdf.js's own
 * minimum-supported-runtime creeps forward — Uint8Array.prototype.toHex()
 * first, now Map.prototype.getOrInsert(Computed)(). Each new one hit gets
 * added here rather than pinning to an older, less-maintained pdfjs-dist
 * release. Idempotent and safe to call from either realm, more than once.
 */
export function applyPdfJsCompatPolyfills(): void {
  if (typeof Uint8Array.prototype.toHex !== 'function') {
    Object.defineProperty(Uint8Array.prototype, 'toHex', {
      value(this: Uint8Array) {
        let hex = ''
        for (const byte of this) hex += byte.toString(16).padStart(2, '0')
        return hex
      },
      writable: true,
      configurable: true
    })
  }
  patchGetOrInsert(Map.prototype)
  patchGetOrInsert(WeakMap.prototype)
}
