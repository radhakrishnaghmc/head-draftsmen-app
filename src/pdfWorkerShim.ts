// Marks this file as a module (needed for both `declare global` below and
// the top-level `await` further down) despite otherwise having no static
// imports/exports of its own.
export {}

// Not yet in this project's configured TS lib version's typings.
declare global {
  interface Uint8Array {
    toHex(): string
  }
}

// Loaded as pdf.js's own dedicated Worker script (see pdfToImages.ts's
// GlobalWorkerOptions.workerSrc) — a separate JS global scope from the main
// thread, so the main-thread Uint8Array.prototype.toHex() polyfill there
// never reaches here. pdf.js calls .toHex() unconditionally while hashing a
// PDF's content, and that hashing happens inside this worker realm, not the
// main thread — so it must be polyfilled again here, before the real worker
// module (dynamically imported below, so this runs first) ever loads.
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

// pdfjs-dist ships no .d.ts for its raw worker build output.
// @ts-expect-error TS7016
await import('pdfjs-dist/build/pdf.worker.min.mjs')
