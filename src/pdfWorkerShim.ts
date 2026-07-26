// Loaded as pdf.js's own dedicated Worker script (see pdfToImages.ts's
// GlobalWorkerOptions.workerSrc) — a separate JS global scope from the main
// thread, so the main-thread polyfills applied there don't reach here. The
// actual PDF parsing/hashing code that needs these compat polyfills (see
// pdfJsCompatPolyfills.ts) runs inside this worker realm, not the main
// thread — so it must be applied again here, before the real worker module
// (dynamically imported below, so this runs first) ever loads.
import { applyPdfJsCompatPolyfills } from './pdfJsCompatPolyfills'

applyPdfJsCompatPolyfills()

// pdfjs-dist ships no .d.ts for its raw worker build output.
// @ts-expect-error TS7016
await import('pdfjs-dist/build/pdf.worker.min.mjs')
