import * as pdfjsLib from 'pdfjs-dist'
import { applyPdfJsCompatPolyfills } from './pdfJsCompatPolyfills'
// `?worker&url` (not plain `?url`): Vite bundles the shim AND the real
// pdf.worker it imports into one self-contained worker file and hands back
// its hosted URL. Plain `?url` instead inlined the shim source as a base64
// data: URL whose bare `import 'pdfjs-dist/…/pdf.worker'` can't resolve in
// the packaged app (and a data: worker trips the app's CSP) — so PDF loading
// worked in dev but broke in the built app. workerSrc (a URL, vs workerPort)
// lets pdf.js spin up one worker per document, which the folder import needs
// for its concurrent parses.
import pdfWorkerUrl from './pdfWorkerShim.ts?worker&url'

// Main-thread compat shims (the worker gets its own via the shim); harmless
// to call more than once.
applyPdfJsCompatPolyfills()
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl

export { pdfjsLib }
