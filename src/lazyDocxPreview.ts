/**
 * Lazy, drop-in wrapper for docx-preview's `renderAsync`. docx-preview (and its
 * bundled dependencies) is one of the heaviest libraries in the app, but it's
 * only needed once the user renders a document (preview / print / thumbnail /
 * PDF) — never on the initial Works List screen. Importing it dynamically here
 * lets the bundler split it into its own chunk that loads on first use, so the
 * app window appears faster at startup.
 *
 * Call sites keep using `renderAsync(...)` exactly as before (they already
 * await it) — only their import path changes from 'docx-preview' to this file.
 */
type RenderAsync = typeof import('docx-preview')['renderAsync']

export const renderAsync: RenderAsync = ((...args: Parameters<RenderAsync>) =>
  import('docx-preview').then((m) => m.renderAsync(...args))) as RenderAsync
