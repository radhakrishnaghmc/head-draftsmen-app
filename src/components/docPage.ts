import { useEffect, useState } from 'react'
import type { RefObject } from 'react'
import { api } from '../ipc'
import { renderAsync } from '../lazyDocxPreview'

// A4 at 96dpi (this app is for a Telangana government department — A4, not
// US Letter) with a 1-inch margin all round, matching Word's default page
// setup.
export const PAGE_WIDTH = 794
export const PAGE_HEIGHT = 1123
const MARGIN = 96

/**
 * CSS injected into a document (the editable canvas, or the filled-document
 * preview) so it reads as a page sitting on a gray desk, like Word — a
 * white A4 sheet, centered, with a 1-inch margin. Deliberately one
 * continuous sheet, not banded into fixed-height "pages": a hard visual
 * break at an arbitrary fixed pixel height doesn't know where the real
 * content's paragraphs/tables actually end, so it used to cut straight
 * through the middle of pasted content and read as corruption. The page
 * count shown while scrolling (usePageScrollTracker) is an estimate based
 * on this same page height, without drawing a false line through the
 * content to back it up.
 *
 * html's overflow-x is hidden because a real page never scrolls sideways:
 * once content is taller than the visible area, the iframe's own vertical
 * scrollbar steals a few pixels of width, which would otherwise push the
 * fixed-width page box into its own horizontal overflow — a second,
 * redundant scrollbar stacked on top of the outer .doc-desk one (the only
 * horizontal scroll that should ever appear, and only for a genuinely
 * too-narrow window).
 */
/** Decode a base64 .docx payload (as returned by the preview IPC calls) into raw bytes for docx-preview's `renderAsync`. */
export function base64ToUint8(b64: string): Uint8Array {
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return bytes
}

/** Shared `renderAsync` options for a faithful, paginated render — used everywhere a real .docx (not the pre-OOXML HTML editor) is shown on screen. */
export const DOCX_PREVIEW_OPTIONS = {
  className: 'docx',
  inWrapper: true,
  ignoreWidth: false,
  ignoreHeight: false,
  ignoreFonts: false,
  breakPages: true,
  experimental: true,
  trimXmlDeclaration: true,
  useBase64URL: true,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  renderChanges: false
}

/**
 * docx-preview mishandles Word text boxes (`wps` shapes — e.g. the WordArt
 * "CYBERABAD MUNICIPAL CORPORATION" letterhead): it doesn't parse them as
 * graphics and forces the drawing wrapper to full page width, which stretches
 * the title across the whole sheet in the preview. Real pictures (the emblems)
 * render fine. This runs after `renderAsync` to constrain each text-box drawing
 * — a docx-preview drawing `div` (marked by its `text-indent: 0px`) that
 * contains no `<img>` — back to its content width and centre it, matching how
 * Word lays it out. Preview-only: the exported .docx is built from the real
 * template and is never touched.
 */
export function normalizeDocxTextboxes(container: HTMLElement): void {
  const drawings = container.querySelectorAll<HTMLElement>('div[style*="text-indent: 0px"]')
  drawings.forEach((d) => {
    if (d.querySelector('img')) return // a real picture (emblem) — leave it alone
    d.style.width = 'auto'
    d.style.maxWidth = '100%'
    d.style.display = 'block'
    d.style.textAlign = 'center'
    d.style.transform = 'none'
    d.querySelectorAll<HTMLElement>('*').forEach((el) => {
      if (el.style.width === '100%') el.style.width = 'auto'
      if (el.style.transform && /scale/i.test(el.style.transform)) el.style.transform = 'none'
    })
  })

  // WordArt ("v:textpath" — e.g. the same "CYBERABAD MUNICIPAL CORPORATION"
  // letterhead, in its older VML form) renders as an <svg> that keeps VML's
  // authored `position:absolute; margin-left:…; margin-top:…` — offsets from
  // wherever docx-preview happens to lay out the run in normal flow, which is
  // never where Word actually placed it, so the title drifts off the page's
  // right edge. Word's real intent is simpler: it's the sole run on an
  // already `text-align:center` line — so drop the absolute positioning and
  // let it flow centered, keeping only its authored size (svg[style*=
  // "v-text-anchor"] is a reliable VML-only marker, never emitted for
  // anything else docx-preview renders).
  container.querySelectorAll<SVGElement>('svg[style*="v-text-anchor"]').forEach((svg) => {
    svg.style.position = 'static'
    svg.style.display = 'block'
    svg.style.margin = '0 auto'
  })
}

/** Converts raw bytes to a base64 string in fixed-size chunks — a plain `String.fromCharCode(...bytes)` call blows the engine's argument-count limit on a multi-hundred-KB PNG. */
function uint8ToBase64(bytes: Uint8Array): string {
  const CHUNK = 8192
  let binary = ''
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  }
  return btoa(binary)
}

/**
 * Renders a filled .docx as an accurate preview — via LibreOffice's own
 * docx→PDF→PNG raster pipeline (core/docxToPdf.ts's docxToPageImages), one
 * image per page — so preview and print show exactly what Word would
 * produce: real font metrics, real page breaks, real WordArt, not
 * docx-preview.js's approximate HTML re-implementation (see
 * normalizeDocxTextboxes above for the kind of gap that leaves).
 *
 * Deliberately renders LibreOffice's PDF output through LibreOffice's OWN
 * rasterizer again, rather than pdf.js: a LibreOffice-produced PDF for a
 * font this machine doesn't have installed (e.g. "Book Antiqua", "Segoe UI")
 * embeds a substituted font whose glyph mapping pdf.js decodes wrong,
 * silently swapping in different characters — confirmed by rendering the
 * exact same, byte-verified-correct PDF through Poppler (correct) and pdf.js
 * (garbled) side by side. LibreOffice's own rasterizer renders that same
 * substituted font correctly, because it's the same engine that chose the
 * substitution in the first place.
 *
 * Falls back to the existing renderAsync/DOCX_PREVIEW_OPTIONS HTML render
 * when LibreOffice isn't installed on this machine (core/docxToPdf.ts throws
 * a clear error in that case) — every call site already expects
 * docx-preview's own output shape (`div.docx-wrapper > section.docx`, one
 * per page), so this builds THAT exact shape for the accurate path too
 * (an `<img>` filling each `section.docx` instead of live DOM content) —
 * every caller's existing CSS (scaling, shadows, page counting via
 * `section.docx`) keeps working unchanged, whichever path actually rendered.
 */
export async function renderDocPreview(
  docxBytes: Uint8Array,
  container: HTMLElement
): Promise<{ pageCount: number; accurate: boolean }> {
  try {
    const images = await api.docxToPageImages(docxBytes)
    if (images.length === 0) throw new Error('LibreOffice produced no page images.')
    container.innerHTML = ''
    const wrapper = document.createElement('div')
    wrapper.className = 'docx-wrapper'
    for (const png of images) {
      const section = document.createElement('section')
      section.className = 'docx'
      section.style.width = `${PAGE_WIDTH}px`
      const img = document.createElement('img')
      img.src = `data:image/png;base64,${uint8ToBase64(png)}`
      img.alt = ''
      img.style.display = 'block'
      img.style.width = '100%'
      img.style.height = 'auto'
      section.appendChild(img)
      wrapper.appendChild(section)
    }
    container.appendChild(wrapper)
    return { pageCount: images.length, accurate: true }
  } catch {
    container.innerHTML = ''
    await renderAsync(docxBytes, container, undefined, DOCX_PREVIEW_OPTIONS)
    normalizeDocxTextboxes(container)
    return { pageCount: container.querySelectorAll('section.docx').length, accurate: false }
  }
}

/** Builds the same `div.docx-wrapper > section.docx > img` shape as renderDocPreview's accurate path, from an already-rendered set of page PNGs, into one container. */
function paintPageImages(images: Uint8Array[], container: HTMLElement): void {
  container.innerHTML = ''
  const wrapper = document.createElement('div')
  wrapper.className = 'docx-wrapper'
  for (const png of images) {
    const section = document.createElement('section')
    section.className = 'docx'
    section.style.width = `${PAGE_WIDTH}px`
    const img = document.createElement('img')
    img.src = `data:image/png;base64,${uint8ToBase64(png)}`
    img.alt = ''
    img.style.display = 'block'
    img.style.width = '100%'
    img.style.height = 'auto'
    section.appendChild(img)
    wrapper.appendChild(section)
  }
  container.appendChild(wrapper)
}

/**
 * Same accurate LibreOffice-backed render as renderDocPreview, but for
 * several documents/containers at once — e.g. a tile grid's live thumbnails,
 * which used to call renderDocPreview once per tile (up to 6-12 independent
 * LibreOffice conversions racing each other for the same profile lock; see
 * core/docxToPdf.ts's docxBuffersToPageImages for the measured cost of that).
 * Converts every document in ONE batched IPC round-trip instead, then paints
 * each container from its own slice of the results. Falls back to rendering
 * each item through the ordinary single-document renderDocPreview (which has
 * its own docx-preview.js fallback) if the batched call itself fails outright
 * — e.g. LibreOffice missing entirely — so a batch failure doesn't blank
 * every tile when the ordinary single-document path would still work.
 */
export async function renderDocPreviewBatch(
  items: { docxBytes: Uint8Array; container: HTMLElement }[]
): Promise<void> {
  if (items.length === 0) return
  try {
    const results = await api.docxToPageImagesBatch(items.map((it) => it.docxBytes))
    items.forEach((it, i) => paintPageImages(results[i] ?? [], it.container))
  } catch {
    await Promise.all(items.map((it) => renderDocPreview(it.docxBytes, it.container)))
  }
}

export function pageShellStyle(): string {
  return `
    html, body { margin: 0; }
    html { background: #d8d8dc; overflow-x: hidden; }
    body {
      width: ${PAGE_WIDTH}px;
      min-height: ${PAGE_HEIGHT}px;
      margin: 0 auto;
      padding: ${MARGIN}px;
      box-sizing: border-box;
      background: #ffffff;
      box-shadow: 0 0 0 1px rgba(16, 24, 40, 0.12), 0 1px 3px rgba(16, 24, 40, 0.08);
      font-family: "Times New Roman", Times, serif;
      font-size: 15px;
      line-height: 1.15;
      color: #1a1a1a;
    }
  `
}

/**
 * Estimates which page is currently at the top of the iframe's viewport
 * (scroll position ÷ page height), so a "Page X of Y" badge can be shown
 * while scrolling a multi-page pasted document — only once there's more
 * than one page's worth of content. This is a position estimate, not a
 * readout of real page breaks (this document has none — see
 * pageShellStyle above).
 *
 * Handles two different ways a caller's iframe document becomes ready:
 * synchronously (document.open()/write()/close(), no 'load' event fires
 * for that) and asynchronously (a `srcDoc` change, which reloads the
 * iframe and fires 'load' once the new document is parsed) — attaching
 * once up front covers the first case, and the 'load' listener covers the
 * second, re-attaching against whatever document is current each time.
 */
export function usePageScrollTracker(frameRef: RefObject<HTMLIFrameElement>, contentVersion: unknown) {
  const [state, setState] = useState({ current: 1, total: 1 })

  useEffect(() => {
    const frame = frameRef.current
    if (!frame) return
    let cleanup = () => {}

    function attach() {
      cleanup()
      const win = frame!.contentWindow
      const doc = frame!.contentDocument
      if (!win || !doc || !doc.documentElement) {
        cleanup = () => {}
        return
      }

      function recompute() {
        const scrollTop = win!.scrollY
        const height = doc!.documentElement.scrollHeight
        const total = Math.max(1, Math.ceil(height / PAGE_HEIGHT))
        const current = Math.min(total, Math.floor(scrollTop / PAGE_HEIGHT) + 1)
        setState({ current, total })
      }

      recompute()
      win.addEventListener('scroll', recompute)
      doc.addEventListener('input', recompute)
      const ro = new ResizeObserver(recompute)
      ro.observe(doc.documentElement)
      cleanup = () => {
        win.removeEventListener('scroll', recompute)
        doc.removeEventListener('input', recompute)
        ro.disconnect()
      }
    }

    attach()
    frame.addEventListener('load', attach)
    return () => {
      frame.removeEventListener('load', attach)
      cleanup()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [contentVersion])

  return state
}
