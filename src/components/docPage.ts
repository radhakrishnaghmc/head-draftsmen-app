import { useEffect, useState } from 'react'
import type { RefObject } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { FONT_SUBSTITUTES, buildFontFallbackCss } from '@core/fontFallbacks'
import { scaleDocxText } from '@core/docxScale'
import { paginateFlowSections } from '@core/paginateFlow'

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
      // Skip anything inside a foreignObject: that's docx-preview's WordArt
      // fit-to-box scaleX (added to replicate Word's letterhead-title
      // stretch), not the stray full-width scale this loop is meant to undo.
      if (el.closest('foreignObject')) return
      if (el.style.transform && /scale/i.test(el.style.transform)) el.style.transform = 'none'
    })
  })

  // WordArt letterhead titles come in two different VML shapes across our
  // bundled templates, and docx-preview renders both as an <svg> that keeps
  // VML's authored `style` string verbatim (svg[style*="v-text-anchor"] is a
  // reliable marker for either kind — never emitted for anything else
  // docx-preview renders):
  //
  // 1. A plain text-box shape (`_x0000_t202`, e.g. work-order-template.docx)
  //    authored with an explicit `left:0; margin-left:…; margin-top:…` — an
  //    offset from the paragraph/cell hosting the drawing's run, exactly as
  //    Word anchors it. But docx-preview never marks that run's own wrapper
  //    `position: relative`, so the browser resolves the explicit `left:0`
  //    against the nearest positioned ancestor up the tree instead — usually
  //    the page root — landing the title far from its authored spot. Marking
  //    the immediate parent `relative` makes the browser resolve it the way
  //    Word does.
  // 2. A curved WordArt shape (`_x0000_t136`, e.g. zonal-work-order-
  //    template.docx) authored with only `margin-left`/`margin-top` and no
  //    `left` at all — its real position is Word's own in-flow placement of
  //    that single centered-text line, which fixing the ancestor can't
  //    recover. Dropping the absolute positioning and centering it in flow
  //    (keeping only its authored size) matches Word here instead.
  container.querySelectorAll<SVGElement>('svg[style*="v-text-anchor"]').forEach((svg) => {
    const style = svg.getAttribute('style') ?? ''
    if (style.includes('left:0')) {
      const parent = svg.parentElement
      if (parent && getComputedStyle(parent).position === 'static') parent.style.position = 'relative'
    } else {
      svg.style.position = 'static'
      svg.style.display = 'block'
      svg.style.margin = '0 auto'
    }
  })
}

// Fetched once (main process reads the bundled Carlito files off disk) and
// reused for every render — see core/fontFallbacks.ts for why this exists.
// A failure here (e.g. an old build missing the bundled files) shouldn't
// break the preview itself, so it just falls back to no substitute CSS.
let fontFallbackCssPromise: Promise<string> | undefined
function getFontFallbackCss(): Promise<string> {
  if (!fontFallbackCssPromise) {
    fontFallbackCssPromise = api
      .fontFallbackFiles()
      .then(buildFontFallbackCss)
      .catch(() => '')
  }
  return fontFallbackCssPromise
}

/**
 * Appends the metric-compatible substitute (see FONT_SUBSTITUTES) after each
 * matching inline font-family docx-preview wrote, and injects the @font-face
 * CSS that defines it — so the substitute is only ever used when the real
 * font (still listed first) isn't installed on the viewing machine. Preview-
 * only: the exported .docx itself never goes through this.
 */
async function addFontFallbacks(container: HTMLElement): Promise<void> {
  container.querySelectorAll<HTMLElement>('[style*="font-family"]').forEach((el) => {
    const primary = el.style.fontFamily.split(',')[0]?.trim().replace(/^['"]|['"]$/g, '')
    const substitute = primary && FONT_SUBSTITUTES[primary.toLowerCase()]
    if (substitute && !el.style.fontFamily.toLowerCase().includes(substitute.toLowerCase())) {
      el.style.fontFamily = `${el.style.fontFamily}, ${substitute}, sans-serif`
    }
  })
  const css = await getFontFallbackCss()
  if (!css) return
  const style = document.createElement('style')
  style.setAttribute('data-font-fallbacks', '')
  style.textContent = css
  container.insertBefore(style, container.firstChild)
}

/**
 * docx-preview doesn't position tab stops while it renders: each tab becomes
 * an empty `.docx-tab-stop` span, and the `word-spacing` that actually pushes
 * the following text to the stop is written later, from a `setTimeout(…, 500)`
 * inside the library's own `refreshTabStops` — its way of waiting until the
 * rendered nodes are in the DOM and measurable. `renderAsync` resolves well
 * before that timer fires, so a caller that snapshots `container.innerHTML`
 * straight after rendering (every print path does) captures tabs that haven't
 * been positioned yet: a right-aligned stop — e.g. the "Dated:" date on the
 * Work Order number line — then prints hard against the text to its left,
 * even though the on-screen preview (which stays live, and so gets the late
 * update) shows it right-aligned. Waiting the timer out before the snapshot is
 * what keeps print matching preview.
 */
const TAB_STOP_SETTLE_MS = 600

async function settleTabStops(container: HTMLElement): Promise<void> {
  if (!container.querySelector('.docx-tab-stop')) return
  await new Promise((resolve) => setTimeout(resolve, TAB_STOP_SETTLE_MS))
}

/**
 * Renders a filled .docx as an HTML preview via docx-preview's `renderAsync`
 * (see normalizeDocxTextboxes above for its text-box gap). `accurate` is
 * always false here — kept in the return type since callers branch on it —
 * this used to also have a PDFium-rasterized image path (LibreOffice
 * docx→PDF, then Electron's Chromium/PDFium PDF→PNG) for pixel-exact
 * previews, but that path produced blank/black pages on some machines and
 * was reverted in favor of this HTML render everywhere.
 */
export async function renderDocPreview(
  docxBytes: Uint8Array,
  container: HTMLElement,
  options?: {
    // Set when the caller is about to read the rendered HTML back out (printing)
    // rather than leave it on screen — see settleTabStops above.
    forSnapshot?: boolean
    // false renders the whole document as one continuous flow instead of
    // docx-preview's default per-real-page-break sections with gaps between
    // them — used by previews (e.g. Issue Tender Notice) that want a single
    // scrollable sheet rather than a stack of separate page blocks.
    breakPages?: boolean
    // docx-preview starts a new page only at an explicit page break, so a document that just flows
    // onto a second page shows as one tall sheet. Set to cut such a sheet into page-height sheets —
    // see paginateFlowSections.
    flowPages?: boolean
  }
): Promise<{ pageCount: number; accurate: boolean }> {
  const renderOptions =
    options?.breakPages === false ? { ...DOCX_PREVIEW_OPTIONS, breakPages: false } : DOCX_PREVIEW_OPTIONS
  if (options?.forSnapshot) {
    container.innerHTML = ''
    // The snapshot is re-laid-out in the print window, a clean page. docx-preview's
    // style rules are unscoped (`.docx p { text-indent: … }`), so another document's
    // preview already on this page leaks into measurements made here — the baked
    // tab widths (right-aligned "Date:" on the Work Order No line) then come out
    // wider than the clean print page has room for, and the line wraps. Measure in
    // a throwaway iframe instead, which has none of the app's other previews.
    const frame = document.createElement('iframe')
    frame.setAttribute('aria-hidden', 'true')
    frame.style.cssText = `position:fixed;top:-99999px;left:-99999px;width:${PAGE_WIDTH}px;height:1200px;border:0`
    document.body.appendChild(frame)
    try {
      const doc = frame.contentDocument!
      doc.open()
      doc.write(
        '<!DOCTYPE html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:#fff}' +
          '.docx-wrapper{background:#fff !important;padding:0 !important;margin:0 !important}' +
          '.docx-wrapper>section.docx{box-shadow:none !important;margin:0 auto !important}</style></head><body><div id="snap"></div></body></html>'
      )
      doc.close()
      const target = doc.getElementById('snap') as HTMLElement
      await renderAsync(docxBytes, target, undefined, renderOptions)
      normalizeDocxTextboxes(target)
      await addFontFallbacks(target)
      await settleTabStops(target)
      // The tall sheet is cut here, in the laid-out frame, so the print window gets real page sheets
      // (each with its own top/bottom padding) instead of one sheet the browser slices with no margins.
      if (options?.flowPages) paginateFlowSections(target)
      container.innerHTML = target.innerHTML
    } finally {
      frame.remove()
    }
    return { pageCount: container.querySelectorAll('section.docx').length, accurate: false }
  }
  // Live previews re-render on every keystroke. Rendering straight into `container` empties it first and
  // then waits on the async layout, so the page flashed blank on each edit (and two overlapping renders
  // could interleave). Render into an offscreen stage instead and swap the finished pages in; a render
  // that was overtaken by a newer one for the same container is dropped, not painted over it.
  const seq = (latestRender.get(container) ?? 0) + 1
  latestRender.set(container, seq)
  const stage = document.createElement('div')
  stage.className = container.className
  stage.style.cssText = `position:absolute;left:-99999px;top:0;visibility:hidden;width:${container.clientWidth || PAGE_WIDTH}px`
  ;(container.parentElement ?? document.body).appendChild(stage)
  try {
    await renderAsync(docxBytes, stage, undefined, renderOptions)
    normalizeDocxTextboxes(stage)
    await addFontFallbacks(stage)
    if (options?.flowPages) paginateFlowSections(stage)
    const pageCount = stage.querySelectorAll('section.docx').length
    if (latestRender.get(container) === seq) container.replaceChildren(...Array.from(stage.childNodes))
    return { pageCount, accurate: false }
  } finally {
    stage.remove()
  }
}

/** Newest render started per preview container — see renderDocPreview. */
const latestRender = new WeakMap<HTMLElement, number>()

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

/**
 * Keeps a one-sheet document (the Agreement Bond) on one A4 page: renders it offscreen, and if the filled-in
 * text runs past the sheet shrinks the body text/spacing (scaleDocxText) and re-measures until it fits.
 * Returns the .docx unchanged when it already fits, or after 5 passes / at a 75% floor — so Word, PDF and
 * print all get the same fitted file.
 */
export async function fitDocxToOnePage(docxB64: string): Promise<string> {
  const stage = document.createElement('div')
  stage.style.cssText = `position:absolute;left:-99999px;top:0;visibility:hidden;width:${PAGE_WIDTH}px`
  document.body.appendChild(stage)
  try {
    let factor = 1
    for (let pass = 0; pass < 5; pass++) {
      const candidate = scaleDocxText(docxB64, factor)
      stage.replaceChildren()
      await renderAsync(base64ToUint8(candidate), stage, undefined, DOCX_PREVIEW_OPTIONS)
      normalizeDocxTextboxes(stage)
      await addFontFallbacks(stage)
      const sheets = stage.querySelectorAll<HTMLElement>('section.docx')
      // An explicit page break is the template's own choice — nothing to fit.
      if (sheets.length !== 1) return docxB64
      const sheet = sheets[0]
      const height = sheet.getBoundingClientRect().height
      const limit = parseFloat(getComputedStyle(sheet).minHeight) || PAGE_HEIGHT
      if (height <= limit + 0.5) return candidate
      if (factor <= 0.75) return candidate
      const article = sheet.querySelector('article')?.getBoundingClientRect().height ?? height
      // Only the article shrinks (the letterhead gap is fixed), so aim its share of the overflow, with 3% headroom for Word's own layout.
      factor = Math.max(0.75, factor * ((article - (height - limit)) / article) * 0.97)
    }
    return scaleDocxText(docxB64, factor)
  } finally {
    stage.remove()
  }
}
