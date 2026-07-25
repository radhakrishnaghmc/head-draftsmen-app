import { useEffect, useState } from 'react'
import type { RefObject } from 'react'

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
