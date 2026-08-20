import { useEffect, useRef } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH, PAGE_HEIGHT } from './docPage'

const DEFAULT_THUMB_WIDTH = 110

interface Props {
  docx: string
  /** Thumbnail width in px (height derives from the page ratio). Defaults to 110. */
  width?: number
}

/**
 * A live, shrunk-down render of the actual document (top of page 1) — the
 * approximate docx-preview.js render (NOT docPage.ts's accurate,
 * LibreOffice-backed renderDocPreview, which every full-size preview modal
 * uses), just displayed at thumbnail size via a CSS transform: scale, and
 * clipped to one page's worth by the fixed-size, overflow:hidden
 * .doc-thumb-box around it. A grid can show dozens of these at once, and each
 * accurate render is its own LibreOffice conversion — using it here would
 * hammer LibreOffice with concurrent conversions for a decorative icon few
 * people zoom in on, so this deliberately trades accuracy for speed. Purely
 * decorative (not interactive), so pointer events are switched off.
 */
export default function DocThumbnail({ docx, width = DEFAULT_THUMB_WIDTH }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const scale = width / PAGE_WIDTH
  const thumbHeight = Math.round(PAGE_HEIGHT * scale)

  useEffect(() => {
    let cancelled = false
    const container = containerRef.current
    if (!container) return
    void (async () => {
      try {
        const bytes = base64ToUint8(docx)
        if (cancelled) return
        container.innerHTML = ''
        await renderAsync(bytes, container, undefined, DOCX_PREVIEW_OPTIONS)
      } catch {
        // Purely decorative — a bad/corrupt buffer just stays blank.
      }
    })()
    return () => {
      cancelled = true
    }
  }, [docx])

  return (
    <div className="doc-thumb-box" style={{ width, height: thumbHeight }}>
      <div
        ref={containerRef}
        className="doc-thumb-frame"
        style={{ width: PAGE_WIDTH, height: PAGE_HEIGHT, transform: `scale(${scale})` }}
      />
    </div>
  )
}
