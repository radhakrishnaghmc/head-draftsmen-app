import { useEffect, useRef } from 'react'
import { renderAsync } from 'docx-preview'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH, PAGE_HEIGHT } from './docPage'

const THUMB_WIDTH = 110
const SCALE = THUMB_WIDTH / PAGE_WIDTH
const THUMB_HEIGHT = Math.round(PAGE_HEIGHT * SCALE)

interface Props {
  docx: string
}

/**
 * A live, shrunk-down render of the actual document (top of page 1) — the
 * same docx-preview render used for the full editor/preview, just displayed
 * at thumbnail size via a CSS transform: scale, and clipped to one page's
 * worth by the fixed-size, overflow:hidden .doc-thumb-box around it. Purely
 * decorative (not interactive), so pointer events are switched off.
 */
export default function DocThumbnail({ docx }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)

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
    <div className="doc-thumb-box" style={{ width: THUMB_WIDTH, height: THUMB_HEIGHT }}>
      <div
        ref={containerRef}
        className="doc-thumb-frame"
        style={{ width: PAGE_WIDTH, height: PAGE_HEIGHT, transform: `scale(${SCALE})` }}
      />
    </div>
  )
}
