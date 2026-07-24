import { pageShellStyle, PAGE_WIDTH, PAGE_HEIGHT } from './docPage'

const THUMB_WIDTH = 110
const SCALE = THUMB_WIDTH / PAGE_WIDTH
const THUMB_HEIGHT = Math.round(PAGE_HEIGHT * SCALE)

interface Props {
  html: string
}

/**
 * A live, shrunk-down render of the actual document (top of page 1) — the
 * same isolated iframe + pageShellStyle used for the full editor/preview,
 * just displayed at thumbnail size via a CSS transform: scale. Purely
 * decorative (not interactive), so scrollbars are hidden and pointer
 * events are switched off.
 */
export default function DocThumbnail({ html }: Props) {
  return (
    <div className="doc-thumb-box" style={{ width: THUMB_WIDTH, height: THUMB_HEIGHT }}>
      <iframe
        className="doc-thumb-frame"
        title="Document thumbnail"
        tabIndex={-1}
        style={{ width: PAGE_WIDTH, height: PAGE_HEIGHT, transform: `scale(${SCALE})` }}
        srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>${pageShellStyle()}\nhtml,body{overflow:hidden!important}</style></head><body>${html}</body></html>`}
      />
    </div>
  )
}
