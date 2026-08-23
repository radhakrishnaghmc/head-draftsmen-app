import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { detectContentBox } from './PhotosToPdfPage'
import { api } from '../ipc'
import { IconCheck, IconCrop, IconRefresh, IconWarn } from './Icons'
import { closeOnBackdropMouseDown } from '../overlayClose'

type Pt = [number, number]

// A crop rectangle expressed as fractions (0..1) of the CURRENT working image's
// pixel size, so it stays valid no matter how large the on-screen preview is drawn.
interface CropRect {
  x: number
  y: number
  w: number
  h: number
}

export type PhotoFilter = 'none' | 'bw' | 'magic' | 'contrast'

const FULL: CropRect = { x: 0, y: 0, w: 1, h: 1 }
const MIN_FRAC = 0.05

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v))
}

function clamp255(v: number): number {
  return v < 0 ? 0 : v > 255 ? 255 : v
}

// `decode()` (not `onload`) is what actually guarantees the image is fully
// rasterised before it's handed off — with `onload`, a progressive JPEG (which
// is exactly what WhatsApp/most phones produce) can still be mid-decode, and
// drawing it to canvas at that point bakes in corrupted blotches over the
// undecoded regions instead of throwing.
async function loadImage(src: string): Promise<HTMLImageElement> {
  const img = new Image()
  img.src = src
  try {
    await img.decode()
  } catch {
    throw new Error('Could not read an image.')
  }
  return img
}

function canvasToDataUrl(canvas: HTMLCanvasElement): string {
  return canvas.toDataURL('image/jpeg', 0.95)
}

// ---- Rectangle crop (axis-aligned, dragged corners) ------------------------

type RectHandle = 'move' | 'nw' | 'ne' | 'sw' | 'se'

// Recomputed from the rect as it was when the drag started plus the total pointer
// delta since then (never incrementally), so nothing drifts from rounding.
function applyRectDrag(handle: RectHandle, start: CropRect, dxFrac: number, dyFrac: number): CropRect {
  if (handle === 'move') {
    return {
      x: clamp(start.x + dxFrac, 0, 1 - start.w),
      y: clamp(start.y + dyFrac, 0, 1 - start.h),
      w: start.w,
      h: start.h
    }
  }
  let x1 = start.x
  let y1 = start.y
  let x2 = start.x + start.w
  let y2 = start.y + start.h
  if (handle === 'nw' || handle === 'sw') x1 = clamp(start.x + dxFrac, 0, x2 - MIN_FRAC)
  if (handle === 'ne' || handle === 'se') x2 = clamp(x2 + dxFrac, x1 + MIN_FRAC, 1)
  if (handle === 'nw' || handle === 'ne') y1 = clamp(start.y + dyFrac, 0, y2 - MIN_FRAC)
  if (handle === 'sw' || handle === 'se') y2 = clamp(y2 + dyFrac, y1 + MIN_FRAC, 1)
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }
}

function cropToCanvas(img: HTMLImageElement, rect: CropRect): HTMLCanvasElement {
  const nw = img.naturalWidth
  const nh = img.naturalHeight
  const sx = Math.round(rect.x * nw)
  const sy = Math.round(rect.y * nh)
  const sw = Math.max(1, Math.round(rect.w * nw))
  const sh = Math.max(1, Math.round(rect.h * nh))
  const canvas = document.createElement('canvas')
  canvas.width = sw
  canvas.height = sh
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not process the photo.')
  ctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh)
  return canvas
}

// ---- Perspective flatten (4 free corners → straight rectangle) -------------

function capSize(w: number, h: number, max: number): { w: number; h: number } {
  const s = Math.min(1, max / Math.max(w, h))
  return { w: Math.max(1, Math.round(w * s)), h: Math.max(1, Math.round(h * s)) }
}

// Solves the 8x8 linear system for a projective transform's coefficients
// [a,b,c,d,e,f,g,h] (with the implicit 3,3 entry = 1) mapping `from` → `to`:
//   toX = (a·x + b·y + c) / (g·x + h·y + 1)
//   toY = (d·x + e·y + f) / (g·x + h·y + 1)
function homographyCoeffs(from: Pt[], to: Pt[]): number[] {
  const A: number[][] = []
  const B: number[] = []
  for (let i = 0; i < 4; i++) {
    const [sx, sy] = from[i]
    const [dx, dy] = to[i]
    A.push([sx, sy, 1, 0, 0, 0, -dx * sx, -dx * sy])
    B.push(dx)
    A.push([0, 0, 0, sx, sy, 1, -dy * sx, -dy * sy])
    B.push(dy)
  }
  const n = 8
  const M = A.map((row, i) => [...row, B[i]])
  for (let col = 0; col < n; col++) {
    let piv = col
    for (let r = col + 1; r < n; r++) if (Math.abs(M[r][col]) > Math.abs(M[piv][col])) piv = r
    ;[M[col], M[piv]] = [M[piv], M[col]]
    const pv = M[col][col] || 1e-9
    for (let c = col; c <= n; c++) M[col][c] /= pv
    for (let r = 0; r < n; r++) {
      if (r === col) continue
      const f = M[r][col]
      if (f === 0) continue
      for (let c = col; c <= n; c++) M[r][c] -= f * M[col][c]
    }
  }
  return M.map((row) => row[n])
}

function quadOutputSize(q: Pt[]): { w: number; h: number } {
  const dist = (a: Pt, b: Pt) => Math.hypot(a[0] - b[0], a[1] - b[1])
  const w = Math.max(dist(q[0], q[1]), dist(q[3], q[2]))
  const h = Math.max(dist(q[0], q[3]), dist(q[1], q[2]))
  return { w: Math.max(1, Math.round(w)), h: Math.max(1, Math.round(h)) }
}

// Flattens the quadrilateral `quad` (corners in the image's OWN pixel space, order
// [topLeft, topRight, bottomRight, bottomLeft]) into a straight W×H rectangle —
// the "perspective correction" step. Renders by inverse-mapping every OUTPUT pixel
// back into the source (via the to-quad homography) and bilinear-sampling it, which
// avoids gaps in the result. The source is capped to a sane resolution first so a
// full-frame photo can't turn a click into a multi-second per-pixel loop.
function warpQuadToRect(img: HTMLImageElement, quad: Pt[]): HTMLCanvasElement {
  const nw = img.naturalWidth
  const nh = img.naturalHeight
  const capped = capSize(nw, nh, 3500)
  const scale = capped.w / nw
  // `quad` holds fractions (0..1) of the natural image — convert to pixel
  // coordinates in the (possibly downscaled) source canvas built below.
  const quadPx: Pt[] = quad.map(([x, y]) => [x * nw * scale, y * nh * scale])

  const src = document.createElement('canvas')
  src.width = capped.w
  src.height = capped.h
  const sctx = src.getContext('2d')
  if (!sctx) throw new Error('Could not process the photo.')
  sctx.drawImage(img, 0, 0, capped.w, capped.h)
  const sdata = sctx.getImageData(0, 0, capped.w, capped.h).data

  const { w, h } = quadOutputSize(quadPx)
  const dst = document.createElement('canvas')
  dst.width = w
  dst.height = h
  const dctx = dst.getContext('2d')
  if (!dctx) throw new Error('Could not process the photo.')
  const out = dctx.createImageData(w, h)

  // Coeffs mapping the FLAT rectangle back onto the source quad, so each output
  // pixel can look up exactly where it came from (inverse warp).
  const [a, b, c, d, e, f, g, hh] = homographyCoeffs(
    [
      [0, 0],
      [w, 0],
      [w, h],
      [0, h]
    ],
    quadPx
  )
  const sw = capped.w
  const sh = capped.h
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const denom = g * x + hh * y + 1
      const sx = (a * x + b * y + c) / denom
      const sy = (d * x + e * y + f) / denom
      const di = (y * w + x) * 4
      if (sx < 0 || sy < 0 || sx > sw - 1 || sy > sh - 1) {
        out.data[di] = 255
        out.data[di + 1] = 255
        out.data[di + 2] = 255
        out.data[di + 3] = 255
        continue
      }
      const x0 = Math.floor(sx)
      const y0 = Math.floor(sy)
      const x1 = Math.min(sw - 1, x0 + 1)
      const y1 = Math.min(sh - 1, y0 + 1)
      const fx = sx - x0
      const fy = sy - y0
      for (let ch = 0; ch < 4; ch++) {
        const i00 = (y0 * sw + x0) * 4 + ch
        const i10 = (y0 * sw + x1) * 4 + ch
        const i01 = (y1 * sw + x0) * 4 + ch
        const i11 = (y1 * sw + x1) * 4 + ch
        const top = sdata[i00] * (1 - fx) + sdata[i10] * fx
        const bot = sdata[i01] * (1 - fx) + sdata[i11] * fx
        out.data[di + ch] = top * (1 - fy) + bot * fy
      }
    }
  }
  dctx.putImageData(out, 0, 0)
  return dst
}

// ---- Filters -----------------------------------------------------------------

// A large-radius blur approximated by downscaling then upscaling with smoothing —
// canvas's own resampler does the blur, so this needs no manual convolution loop.
// Used as a "local background" estimate so filters cope with uneven lighting.
function localBackground(canvas: HTMLCanvasElement): Uint8ClampedArray {
  const w = canvas.width
  const h = canvas.height
  const sw = Math.max(8, Math.round(w / 24))
  const sh = Math.max(8, Math.round(h / 24))
  const small = document.createElement('canvas')
  small.width = sw
  small.height = sh
  const sctx = small.getContext('2d')
  if (!sctx) throw new Error('Could not process the photo.')
  sctx.drawImage(canvas, 0, 0, sw, sh)
  const big = document.createElement('canvas')
  big.width = w
  big.height = h
  const bctx = big.getContext('2d')
  if (!bctx) throw new Error('Could not process the photo.')
  bctx.imageSmoothingEnabled = true
  bctx.imageSmoothingQuality = 'high'
  bctx.drawImage(small, 0, 0, w, h)
  return bctx.getImageData(0, 0, w, h).data
}

function applyFilter(canvas: HTMLCanvasElement, filter: 'contrast' | 'bw'): HTMLCanvasElement {
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not process the photo.')
  const w = canvas.width
  const h = canvas.height
  const id = ctx.getImageData(0, 0, w, h)
  const d = id.data

  if (filter === 'contrast') {
    const C = 55
    const factor = (259 * (C + 255)) / (255 * (259 - C))
    for (let i = 0; i < d.length; i += 4) {
      d[i] = clamp255(factor * (d[i] - 128) + 128)
      d[i + 1] = clamp255(factor * (d[i + 1] - 128) + 128)
      d[i + 2] = clamp255(factor * (d[i + 2] - 128) + 128)
    }
  } else {
    // Adaptive threshold: each pixel is compared to its own local neighbourhood's
    // brightness (not one global cutoff), so shadows/uneven lighting still binarize
    // cleanly to crisp black text on pure white.
    const bg = localBackground(canvas)
    for (let i = 0; i < d.length; i += 4) {
      const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]
      const bgG = 0.299 * bg[i] + 0.587 * bg[i + 1] + 0.114 * bg[i + 2]
      const v = g < bgG - 14 ? 0 : 255
      d[i] = v
      d[i + 1] = v
      d[i + 2] = v
    }
  }
  ctx.putImageData(id, 0, 0)
  return canvas
}

// ---- Magic color (OCR-driven) -------------------------------------------------

// A detected text line's box, always in the FULL working image's own pixel space
// (the OCR pass runs on a downscaled copy for speed; boxes are scaled back up
// once, right after the call, so every consumer works in one coordinate system).
interface OcrBox {
  x: number
  y: number
  w: number
  h: number
}

const OCR_ANALYSIS_MAX = 1600

// Runs OCR once per working image (cached — OCR is a multi-second main-process
// call, not something to repeat on every preview tick) and returns each detected
// text line's box in the full-resolution image's own pixel space.
function useOcrLines() {
  const cache = useRef<{ url: string; lines: OcrBox[] } | null>(null)
  async function ensure(url: string): Promise<OcrBox[]> {
    if (cache.current?.url === url) return cache.current.lines
    const img = await loadImage(url)
    const full = document.createElement('canvas')
    full.width = img.naturalWidth
    full.height = img.naturalHeight
    const fctx = full.getContext('2d')
    if (!fctx) throw new Error('Could not process the photo.')
    fctx.drawImage(img, 0, 0)
    const small = scaledCopy(full, OCR_ANALYSIS_MAX)
    const pages = await api.ocrPhotosToLayout([canvasToDataUrl(small)])
    const scaleX = full.width / small.width
    const scaleY = full.height / small.height
    const lines: OcrBox[] = (pages[0]?.lines ?? []).map((l) => ({
      x: l.x * scaleX,
      y: l.y * scaleY,
      w: l.w * scaleX,
      h: l.h * scaleY
    }))
    cache.current = { url, lines }
    return lines
  }
  return ensure
}

// "Magic color": OCR finds where the text actually is, then everything OUTSIDE
// those lines is forced pure white (background, shadows, stray marks all gone)
// and everything INSIDE them is thresholded against its own local brightness —
// dark ink turns solid black, the paper between letters turns white. `lines` are
// in the natural (full) image's pixel space; naturalW/H is that same space, so
// this scales correctly whether `canvas` is a small live-preview or the full-res
// output canvas.
function applyMagicColorMask(canvas: HTMLCanvasElement, lines: OcrBox[], naturalW: number, naturalH: number): void {
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not process the photo.')
  const w = canvas.width
  const h = canvas.height
  const sx = w / naturalW
  const sy = h / naturalH
  const id = ctx.getImageData(0, 0, w, h)
  const d = id.data

  const mask = new Uint8Array(w * h)
  const pad = Math.max(1, Math.round(Math.min(w, h) * 0.004))
  for (const l of lines) {
    const x0 = Math.max(0, Math.floor(l.x * sx) - pad)
    const y0 = Math.max(0, Math.floor(l.y * sy) - pad)
    const x1 = Math.min(w, Math.ceil((l.x + l.w) * sx) + pad)
    const y1 = Math.min(h, Math.ceil((l.y + l.h) * sy) + pad)
    for (let y = y0; y < y1; y++) {
      const row = y * w
      for (let x = x0; x < x1; x++) mask[row + x] = 1
    }
  }

  const bg = localBackground(canvas)
  for (let i = 0, p = 0; i < d.length; i += 4, p++) {
    if (!mask[p]) {
      d[i] = 255
      d[i + 1] = 255
      d[i + 2] = 255
      continue
    }
    const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]
    const bgG = 0.299 * bg[i] + 0.587 * bg[i + 1] + 0.114 * bg[i + 2]
    const v = g < bgG - 14 ? 0 : 255
    d[i] = v
    d[i + 1] = v
    d[i + 2] = v
  }
  ctx.putImageData(id, 0, 0)
}

function scaledCopy(canvas: HTMLCanvasElement, maxSide: number): HTMLCanvasElement {
  const { w, h } = capSize(canvas.width, canvas.height, maxSide)
  const out = document.createElement('canvas')
  out.width = w
  out.height = h
  const ctx = out.getContext('2d')
  if (!ctx) throw new Error('Could not process the photo.')
  ctx.drawImage(canvas, 0, 0, w, h)
  return out
}

// ---- Component -----------------------------------------------------------------

interface Props {
  name: string
  dataUrl: string
  onCancel: () => void
  onApply: (dataUrl: string) => void
}

type Tab = 'geometry' | 'filter'
type GeoMode = 'rect' | 'quad'

const FILTERS: { id: PhotoFilter; label: string; hint: string }[] = [
  { id: 'none', label: 'None', hint: 'Keep the photo as-is' },
  { id: 'contrast', label: 'Contrast boost', hint: 'Darkens faint ink so text reads clearly' },
  { id: 'magic', label: 'Magic color', hint: 'OCR finds the text, then whitens everything else and darkens the text' },
  { id: 'bw', label: 'B & W', hint: 'Scan-style black text on pure white' }
]

/**
 * "Adjust photo" — a small on-the-spot scanner: crop by a rectangle OR flatten a
 * skewed shot by dragging the document's 4 corners onto a straight rectangle
 * (perspective correction), then optionally punch up the result with a filter
 * (contrast boost / magic color flat-fielding / B&W). Each step is applied against
 * a local working copy that lives only in this modal — nothing touches the photo in
 * the batch until "Save changes".
 */
export default function PhotoAdjustModal({ name, dataUrl, onCancel, onApply }: Props) {
  const imgRef = useRef<HTMLImageElement>(null)
  const stageRef = useRef<HTMLDivElement>(null)
  const [tab, setTab] = useState<Tab>('geometry')
  const [geoMode, setGeoMode] = useState<GeoMode>('rect')
  const [workingUrl, setWorkingUrl] = useState(dataUrl)
  const [display, setDisplay] = useState<{ w: number; h: number }>({ w: 0, h: 0 })
  const [rect, setRect] = useState<CropRect>(FULL)
  const [quad, setQuad] = useState<Pt[]>([
    [0.08, 0.08],
    [0.92, 0.08],
    [0.92, 0.92],
    [0.08, 0.92]
  ])
  const [filter, setFilter] = useState<PhotoFilter>('none')
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const [ocrBusy, setOcrBusy] = useState(false)
  const [notice, setNotice] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const ensureOcrLines = useOcrLines()
  const rectDrag = useRef<{ handle: RectHandle; startX: number; startY: number; start: CropRect } | null>(null)
  const quadDrag = useRef<number | null>(null)

  function onImgLoad() {
    const img = imgRef.current
    if (!img) return
    const maxW = Math.min(760, window.innerWidth * 0.7)
    const maxH = Math.min(520, window.innerHeight * 0.56)
    const scale = Math.min(maxW / img.naturalWidth, maxH / img.naturalHeight, 1)
    setDisplay({
      w: Math.max(1, Math.round(img.naturalWidth * scale)),
      h: Math.max(1, Math.round(img.naturalHeight * scale))
    })
  }

  // Re-run the filter preview (on a downscaled copy, so switching filters stays
  // snappy) whenever the working image or the chosen filter changes. Magic color
  // needs an OCR pass first (main-process, multi-second) — that's cached per
  // working image by ensureOcrLines, so switching filters back and forth doesn't
  // re-run it.
  useEffect(() => {
    let cancelled = false
    setNotice((n) => (n?.startsWith('No text detected') ? null : n))
    if (filter === 'none') {
      setPreviewUrl(null)
      setOcrBusy(false)
      return
    }
    if (filter === 'magic') {
      setOcrBusy(true)
      void ensureOcrLines(workingUrl)
        .then(async (lines) => {
          if (cancelled) return
          if (lines.length === 0) {
            setNotice('No text detected in this photo — Magic color left it unchanged.')
            setPreviewUrl(null)
            return
          }
          const img = await loadImage(workingUrl)
          if (cancelled) return
          const base = document.createElement('canvas')
          base.width = img.naturalWidth
          base.height = img.naturalHeight
          base.getContext('2d')?.drawImage(img, 0, 0)
          const small = scaledCopy(base, 700)
          applyMagicColorMask(small, lines, img.naturalWidth, img.naturalHeight)
          if (!cancelled) setPreviewUrl(canvasToDataUrl(small))
        })
        .catch((e) => !cancelled && setNotice(e instanceof Error ? e.message : String(e)))
        .finally(() => !cancelled && setOcrBusy(false))
      return () => {
        cancelled = true
      }
    }
    void loadImage(workingUrl).then((img) => {
      if (cancelled) return
      const base = document.createElement('canvas')
      base.width = img.naturalWidth
      base.height = img.naturalHeight
      const ctx = base.getContext('2d')
      if (!ctx) return
      ctx.drawImage(img, 0, 0)
      const small = scaledCopy(base, 700)
      applyFilter(small, filter)
      if (!cancelled) setPreviewUrl(canvasToDataUrl(small))
    })
    return () => {
      cancelled = true
    }
  }, [workingUrl, filter])

  function beginRectDrag(e: React.PointerEvent, handle: RectHandle) {
    e.preventDefault()
    e.stopPropagation()
    ;(e.currentTarget as Element).setPointerCapture(e.pointerId)
    rectDrag.current = { handle, startX: e.clientX, startY: e.clientY, start: rect }
  }
  function onRectPointerMove(e: React.PointerEvent) {
    const drag = rectDrag.current
    if (!drag || !display.w || !display.h) return
    const dxFrac = (e.clientX - drag.startX) / display.w
    const dyFrac = (e.clientY - drag.startY) / display.h
    setRect(applyRectDrag(drag.handle, drag.start, dxFrac, dyFrac))
  }
  function endRectDrag() {
    rectDrag.current = null
  }

  function beginQuadDrag(e: React.PointerEvent, idx: number) {
    e.preventDefault()
    e.stopPropagation()
    ;(e.currentTarget as Element).setPointerCapture(e.pointerId)
    quadDrag.current = idx
  }
  function onQuadPointerMove(e: React.PointerEvent) {
    const idx = quadDrag.current
    if (idx === null || !stageRef.current) return
    const box = stageRef.current.getBoundingClientRect()
    const x = clamp((e.clientX - box.left) / display.w, 0, 1)
    const y = clamp((e.clientY - box.top) / display.h, 0, 1)
    setQuad((q) => q.map((p, i) => (i === idx ? ([x, y] as Pt) : p)))
  }
  function endQuadDrag() {
    quadDrag.current = null
  }

  async function autoDetect() {
    const img = imgRef.current
    if (!img) return
    // Guards against the same progressive-JPEG decode race applyGeometry() below
    // guards against — img is a live DOM <img>, not something routed through
    // loadImage(), so it needs its own decode() before any pixel read.
    await img.decode().catch(() => {})
    const box = detectContentBox(img)
    if (!box) {
      setNotice('No clear document edge found — adjust the corners/box by hand.')
      return
    }
    setNotice(null)
    const x = box.x / img.naturalWidth
    const y = box.y / img.naturalHeight
    const w = box.w / img.naturalWidth
    const h = box.h / img.naturalHeight
    if (geoMode === 'rect') {
      setRect({ x, y, w, h })
    } else {
      setQuad([
        [x, y],
        [x + w, y],
        [x + w, y + h],
        [x, y + h]
      ])
    }
  }

  function resetGeometry() {
    setNotice(null)
    setRect(FULL)
    setQuad([
      [0.08, 0.08],
      [0.92, 0.08],
      [0.92, 0.92],
      [0.08, 0.92]
    ])
  }

  function applyGeometry() {
    const img = imgRef.current
    if (!img) return
    setBusy(true)
    setNotice(null)
    // Deferred so the "Working…" state paints before the (possibly heavy,
    // synchronous) per-pixel crop/warp runs.
    setTimeout(() => {
      void (async () => {
        try {
          // Same progressive-JPEG decode race as loadImage() — img is a live DOM
          // <img>, so it gets its own guard rather than routing through that helper.
          await img.decode().catch(() => {})
          const canvas = geoMode === 'rect' ? cropToCanvas(img, rect) : warpQuadToRect(img, quad)
          setWorkingUrl(canvasToDataUrl(canvas))
          setRect(FULL)
          setQuad([
            [0.08, 0.08],
            [0.92, 0.08],
            [0.92, 0.92],
            [0.08, 0.92]
          ])
        } catch (e) {
          setNotice(e instanceof Error ? e.message : String(e))
        } finally {
          setBusy(false)
        }
      })()
    }, 0)
  }

  function resetAll() {
    setWorkingUrl(dataUrl)
    setFilter('none')
    resetGeometry()
  }

  function save() {
    setBusy(true)
    setNotice(null)
    setTimeout(() => {
      void (async () => {
        try {
          const img = await loadImage(workingUrl)
          const canvas = document.createElement('canvas')
          canvas.width = img.naturalWidth
          canvas.height = img.naturalHeight
          const ctx = canvas.getContext('2d')
          if (!ctx) throw new Error('Could not process the photo.')
          ctx.drawImage(img, 0, 0)
          if (filter === 'contrast' || filter === 'bw') {
            applyFilter(canvas, filter)
          } else if (filter === 'magic') {
            // Reuses the OCR pass the preview already ran for this working image,
            // if any — only a fresh main-process call when nothing's cached yet.
            const lines = await ensureOcrLines(workingUrl)
            if (lines.length === 0) {
              setNotice('No text detected in this photo — Magic color left it unchanged.')
            } else {
              applyMagicColorMask(canvas, lines, img.naturalWidth, img.naturalHeight)
            }
          }
          onApply(canvasToDataUrl(canvas))
        } catch (e) {
          setNotice(e instanceof Error ? e.message : String(e))
        } finally {
          setBusy(false)
        }
      })()
    }, 0)
  }

  const rectStyle = {
    left: `${rect.x * display.w}px`,
    top: `${rect.y * display.h}px`,
    width: `${rect.w * display.w}px`,
    height: `${rect.h * display.h}px`
  }
  const rectHandles: RectHandle[] = ['nw', 'ne', 'sw', 'se']

  return createPortal(
    <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(onCancel)}>
      <div className="photo-adjust-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-preview-head">
          <span className="pdf-preview-title">
            <IconCrop /> Adjust photo — {name}
          </span>
          <div className="pdf-preview-head-actions">
            <button className="ghost" onClick={resetAll} disabled={busy}>
              <IconRefresh /> Reset
            </button>
            <button className="ghost" onClick={onCancel} disabled={busy}>
              Cancel
            </button>
            <button className="primary" onClick={save} disabled={busy || ocrBusy}>
              <IconCheck /> {busy ? 'Working…' : 'Save changes'}
            </button>
          </div>
        </div>

        <div className="photo-adjust-tabs">
          <button className={tab === 'geometry' ? 'active' : ''} onClick={() => setTab('geometry')}>
            Crop &amp; flatten
          </button>
          <button className={tab === 'filter' ? 'active' : ''} onClick={() => setTab('filter')}>
            Filter
          </button>
        </div>

        {notice && (
          <div className="notice error tool-outcome photo-crop-notice">
            <IconWarn /> {notice}
          </div>
        )}

        {tab === 'geometry' ? (
          <>
            <div className="photo-adjust-subbar">
              <div className="photo-adjust-modebtns">
                <button className={geoMode === 'rect' ? 'active' : ''} onClick={() => setGeoMode('rect')}>
                  Rectangle crop
                </button>
                <button className={geoMode === 'quad' ? 'active' : ''} onClick={() => setGeoMode('quad')}>
                  Flatten (4 corners)
                </button>
              </div>
              <div className="photo-adjust-modebtns">
                <button onClick={autoDetect} disabled={busy}>
                  Auto-detect
                </button>
                <button onClick={resetGeometry} disabled={busy}>
                  Reset box
                </button>
                <button className="primary" onClick={applyGeometry} disabled={busy}>
                  {busy ? 'Working…' : geoMode === 'rect' ? 'Apply crop' : 'Flatten photo'}
                </button>
              </div>
            </div>

            <div className="photo-crop-body">
              <div
                ref={stageRef}
                className="photo-crop-stage"
                style={{ width: display.w || undefined, height: display.h || undefined }}
                onPointerMove={geoMode === 'rect' ? onRectPointerMove : onQuadPointerMove}
                onPointerUp={geoMode === 'rect' ? endRectDrag : endQuadDrag}
              >
                <img
                  ref={imgRef}
                  className="photo-crop-img"
                  src={workingUrl}
                  alt={name}
                  onLoad={onImgLoad}
                  style={{ width: display.w || undefined, height: display.h || undefined }}
                  draggable={false}
                />
                {display.w > 0 && geoMode === 'rect' && (
                  <>
                    <div className="photo-crop-mask" style={{ clipPath: rectMaskClipPath(rect) }} />
                    <div className="photo-crop-box" style={rectStyle} onPointerDown={(e) => beginRectDrag(e, 'move')}>
                      {rectHandles.map((h) => (
                        <span
                          key={h}
                          className={`photo-crop-handle photo-crop-handle-${h}`}
                          onPointerDown={(e) => beginRectDrag(e, h)}
                        />
                      ))}
                    </div>
                  </>
                )}
                {display.w > 0 && geoMode === 'quad' && (
                  <svg className="photo-crop-quad-svg" width={display.w} height={display.h}>
                    <path d={quadMaskPath(quad, display.w, display.h)} fillRule="evenodd" fill="rgba(0,0,0,0.55)" />
                    <polygon
                      points={quad.map(([x, y]) => `${x * display.w},${y * display.h}`).join(' ')}
                      fill="none"
                      stroke="var(--rose-500)"
                      strokeWidth={2}
                    />
                    {quad.map(([x, y], i) => (
                      <circle
                        key={i}
                        cx={x * display.w}
                        cy={y * display.h}
                        r={9}
                        className="photo-crop-quad-handle"
                        onPointerDown={(e) => beginQuadDrag(e, i)}
                      />
                    ))}
                  </svg>
                )}
              </div>
            </div>
            <p className="pdf-ws-rail-hint photo-crop-hint">
              {geoMode === 'rect'
                ? 'Drag inside the box to move it, drag a corner to resize.'
                : 'Drag each corner onto the document\'s actual corners, then "Flatten photo" to straighten it like a scan.'}
            </p>
          </>
        ) : (
          <div className="photo-adjust-filter">
            <div className="photo-adjust-filter-preview">
              <img src={previewUrl ?? workingUrl} alt={`${name} preview`} />
              {ocrBusy && (
                <div className="photo-adjust-filter-busy">
                  <span className="photo-adjust-spinner" />
                  Detecting text…
                </div>
              )}
            </div>
            <div className="photo-adjust-filter-opts">
              {FILTERS.map((f) => (
                <label key={f.id} className={`photo-adjust-filter-opt${filter === f.id ? ' active' : ''}`}>
                  <input type="radio" name="photo-filter" checked={filter === f.id} onChange={() => setFilter(f.id)} />
                  <span className="photo-adjust-filter-label">{f.label}</span>
                  <span className="photo-adjust-filter-hint">{f.hint}</span>
                </label>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>,
    document.body
  )
}

// Dims everything outside the crop box using a clip-path "donut" (evenodd fill
// between the full-stage rect and the crop rect).
function rectMaskClipPath(c: CropRect): string {
  const x1 = c.x * 100
  const y1 = c.y * 100
  const x2 = (c.x + c.w) * 100
  const y2 = (c.y + c.h) * 100
  return `polygon(evenodd, 0% 0%, 100% 0%, 100% 100%, 0% 100%, 0% 0%, ${x1}% ${y1}%, ${x1}% ${y2}%, ${x2}% ${y2}%, ${x2}% ${y1}%, ${x1}% ${y1}%)`
}

function quadMaskPath(q: Pt[], w: number, h: number): string {
  const pts = q.map(([x, y]) => `${x * w},${y * h}`).join(' L')
  return `M0,0 H${w} V${h} H0 Z M${pts} Z`
}
