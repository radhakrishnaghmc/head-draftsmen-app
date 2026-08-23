import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { PDFDocument } from 'pdf-lib'
import { api } from '../ipc'
import { pdfPagesToDataUrls } from '../pdfToImages'
import { IconChevronLeft, IconImage, IconFolder, IconDoc, IconEye, IconDownload, IconTrash, IconWarn, IconOpen, IconCrop } from './Icons'
import PhotoAdjustModal from './PhotoAdjustModal'
import { closeOnBackdropMouseDown } from '../overlayClose'

// One picked image, held as a data URL (same-origin, so the canvas we analyse it
// on is never tainted). `edited` is set once the user has manually cropped,
// flattened or filtered the photo in the Adjust modal — that photo's dataUrl is
// already framed the way they want, so the global auto-crop toggle skips it.
interface Pic {
  id: string
  name: string
  dataUrl: string
  edited?: boolean
}

type PageSize = 'a4' | 'letter' | 'fit'
type Orientation = 'auto' | 'portrait' | 'landscape'

// Points per millimetre (PDF user space is 72 dpi → 1pt).
const MM = 2.834645669
// Portrait [short, long] edges in points.
const SIZES: Record<Exclude<PageSize, 'fit'>, [number, number]> = {
  a4: [595.28, 841.89],
  letter: [612, 792]
}

function nextId(): string {
  return `pic-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(file)
  })
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

function dataUrlToBytes(dataUrl: string): Uint8Array {
  const b64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

// The "auto margins" feature: find the bounding box of the actual document/content
// inside a photo and drop the surrounding background, so each page comes out
// tight like a scan instead of a small sheet floating on a desk.
//
// How: shrink the image to a small analysis canvas, estimate the background colour
// from the median of a strip running along all four edges (robust even if one
// corner is occluded by a hand or shadow — a single corner pixel was too fragile),
// then mark every pixel that differs from it by more than a threshold as "content".
// A row/column counts as content only once enough of its pixels are content (so
// specks and JPEG noise don't blow the box out to the full frame). The box is
// padded slightly and mapped back to the full-resolution image. Returns the crop
// rectangle in the ORIGINAL image's pixels, or null when nothing clear was found
// (then the whole image is used unchanged).
export function detectContentBox(img: HTMLImageElement): { x: number; y: number; w: number; h: number } | null {
  const nw = img.naturalWidth || img.width
  const nh = img.naturalHeight || img.height
  const scale = Math.min(1, 1100 / Math.max(nw, nh))
  const aw = Math.max(1, Math.round(nw * scale))
  const ah = Math.max(1, Math.round(nh * scale))
  const canvas = document.createElement('canvas')
  canvas.width = aw
  canvas.height = ah
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.drawImage(img, 0, 0, aw, ah)
  const px = ctx.getImageData(0, 0, aw, ah).data

  const at = (x: number, y: number) => {
    const i = (y * aw + x) * 4
    return [px[i], px[i + 1], px[i + 2]] as const
  }
  // Background colour: median of pixels sampled along a strip on all four edges,
  // so one occluded/shadowed corner can't skew the whole estimate.
  const strip = Math.max(2, Math.round(Math.min(aw, ah) * 0.02))
  const rs: number[] = []
  const gs: number[] = []
  const bs: number[] = []
  const sample = (x: number, y: number) => {
    const [r, g, b] = at(x, y)
    rs.push(r)
    gs.push(g)
    bs.push(b)
  }
  for (let x = 0; x < aw; x += 2) {
    for (let y = 0; y < strip; y++) sample(x, y)
    for (let y = ah - strip; y < ah; y++) sample(x, y)
  }
  for (let y = 0; y < ah; y += 2) {
    for (let x = 0; x < strip; x++) sample(x, y)
    for (let x = aw - strip; x < aw; x++) sample(x, y)
  }
  const median = (arr: number[]) => {
    arr.sort((a, b) => a - b)
    return arr[Math.floor(arr.length / 2)]
  }
  const bg = [median(rs), median(gs), median(bs)]
  const THRESH = 60 // per-channel-sum difference that counts as "not background"
  const isContent = (x: number, y: number) => {
    const [r, g, b] = at(x, y)
    return Math.abs(r - bg[0]) + Math.abs(g - bg[1]) + Math.abs(b - bg[2]) > THRESH
  }

  // Need a meaningful share of a row/column to be content before we trust it.
  const rowMin = Math.max(3, Math.floor(aw * 0.02))
  const colMin = Math.max(3, Math.floor(ah * 0.02))
  let top = -1
  let bottom = -1
  for (let y = 0; y < ah; y++) {
    let n = 0
    for (let x = 0; x < aw; x++) if (isContent(x, y)) n++
    if (n >= rowMin) {
      if (top === -1) top = y
      bottom = y
    }
  }
  let left = -1
  let right = -1
  for (let x = 0; x < aw; x++) {
    let n = 0
    for (let y = 0; y < ah; y++) if (isContent(x, y)) n++
    if (n >= colMin) {
      if (left === -1) left = x
      right = x
    }
  }
  if (top === -1 || left === -1 || bottom <= top || right <= left) return null

  // A little breathing room, then map back to full resolution.
  const pad = Math.round(Math.max(aw, ah) * 0.01)
  top = Math.max(0, top - pad)
  left = Math.max(0, left - pad)
  bottom = Math.min(ah - 1, bottom + pad)
  right = Math.min(aw - 1, right + pad)

  const box = {
    x: Math.floor(left / scale),
    y: Math.floor(top / scale),
    w: Math.ceil((right - left + 1) / scale),
    h: Math.ceil((bottom - top + 1) / scale)
  }
  // If the box is basically the whole frame, there was nothing to trim.
  if (box.w >= nw * 0.97 && box.h >= nh * 0.97) return null
  return box
}

// Normalise every image to JPEG bytes (pdf-lib only embeds JPEG/PNG), optionally
// cropped to the detected content box. Returns the bytes and the pixel size drawn.
async function toJpegBytes(dataUrl: string, autoCrop: boolean): Promise<{ bytes: Uint8Array; w: number; h: number }> {
  const img = await loadImage(dataUrl)
  const nw = img.naturalWidth || img.width
  const nh = img.naturalHeight || img.height
  const box = autoCrop ? detectContentBox(img) : null
  const sx = box?.x ?? 0
  const sy = box?.y ?? 0
  const w = box?.w ?? nw
  const h = box?.h ?? nh
  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not process an image.')
  // White backing, so any transparency (e.g. PNG) flattens cleanly for print.
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  ctx.drawImage(img, sx, sy, w, h, 0, 0, w, h)
  return { bytes: dataUrlToBytes(canvas.toDataURL('image/jpeg', 0.92)), w, h }
}

interface Props {
  onBack: () => void
}

/**
 * Tool: combine a batch of photos (or the pages of an uploaded PDF, re-imaged)
 * into a single PDF — one image per page, in the order shown (drag to reorder).
 * Each page's size (A4 / Letter / fit-to-photo), orientation and margin are
 * chosen in the rail. "Auto-crop borders" (the auto-margins feature) detects the
 * document/content edges in each photo and trims the surrounding background so
 * pages come out tight like a scan. Any photo can also be hand-adjusted (the
 * "Adjust" button on its tile, PhotoAdjustModal) — crop by a rectangle or flatten
 * a skewed shot by dragging its 4 corners onto a straight rectangle (perspective
 * correction), then optionally apply a filter (contrast boost / magic color /
 * B&W). An adjusted photo skips the global auto-crop, since it's already framed.
 * All image work is done here in the renderer (canvas + pdf-lib); the main
 * process only writes the finished bytes to disk.
 */
export default function PhotosToPdfPage({ onBack }: Props) {
  const [pics, setPics] = useState<Pic[]>([])
  const [pageSize, setPageSize] = useState<PageSize>('a4')
  const [orientation, setOrientation] = useState<Orientation>('auto')
  const [marginMm, setMarginMm] = useState(10)
  const [autoCrop, setAutoCrop] = useState(true)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  // A blob: URL for the built PDF, shown in the preview modal's iframe (Chromium's
  // own PDF viewer). Revoked whenever it's replaced or the tool unmounts.
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)
  const [cropPicId, setCropPicId] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => () => { if (previewUrl) URL.revokeObjectURL(previewUrl) }, [previewUrl])

  const hasPics = pics.length > 0

  async function addFiles(files: File[]) {
    setError(null)
    setSaved(null)
    setBusy(true)
    try {
      const added: Pic[] = []
      for (const file of files) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // A PDF is expanded to one image per page, so its pages join the batch.
          const pages = await pdfPagesToDataUrls(file)
          const base = file.name.replace(/\.pdf$/i, '')
          pages.forEach((dataUrl, i) => added.push({ id: nextId(), name: `${base} (page ${i + 1})`, dataUrl }))
        } else if (file.type.startsWith('image/') || /\.(jpe?g|png|gif|bmp|webp)$/i.test(file.name)) {
          added.push({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) })
        }
      }
      if (added.length === 0) throw new Error('No photos or PDFs were found in the selection.')
      setPics((cur) => [...cur, ...added])
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  function onPick(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files ? Array.from(e.target.files) : []
    e.target.value = ''
    if (files.length) void addFiles(files)
  }

  function removePic(id: string) {
    setPics((cur) => cur.filter((p) => p.id !== id))
  }

  function applyAdjust(id: string, dataUrl: string) {
    setPics((cur) => cur.map((p) => (p.id === id ? { ...p, dataUrl, edited: true } : p)))
    setCropPicId(null)
  }

  const cropPic = cropPicId ? pics.find((p) => p.id === cropPicId) ?? null : null

  function clearAll() {
    setPics([])
    setError(null)
    setSaved(null)
  }

  function onDragStart(e: React.DragEvent, i: number) {
    setDragIndex(i)
    e.dataTransfer.effectAllowed = 'move'
  }
  function onDragOver(e: React.DragEvent, i: number) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    if (i !== overIndex) setOverIndex(i)
  }
  function onDrop(e: React.DragEvent, i: number) {
    e.preventDefault()
    if (dragIndex !== null && dragIndex !== i) {
      setPics((cur) => {
        const next = [...cur]
        const [moved] = next.splice(dragIndex, 1)
        next.splice(i, 0, moved)
        return next
      })
    }
    setDragIndex(null)
    setOverIndex(null)
  }
  function onDragEnd() {
    setDragIndex(null)
    setOverIndex(null)
  }

  // Build the PDF bytes from the current photos + options — shared by Preview
  // (shows them in an iframe) and Export (writes them to disk).
  async function buildPdfBytes(): Promise<Uint8Array> {
    const pdf = await PDFDocument.create()
    const margin = Math.max(0, marginMm) * MM
    for (const pic of pics) {
      const { bytes, w: iw, h: ih } = await toJpegBytes(pic.dataUrl, autoCrop && !pic.edited)
      const image = await pdf.embedJpg(bytes)

      let pw: number
      let ph: number
      if (pageSize === 'fit') {
        // Page hugs the image (1px → 1pt), plus the chosen margin all round.
        pw = iw + margin * 2
        ph = ih + margin * 2
      } else {
        const [short, long] = SIZES[pageSize]
        const landscape = orientation === 'landscape' || (orientation === 'auto' && iw > ih)
        pw = landscape ? long : short
        ph = landscape ? short : long
      }

      const page = pdf.addPage([pw, ph])
      const availW = Math.max(1, pw - margin * 2)
      const availH = Math.max(1, ph - margin * 2)
      const s = Math.min(availW / iw, availH / ih)
      const dw = iw * s
      const dh = ih * s
      page.drawImage(image, { x: (pw - dw) / 2, y: (ph - dh) / 2, width: dw, height: dh })
    }
    return pdf.save()
  }

  async function previewPdf() {
    if (!hasPics || busy) return
    setBusy(true)
    setError(null)
    try {
      const out = await buildPdfBytes()
      const url = URL.createObjectURL(new Blob([out as unknown as BlobPart], { type: 'application/pdf' }))
      setPreviewUrl(url) // the unmount/replace effect revokes the previous one
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  function closePreview() {
    setPreviewUrl(null)
  }

  async function exportPdf() {
    if (!hasPics || busy) return
    setBusy(true)
    setError(null)
    setSaved(null)
    try {
      const out = await buildPdfBytes()
      const res = await api.savePdf(out, 'Photos.pdf')
      if (res) {
        setSaved(res.file)
        closePreview()
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconDoc /> <span className="pdf-ws-title-label">Tool:</span> Photos → PDF
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={busy}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <input ref={fileInputRef} type="file" accept="image/*,application/pdf,.pdf" multiple hidden onChange={onPick} />

      <div className="card pdf-workspace">
        <div className="pdf-ws-body">
          <div className="pdf-ws-main">
            {error && (
              <div className="notice error tool-outcome">
                <IconWarn /> {error}
              </div>
            )}

            {!hasPics ? (
              <button className="pdf-ws-dropzone" onClick={() => fileInputRef.current?.click()} disabled={busy}>
                <IconImage />
                <span className="pdf-ws-dropzone-title">{busy ? 'Reading…' : 'Upload photos or PDF'}</span>
                <span className="pdf-ws-dropzone-sub">Each photo (and each PDF page) becomes one page of the new PDF</span>
              </button>
            ) : (
              <div className="photo-tile-grid">
                {pics.map((pic, i) => (
                  <div
                    key={pic.id}
                    className={[
                      'photo-tile-card',
                      dragIndex === i ? 'dragging' : '',
                      overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
                    ]
                      .filter(Boolean)
                      .join(' ')}
                    draggable
                    onDragStart={(e) => onDragStart(e, i)}
                    onDragOver={(e) => onDragOver(e, i)}
                    onDrop={(e) => onDrop(e, i)}
                    onDragEnd={onDragEnd}
                  >
                    <span className="photo-tile-number">{i + 1}</span>
                    <button className="photo-tile-remove" title="Remove" onClick={() => removePic(pic.id)}>
                      <IconTrash />
                    </button>
                    <img className="photo-tile-img" src={pic.dataUrl} alt={pic.name} />
                    <div className="photo-tile-actions">
                      <button className="photo-tile-crop" title="Crop, flatten or filter this photo" onClick={() => setCropPicId(pic.id)}>
                        <IconCrop /> {pic.edited ? 'Edit' : 'Adjust'}
                      </button>
                      {pic.edited && <span className="photo-tile-cropped-badge">Edited</span>}
                    </div>
                    <div className="photo-tile-name" title={pic.name}>
                      {pic.name}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <aside className="pdf-ws-rail">
            <button
              className="primary pdf-ws-upload"
              onClick={() => fileInputRef.current?.click()}
              disabled={busy}
            >
              <IconFolder /> {busy ? 'Working…' : hasPics ? 'Add more photos' : 'Upload photos or PDF'}
            </button>

            {hasPics && (
              <>
                <div className="pdf-ws-rail-count">
                  <strong>{pics.length}</strong> photo{pics.length === 1 ? '' : 's'} — drag to reorder
                </div>

                <div className="photos-pdf-opts">
                  <label className="photos-pdf-opt">
                    <span>Page size</span>
                    <select value={pageSize} onChange={(e) => setPageSize(e.target.value as PageSize)}>
                      <option value="a4">A4</option>
                      <option value="letter">Letter</option>
                      <option value="fit">Fit to photo</option>
                    </select>
                  </label>
                  <label className="photos-pdf-opt">
                    <span>Orientation</span>
                    <select
                      value={orientation}
                      onChange={(e) => setOrientation(e.target.value as Orientation)}
                      disabled={pageSize === 'fit'}
                    >
                      <option value="auto">Auto (per photo)</option>
                      <option value="portrait">Portrait</option>
                      <option value="landscape">Landscape</option>
                    </select>
                  </label>
                  <label className="photos-pdf-opt">
                    <span>Margin (mm)</span>
                    <input
                      type="number"
                      min={0}
                      max={40}
                      value={marginMm}
                      onChange={(e) => setMarginMm(Math.max(0, Math.min(40, Number(e.target.value) || 0)))}
                    />
                  </label>
                  <label className="photos-pdf-check" title="Detect each photo's document edges and trim the background around it">
                    <input type="checkbox" checked={autoCrop} onChange={(e) => setAutoCrop(e.target.checked)} />
                    <span>Auto-crop borders (auto margins)</span>
                  </label>
                </div>

                <div className="pdf-ws-rail-actions">
                  <button className="primary" onClick={previewPdf} disabled={busy}>
                    <IconEye /> {busy ? 'Building…' : 'Preview PDF'}
                  </button>
                  <button className="pdf-ws-railbtn" onClick={exportPdf} disabled={busy}>
                    <IconDownload /> {busy ? 'Building…' : 'Export to PDF'}
                  </button>
                </div>

                <button className="pdf-ws-clearall" onClick={clearAll} disabled={busy}>
                  <IconTrash /> Remove all photos
                </button>
              </>
            )}

            {saved && (
              <div className="notice ok tool-result pdf-ws-rail-result">
                <span>
                  <IconDoc /> Saved: {saved}
                </span>
                <button className="ghost" onClick={() => api.openPath(saved)}>
                  <IconOpen /> Open
                </button>
              </div>
            )}

            <p className="pdf-ws-rail-hint">
              Pages follow the order shown. Auto-crop trims the plain background around each photo; turn it off to keep the
              whole frame. Use "Adjust" on a photo to crop it by hand, flatten an angled shot, or apply a scan filter.
            </p>
          </aside>
        </div>
      </div>

      {previewUrl &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(closePreview)}>
            <div className="pdf-preview-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
              <div className="pdf-preview-head">
                <span className="pdf-preview-title">
                  <IconDoc /> PDF preview — {pics.length} page{pics.length === 1 ? '' : 's'}
                </span>
                <div className="pdf-preview-head-actions">
                  <button className="pdf-ws-railbtn" onClick={exportPdf} disabled={busy}>
                    <IconDownload /> {busy ? 'Saving…' : 'Save PDF'}
                  </button>
                  <button className="ghost" onClick={closePreview}>
                    Close
                  </button>
                </div>
              </div>
              <iframe className="pdf-preview-frame" src={previewUrl} title="PDF preview" />
            </div>
          </div>,
          document.body
        )}

      {cropPic && (
        <PhotoAdjustModal
          name={cropPic.name}
          dataUrl={cropPic.dataUrl}
          onCancel={() => setCropPicId(null)}
          onApply={(dataUrl) => applyAdjust(cropPic.id, dataUrl)}
        />
      )}
    </>
  )
}
