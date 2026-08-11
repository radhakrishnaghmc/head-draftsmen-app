import { useEffect, useRef, useState } from 'react'
import exifr from 'exifr'
import { api } from '../ipc'
import type { ExcelTable } from '@core/types'
import { parseGpsOverlay } from '@core/gpsOverlay'
import { IconChevronLeft, IconImage, IconFolder, IconTable, IconTrash, IconWarn, IconOpen } from './Icons'

// One photo's extracted GPS row. `source` records where the coordinates came
// from: EXIF metadata (exact), the OCR'd overlay stamp (best-effort), or none.
// `confidence` is 'high' for EXIF and for a clean overlay read, 'low' for a
// shaky overlay read (flagged for the user to verify), 'none' when nothing found.
interface PhotoRow {
  id: string
  // While a photo is still being read, its row is shown as a placeholder so the
  // table fills in live, row by row, alongside the batch being processed.
  pending?: boolean
  name: string
  lat?: number
  lon?: number
  latDMS?: string
  lonDMS?: string
  altitude?: string
  address?: string
  place?: string
  pincode?: string
  state?: string
  taken?: string
  source: 'EXIF' | 'OCR' | 'none'
  confidence: 'high' | 'low' | 'none'
  raw?: string
}

// Read GPS straight from a photo's EXIF metadata. Returns undefined (→ fall back
// to OCR) when the photo carries no embedded GPS.
type PhotoData = Omit<PhotoRow, 'id' | 'pending'>

async function readExif(file: File): Promise<PhotoData | undefined> {
  try {
    const d = await exifr.parse(file, { tiff: true, exif: true, gps: true })
    if (!d || typeof d.latitude !== 'number' || typeof d.longitude !== 'number') return undefined
    const taken = d.DateTimeOriginal ?? d.CreateDate ?? d.ModifyDate
    return {
      name: file.name,
      lat: d.latitude,
      lon: d.longitude,
      altitude: typeof d.GPSAltitude === 'number' ? `${d.GPSAltitude.toFixed(1)} m` : undefined,
      taken: taken instanceof Date && !isNaN(taken.getTime()) ? fmtDate(taken) : undefined,
      source: 'EXIF',
      confidence: 'high'
    }
  } catch {
    return undefined
  }
}

// The overlay path: OCR the stamped text (in the main process) and parse it.
async function readOverlay(file: File): Promise<PhotoData> {
  try {
    const bytes = new Uint8Array(await file.arrayBuffer())
    const lines = await api.ocrGpsOverlay(bytes)
    const g = parseGpsOverlay(lines)
    const found = g.lat != null || g.lon != null
    return {
      name: file.name,
      lat: g.lat,
      lon: g.lon,
      latDMS: g.latDMS,
      lonDMS: g.lonDMS,
      altitude: g.altitude,
      address: g.address,
      place: g.place,
      pincode: g.pincode,
      state: g.state,
      taken: g.dateTime,
      source: found ? 'OCR' : 'none',
      confidence: found ? g.confidence : 'none',
      raw: g.raw
    }
  } catch {
    return { name: file.name, source: 'none', confidence: 'none' }
  }
}

async function processFile(file: File): Promise<PhotoData> {
  return (await readExif(file)) ?? (await readOverlay(file))
}

function fmtDate(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`
}

const dec = (n: number | undefined): string => (typeof n === 'number' ? n.toFixed(6) : '')

// Natural width/height ratio of an image, so a Cover Flow card can be sized to the
// photo's shape (no letterbox bars behind it). Falls back to 1 if it can't load.
function imageRatio(url: string): Promise<number> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve(img.naturalWidth && img.naturalHeight ? img.naturalWidth / img.naturalHeight : 1)
    img.onerror = () => resolve(1)
    img.src = url
  })
}

interface Props {
  onBack: () => void
}

/**
 * Tool: read the GPS coordinates from a batch of photos and list them, then
 * export to Excel. First tries each photo's EXIF geotag (exact); for photos with
 * none — e.g. GPS-Map-Camera shots whose coordinates are only stamped on the
 * image — it OCRs the overlay stamp and parses the latitude/longitude out of it,
 * flagging shaky reads as "check" so the user can verify. Everything runs from
 * the app (EXIF in the renderer, OCR in the main process); nothing is uploaded.
 */
export default function GpsPhotosPage({ onBack }: Props) {
  const [rows, setRows] = useState<PhotoRow[]>([])
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  // Object URLs for the batch currently being read, shown as a Cover Flow while
  // OCR/EXIF runs. Revoked by the effect below when the batch is cleared/unmounts.
  const [procPreviews, setProcPreviews] = useState<{ name: string; url: string; ratio: number }[]>([])
  const fileInputRef = useRef<HTMLInputElement>(null)
  const idRef = useRef(0)

  useEffect(() => () => procPreviews.forEach((p) => URL.revokeObjectURL(p.url)), [procPreviews])

  // Counts (and the export) look only at rows already read — placeholders still
  // being processed are excluded so the tallies never flicker mid-batch.
  const done = rows.filter((r) => !r.pending)
  const withGps = done.filter((r) => r.lat != null && r.lon != null).length
  const needCheck = done.filter((r) => r.confidence === 'low').length
  const withoutGps = done.filter((r) => r.source === 'none').length

  async function addFiles(files: File[]) {
    const images = files.filter((f) => /\.(jpe?g|png|tiff?|heic|heif|webp|dng)$/i.test(f.name) || f.type.startsWith('image/'))
    if (images.length === 0) return
    setError(null)
    setSaved(null)
    setBusy(true)
    setProgress({ done: 0, total: images.length })
    setProcPreviews(
      await Promise.all(
        images.map(async (f) => {
          const url = URL.createObjectURL(f)
          return { name: f.name, url, ratio: await imageRatio(url) }
        })
      )
    )
    // Insert a placeholder row for every photo up front, so the table appears
    // immediately and each row is filled in the moment its photo is read — the
    // reading (cover flow) and the writing (table) run side by side.
    const ids = images.map(() => `g${idRef.current++}`)
    setRows((cur) => [...cur, ...images.map((f, i) => ({ id: ids[i], name: f.name, source: 'none' as const, confidence: 'none' as const, pending: true }))])
    try {
      // A small pool: EXIF photos finish instantly; overlay-OCR photos each fan
      // their preprocessing passes out to the OCR child pool in the main process
      // (electron/ocr.ts), so a handful of photos in flight already saturates it.
      const POOL = 4
      let doneCount = 0
      let next = 0
      async function worker() {
        while (next < images.length) {
          const i = next++
          const data = await processFile(images[i])
          // Replace this photo's placeholder in place — the row goes from
          // "Reading…" to its coordinates without waiting for the rest.
          setRows((cur) => cur.map((r) => (r.id === ids[i] ? { ...data, id: ids[i] } : r)))
          doneCount++
          // Update on every photo — OCR reads are slow, so a stale bar looks stuck.
          setProgress({ done: doneCount, total: images.length })
        }
      }
      await Promise.all(Array.from({ length: Math.min(POOL, images.length) }, worker))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
      // Clear any placeholders left hanging so no row stays stuck on "Reading…".
      setRows((cur) => cur.map((r) => (r.pending && ids.includes(r.id) ? { ...r, pending: false } : r)))
    } finally {
      setBusy(false)
      setProgress(null)
      setProcPreviews([]) // the effect revokes the batch's object URLs
    }
  }

  function onPick(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files ? Array.from(e.target.files) : []
    e.target.value = ''
    if (files.length) void addFiles(files)
  }

  function clearAll() {
    setRows([])
    setError(null)
    setSaved(null)
  }

  async function exportExcel() {
    if (rows.length === 0 || busy) return
    setError(null)
    setSaved(null)
    // Export layout: S.No, Address, Latitude, Longitude, Image name. Address is
    // filled from any place/pincode/state the overlay yielded (blank otherwise —
    // it's often absent). Raw OCR text trails as a verification helper.
    const headers = ['S.No', 'Address', 'Latitude', 'Longitude', 'Image name', 'Raw OCR text']
    const table: ExcelTable = {
      id: 'gps-photos',
      name: 'GPS Coordinates',
      path: '',
      headers,
      rows: rows.map((r, i) => ({
        'S.No': String(i + 1),
        Address: r.address ?? [r.place, r.pincode, r.state].filter(Boolean).join(', '),
        Latitude: dec(r.lat),
        Longitude: dec(r.lon),
        'Image name': r.name,
        'Raw OCR text': r.source === 'OCR' ? r.raw ?? '' : ''
      }))
    }
    try {
      const file = await api.exportTable(table, 'GPS Coordinates')
      if (file) setSaved(file)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  const hasRows = rows.length > 0

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconImage /> <span className="pdf-ws-title-label">Tool:</span> Extract Latitude / Longitude from GPS Photos
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={busy}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <input ref={fileInputRef} type="file" accept="image/*,.heic,.heif,.dng" multiple hidden onChange={onPick} />

      <div className="card pdf-workspace">
        <div className="gps-toolbar">
          <button className="primary" onClick={() => fileInputRef.current?.click()} disabled={busy}>
            <IconFolder /> {busy ? 'Reading…' : hasRows ? 'Add more photos' : 'Upload photos'}
          </button>
          <button className="primary" onClick={exportExcel} disabled={!hasRows || busy}>
            <IconTable /> Export to Excel
          </button>
          {hasRows && (
            <button className="pdf-ws-clearbtn" onClick={clearAll} disabled={busy}>
              <IconTrash /> Clear
            </button>
          )}
          {hasRows && (
            <span className="gps-counts">
              <strong>{rows.length}</strong> photo{rows.length === 1 ? '' : 's'} · <strong>{withGps}</strong> with GPS
              {needCheck > 0 && (
                <>
                  {' · '}
                  <span className="gps-warn">{needCheck} to check</span>
                </>
              )}
              {withoutGps > 0 && ` · ${withoutGps} not found`}
            </span>
          )}
        </div>

        {busy && procPreviews.length > 0 && (() => {
          // A windowed coverflow: only the photo being read (centre) plus a few on
          // each side are shown, so cards stay large and clearly separated instead
          // of the whole batch cramming onto the screen. Indices wrap, so the card
          // to the LEFT of the first photo is the last photo (a continuous loop) —
          // and because the window (W each side) is smaller than the batch, the
          // loop's seam stays off-screen and no card ever jumps across.
          const center = Math.min(progress ? progress.done : 0, procPreviews.length - 1)
          const N = procPreviews.length
          const W = Math.max(0, Math.min(3, Math.floor(N / 2) - 1)) // cards each side
          const STEP = 30 // degrees between neighbouring cards
          const RADIUS = 360 // px toward the viewer (front card sits largest)
          const MAXW = 380
          const MAXH = 300
          const cards: { off: number; idx: number; p: { name: string; url: string; ratio: number } }[] = []
          for (let off = -W; off <= W; off++) {
            const idx = ((center + off) % N + N) % N
            cards.push({ off, idx, p: procPreviews[idx] })
          }
          return (
            <div className="gps-coverflow">
              <div className="gps-coverflow-stage">
                {cards.map(({ off, idx, p }) => {
                  const abs = Math.abs(off)
                  // Size the card to the photo's shape (fit inside MAXW×MAXH), so
                  // the image fills it with no letterbox/frame behind it.
                  let w = MAXW
                  let h = MAXW / (p.ratio || 1)
                  if (h > MAXH) {
                    h = MAXH
                    w = MAXH * (p.ratio || 1)
                  }
                  return (
                    <div
                      key={idx}
                      className={`gps-cf-card${off === 0 ? ' active' : ''}`}
                      style={{
                        width: `${Math.round(w)}px`,
                        height: `${Math.round(h)}px`,
                        marginLeft: `${-Math.round(w) / 2}px`,
                        marginTop: `${-Math.round(h) / 2}px`,
                        transform: `rotateY(${off * STEP}deg) translateZ(${RADIUS}px)`,
                        zIndex: 100 - abs,
                        opacity: abs >= 3 ? 0.55 : 1
                      }}
                    >
                      <img src={p.url} alt={p.name} />
                    </div>
                  )
                })}
              </div>
              <div className="gps-coverflow-name" title={procPreviews[center]?.name}>
                Reading {procPreviews[center]?.name}
              </div>
            </div>
          )
        })()}

        {busy && progress && (
          <div className="gps-progress">
            <span className="gps-progress-label">
              Reading {progress.done} / {progress.total} photos (
              {progress.total ? Math.round((progress.done / progress.total) * 100) : 0}%) — overlay photos are OCR'd, which
              is slower
            </span>
            <span className="gps-progress-track" aria-hidden>
              <span
                className="gps-progress-bar"
                style={{ width: `${progress.total ? (progress.done / progress.total) * 100 : 0}%` }}
              />
            </span>
          </div>
        )}

        {error && (
          <div className="notice error tool-outcome">
            <IconWarn /> {error}
          </div>
        )}
        {saved && (
          <div className="notice ok tool-result tool-outcome">
            <span>
              <IconTable /> Saved Excel: {saved}
            </span>
            <button className="ghost" onClick={() => api.openPath(saved)}>
              <IconOpen /> Open
            </button>
          </div>
        )}

        {!hasRows ? (
          <button className="pdf-ws-dropzone" onClick={() => fileInputRef.current?.click()} disabled={busy}>
            <IconImage />
            <span className="pdf-ws-dropzone-title">{busy ? 'Reading…' : 'Upload photos'}</span>
            <span className="pdf-ws-dropzone-sub">
              GPS is read from each photo's metadata; photos with the coordinates only stamped on them are OCR'd
            </span>
          </button>
        ) : (
          <>
            <p className="gps-hint">
              <strong>Exact</strong> = read from the photo's EXIF geotag (trustworthy). <strong>OCR</strong> = read off the
              stamped overlay by consensus — usually right, but verify anything critical, as OCR can occasionally miss a
              digit. <span className="gps-warn">Check</span> = a shaky read to verify against the photo. The raw OCR text is
              in the exported sheet for checking.
            </p>
            <div className="gps-table-wrap">
              <table className="gps-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>File name</th>
                    <th>Latitude</th>
                    <th>Longitude</th>
                    <th>Altitude</th>
                    <th>Place</th>
                    <th>Date / Time</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => {
                    const has = r.lat != null && r.lon != null
                    if (r.pending) {
                      return (
                        <tr key={r.id} className="gps-row-pending">
                          <td>{i + 1}</td>
                          <td className="gps-name" title={r.name}>
                            {r.name}
                          </td>
                          <td colSpan={5}>
                            <span className="gps-reading">
                              <span className="gps-reading-dot" aria-hidden /> Reading…
                            </span>
                          </td>
                        </tr>
                      )
                    }
                    return (
                      <tr key={r.id} className={r.confidence === 'low' ? 'gps-row-check' : r.source === 'none' ? 'gps-row-missing' : ''}>
                        <td>{i + 1}</td>
                        <td className="gps-name" title={r.source === 'OCR' ? r.raw : r.name}>
                          {r.name}
                        </td>
                        <td>{has ? dec(r.lat) : <span className="gps-warn">—</span>}</td>
                        <td>{has ? dec(r.lon) : ''}</td>
                        <td>{r.altitude ?? ''}</td>
                        <td>{r.address ?? r.place ?? ''}</td>
                        <td>{r.taken ?? ''}</td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </>
  )
}
