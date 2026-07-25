import { useRef, useState } from 'react'
import { api } from '../ipc'
import { extractWorkName, extractEstimateItemsFromLines } from '@core/estimateExtract'
import { extractEstimateAmountLakhs } from '@core/deviation'
import { computeEcvFromItems } from '../boqTransform'
import { formatRupees } from '@core/worksAmounts'
import { pdfPagesToDataUrls } from '../pdfToImages'
import {
  matchWorksRow,
  saveEcvToWorksList,
  downloadBoqFromItems,
  downloadScheduleAFromItems,
  downloadDeviationFromItems,
  downloadMaterialFromItems
} from '../estimateDownloads'
import { IconFolder, IconImage, IconTrash, IconDownload, IconWarn } from './Icons'
import ExcelInline from './ExcelInline'
import type { ExcelTable } from '@core/types'
import type { EstimateWorkItem } from '@core/estimateExtract'

interface Photo {
  id: string
  name: string
  dataUrl: string
}

interface Props {
  /** The Works List database — tables[0], by the app's own convention — used for ECV write-back and Circle/Agency lookups. */
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
}

type ActionKey = 'excel' | 'boq' | 'scheduleA' | 'deviation' | 'material'

function nextId(): string {
  return `photo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

/**
 * Sl No / Description / No's / L / B / D / Unit / Quantity / Rate / Amount —
 * a plain, editable review table, not the app's specialized BOQ/Schedule A
 * template. No's/L/B/D are included editable here (rather than only shown on
 * the final workbook) because OCR can only reconstruct them when the source
 * photo's dimension row happens to resolve unambiguously — this is the
 * user's chance to fill in or correct them by eye before export.
 */
function itemsToTable(items: EstimateWorkItem[]): ExcelTable {
  const headers = ['Sl No', 'Description', "No's", 'L', 'B', 'D', 'Unit', 'Quantity', 'Rate', 'Amount']
  const rows = items.map((it, i) => {
    const qty = Number(it.quantity.replace(/,/g, ''))
    const rate = Number(it.rate.replace(/,/g, ''))
    const amount = Number.isFinite(qty) && Number.isFinite(rate) && qty && rate ? (qty * rate).toFixed(2) : ''
    return {
      'Sl No': String(i + 1),
      Description: it.description,
      "No's": it.nos ?? '',
      L: it.l ?? '',
      B: it.b ?? '',
      D: it.d ?? '',
      Unit: it.unit,
      Quantity: it.quantity,
      Rate: it.rate,
      Amount: amount
    }
  })
  return { id: `photo-estimate-${Date.now()}`, name: 'Estimate from photos', path: '', headers, rows }
}

/** Reverses itemsToTable — reads back whatever the user has (possibly) edited in the review grid. */
function tableToItems(table: ExcelTable): EstimateWorkItem[] {
  return table.rows
    .filter((r) => (r['Description'] ?? '').trim() && (r['Quantity'] ?? '').trim())
    .map((r) => ({
      description: r['Description'] ?? '',
      quantity: r['Quantity'] ?? '',
      rate: r['Rate'] ?? '',
      unit: r['Unit'] ?? '',
      nos: (r["No's"] ?? '').trim() || undefined,
      l: (r['L'] ?? '').trim() || undefined,
      b: (r['B'] ?? '').trim() || undefined,
      d: (r['D'] ?? '').trim() || undefined
    }))
}

/**
 * Upload photos of a paper estimate, or a PDF of scanned pages (each PDF
 * page is rendered to an image client-side via pdfjs-dist — see
 * ../pdfToImages.ts — and dropped into the same photo list), in page order
 * — drag to reorder — and convert them into an editable spreadsheet: each
 * photo is read with local OCR (PaddleOCR via @gutenye/ocr-node, fully
 * offline — see electron/ocr.ts),
 * whose line detector returns each printed line as one clean unit of text in
 * reading order (electron/main.ts's ocrEstimatePhotos). Those lines are
 * parsed directly by core/estimateExtract.ts's extractEstimateItemsFromLines
 * — no column-position reconstruction involved, since this document's
 * merged-description-cell layout defeats that almost every time regardless
 * of photo quality or OCR engine. OCR on a real photographed page is
 * inherently imperfect — the result is always shown for review/editing
 * before export, never saved directly.
 *
 * Once reviewed, the same items feed straight into the app's existing
 * BOQ / Schedule A / Deviation Statement / Material Quantity generators (see
 * ../estimateDownloads.ts) — the photos are just an alternative way to get
 * an estimate's items into the app, not a separate feature with its own
 * output logic.
 */
export default function UploadPhotosTab({ tables, onChange }: Props) {
  const [photos, setPhotos] = useState<Photo[]>([])
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const [converting, setConverting] = useState(false)
  const [convertError, setConvertError] = useState<string | null>(null)
  const [resultTable, setResultTable] = useState<ExcelTable | null>(null)
  // The raw OCR'd lines, kept alongside the editable preview so BOQ/Schedule
  // A/Deviation generation (below) can reuse them exactly like the
  // Excel-upload versions of these features do, and so a failed conversion
  // still shows what was actually read (for troubleshooting a bad photo).
  const [ocrGrid, setOcrGrid] = useState<string[][] | null>(null)
  const [workName, setWorkName] = useState<string | undefined>()
  const [agencyName, setAgencyName] = useState('')
  const [departmentName, setDepartmentName] = useState('')
  const [district, setDistrict] = useState('')

  const [actionBusy, setActionBusy] = useState<ActionKey | null>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [ecvConfirm, setEcvConfirm] = useState<{ workName: string; ecvRupees: number; matchedName: string } | null>(
    null
  )

  const worksTable = tables[0] ?? null

  async function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    const files = Array.from(fileList)
    setConvertError(null)
    try {
      const added: Photo[] = []
      for (const file of files) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // A PDF here is a scanned/photographed multi-page estimate saved as
          // one file — split it into one page-image per photo so it feeds
          // the exact same per-photo OCR pipeline as a direct photo upload.
          const pages = await pdfPagesToDataUrls(file)
          const base = stripExt(file.name)
          pages.forEach((dataUrl, i) => added.push({ id: nextId(), name: `${base} (page ${i + 1})`, dataUrl }))
        } else if (file.type.startsWith('image/')) {
          added.push({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) })
        }
      }
      setPhotos((prev) => [...prev, ...added])
    } catch (e) {
      setConvertError(e instanceof Error ? e.message : String(e))
    }
  }

  function removePhoto(id: string) {
    setPhotos((prev) => prev.filter((p) => p.id !== id))
  }

  function handleDragStart(e: React.DragEvent, index: number) {
    setDragIndex(index)
    e.dataTransfer.effectAllowed = 'move'
  }

  function handleDragOver(e: React.DragEvent, index: number) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    if (index !== overIndex) setOverIndex(index)
  }

  function handleDrop(e: React.DragEvent, index: number) {
    e.preventDefault()
    if (dragIndex !== null && dragIndex !== index) {
      setPhotos((prev) => {
        const next = [...prev]
        const [moved] = next.splice(dragIndex, 1)
        next.splice(index, 0, moved)
        return next
      })
    }
    setDragIndex(null)
    setOverIndex(null)
  }

  function handleDragEnd() {
    setDragIndex(null)
    setOverIndex(null)
  }

  async function convertToEstimate() {
    if (photos.length === 0) return
    setConverting(true)
    setConvertError(null)
    setResultTable(null)
    setActionError(null)
    setActionSaved(null)
    setOcrGrid(null)
    try {
      const sheet = await api.ocrEstimatePhotos(photos.map((p) => p.dataUrl))
      setOcrGrid(sheet.grid)
      const lines = sheet.grid.map((r) => r[0] ?? '')
      const items = extractEstimateItemsFromLines(lines)
      if (items.length === 0) {
        throw new Error(
          'No work items with a quantity, rate, and unit were recognized in these photos. Clearer, well-lit, straight-on photos read better — see the recognized text below.'
        )
      }
      setResultTable(itemsToTable(items))
      // No real "header row" boundary in a line-based grid — search the
      // whole thing for the title block's "Name of Work"/"Estimate Amount"
      // labels, same as extractEstimateItemsFromLines finds its own anchor.
      setWorkName(extractWorkName(sheet.grid, sheet.grid.length))
    } catch (e) {
      setConvertError(e instanceof Error ? e.message : String(e))
    } finally {
      setConverting(false)
    }
  }

  async function download() {
    if (!resultTable) return
    setActionBusy('excel')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const match = workName && worksTable ? await matchWorksRow(workName, worksTable) : undefined
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportDetailedEstimate(
        items,
        { zone: match?.row['Zone'], circle: match?.row['Circle'], workName },
        `${base} Estimate`
      )
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  async function downloadBoq() {
    if (!resultTable) return
    setActionBusy('boq')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await downloadBoqFromItems(items, workName, base)
      if (path) setActionSaved(path)

      if (workName && worksTable) {
        const match = await matchWorksRow(workName, worksTable)
        if (match) {
          setEcvConfirm({
            workName,
            ecvRupees: computeEcvFromItems(items),
            matchedName: match.row['Name of the work'] || workName
          })
        }
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  async function confirmEcvUpdate() {
    const pending = ecvConfirm
    setEcvConfirm(null)
    if (!pending || !worksTable) return
    const updated = await saveEcvToWorksList(pending.workName, pending.ecvRupees, worksTable)
    if (updated) onChange(updated)
  }

  async function downloadScheduleA() {
    if (!resultTable) return
    setActionBusy('scheduleA')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await downloadScheduleAFromItems(items, workName, worksTable, base)
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  async function generateDeviation() {
    if (!resultTable) return
    if (!agencyName.trim()) {
      setActionError("Enter the agency's name before generating the Deviation Statement.")
      return
    }
    setActionBusy('deviation')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const estimateAmountLakhs = extractEstimateAmountLakhs(ocrGrid ?? [], (ocrGrid ?? []).length, items)
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await downloadDeviationFromItems(items, workName, agencyName, estimateAmountLakhs, worksTable, base)
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  async function downloadMaterial() {
    if (!resultTable) return
    setActionBusy('material')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const ecvRupees = computeEcvFromItems(items)
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await downloadMaterialFromItems(items, workName, ecvRupees, departmentName, district, base)
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconImage />
        <p>
          {photos.length === 0
            ? 'Upload photos of a paper estimate (or a PDF of scanned pages) to convert them into a spreadsheet.'
            : `${photos.length} photo${photos.length === 1 ? '' : 's'} added — drag to reorder into page order.`}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={() => fileInputRef.current?.click()}>
            <IconFolder /> Upload Photos or PDF
          </button>
          {photos.length > 0 && (
            <button className="primary" onClick={convertToEstimate} disabled={converting}>
              <IconImage /> {converting ? 'Reading photos…' : 'Convert to Estimate'}
            </button>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*,application/pdf"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              void handleFiles(e.target.files)
              e.target.value = ''
            }}
          />
        </div>
      </div>

      {photos.length > 0 && (
        <div className="photo-tile-grid">
          {photos.map((photo, i) => (
            <div
              key={photo.id}
              className={[
                'photo-tile-card',
                dragIndex === i ? 'dragging' : '',
                overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
              ]
                .filter(Boolean)
                .join(' ')}
              draggable
              onDragStart={(e) => handleDragStart(e, i)}
              onDragOver={(e) => handleDragOver(e, i)}
              onDrop={(e) => handleDrop(e, i)}
              onDragEnd={handleDragEnd}
            >
              <span className="photo-tile-number">{i + 1}</span>
              <button className="photo-tile-remove" title="Remove" onClick={() => removePhoto(photo.id)}>
                <IconTrash />
              </button>
              <img className="photo-tile-img" src={photo.dataUrl} alt={photo.name} />
              <div className="photo-tile-name" title={photo.name}>
                {photo.name}
              </div>
            </div>
          ))}
        </div>
      )}

      {convertError && (
        <div className="notice error">
          <IconWarn />
          {convertError}
        </div>
      )}

      {convertError && ocrGrid && (
        <div className="ocr-grid-preview" style={{ marginTop: 14 }}>
          <table>
            <tbody>
              {ocrGrid.map((row, i) => (
                <tr key={i}>
                  <td className="row-num">{i}</td>
                  <td>{row[0]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {resultTable && (
        <>
          <div className="notice ok" style={{ marginTop: 14 }}>
            Read {resultTable.rows.length} row{resultTable.rows.length === 1 ? '' : 's'} from {photos.length} photo
            {photos.length === 1 ? '' : 's'} — OCR is never perfect, please check the rows below before generating
            anything.
          </div>
          <ExcelInline table={resultTable} onChange={setResultTable} />

          {actionError && (
            <div className="notice error" style={{ marginTop: 14 }}>
              <IconWarn />
              {actionError}
            </div>
          )}

          <div className="boq-actions" style={{ marginTop: 14, flexWrap: 'wrap' }}>
            <button className="primary" onClick={download} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'excel' ? 'Saving…' : 'Download Estimate'}
            </button>
            <button className="primary" onClick={downloadBoq} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'boq' ? 'Saving…' : 'Download BOQ'}
            </button>
            <button className="primary" onClick={downloadScheduleA} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'scheduleA' ? 'Saving…' : 'Download Schedule A'}
            </button>
          </div>
          <div className="boq-actions" style={{ marginTop: 8, flexWrap: 'wrap' }}>
            <input
              className="editor-name"
              style={{ maxWidth: 180 }}
              placeholder="Name of the Agency"
              value={agencyName}
              onChange={(e) => setAgencyName(e.target.value)}
            />
            <button className="primary" onClick={generateDeviation} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'deviation' ? 'Saving…' : 'Generate Deviation Statement'}
            </button>
            <input
              className="editor-name"
              style={{ maxWidth: 150 }}
              placeholder="Department Name"
              value={departmentName}
              onChange={(e) => setDepartmentName(e.target.value)}
            />
            <input
              className="editor-name"
              style={{ maxWidth: 110 }}
              placeholder="District"
              value={district}
              onChange={(e) => setDistrict(e.target.value)}
            />
            <button className="primary" onClick={downloadMaterial} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'material' ? 'Saving…' : 'Download Material Quantity'}
            </button>
          </div>
          {actionSaved && <p className="hint">Saved to {actionSaved}</p>}
        </>
      )}

      {ecvConfirm && (
        <div className="editor-overlay" onClick={() => setEcvConfirm(null)}>
          <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <h3>Update ECV in the Works List?</h3>
            <p className="confirm-warn">
              "{ecvConfirm.matchedName}" is already in the Works List. Update its ECV to{' '}
              <strong>{formatRupees(ecvConfirm.ecvRupees)}</strong>, computed from this BOQ?
            </p>
            <p className="confirm-hint">Every other field on that row stays untouched.</p>
            <div className="confirm-actions">
              <button className="ghost" onClick={() => setEcvConfirm(null)}>
                No, keep existing
              </button>
              <button className="primary" onClick={confirmEcvUpdate}>
                Update ECV
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
