import { useRef, useState } from 'react'
import { api } from '../ipc'
import { extractWorkName, extractEstimateItemsFromLines } from '@core/estimateExtract'
import { extractEstimateAmountLakhs } from '@core/deviation'
import { findWorksRowByName, metaFromWorksRow } from '@core/scheduleA'
import { applyEcvFromBoq } from '@core/worksAmounts'
import { buildBoqFromEstimate, computeEcvFromItems, boqToScheduleA } from '../boqTransform'
import { IconFolder, IconImage, IconTrash, IconDownload, IconWarn } from './Icons'
import ExcelInline from './ExcelInline'
import type { ExcelTable } from '@core/types'
import type { EstimateWorkItem } from '@core/estimateExtract'
import type { DeviationItem } from '../../electron/ipc-contract'

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

type ActionKey = 'excel' | 'boq' | 'scheduleA' | 'deviation'

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

/** Sl No / Description / Unit / Quantity / Rate / Amount — a plain, editable review table, not the app's specialized BOQ/Schedule A template. */
function itemsToTable(items: EstimateWorkItem[]): ExcelTable {
  const headers = ['Sl No', 'Description', 'Unit', 'Quantity', 'Rate', 'Amount']
  const rows = items.map((it, i) => {
    const qty = Number(it.quantity.replace(/,/g, ''))
    const rate = Number(it.rate.replace(/,/g, ''))
    const amount = Number.isFinite(qty) && Number.isFinite(rate) && qty && rate ? (qty * rate).toFixed(2) : ''
    return {
      'Sl No': String(i + 1),
      Description: it.description,
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
      unit: r['Unit'] ?? ''
    }))
}

/**
 * Upload photos of a paper estimate (in page order — drag to reorder) and
 * convert them into an editable spreadsheet: each photo is read with local
 * OCR (PaddleOCR via @gutenye/ocr-node, fully offline — see electron/ocr.ts),
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
 * BOQ / Schedule A / Deviation Statement generators — the photos are just
 * an alternative way to get an estimate's items into the app, not a
 * separate feature with its own output logic.
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

  const [actionBusy, setActionBusy] = useState<ActionKey | null>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)

  const worksTable = tables[0] ?? null

  async function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    const files = Array.from(fileList).filter((f) => f.type.startsWith('image/'))
    const added = await Promise.all(
      files.map(async (file) => ({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) }))
    )
    setPhotos((prev) => [...prev, ...added])
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
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportTable(resultTable, `${base} Estimate`)
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  // Matches a work name against the Works List by exact name first, falling
  // back to the local embedding model when that fails (an estimate's
  // title-block wording very often differs slightly from the Works List's
  // own entry) — shared by the BOQ/Schedule A/Deviation actions below.
  async function matchWorksRow(name: string, table: ExcelTable) {
    const exact = findWorksRowByName(table, name)
    if (exact) return exact
    const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
    if (!nameHeader) return undefined
    try {
      const rowNames = table.rows.map((r) => r[nameHeader] ?? '')
      const vectors = await api.embedTexts([name, ...rowNames])
      return findWorksRowByName(table, name, { workNameVector: vectors[0], rowNameVectors: vectors.slice(1) })
    } catch {
      return undefined
    }
  }

  async function downloadBoq() {
    if (!resultTable) return
    setActionBusy('boq')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      const boq = buildBoqFromEstimate(items)
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportBoq(boq, `${base} BOQ`, workName)
      if (path) setActionSaved(path)

      if (workName && worksTable) {
        const ecvRupees = computeEcvFromItems(items)
        let result = applyEcvFromBoq(worksTable, workName, ecvRupees)
        if (!result.matched) {
          const nameHeader = worksTable.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
          if (nameHeader) {
            try {
              const rowNames = worksTable.rows.map((r) => r[nameHeader] ?? '')
              const vectors = await api.embedTexts([workName, ...rowNames])
              result = applyEcvFromBoq(worksTable, workName, ecvRupees, {
                workNameVector: vectors[0],
                rowNameVectors: vectors.slice(1)
              })
            } catch {
              // Embeddings unavailable — leave the Works List untouched.
            }
          }
        }
        if (result.matched) onChange(result.table)
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setActionBusy(null)
    }
  }

  async function downloadScheduleA() {
    if (!resultTable) return
    setActionBusy('scheduleA')
    setActionError(null)
    setActionSaved(null)
    try {
      const items = tableToItems(resultTable)
      // The app's own BOQ column names already match Schedule A's expected
      // patterns exactly, so no embedding fallback is needed for this step.
      const scheduleARows = boqToScheduleA(buildBoqFromEstimate(items))

      const meta = workName && worksTable ? metaFromWorksRow((await matchWorksRow(workName, worksTable))?.row ?? {}) : undefined
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportScheduleA(scheduleARows, `${base} Schedule A`, meta)
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
      const deviationItems: DeviationItem[] = items.map((it) => ({
        description: it.description,
        unit: it.unit,
        quantity: it.quantity,
        rate: it.rate
      }))
      const estimateAmountLakhs = extractEstimateAmountLakhs(ocrGrid ?? [], (ocrGrid ?? []).length, items)
      const circle = workName && worksTable ? (await matchWorksRow(workName, worksTable))?.row['Circle'] : undefined

      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportDeviation(
        deviationItems,
        { circle, nameOfWork: workName, agencyName: agencyName.trim(), estimateAmountLakhs },
        `${base} Deviation`
      )
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
            ? 'Upload photos of a paper estimate to convert them into a spreadsheet.'
            : `${photos.length} photo${photos.length === 1 ? '' : 's'} added — drag to reorder into page order.`}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={() => fileInputRef.current?.click()}>
            <IconFolder /> Upload Photos
          </button>
          {photos.length > 0 && (
            <button className="primary" onClick={convertToEstimate} disabled={converting}>
              <IconImage /> {converting ? 'Reading photos…' : 'Convert to Estimate'}
            </button>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
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
              <IconDownload /> {actionBusy === 'excel' ? 'Saving…' : 'Download Excel'}
            </button>
            <button className="primary" onClick={downloadBoq} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'boq' ? 'Saving…' : 'Download BOQ'}
            </button>
            <button className="primary" onClick={downloadScheduleA} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'scheduleA' ? 'Saving…' : 'Download Schedule A'}
            </button>
          </div>
          <div className="boq-actions" style={{ marginTop: 8 }}>
            <input
              className="editor-name"
              style={{ maxWidth: 220 }}
              placeholder="Name of the Agency"
              value={agencyName}
              onChange={(e) => setAgencyName(e.target.value)}
            />
            <button className="primary" onClick={generateDeviation} disabled={actionBusy !== null}>
              <IconDownload /> {actionBusy === 'deviation' ? 'Saving…' : 'Generate Deviation Statement'}
            </button>
          </div>
          {actionSaved && <p className="hint">Saved to {actionSaved}</p>}
        </>
      )}
    </div>
  )
}
