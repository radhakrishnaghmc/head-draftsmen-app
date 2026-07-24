import { useRef, useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { IconFolder, IconImage, IconTrash, IconDownload, IconWarn } from './Icons'
import ExcelInline from './ExcelInline'
import type { ExcelTable } from '@core/types'
import type { EstimateWorkItem } from '@core/estimateExtract'

interface Photo {
  id: string
  name: string
  dataUrl: string
}

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

/**
 * Upload photos of a paper estimate (in page order — drag to reorder) and
 * convert them into an editable spreadsheet: each photo is read with local
 * OCR (Tesseract, fully offline — see electron/ocr.ts) and reconstructed
 * into a table (electron/main.ts's ocrEstimatePhotos + core/ocrTableReconstruct.ts),
 * then run through the same estimate-column matching used for uploaded
 * Excel estimates. OCR on a real photographed page is inherently
 * imperfect — the result is always shown for review/editing before export,
 * never saved directly.
 */
export default function UploadPhotosTab() {
  const [photos, setPhotos] = useState<Photo[]>([])
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const [converting, setConverting] = useState(false)
  const [convertError, setConvertError] = useState<string | null>(null)
  const [resultTable, setResultTable] = useState<ExcelTable | null>(null)
  const [aiAssisted, setAiAssisted] = useState<string[]>([])
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState<string | null>(null)

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
    setSaved(null)
    try {
      const sheet = await api.ocrEstimatePhotos(photos.map((p) => p.dataUrl))
      const headerRow = guessHeaderRow(sheet.grid)
      const { items, aiAssisted: assisted } = await extractEstimateItemsWithAi(sheet.grid, headerRow)
      if (items.length === 0) {
        throw new Error(
          'No work items with a quantity, rate, and unit were recognized in these photos. Clearer, well-lit, straight-on photos read better — you can also fix up rows manually below once something is recognized.'
        )
      }
      setAiAssisted(assisted)
      setResultTable(itemsToTable(items))
    } catch (e) {
      setConvertError(e instanceof Error ? e.message : String(e))
    } finally {
      setConverting(false)
    }
  }

  async function download() {
    if (!resultTable) return
    setSaving(true)
    setSaved(null)
    try {
      const base = photos[0] ? stripExt(photos[0].name) : 'Estimate'
      const path = await api.exportTable(resultTable, `${base} Estimate`)
      if (path) setSaved(path)
    } catch (e) {
      setConvertError(e instanceof Error ? e.message : String(e))
    } finally {
      setSaving(false)
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

      {resultTable && (
        <>
          <div className="notice ok" style={{ marginTop: 14 }}>
            Read {resultTable.rows.length} row{resultTable.rows.length === 1 ? '' : 's'} from {photos.length} photo
            {photos.length === 1 ? '' : 's'} — OCR is never perfect, please check the rows below before downloading.
            {aiAssisted.length > 0 &&
              ` ${aiAssisted.join(', ')} column${aiAssisted.length === 1 ? '' : 's'} matched by AI — please double-check.`}
          </div>
          <ExcelInline table={resultTable} onChange={setResultTable} />
          <div className="boq-actions" style={{ marginTop: 14 }}>
            <button className="primary" onClick={download} disabled={saving}>
              <IconDownload /> {saving ? 'Saving…' : 'Download Excel'}
            </button>
            {saved && <span className="hint">Saved to {saved}</span>}
          </div>
        </>
      )}
    </div>
  )
}
