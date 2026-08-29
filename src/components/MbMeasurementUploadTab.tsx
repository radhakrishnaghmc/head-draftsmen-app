import { useEffect, useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfPagesToDataUrls } from '../pdfToImages'
import { MB_MEASUREMENT_HEADERS } from '@core/mbMeasurementExtract'
import ExcelInline from './ExcelInline'
import type { ExcelTable } from '@core/types'
import type { MbMeasurementProgress } from '../../electron/ipc-contract'
import { IconFolder, IconDownload, IconWarn, IconTable } from './Icons'

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

function sheetGridToTable(id: string, name: string, grid: string[][]): ExcelTable {
  const headers = grid[0] ?? [...MB_MEASUREMENT_HEADERS]
  const rows = grid.slice(1).map((cells) => {
    const row: Record<string, string> = {}
    headers.forEach((h, i) => {
      row[h] = cells[i] ?? ''
    })
    return row
  })
  return { id, name, path: '', headers, rows }
}

/**
 * Upload photos or a scanned PDF of MB "L.F. No. 83" measurement sheet pages
 * and read them into an editable grid (Date / Description / No / L1-3 /
 * B1-3 / D1-3 / Contents), best-effort via local OCR — always review before
 * downloading, especially handwritten figures.
 */
export default function MbMeasurementUploadTab() {
  const [table, setTable] = useState<ExcelTable | null>(null)
  const [fileNames, setFileNames] = useState<string[]>([])
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState<MbMeasurementProgress | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => api.onMbMeasurementProgress(setProgress), [])

  async function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    setBusy(true)
    setProgress(null)
    setError(null)
    setSaved(null)
    try {
      const files = Array.from(fileList)
      const dataUrls: string[] = []
      for (const file of files) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          dataUrls.push(...(await pdfPagesToDataUrls(file)))
        } else {
          dataUrls.push(await readAsDataUrl(file))
        }
      }
      if (dataUrls.length === 0) throw new Error('No pages found in the selected file(s).')
      const grid = await api.ocrMbMeasurementSheet(dataUrls)
      setTable(sheetGridToTable(`mb-measurement-${Date.now()}`, 'MB Measurement Sheet', grid.grid))
      setFileNames(files.map((f) => f.name))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      setProgress(null)
    }
  }

  async function download() {
    if (!table) return
    setError(null)
    try {
      const path = await api.exportTable(table, 'MB Measurement Sheet')
      setSaved(path ? `Saved to ${path}` : 'Cancelled.')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  const pct = progress && progress.cellsTotal > 0 ? Math.round((progress.cellsDone / progress.cellsTotal) * 100) : null
  const label = !busy
    ? table
      ? 'Upload another MB sheet'
      : 'Upload MB measurement sheet'
    : !progress
      ? 'Reading…'
      : progress.totalPages > 1
        ? `Reading page ${progress.page}/${progress.totalPages}${pct !== null ? ` (${pct}%)` : ''}…`
        : pct !== null
          ? `Reading… ${pct}%`
          : 'Reading…'

  return (
    <div className="card">
      <div className="empty empty--tight">
        <IconTable />
        <div className="boq-actions">
          <button
            className="primary upload-btn scan-progress-btn"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy}
          >
            {busy && pct !== null && <span className="scan-progress-fill" style={{ width: `${pct}%` }} />}
            <span className="scan-progress-label">
              <IconFolder /> {label}
            </span>
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*,application/pdf,.pdf"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              void handleFiles(e.target.files)
              e.target.value = ''
            }}
          />
          {table && (
            <button className="primary" onClick={download}>
              <IconDownload /> Download as Excel
            </button>
          )}
        </div>
        {fileNames.length > 0 && <p className="estimate-hint">Read from {fileNames.join(', ')} — review before downloading, especially handwritten figures.</p>}
        {error && (
          <div className="notice error">
            <IconWarn /> {error}
          </div>
        )}
        {saved && <div className="notice ok">{saved}</div>}
      </div>

      {table && (
        <div className="estimate-preview-scroll">
          <ExcelInline table={table} onChange={setTable} />
        </div>
      )}
    </div>
  )
}
