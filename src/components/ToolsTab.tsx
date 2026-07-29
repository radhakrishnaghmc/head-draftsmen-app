import { useState } from 'react'
import { api } from '../ipc'
import { IconTable, IconFolder, IconWarn, IconOpen, IconImage } from './Icons'
import UploadPhotosTab from './UploadPhotosTab'
import type { ExcelTable } from '@core/types'

interface SplitState {
  busy: boolean
  result: { dir: string; files: string[] } | null
  error: string | null
}

interface Props {
  /** The Works List database — passed through to the photo-estimate tool for ECV write-back and Circle/Agency lookups. */
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
}

/**
 * Utility tools that sit outside the main tender/estimate workflow, shown as a
 * grid of tiles (the app's frosted-glass doc-tile look) plus full-width tool
 * panels below. Today: the Excel Sheet Separator, and reading an estimate from
 * photos / a scanned PDF. New tools slot in as additional tiles or panels.
 */
export default function ToolsTab({ tables, onChange }: Props) {
  const [split, setSplit] = useState<SplitState>({ busy: false, result: null, error: null })
  // The photo-estimate tool opens as a panel below the tiles when its tile is
  // clicked — the same tile look as the Excel Separator, but it reveals a
  // full workflow (upload → OCR → review → downloads) rather than a one-shot dialog.
  const [showPhotos, setShowPhotos] = useState(false)

  async function runSeparator() {
    if (split.busy) return
    setSplit({ busy: true, result: null, error: null })
    try {
      // null when the user cancels either dialog — leave the tile idle.
      const result = await api.splitExcelSheets()
      setSplit({ busy: false, result, error: null })
    } catch (e) {
      setSplit({ busy: false, result: null, error: e instanceof Error ? e.message : String(e) })
    }
  }

  return (
    <div className="card">
      <div className="doc-tile-grid">
        <button
          className="doc-tile-card tone-teal tool-card"
          onClick={runSeparator}
          disabled={split.busy}
          title="Split a multi-sheet workbook into one file per sheet"
        >
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Excel Sheet Separator</span>
          <span className="doc-tile-card-meta">
            {split.busy ? 'Working…' : 'One file per sheet, named after each tab'}
          </span>
        </button>

        <button
          className={`doc-tile-card tone-amber tool-card ${showPhotos ? 'on' : ''}`}
          onClick={() => setShowPhotos((v) => !v)}
          title="Read an estimate from photos or a scanned PDF"
          aria-expanded={showPhotos}
        >
          <span className="tool-card-ic">
            <IconImage />
          </span>
          <span className="doc-tile-card-name">Estimate from Photos / PDF</span>
          <span className="doc-tile-card-meta">
            {showPhotos ? 'Open below — click to hide' : 'Photos / scanned PDF → BOQ, Schedule A, Deviation, Material'}
          </span>
        </button>
      </div>

      {split.error && (
        <div className="notice error">
          <IconWarn />
          {split.error}
        </div>
      )}

      {split.result && (
        <div className="notice ok tool-result">
          <span>
            <IconFolder /> Saved {split.result.files.length} sheet
            {split.result.files.length === 1 ? '' : 's'} to {split.result.dir}
          </span>
          <button className="ghost" onClick={() => api.openPath(split.result!.dir)}>
            <IconOpen /> Open folder
          </button>
        </div>
      )}

      {showPhotos && (
        <div className="workspace-section">
          <div className="workspace-section-hint">
            Read an estimate from photos or a scanned PDF (OCR'd, then reviewed before downloading) — its BOQ, Schedule
            A, Deviation, and Material Quantity.
          </div>
          <UploadPhotosTab tables={tables} onChange={onChange} />
        </div>
      )}
    </div>
  )
}
