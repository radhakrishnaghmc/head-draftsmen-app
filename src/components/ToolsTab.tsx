import { useState } from 'react'
import { api } from '../ipc'
import { IconTable, IconFolder, IconWarn, IconOpen, IconImage, IconClipboard } from './Icons'
import UploadPhotosTab from './UploadPhotosTab'
import WorkOrderAgreementTab from './WorkOrderAgreementTab'
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
  // Standalone document generators — Work Order, Agreement Bond and Schedule A
  // built ONLY from uploaded files, with no Works List link and no Zone/Circle
  // checks. All three live in one panel (one L-1 + Intimation upload drives the
  // Work Order and Agreement; an estimate upload drives the Schedule A), so the
  // three tiles are labelled entry points that reveal the same panel.
  const [showDocs, setShowDocs] = useState(false)
  // The Schedule A tile opens its own estimate-only panel (no L1/Intimation).
  const [showScheduleA, setShowScheduleA] = useState(false)

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

        {(['Work Order', 'Agreement Bond'] as const).map((name) => (
          <button
            key={name}
            className={`doc-tile-card tone-sky tool-card ${showDocs ? 'on' : ''}`}
            onClick={() => setShowDocs((v) => !v)}
            title="Generate from L1 + Intimation only — no Works List, no Zone/Circle check"
            aria-expanded={showDocs}
          >
            <span className="tool-card-ic">
              <IconClipboard />
            </span>
            <span className="doc-tile-card-name">{name}</span>
            <span className="doc-tile-card-meta">
              {showDocs ? 'Open below — click to hide' : 'From L1 + Intimation — any circle/zone'}
            </span>
          </button>
        ))}

        <button
          className={`doc-tile-card tone-sky tool-card ${showScheduleA ? 'on' : ''}`}
          onClick={() => setShowScheduleA((v) => !v)}
          title="Generate a Schedule A from an estimate / BOQ only — no Works List"
          aria-expanded={showScheduleA}
        >
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Schedule A</span>
          <span className="doc-tile-card-meta">
            {showScheduleA ? 'Open below — click to hide' : 'From an uploaded estimate / BOQ'}
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

      {showDocs && (
        <div className="workspace-section">
          <div className="workspace-section-hint">
            Standalone Work Order & Agreement Bond — built only from the uploaded <strong>L1 selection form</strong> and{' '}
            <strong>Online Intimation</strong>, for any circle/zone. Nothing is read from or written to the Works List,
            and no Zone/Circle check is applied.
          </div>
          <WorkOrderAgreementTab standalone tables={[]} onChange={() => {}} />
        </div>
      )}

      {showScheduleA && (
        <div className="workspace-section">
          <div className="workspace-section-hint">
            Standalone Schedule A — built only from an uploaded <strong>estimate / BOQ</strong>, for any circle/zone. No
            L1 or Intimation needed, and nothing is read from or written to the Works List.
          </div>
          <WorkOrderAgreementTab scheduleAOnly tables={[]} onChange={() => {}} />
        </div>
      )}
    </div>
  )
}
