import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconTable, IconFolder, IconWarn, IconOpen, IconImage, IconClipboard, IconBolt } from './Icons'
import UploadPhotosTab from './UploadPhotosTab'
import WorkOrderAgreementTab from './WorkOrderAgreementTab'
import ElectricalEstimateTab from './ElectricalEstimateTab'
import type { ExcelTable } from '@core/types'

interface SplitState {
  busy: boolean
  result: { dir: string; files: string[] } | null
  error: string | null
}

interface SplitProgress {
  done: number
  total: number
  sheet: string
}

// The tool tiles whose panel opens beneath them (the Excel Separator runs a
// one-shot dialog instead of opening a panel, so it isn't one of these).
type Panel = 'photos' | 'workOrder' | 'agreement' | 'scheduleA' | 'electrical'

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
  // Live per-sheet progress pushed from the main process while a split runs.
  const [progress, setProgress] = useState<SplitProgress | null>(null)
  // Which tool's panel is expanded, if any. One at a time (accordion): opening a
  // tile reveals its panel directly beneath that tile and closes any other, so
  // the workspace stays focused on the one tool the user picked. Each tile is
  // its own focused entry point — Work Order and Agreement Bond ask only for the
  // L-1 + Intimation and each show only their own document.
  const [open, setOpen] = useState<Panel | null>(null)
  // Whether the currently-open single-upload tool has picked a file yet. Until
  // it has, its panel must take up no grid row (else the empty full-width row
  // would push the remaining tiles onto the next line — see the wrappers below).
  const [panelFilled, setPanelFilled] = useState(false)

  // Open (or toggle shut) a tool's panel. Also clears the Excel Separator's
  // "Saved N sheets…" result so that notice doesn't linger once the user has
  // moved on to another tile, and resets the single-upload "has content" flag.
  function toggle(panel: Panel) {
    setOpen((cur) => (cur === panel ? null : panel))
    setPanelFilled(false)
    setSplit({ busy: false, result: null, error: null })
    setProgress(null)
  }

  // Subscribe once to the main process's per-sheet progress events; the split
  // itself runs in the main process, so the whole UI (this and every other tab)
  // stays usable while it works.
  useEffect(() => api.onSplitProgress(setProgress), [])

  async function runSeparator() {
    if (split.busy) return
    // Close any open tool panel (e.g. Work Order / Agreement) — picking the
    // separator is a switch to a different tool, so its panel shouldn't linger.
    setOpen(null)
    setPanelFilled(false)
    setSplit({ busy: true, result: null, error: null })
    setProgress(null)
    try {
      // null when the user cancels either dialog — leave the tile idle.
      const result = await api.splitExcelSheets()
      setSplit({ busy: false, result, error: null })
    } catch (e) {
      setSplit({ busy: false, result: null, error: e instanceof Error ? e.message : String(e) })
    } finally {
      setProgress(null)
    }
  }

  return (
    <div className="card">
      <div className="doc-tile-grid">
        <button
          className="doc-tile-card tone-teal tool-card"
          onClick={runSeparator}
          disabled={split.busy}
        >
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Excel Sheet Separator</span>
          <span className="doc-tile-card-meta">
            {split.busy
              ? progress
                ? `Splitting ${progress.done} / ${progress.total} sheets…`
                : 'Reading workbook…'
              : 'One file per sheet, named after each tab'}
          </span>
          <span className="tool-card-cta">
            <IconFolder /> Choose workbook
          </span>
          {/* Progress lives inside the tile itself — a thin bar pinned to the
              bottom edge while a split runs — so nothing pushes the other tiles
              around. The per-sheet count shows in the meta line above. */}
          {split.busy && (
            <span className="tool-card-progress" aria-hidden>
              <span
                className="tool-card-progress-bar"
                style={{ width: progress && progress.total > 0 ? `${(progress.done / progress.total) * 100}%` : '8%' }}
              />
            </span>
          )}
        </button>

        <button
          className={`doc-tile-card tone-amber tool-card ${open === 'photos' ? 'on' : ''}`}
          onClick={() => toggle('photos')}
          aria-expanded={open === 'photos'}
        >
          <span className="tool-card-ic">
            <IconImage />
          </span>
          <span className="doc-tile-card-name">Estimate from Photos / PDF</span>
          <span className="doc-tile-card-meta">
            {open === 'photos' ? 'Open below — click to hide' : 'Photos / scanned PDF → BOQ, Schedule A, Deviation, Material'}
          </span>
          <span className="tool-card-cta">
            <IconImage /> Upload photos / PDF
          </span>
        </button>
        {open === 'photos' && (
          <div
            className={panelFilled ? 'tool-panel-row workspace-section' : ''}
            style={panelFilled ? undefined : { display: 'contents' }}
          >
            <UploadPhotosTab tables={tables} onChange={onChange} autoOpen onContent={setPanelFilled} />
          </div>
        )}

        <div className="tool-cell">
          <button
            className={`doc-tile-card tone-sky tool-card ${open === 'workOrder' ? 'on' : ''}`}
            onClick={() => toggle('workOrder')}
            aria-expanded={open === 'workOrder'}
          >
            <span className="tool-card-ic">
              <IconClipboard />
            </span>
            <span className="doc-tile-card-name">Work Order</span>
            <span className="doc-tile-card-meta">
              {open === 'workOrder' ? 'Open below — click to hide' : 'From L1 + Intimation — any circle/zone'}
            </span>
            <span className="tool-card-cta">
              <IconFolder /> Upload L1 + Intimation
            </span>
          </button>
          {open === 'workOrder' && (
            <div className="tool-inline-panel">
              <WorkOrderAgreementTab standalone only="workOrder" tables={[]} onChange={() => {}} />
            </div>
          )}
        </div>

        <div className="tool-cell">
          <button
            className={`doc-tile-card tone-sky tool-card ${open === 'agreement' ? 'on' : ''}`}
            onClick={() => toggle('agreement')}
            aria-expanded={open === 'agreement'}
          >
            <span className="tool-card-ic">
              <IconClipboard />
            </span>
            <span className="doc-tile-card-name">Agreement Bond</span>
            <span className="doc-tile-card-meta">
              {open === 'agreement' ? 'Open below — click to hide' : 'From L1 + Intimation — any circle/zone'}
            </span>
            <span className="tool-card-cta">
              <IconFolder /> Upload L1 + Intimation
            </span>
          </button>
          {open === 'agreement' && (
            <div className="tool-inline-panel">
              <WorkOrderAgreementTab standalone only="agreement" tables={[]} onChange={() => {}} />
            </div>
          )}
        </div>

        <button
          className={`doc-tile-card tone-sky tool-card ${open === 'scheduleA' ? 'on' : ''}`}
          onClick={() => toggle('scheduleA')}
          aria-expanded={open === 'scheduleA'}
        >
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Schedule A</span>
          <span className="doc-tile-card-meta">
            {open === 'scheduleA' ? 'Open below — click to hide' : 'From an uploaded estimate / BOQ'}
          </span>
          <span className="tool-card-cta">
            <IconTable /> Upload estimate / BOQ
          </span>
        </button>
        {open === 'scheduleA' && (
          <div
            className={panelFilled ? 'tool-panel-row workspace-section' : ''}
            style={panelFilled ? undefined : { display: 'contents' }}
          >
            <WorkOrderAgreementTab scheduleAOnly autoOpen onContent={setPanelFilled} tables={[]} onChange={() => {}} />
          </div>
        )}

        <button
          className={`doc-tile-card tone-green tool-card ${open === 'electrical' ? 'on' : ''}`}
          onClick={() => toggle('electrical')}
          aria-expanded={open === 'electrical'}
        >
          <span className="tool-card-ic">
            <IconBolt />
          </span>
          <span className="doc-tile-card-name">Electrical Estimate</span>
          <span className="doc-tile-card-meta">
            {open === 'electrical' ? 'Open below — click to hide' : 'Electrical estimate → BOQ + Schedule A'}
          </span>
          <span className="tool-card-cta">
            <IconBolt /> Upload electrical estimate
          </span>
        </button>
        {open === 'electrical' && (
          <div
            className={panelFilled ? 'tool-panel-row workspace-section' : ''}
            style={panelFilled ? undefined : { display: 'contents' }}
          >
            <ElectricalEstimateTab autoOpen onContent={setPanelFilled} />
          </div>
        )}
      </div>

      {/* Separator outcome sits below the whole grid, so it never pushes the
          tiles around. Progress itself shows in the tile (above). */}
      {split.error && (
        <div className="notice error tool-outcome">
          <IconWarn />
          {split.error}
        </div>
      )}
      {split.result && (
        <div className="notice ok tool-result tool-outcome">
          <span>
            <IconFolder /> Saved {split.result.files.length} sheet
            {split.result.files.length === 1 ? '' : 's'} to {split.result.dir}
          </span>
          <button className="ghost" onClick={() => api.openPath(split.result!.dir)}>
            <IconOpen /> Open folder
          </button>
        </div>
      )}
    </div>
  )
}
