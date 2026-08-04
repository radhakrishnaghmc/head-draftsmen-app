import { useEffect, useRef, useState } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { IconTable, IconFolder, IconWarn, IconOpen, IconImage, IconClipboard, IconBolt, IconPrint, IconDownload } from './Icons'
import UploadPhotosTab from './UploadPhotosTab'
import WorkOrderAgreementTab from './WorkOrderAgreementTab'
import ElectricalEstimateTab from './ElectricalEstimateTab'
import DocThumbnail from './DocThumbnail'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH } from './docPage'
import { circlesOf, corporationByName } from '../zoneCircleDirectory'
import type { CreatedDocument, ExcelTable } from '@core/types'
import type { Office } from '../office'

// Tones cycled across the document tiles so the row of blank-form tiles reads
// like the rest of the frosted-glass grid.
const DOC_TILE_TONES = ['tone-teal', 'tone-amber', 'tone-sky', 'tone-rose', 'tone-green']

// What a document tile can do with its blank form: send to the printer, or save
// as Word / PDF ('docx' | 'pdf' match exportCreatedDocument's format union).
type DocAction = 'print' | 'docx' | 'pdf'

// The office's Corporation/Zone/Circle/CNO as the fixed-placeholder values baked
// into a document (matching bakeLoginPlaceholders' keys), skipping any unset.
// `circle`/`cno` can be overridden (a zone-only office picks a circle at print).
function officeValues(office: Office, circle?: string, cno?: string): Record<string, string> {
  const v: Record<string, string> = {}
  if (office.corporation) {
    v.corporation = office.corporation
    const full = corporationByName(office.corporation)?.fullName
    if (full) v['corporation full name'] = full.toUpperCase()
  }
  if (office.zone) v.zone = office.zone
  const c = circle ?? office.circle
  if (c) v.circle = c
  const n = cno ?? office.circleNumber
  if (n) v.cno = n
  return v
}

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
  /** The chosen office — so the Work Order / Agreement "Fill details manually" form can take Circle/Zone/Corporation from it instead of re-asking. */
  office?: Office
  /** The Issue-Documents templates — shown here as blank-form tiles that print
   * empty (Zone/Circle already baked in, other placeholders left to hand-fill),
   * regardless of the selected office. */
  documents?: CreatedDocument[]
}

/**
 * Utility tools that sit outside the main tender/estimate workflow, shown as a
 * grid of tiles (the app's frosted-glass doc-tile look) plus full-width tool
 * panels below. Today: the Excel Sheet Separator, and reading an estimate from
 * photos / a scanned PDF. New tools slot in as additional tiles or panels.
 */
export default function ToolsTab({ tables, onChange, office, documents = [] }: Props) {
  // Off-screen holder the docx is rendered into before it's handed to the OS
  // print dialog — printCreatedDocument needs plain HTML, not a docx buffer.
  const printScratchRef = useRef<HTMLDivElement>(null)
  // Which document tile is mid-print (renders its spinner/label), if any.
  const [printingId, setPrintingId] = useState<string | null>(null)
  const [printError, setPrintError] = useState<string | null>(null)
  // A zone-only office spans many circles, so before producing a document we ask
  // which circle's office details to stamp on it. Set while that picker is open,
  // remembering which action (print / Word / PDF) to run once a circle is chosen.
  const [circlePrompt, setCirclePrompt] = useState<{ doc: CreatedDocument; action: DocAction } | null>(null)
  const [chosenCircle, setChosenCircle] = useState<string>('')

  const zoneOnly = !!office?.zone && !office?.circle

  // Run a tile action (print, or download as Word / PDF). A zone-only office has
  // no circle, so first ask which circle's office details to stamp on.
  function runDocAction(doc: CreatedDocument, action: DocAction) {
    if (zoneOnly) {
      const circles = circlesOf(office?.corporation, office?.zone)
      setChosenCircle(circles[0]?.circle ?? '')
      setCirclePrompt({ doc, action })
      return
    }
    void doDocAction(doc, office ? officeValues(office) : {}, action)
  }

  // Stamp the chosen circle's office details, then run the pending action.
  function confirmCircleAndRun() {
    if (!circlePrompt || !office?.zone) return
    const entry = circlesOf(office.corporation, office.zone).find((e) => e.circle === chosenCircle)
    const { doc, action } = circlePrompt
    setCirclePrompt(null)
    void doDocAction(doc, officeValues(office, chosenCircle, entry?.cno), action)
  }

  // Build the blank form: stamp the office details (Corporation/Zone/Circle/CNO —
  // a no-op if they were already baked at login), then blank out every remaining
  // per-row placeholder so the form has empty fields to hand-fill, not literal
  // "{{...}}" text. The Dy. EE Forwarding Note gets three dotted lines where the
  // (often long) work name goes, so it can be hand-written — the fill turns each
  // "\n" into a real <w:br/>, so it holds up in Word/PDF, not just on print.
  async function buildBlankDocx(doc: CreatedDocument, office: Record<string, string>): Promise<string> {
    let docx = doc.docx
    if (Object.keys(office).length > 0) docx = await api.bakeFixedPlaceholdersInDocument(docx, office)
    const labels = await api.findPlaceholdersInDocument(docx)
    const nameLines = doc.id === 'doc_dy_ee_forwarding_note'
    const NAME_RE = /name of (the )?work/i
    const dottedLine = '.'.repeat(60)
    const dotted = [dottedLine, dottedLine, dottedLine].join('\n')
    const resolved = labels.map((label) =>
      nameLines && NAME_RE.test(label) ? { label, column: '__nameOfWork__', score: 1 } : { label, column: null, score: 0 }
    )
    return api.fillPlaceholdersInDocument(docx, resolved, nameLines ? { __nameOfWork__: dotted } : {})
  }

  async function doDocAction(doc: CreatedDocument, office: Record<string, string>, action: DocAction) {
    setPrintingId(doc.id)
    setPrintError(null)
    try {
      const docx = await buildBlankDocx(doc, office)
      if (action === 'print') {
        const container = printScratchRef.current
        if (!container) throw new Error('Print failed to initialize.')
        container.innerHTML = ''
        await renderAsync(base64ToUint8(docx), container, undefined, DOCX_PREVIEW_OPTIONS)
        await api.printCreatedDocument(container.innerHTML)
      } else {
        await api.exportCreatedDocument(docx, doc.name, [action])
      }
    } catch (e) {
      setPrintError(e instanceof Error ? e.message : String(e))
    } finally {
      setPrintingId(null)
    }
  }
  const [split, setSplit] = useState<SplitState>({ busy: false, result: null, error: null })
  // Live per-sheet progress pushed from the main process while a split runs.
  const [progress, setProgress] = useState<SplitProgress | null>(null)
  // A workbook the user picked but hasn't split yet — while set, a small chooser
  // (this sheet, or all) is shown so they can separate one sheet or every sheet.
  const [splitPick, setSplitPick] = useState<{ path: string; name: string; sheets: string[] } | null>(null)
  const [splitChoice, setSplitChoice] = useState<string>('all')
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
    setSplitPick(null)
  }

  // Subscribe once to the main process's per-sheet progress events; the split
  // itself runs in the main process, so the whole UI (this and every other tab)
  // stays usable while it works.
  useEffect(() => api.onSplitProgress(setProgress), [])

  // Step 1 — pick a workbook and read its sheet names, then show the chooser
  // (which sheet, or all) below. Picking the separator is a switch to a
  // different tool, so any open tool panel is closed.
  async function chooseWorkbook() {
    if (split.busy) return
    setOpen(null)
    setPanelFilled(false)
    setSplit({ busy: false, result: null, error: null })
    setProgress(null)
    try {
      const picked = await api.pickWorkbookForSplit() // null when the user cancels
      if (!picked) return
      setSplitPick(picked)
      setSplitChoice('all')
    } catch (e) {
      setSplit({ busy: false, result: null, error: e instanceof Error ? e.message : String(e) })
    }
  }

  // Step 2 — separate the chosen sheet (or all sheets) into the folder the user
  // then picks.
  async function doSplit() {
    if (!splitPick || split.busy) return
    const srcPath = splitPick.path
    const sheetNames = splitChoice === 'all' ? null : [splitChoice]
    setSplitPick(null)
    setSplit({ busy: true, result: null, error: null })
    setProgress(null)
    try {
      const result = await api.splitWorkbook(srcPath, sheetNames) // null when the folder dialog is cancelled
      setSplit({ busy: false, result, error: null })
    } catch (e) {
      setSplit({ busy: false, result: null, error: e instanceof Error ? e.message : String(e) })
    } finally {
      setProgress(null)
    }
  }

  return (
    <div className="card">
      <div className="doc-tile-grid tools-grid">
        <button
          className={`doc-tile-card tone-teal tool-card ${splitPick ? 'on' : ''}`}
          onClick={chooseWorkbook}
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
              : splitPick
                ? 'Choose a sheet below — or all'
                : 'One sheet, or all — one file per tab'}
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
              <WorkOrderAgreementTab standalone only="workOrder" tables={[]} onChange={() => {}} office={office} />
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
              <WorkOrderAgreementTab standalone only="agreement" tables={[]} onChange={() => {}} office={office} />
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
            <WorkOrderAgreementTab scheduleAOnly autoOpen onContent={setPanelFilled} tables={tables} onChange={() => {}} />
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

        {/* Every Issue-Documents template, as a blank-form tile — office-
            independent, so all show regardless of the selected office. Hover
            reveals Print / Word / PDF; each stamps the office details and blanks
            the rest of the fields for hand-filling. */}
        {documents.map((doc, i) => (
          <div
            key={doc.id}
            className={`doc-tile-card ${DOC_TILE_TONES[i % DOC_TILE_TONES.length]} tool-card doc-blank-card${
              printingId === doc.id ? ' busy' : ''
            }`}
            title={doc.name}
          >
            <DocThumbnail docx={doc.docx} width={72} />
            <span className="doc-tile-card-name">{doc.name}</span>
            <div className="doc-blank-actions">
              <button
                type="button"
                onClick={() => runDocAction(doc, 'print')}
                disabled={printingId === doc.id}
                title={`Print blank ${doc.name}`}
              >
                <IconPrint /> Print
              </button>
              <button
                type="button"
                onClick={() => runDocAction(doc, 'docx')}
                disabled={printingId === doc.id}
                title={`Download ${doc.name} as Word`}
              >
                <IconDownload /> Word
              </button>
              <button
                type="button"
                onClick={() => runDocAction(doc, 'pdf')}
                disabled={printingId === doc.id}
                title={`Download ${doc.name} as PDF`}
              >
                <IconDownload /> PDF
              </button>
            </div>
          </div>
        ))}
      </div>

      {/* Off-screen render target for the OS print dialog. */}
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />
      {printError && (
        <div className="notice error tool-outcome">
          <IconWarn />
          {printError}
        </div>
      )}

      {/* Zone-only office: ask which circle's office details to stamp, since a
          zonal Head Draughtsman spans every circle in the zone. */}
      {circlePrompt && (
        <div className="editor-overlay" onClick={() => setCirclePrompt(null)}>
          <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <h3>Which circle?</h3>
            <p className="confirm-hint">
              {office?.zone} zone spans several circles. Choose the circle to stamp on{' '}
              <strong>{circlePrompt.doc.name}</strong>.
            </p>
            <label className="gen-row-label">
              Circle:{' '}
              <select value={chosenCircle} onChange={(e) => setChosenCircle(e.target.value)}>
                {circlesOf(office?.corporation, office?.zone).map((e) => (
                  <option key={e.circle} value={e.circle}>
                    {e.circle} — {e.cno}
                  </option>
                ))}
              </select>
            </label>
            <div className="confirm-actions">
              <button className="ghost" onClick={() => setCirclePrompt(null)}>
                Cancel
              </button>
              <button className="primary" disabled={!chosenCircle} onClick={confirmCircleAndRun}>
                {circlePrompt.action === 'print' ? (
                  <>
                    <IconPrint /> Print
                  </>
                ) : (
                  <>
                    <IconDownload /> {circlePrompt.action === 'docx' ? 'Word' : 'PDF'}
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Sheet chooser for the Excel Separator: pick one sheet or separate all. */}
      {splitPick && (
        <div className="notice split-picker tool-outcome">
          <div className="split-picker-head">
            <IconTable />
            <strong>{splitPick.name}</strong>
            <span className="split-picker-count">
              {splitPick.sheets.length} sheet{splitPick.sheets.length === 1 ? '' : 's'}
            </span>
          </div>
          <label className="split-picker-row">
            <span>Separate</span>
            <select value={splitChoice} onChange={(e) => setSplitChoice(e.target.value)}>
              <option value="all">All sheets ({splitPick.sheets.length})</option>
              {splitPick.sheets.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>
          <div className="split-picker-actions">
            <button className="ghost" onClick={() => setSplitPick(null)}>
              Cancel
            </button>
            <button className="primary" onClick={doSplit}>
              <IconFolder /> {splitChoice === 'all' ? 'Separate all' : 'Separate sheet'}
            </button>
          </div>
        </div>
      )}

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
