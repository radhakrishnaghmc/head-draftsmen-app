import { useRef, useState } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { IconTable, IconFolder, IconWarn, IconImage, IconClipboard, IconBolt, IconPrint, IconDownload, IconBell, IconDoc, IconChevronLeft } from './Icons'
import UploadPhotosTab from './UploadPhotosTab'
import WorkOrderAgreementTab from './WorkOrderAgreementTab'
import IntimationToolTab from './IntimationToolTab'
import ElectricalEstimateTab from './ElectricalEstimateTab'
import PdfToolPage from './PdfToolPage'
import ExcelSeparatorPage from './ExcelSeparatorPage'
import WordToolPage from './WordToolPage'
import GpsPhotosPage from './GpsPhotosPage'
import PhotosToPdfPage from './PhotosToPdfPage'
import PhotosToDocPage from './PhotosToDocPage'
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

// Every tool tile opens its own full page (Back button → grid) in place of the
// grid. The file tools (pdf/excel/word) are self-contained page components; the
// rest host an existing workspace component under a shared Back-button header.
type FullPage =
  | 'pdf'
  | 'excel'
  | 'word'
  | 'gps'
  | 'photosToPdf'
  | 'photosToDoc'
  | 'photos'
  | 'workOrder'
  | 'agreement'
  | 'intimation'
  | 'scheduleA'
  | 'electrical'

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
  // Which full-page tool (if any) has taken over the tab: the PDF workspace
  // (upload PDFs, pick pages, merge/separate) or the Excel Sheet Separator (pick
  // a workbook, tick sheets, separate). Each renders in place of the grid with a
  // Back button.
  const [fullPage, setFullPage] = useState<FullPage | null>(null)
  // Which tool's panel is expanded, if any. One at a time (accordion): opening a
  // tile reveals its panel directly beneath that tile and closes any other, so
  // the workspace stays focused on the one tool the user picked. Each tile is
  // its own focused entry point — Work Order and Agreement Bond ask only for the
  // L-1 + Intimation and each show only their own document.
  function openFullPage(page: FullPage) {
    setFullPage(page)
  }
  const backToTools = () => setFullPage(null)

  // The self-contained file-tool pages.
  if (fullPage === 'pdf') return <PdfToolPage onBack={backToTools} />
  if (fullPage === 'excel') return <ExcelSeparatorPage onBack={backToTools} />
  if (fullPage === 'word') return <WordToolPage onBack={backToTools} />
  if (fullPage === 'gps') return <GpsPhotosPage onBack={backToTools} />
  if (fullPage === 'photosToPdf') return <PhotosToPdfPage onBack={backToTools} />
  if (fullPage === 'photosToDoc') return <PhotosToDocPage onBack={backToTools} />

  // The remaining tools each host an existing workspace component under a shared
  // Back-button header, so every tile opens as its own full page.
  if (fullPage) {
    const hosted: Record<string, { icon: JSX.Element; label: string; body: JSX.Element }> = {
      photos: {
        icon: <IconImage />,
        label: 'Estimate from Photos / PDF',
        // No autoOpen: open the workspace first (with its "Upload Photos or PDF"
        // button) like the other file tools, rather than popping the OS file
        // dialog straight away on the tile click.
        body: <UploadPhotosTab tables={tables} onChange={onChange} onContent={() => {}} />
      },
      workOrder: {
        icon: <IconClipboard />,
        label: 'Work Order',
        body: <WorkOrderAgreementTab standalone only="workOrder" tables={[]} onChange={() => {}} office={office} />
      },
      agreement: {
        icon: <IconClipboard />,
        label: 'Agreement Bond',
        body: <WorkOrderAgreementTab standalone only="agreement" tables={[]} onChange={() => {}} office={office} />
      },
      intimation: {
        icon: <IconBell />,
        label: 'Intimation',
        body: <IntimationToolTab office={office} />
      },
      scheduleA: {
        icon: <IconTable />,
        label: 'Schedule A',
        // No autoOpen: open the workspace first (with its upload button), like
        // the other file tools, instead of popping the file dialog immediately.
        body: <WorkOrderAgreementTab scheduleAOnly onContent={() => {}} tables={tables} onChange={() => {}} />
      },
      electrical: {
        icon: <IconBolt />,
        label: 'Electrical Estimate',
        // No autoOpen: open the workspace first (with its upload button).
        body: <ElectricalEstimateTab onContent={() => {}} />
      }
    }
    const tool = hosted[fullPage]
    return (
      <>
        <div className="pdf-ws-topbar">
          <div className="pdf-ws-title">
            {tool.icon} <span className="pdf-ws-title-label">Tool:</span> {tool.label}
          </div>
          <button className="ghost pdf-ws-back" onClick={backToTools}>
            <IconChevronLeft /> Back to Tools
          </button>
        </div>
        <div className="card pdf-workspace">
          <div className="tool-fullpage-body">{tool.body}</div>
        </div>
      </>
    )
  }

  return (
    <div className="card">
      <div className="doc-tile-grid tools-grid">
        <button className="doc-tile-card tone-teal tool-card" onClick={() => openFullPage('excel')}>
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Excel Sheet Separator</span>
          <span className="doc-tile-card-meta">Pick a workbook, tick sheets, separate selected or all</span>
          <span className="tool-card-cta">
            <IconFolder /> Open workspace
          </span>
        </button>

        <button className="doc-tile-card tone-sky tool-card" onClick={() => openFullPage('pdf')}>
          <span className="tool-card-ic">
            <IconDoc />
          </span>
          <span className="doc-tile-card-name">PDF Merge / Separator</span>
          <span className="doc-tile-card-meta">Upload PDFs, pick pages, merge into one or save each separately</span>
          <span className="tool-card-cta">
            <IconFolder /> Open workspace
          </span>
        </button>

        <button className="doc-tile-card tone-rose tool-card" onClick={() => openFullPage('word')}>
          <span className="tool-card-ic">
            <IconDoc />
          </span>
          <span className="doc-tile-card-name">Word Merge / Separator</span>
          <span className="doc-tile-card-meta">Upload Word files, pick pages (PDF), or merge whole files into one .docx</span>
          <span className="tool-card-cta">
            <IconFolder /> Open workspace
          </span>
        </button>

        <button className="doc-tile-card tone-amber tool-card" onClick={() => openFullPage('photosToPdf')}>
          <span className="tool-card-ic">
            <IconDoc />
          </span>
          <span className="doc-tile-card-name">Photos → PDF</span>
          <span className="doc-tile-card-meta">Combine photos (or PDF pages) into one PDF — auto-crops borders</span>
          <span className="tool-card-cta">
            <IconFolder /> Upload photos or PDF
          </span>
        </button>

        <button className="doc-tile-card tone-sky tool-card" onClick={() => openFullPage('photosToDoc')}>
          <span className="tool-card-ic">
            <IconDoc />
          </span>
          <span className="doc-tile-card-name">Photos / PDF → Word / Excel</span>
          <span className="doc-tile-card-meta">OCR photos or a scanned PDF into editable text, export to .docx or .xlsx</span>
          <span className="tool-card-cta">
            <IconFolder /> Upload photos / PDF
          </span>
        </button>

        <button className="doc-tile-card tone-green tool-card" onClick={() => openFullPage('gps')}>
          <span className="tool-card-ic">
            <IconImage />
          </span>
          <span className="doc-tile-card-name">GPS Photos → Latitude / Longitude</span>
          <span className="doc-tile-card-meta">Upload photos, read each one's GPS from its metadata, export to Excel</span>
          <span className="tool-card-cta">
            <IconImage /> Upload GPS photos
          </span>
        </button>

        <button className="doc-tile-card tone-amber tool-card" onClick={() => openFullPage('photos')}>
          <span className="tool-card-ic">
            <IconImage />
          </span>
          <span className="doc-tile-card-name">Estimate from Photos / PDF</span>
          <span className="doc-tile-card-meta">Photos / scanned PDF → BOQ, Schedule A, Deviation, Material</span>
          <span className="tool-card-cta">
            <IconImage /> Upload photos / PDF
          </span>
        </button>

        <button className="doc-tile-card tone-sky tool-card" onClick={() => openFullPage('workOrder')}>
          <span className="tool-card-ic">
            <IconClipboard />
          </span>
          <span className="doc-tile-card-name">Work Order</span>
          <span className="doc-tile-card-meta">From L1 + Intimation — any circle/zone</span>
          <span className="tool-card-cta">
            <IconFolder /> Upload L1 + Intimation
          </span>
        </button>

        <button className="doc-tile-card tone-sky tool-card" onClick={() => openFullPage('agreement')}>
          <span className="tool-card-ic">
            <IconClipboard />
          </span>
          <span className="doc-tile-card-name">Agreement Bond</span>
          <span className="doc-tile-card-meta">From L1 + Intimation — any circle/zone</span>
          <span className="tool-card-cta">
            <IconFolder /> Upload L1 + Intimation
          </span>
        </button>

        <button className="doc-tile-card tone-rose tool-card" onClick={() => openFullPage('intimation')}>
          <span className="tool-card-ic">
            <IconBell />
          </span>
          <span className="doc-tile-card-name">Intimation</span>
          <span className="doc-tile-card-meta">From L1 — office by circle in the work name — any circle</span>
          <span className="tool-card-cta">
            <IconFolder /> Upload L1 selection form
          </span>
        </button>

        <button className="doc-tile-card tone-sky tool-card" onClick={() => openFullPage('scheduleA')}>
          <span className="tool-card-ic">
            <IconTable />
          </span>
          <span className="doc-tile-card-name">Schedule A</span>
          <span className="doc-tile-card-meta">From an uploaded estimate / BOQ</span>
          <span className="tool-card-cta">
            <IconTable /> Upload estimate / BOQ
          </span>
        </button>

        <button className="doc-tile-card tone-green tool-card" onClick={() => openFullPage('electrical')}>
          <span className="tool-card-ic">
            <IconBolt />
          </span>
          <span className="doc-tile-card-name">Electrical Estimate</span>
          <span className="doc-tile-card-meta">Electrical estimate → BOQ + Schedule A</span>
          <span className="tool-card-cta">
            <IconBolt /> Upload electrical estimate
          </span>
        </button>

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

    </div>
  )
}
