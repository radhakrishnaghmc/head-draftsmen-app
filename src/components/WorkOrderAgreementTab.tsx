import { useEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { parseIntimationNotice, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { deriveFields, workOrderPlaceholders, agreementPlaceholders } from '@core/workOrderAgreement'
import { boqToScheduleA, extractWorkNameFromBoq } from '../boqTransform'
import { buildScheduleARows, rowsToScheduleAItems, metaFromWorksRow, findWorksRowByName } from '@core/scheduleA'
import { mismatchHint } from '../docClassify'
import type { PlaceholderMatch } from '@core/createDocument'
import type { ScheduleAMeta } from '../../electron/ipc-contract'
import type { ExcelTable } from '@core/types'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconCheck, IconClipboard, IconTable } from './Icons'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

// The uploaded file is a BOQ/estimate, but the saved document is Schedule A —
// drop any "BOQ" wording from its name so the two don't end up side by side.
function stripBoqWord(name: string): string {
  return name
    .replace(/\bboq\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
}

type DocKind = 'workOrder' | 'agreement'
type Output = DocKind | 'scheduleA'

const DOC_LABEL: Record<Output, string> = {
  workOrder: 'Work Order',
  agreement: 'Agreement Bond',
  scheduleA: 'Schedule A'
}

/**
 * Work order and agreement — produces this work's three outputs (Work Order,
 * Agreement Bond, Schedule A) as a gallery of tiles, each showing a live
 * preview. Mirrors Give Intimation's flow: upload the Online Intimation
 * (.html) and the L-1 selection form (.pdf) and the two document tiles appear;
 * upload the technical-sanctioned estimate / BOQ and the Schedule A tile
 * appears. Clicking a tile opens the full preview with Word / PDF / Print
 * (Excel for Schedule A). The L-1 form both fills the documents and updates
 * that work's Works List row. Reuses the docx-preview + export/print pipeline
 * of Give Intimation.
 */
export default function WorkOrderAgreementTab({ tables, onChange }: Props) {
  const table = tables[0] ?? null

  const [workOrderB64, setWorkOrderB64] = useState<string | null>(null)
  const [agreementB64, setAgreementB64] = useState<string | null>(null)
  const [workOrderLabels, setWorkOrderLabels] = useState<string[]>([])
  const [agreementLabels, setAgreementLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)

  // Which output's preview is expanded to the full-size modal, if any.
  const [expanded, setExpanded] = useState<Output | null>(null)
  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [expandedPages, setExpandedPages] = useState(0)

  // Schedule A (from an uploaded technical-sanctioned estimate / BOQ).
  const [boq, setBoq] = useState<ExcelTable | null>(null)
  const [scheduleA, setScheduleA] = useState<ExcelTable | null>(null)
  const [scheduleAError, setScheduleAError] = useState<string | null>(null)
  const [scheduleABusy, setScheduleABusy] = useState(false)

  const pdfInputRef = useRef<HTMLInputElement>(null)
  const noticeInputRef = useRef<HTMLInputElement>(null)
  const woTileRef = useRef<HTMLDivElement>(null)
  const agTileRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null

  // Load both bundled formats once, and read their placeholders.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const [woB64, agB64] = await Promise.all([api.workOrderTemplate(), api.agreementTemplate()])
        const [woLabels, agLabels] = await Promise.all([
          api.findPlaceholdersInDocument(woB64),
          api.findPlaceholdersInDocument(agB64)
        ])
        if (cancelled) return
        setWorkOrderB64(woB64)
        setAgreementB64(agB64)
        setWorkOrderLabels(woLabels)
        setAgreementLabels(agLabels)
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  // Everything the two documents print, resolved from the uploaded Online
  // Intimation + L-1 selection form + the matched Works List row.
  const fields = useMemo(
    () => deriveFields(notice ?? {}, pdfEval ?? {}, selectedRow ?? {}),
    [notice, pdfEval, selectedRow]
  )

  async function handleNoticeFile(file: File) {
    setActionError(null)
    try {
      const text = await file.text()
      setNotice(parseIntimationNotice(text))
      setNoticeName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  // The L-1 selection / evaluation PDF fills the tender fields (ECV / Tender %/
  // contract / agency) and updates that work's Works List row — matched by
  // Name of Work — so the database reflects the award in one step. The matched
  // work is auto-selected so both documents show it.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)

      if (table && ev.nameOfWork) {
        let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
        const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
        if (nameHeader) {
          try {
            const [rowNameVectors, evalNameVectors] = await Promise.all([
              api.embedTexts(table.rows.map((r) => r[nameHeader] ?? '')),
              api.embedTexts([ev.nameOfWork])
            ])
            embeddings = { rowNameVectors, evalNameVectors }
          } catch {
            embeddings = undefined
          }
        }
        const { table: updated, matchedCount } = updateWorksListFromEvaluations(table, [ev], embeddings)
        if (matchedCount > 0) {
          onChange(updated)
          const nh = updated.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
          const idx = nh ? updated.rows.findIndex((r) => norm(r[nh] ?? '') === norm(ev.nameOfWork!)) : -1
          if (idx >= 0) setRowIndex(idx)
          setPdfStatus(`Updated "${ev.nameOfWork}" in the Works List (Tender ID, Notice No/Date, ECV, Agency, Tender %, Contract Amount).`)
        } else {
          setPdfStatus(`Read the PDF, but no Works List row matched "${ev.nameOfWork}" — the documents are filled, the database wasn't changed.`)
        }
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function fillDoc(kind: DocKind): Promise<string> {
    const b64 = kind === 'workOrder' ? workOrderB64 : agreementB64
    const labels = kind === 'workOrder' ? workOrderLabels : agreementLabels
    if (!b64) throw new Error('Format not loaded yet.')
    const values = kind === 'workOrder' ? workOrderPlaceholders(fields) : agreementPlaceholders(fields)
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(b64, resolved, values)
  }

  async function renderDocInto(kind: DocKind, container: HTMLElement): Promise<number> {
    const filled = await fillDoc(kind)
    container.innerHTML = ''
    await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
    return container.querySelectorAll('section.docx').length
  }

  // Guard: the Online Intimation and the L1 selection form must describe the
  // same work (matched by NIT No, or agency name when the NIT No is absent) —
  // otherwise the documents would splice one work's agency onto another's
  // tender.
  const workMatch = useMemo(
    () => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null),
    [notice, pdfEval]
  )
  const workMismatch = workMatch?.status === 'mismatch'

  // Only build the documents once BOTH the Online Intimation and the L1
  // selection form are uploaded, and they belong to the same work — no tiles
  // are shown before that (same gate as Give Intimation).
  const bothUploaded = !!notice && !!pdfEval && !workMismatch
  const templatesReady = !!workOrderB64 && !!agreementB64
  const docsReady = templatesReady && bothUploaded

  // Live thumbnails in the two document tiles, refreshed whenever the filled
  // values change.
  useEffect(() => {
    if (!docsReady) return
    if (woTileRef.current) void renderDocInto('workOrder', woTileRef.current).catch(() => {})
    if (agTileRef.current) void renderDocInto('agreement', agTileRef.current).catch(() => {})
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [docsReady, fields])

  // The full-size preview inside the expanded modal (documents only — Schedule
  // A renders its table as JSX below).
  useEffect(() => {
    if (expanded !== 'workOrder' && expanded !== 'agreement') return
    const container = expandedRef.current
    if (!container) return
    setActionError(null)
    void renderDocInto(expanded, container)
      .then((pages) => setExpandedPages(pages))
      .catch((e) => setActionError(e instanceof Error ? e.message : String(e)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [expanded, fields])

  function docName(kind: DocKind): string {
    const agency = fields.agencyName ? ` - ${fields.agencyName}` : ''
    return `${DOC_LABEL[kind]}${agency}`
  }

  async function download(kind: DocKind, formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillDoc(kind)
      const res = await api.exportCreatedDocument(filled, docName(kind), formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function print(kind: DocKind) {
    setBusy('print')
    setActionError(null)
    try {
      const filled = await fillDoc(kind)
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // --- Schedule A (uploaded estimate / BOQ) ---

  async function uploadBoq() {
    setScheduleAError(null)
    const uploaded = await api.pickExcels()
    if (uploaded.length === 0) return
    const t = uploaded[0]
    setBoq(t)
    try {
      setScheduleA(boqToScheduleA(t))
      // If a BOQ names its own work, and it matches a Works List row, hop the
      // selection there so the whole tab (documents + Schedule A) lines up.
      const detected = extractWorkNameFromBoq(t)
      if (detected && table) {
        const match = findWorksRowByName(table, detected)
        if (match) {
          const idx = table.rows.indexOf(match.row)
          if (idx >= 0) setRowIndex(idx)
        }
      }
    } catch (e) {
      setScheduleA(null)
      const hint = mismatchHint(t.headers, 'boq')
      setScheduleAError((e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : ''))
    }
  }

  const scheduleAMeta: ScheduleAMeta | undefined = useMemo(
    () => (selectedRow ? metaFromWorksRow(selectedRow) : undefined),
    [selectedRow]
  )

  const scheduleAPreview = useMemo(
    () => (scheduleA ? buildScheduleARows(rowsToScheduleAItems(scheduleA), scheduleAMeta) : null),
    [scheduleA, scheduleAMeta]
  )

  async function downloadScheduleA() {
    if (!scheduleA) return
    setScheduleABusy(true)
    setScheduleAError(null)
    try {
      const base = boq ? stripBoqWord(stripExt(boq.name)) : ''
      const suggestedName = base ? `${base} Schedule A` : 'Schedule A'
      const savedPath = await api.exportScheduleA(scheduleA, suggestedName, scheduleAMeta)
      if (savedPath) setActionSaved(savedPath ? `Saved: ${savedPath}` : 'Cancelled.')
    } catch (e) {
      setScheduleAError(e instanceof Error ? e.message : String(e))
    } finally {
      setScheduleABusy(false)
    }
  }

  function scheduleATableHtml(): string {
    if (!scheduleAPreview) return ''
    const rows = scheduleAPreview
      .map((row) => {
        const cells = Array.from({ length: 6 }, (_, ci) => String(row[ci] ?? ''))
        const wide = cells.slice(1).every((c) => c === '')
        if (wide) return `<tr><td colspan="6" style="font-weight:600">${cells[0]}</td></tr>`
        return `<tr>${cells.map((c) => `<td>${c}</td>`).join('')}</tr>`
      })
      .join('')
    return `<table style="border-collapse:collapse;width:100%;font-family:'Times New Roman',serif;font-size:12px">
      <style>td{border:1px solid #333;padding:5px 7px;vertical-align:top}</style>${rows}</table>`
  }

  async function printScheduleA() {
    setScheduleABusy(true)
    setScheduleAError(null)
    try {
      await api.printCreatedDocument(scheduleATableHtml())
    } catch (e) {
      setScheduleAError(e instanceof Error ? e.message : String(e))
    } finally {
      setScheduleABusy(false)
    }
  }

  const noWorks = !table || table.rows.length === 0
  const anyOutput = docsReady || !!scheduleAPreview

  return (
    <div className="card">
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty">
        <IconClipboard />
        <p>
          Upload the <strong>Online Intimation</strong> and the <strong>L1 selection form</strong> to build the Work
          Order and Agreement Bond. The L1 form also updates that work's row in the Works List. Upload the
          technical-sanctioned <strong>estimate / BOQ</strong> to also generate this work's Schedule&nbsp;A. Each
          output appears below as a tile — click one to preview it full size and print / save it.
        </p>
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={() => noticeInputRef.current?.click()} disabled={!templatesReady}>
            <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
          </button>
          <button className="primary upload-btn" onClick={() => pdfInputRef.current?.click()} disabled={!templatesReady || busy === 'pdf'}>
            <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
          </button>
          <button className="primary upload-btn" onClick={uploadBoq} disabled={noWorks}>
            <IconTable /> {boq ? 'Change estimate / BOQ' : 'Upload estimate / BOQ (Schedule A)'}
          </button>
          <input
            ref={noticeInputRef}
            type="file"
            accept=".html,.htm,text/html"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNoticeFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={pdfInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handlePdfFile(file)
              e.target.value = ''
            }}
          />
        </div>
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
        {pdfName && <p className="estimate-hint">Tender details read from {pdfName}</p>}
        {pdfStatus && (
          <div className="notice ok">
            <IconCheck /> {pdfStatus}
          </div>
        )}
      </div>

      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}
      {scheduleAError && (
        <div className="notice error">
          <IconWarn /> {scheduleAError}
        </div>
      )}
      {workMismatch && workMatch && (
        <div className="notice error">
          <IconWarn /> {sameWorkMismatchMessage(workMatch)}
        </div>
      )}
      {actionSaved && !expanded && (
        <div className="notice ok">
          <IconCheck /> {actionSaved}
        </div>
      )}

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the outputs are filled from a work's row.</div>
      ) : anyOutput ? (
        <div className="wo-tiles">
          {docsReady && (
            <button className="wo-tile" onClick={() => setExpanded('workOrder')}>
              <div className="wo-tile-preview">
                <div ref={woTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.workOrder}</div>
            </button>
          )}
          {docsReady && (
            <button className="wo-tile" onClick={() => setExpanded('agreement')}>
              <div className="wo-tile-preview">
                <div ref={agTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.agreement}</div>
            </button>
          )}
          {scheduleAPreview && (
            <button className="wo-tile" onClick={() => setExpanded('scheduleA')}>
              <div className="wo-tile-preview sched">
                <table className="sa-preview-table">
                  <tbody>
                    {scheduleAPreview.slice(0, 14).map((row, ri) => {
                      const cells = Array.from({ length: 6 }, (_, ci) => row[ci] ?? '')
                      const wide = cells.slice(1).every((c) => c === '')
                      return (
                        <tr key={ri}>
                          {wide ? (
                            <td colSpan={6} className="sa-wide">
                              {String(cells[0])}
                            </td>
                          ) : (
                            cells.map((c, ci) => <td key={ci}>{String(c)}</td>)
                          )}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.scheduleA}</div>
            </button>
          )}
        </div>
      ) : null}

      {expanded && (
        <div className="wo-modal-overlay" onClick={() => setExpanded(null)}>
          <div className="wo-modal" onClick={(e) => e.stopPropagation()}>
            <div className="wo-modal-head">
              <span className="wo-modal-title">
                {DOC_LABEL[expanded]}
                {expanded !== 'scheduleA' && expandedPages > 1 ? ` — ${expandedPages} pages` : ''}
              </span>
              <button className="wo-modal-close" onClick={() => setExpanded(null)} title="Close" aria-label="Close">
                ×
              </button>
            </div>

            <div className="wo-modal-body">
              {expanded === 'scheduleA' ? (
                <div className="sheet-wrap">
                  <table className="sa-preview-table">
                    <tbody>
                      {scheduleAPreview?.map((row, ri) => {
                        const cells = Array.from({ length: 6 }, (_, ci) => row[ci] ?? '')
                        const wide = cells.slice(1).every((c) => c === '')
                        return (
                          <tr key={ri}>
                            {wide ? (
                              <td colSpan={6} className="sa-wide">
                                {String(cells[0])}
                              </td>
                            ) : (
                              cells.map((c, ci) => <td key={ci}>{String(c)}</td>)
                            )}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div ref={expandedRef} className="intimation-docx-preview" />
              )}
            </div>

            <div className="wo-modal-foot">
              {actionError && (
                <div className="notice error" style={{ marginRight: 'auto' }}>
                  <IconWarn /> {actionError}
                </div>
              )}
              {expanded === 'scheduleA' ? (
                <>
                  <button className="primary" onClick={downloadScheduleA} disabled={scheduleABusy}>
                    <IconDownload /> {scheduleABusy ? 'Working…' : 'Excel'}
                  </button>
                  <button className="primary" onClick={printScheduleA} disabled={scheduleABusy}>
                    <IconPrint /> Print
                  </button>
                </>
              ) : (
                <>
                  <button className="primary" onClick={() => download(expanded, ['docx'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                  </button>
                  <button className="primary" onClick={() => download(expanded, ['pdf'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                  </button>
                  <button className="primary" onClick={() => print(expanded)} disabled={busy !== null}>
                    <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                  </button>
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
