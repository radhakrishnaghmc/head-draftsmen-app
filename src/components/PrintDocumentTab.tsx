import { useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import type { PlaceholderMatch } from '@core/createDocument'
import { withComputedAmounts } from '@core/worksAmounts'
import type { CreatedDocument, ExcelTable, QcOfficeParties } from '@core/types'
import type { Office } from '../office'
import { IconDoc, IconEye, IconPrint, IconDownload, IconPlus, IconSearch } from './Icons'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH } from './docPage'
import DocThumbnail from './DocThumbnail'

interface Props {
  tables: ExcelTable[]
  documents: CreatedDocument[]
  onChange: (docs: CreatedDocument[]) => void
  /** Switches to the Works List tab — offered when there are no rows to pick from. */
  onGoToWorksList: () => void
  /** The chosen office — some documents are offered only for a zonal (SE) or circle (EE) office. */
  office: Office
  /** The current office's 3rd/4th-party QC agencies (set on the Works List page) — fills those letters' "To" block. */
  qcParties?: QcOfficeParties
}

/**
 * Whether a document is offered for the current office. Unscoped documents show
 * everywhere; 'zonal' ones only for a Zone-level office (no circle picked), and
 * 'circle' ones only once a circle is selected. See CreatedDocument.officeScope.
 */
function isDocForOffice(doc: CreatedDocument, office: Office): boolean {
  if (!doc.officeScope) return true
  if (doc.officeScope === 'zonal') return !!office.zone && !office.circle
  return !!office.circle
}

function rowLabel(row: Record<string, string>, headers: string[], index: number): string {
  const nameCol = headers.find((h) => /name of (the )?work/i.test(h)) ?? headers[0]
  const v = nameCol ? (row[nameCol] ?? '').trim() : ''
  return v ? `${index + 1}. ${v}` : `Row ${index + 1}`
}

const TILE_TONES = ['tone-indigo', 'tone-sky', 'tone-rose', 'tone-amber', 'tone-teal', 'tone-green']

// The 3rd/4th-party QC letters are addressed to an outside agency whose
// name/address/phone (and, for the 4th party, the test type + AE contact) are
// not on the Works List and change year to year — so they're typed in at issue
// time. Their placeholders are filled from those inputs (and a computed Lakhs
// estimate), never matched to a Works column. Keyed lowercase for matching.
const PARTY_3RD_ID = 'doc_3rd_party_qc'
const PARTY_4TH_ID = 'doc_4th_party_intimation'
const MANUAL_LABELS = new Set([
  'party name',
  'party address',
  'party phone',
  'test type',
  'ae name phone',
  'estimate lakhs'
])
function isPartyDoc(doc: CreatedDocument | null): boolean {
  return doc?.id === PARTY_3RD_ID || doc?.id === PARTY_4TH_ID
}

export default function PrintDocumentTab({ tables, documents, onChange, onGoToWorksList, office, qcParties }: Props) {
  const table = tables[0] ?? null

  // Reorder/delete still act on the full synced list; only display is filtered.
  const visibleCount = useMemo(() => documents.filter((d) => isDocForOffice(d, office)).length, [documents, office])

  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)

  const [expandedId, setExpandedId] = useState<string | null>(null)
  const [rowIndex, setRowIndex] = useState(0)
  // Page-level work picker: the chosen Works List row every document is issued
  // against, found by typing part of its name rather than scrolling all rows.
  const [workSearch, setWorkSearch] = useState('')
  const nameCol = useMemo(
    () => (table ? table.headers.find((h) => /name of (the )?work/i.test(h)) ?? table.headers[0] : undefined),
    [table]
  )
  const workNameOf = (row: Record<string, string>) => (nameCol ? row[nameCol] ?? '' : '')
  const filteredRows = useMemo(() => {
    if (!table) return [] as { row: Record<string, string>; i: number }[]
    const q = workSearch.trim().toLowerCase()
    const all = table.rows.map((row, i) => ({ row, i }))
    return q ? all.filter(({ row }) => workNameOf(row).toLowerCase().includes(q)) : all
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table, workSearch, nameCol])

  // Keep the selected row valid as the filter narrows: if the current pick falls
  // outside the matches, jump to the first match so Preview/Create stay in sync.
  function onWorkSearch(v: string) {
    setWorkSearch(v)
    if (!table) return
    const q = v.trim().toLowerCase()
    const matches = table.rows.map((row, i) => ({ row, i })).filter(({ row }) => !q || workNameOf(row).toLowerCase().includes(q))
    if (matches.length > 0 && !matches.some((m) => m.i === rowIndex)) setRowIndex(matches[0].i)
  }
  // How the Issue button outputs: download Word, download PDF, or print. Pick
  // one, then click Issue.
  const [outputMode, setOutputMode] = useState<'word' | 'pdf' | 'print'>('word')
  const [genBusy, setGenBusy] = useState(false)
  const [genError, setGenError] = useState<string | null>(null)
  const [genNotice, setGenNotice] = useState<string | null>(null)
  const [preview, setPreview] = useState<{ docx: string; resolved: PlaceholderMatch[] } | null>(null)
  const [previewPages, setPreviewPages] = useState(0)
  const previewRef = useRef<HTMLDivElement>(null)
  // Off-screen render target used only to turn a filled .docx into HTML for
  // printing (see handlePrint) — printCreatedDocument needs plain HTML, not
  // a docx buffer, since printing goes through the OS print dialog against
  // a temp HTML file rather than requiring LibreOffice the way PDF export does.
  const printScratchRef = useRef<HTMLDivElement>(null)

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
      const next = [...documents]
      const [moved] = next.splice(dragIndex, 1)
      next.splice(index, 0, moved)
      onChange(next)
    }
    setDragIndex(null)
    setOverIndex(null)
  }

  function handleDragEnd() {
    setDragIndex(null)
    setOverIndex(null)
  }

  function toggleExpand(doc: CreatedDocument) {
    setGenError(null)
    setGenNotice(null)
    setPreview(null)
    setExpandedId((id) => (id === doc.id ? null : doc.id))
  }

  async function resolveForRow(doc: CreatedDocument): Promise<{ docx: string; resolved: PlaceholderMatch[] }> {
    const labels = await api.findPlaceholdersInDocument(doc.docx)
    // Amount-bearing columns (Amount of estimate, ECV, EMD @
    // 1%/1.5%, ASD, Contract Amount) resolve to their computed, Indian-
    // formatted "Rs 1,00,000/-" value rather than the raw spreadsheet figure
    // — see core/worksAmounts.ts.
    const rawRow = table?.rows[rowIndex] ?? {}
    const rawEstimate = (rawRow['Amount of estimate'] ?? '').trim()
    // The 3rd/4th-party letters' "To" agency comes from the office's saved QC
    // parties (set once on the Works List), not typed in per issue: the 3rd-party
    // letter uses the QC college, the 4th-party the testing lab.
    const party = doc.id === PARTY_3RD_ID ? qcParties?.third : doc.id === PARTY_4TH_ID ? qcParties?.fourth : undefined
    // Manual/computed values for the party-letter placeholders. Keys equal their
    // {{placeholder}} label so they self-resolve below. Test Type / AE contact
    // stay blank (hand-filled per test); Estimate Lakhs is the plain Lakhs figure
    // (the 4th-party table column is "Rs. in Lakhs", not the "Rs …/-" form).
    const manualValues: Record<string, string> = {
      'Party Name': party?.name ?? '',
      'Party Address': party?.address ?? '',
      'Party Phone': party?.phone ?? '',
      'Test Type': '',
      'AE Name Phone': '',
      'Estimate Lakhs': rawEstimate && !Number.isNaN(Number(rawEstimate)) ? Number(rawEstimate).toFixed(2) : rawEstimate
    }
    const row = { ...withComputedAmounts(rawRow), ...manualValues }
    const columns = table?.headers ?? []
    // Split off the manual labels so they map to themselves (their value key)
    // instead of being fuzzy-matched to a Works column (e.g. {{Party Name}}
    // must NOT land on "Name of the Agency").
    const manualResolved: PlaceholderMatch[] = labels
      .filter((l) => MANUAL_LABELS.has(l.trim().toLowerCase()))
      .map((label) => ({ label, column: label, score: 1 }))
    const otherLabels = labels.filter((l) => !MANUAL_LABELS.has(l.trim().toLowerCase()))
    if (otherLabels.length === 0 || columns.length === 0) {
      const resolved = [...otherLabels.map((label) => ({ label, column: null, score: 0 })), ...manualResolved]
      return { docx: await api.fillPlaceholdersInDocument(doc.docx, resolved, row), resolved }
    }
    let embeddings: { labelVectors: number[][]; columnVectors: number[][] } | undefined
    try {
      const [labelVectors, columnVectors] = await Promise.all([
        api.embedTexts(otherLabels),
        api.embedTexts(columns)
      ])
      embeddings = { labelVectors, columnVectors }
    } catch {
      // Neural matching unavailable — matchPlaceholdersToColumns falls back
      // to plain token overlap automatically when no embeddings are passed.
      embeddings = undefined
    }
    const resolved = [...matchPlaceholdersToColumns(otherLabels, columns, embeddings), ...manualResolved]
    return { docx: await api.fillPlaceholdersInDocument(doc.docx, resolved, row), resolved }
  }

  async function handlePreview(doc: CreatedDocument) {
    setGenBusy(true)
    setGenError(null)
    try {
      const result = await resolveForRow(doc)
      setPreview(result)
      requestAnimationFrame(() => {
        void (async () => {
          const container = previewRef.current
          if (!container) return
          container.innerHTML = ''
          await renderAsync(base64ToUint8(result.docx), container, undefined, DOCX_PREVIEW_OPTIONS)
          setPreviewPages(container.querySelectorAll('section.docx').length)
        })()
      })
    } catch (e) {
      setGenError(e instanceof Error ? e.message : String(e))
    } finally {
      setGenBusy(false)
    }
  }

  async function handlePrint(doc: CreatedDocument) {
    setGenBusy(true)
    setGenError(null)
    setGenNotice(null)
    try {
      const { docx } = await resolveForRow(doc)
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      container.innerHTML = ''
      await renderAsync(base64ToUint8(docx), container, undefined, DOCX_PREVIEW_OPTIONS)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setGenError(e instanceof Error ? e.message : String(e))
    } finally {
      setGenBusy(false)
    }
  }

  async function handleCreate(doc: CreatedDocument, formats: ('docx' | 'pdf')[]) {
    setGenBusy(true)
    setGenError(null)
    setGenNotice(null)
    try {
      const { docx } = await resolveForRow(doc)
      const res = await api.exportCreatedDocument(docx, doc.name, formats)
      setGenNotice(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setGenError(e instanceof Error ? e.message : String(e))
    } finally {
      setGenBusy(false)
    }
  }

  // The Issue button: generate in whichever output mode is selected.
  function handleIssue(doc: CreatedDocument) {
    if (outputMode === 'print') return handlePrint(doc)
    return handleCreate(doc, [outputMode === 'word' ? 'docx' : 'pdf'])
  }

  const expandedDoc = useMemo(() => documents.find((d) => d.id === expandedId) ?? null, [documents, expandedId])

  return (
    <>
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <section className="card">
        <div className="card-head">
          <div className="head-ic">
            <IconDoc />
          </div>
          <div className="titles">
            <h2>Document</h2>
            <p className="sub">Pick a saved document, choose a Works List row, and create the filled output.</p>
          </div>
        </div>

        {table && table.rows.length > 0 && (
          <div className="doc-work-picker">
            <span className="doc-work-picker-label">Work</span>
            <div className="doc-work-search">
              <IconSearch />
              <input
                type="text"
                placeholder="Search work by name…"
                value={workSearch}
                onChange={(e) => onWorkSearch(e.target.value)}
              />
              {workSearch && (
                <button className="tsearch-clear" onClick={() => onWorkSearch('')}>
                  Clear
                </button>
              )}
            </div>
            <select className="doc-work-select" value={rowIndex} onChange={(e) => setRowIndex(Number(e.target.value))}>
              {filteredRows.length === 0 ? (
                <option value={rowIndex} disabled>
                  No work matches “{workSearch}”
                </option>
              ) : (
                filteredRows.map(({ row, i }) => (
                  <option value={i} key={i}>
                    {rowLabel(row, table.headers, i)}
                  </option>
                ))
              )}
            </select>
          </div>
        )}

        {visibleCount === 0 ? (
          <div className="empty">
            <IconDoc />
            <p>No document templates available yet.</p>
          </div>
        ) : (
          <div className="doc-tile-grid">
            {documents.map((doc, i) =>
              !isDocForOffice(doc, office) ? null : (
              <div
                className={[
                  'doc-tile-card',
                  'tool-card',
                  TILE_TONES[i % TILE_TONES.length],
                  dragIndex === i ? 'dragging' : '',
                  overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
                ]
                  .filter(Boolean)
                  .join(' ')}
                key={doc.id}
                role="button"
                tabIndex={0}
                title={`Issue ${doc.name}`}
                onClick={() => toggleExpand(doc)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault()
                    toggleExpand(doc)
                  }
                }}
                draggable
                onDragStart={(e) => handleDragStart(e, i)}
                onDragOver={(e) => handleDragOver(e, i)}
                onDrop={(e) => handleDrop(e, i)}
                onDragEnd={handleDragEnd}
              >
                <DocThumbnail docx={doc.docx} />
                <span className="doc-tile-card-name">{doc.name}</span>
                <span className="doc-tile-card-meta">Added {doc.createdDate}</span>
                <span className="tool-card-cta">
                  <IconDoc /> Issue {doc.name}
                </span>
              </div>
            ))}
          </div>
        )}
      </section>

      {expandedDoc && !preview && (
        <div className="editor-overlay" onClick={() => toggleExpand(expandedDoc)}>
          <div className="confirm-modal gen-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <h3>{expandedDoc.name}</h3>
            <div className="gen-panel">
              {!table || table.rows.length === 0 ? (
                <button className="primary" onClick={onGoToWorksList}>
                  <IconPlus /> Add works
                </button>
              ) : (
                <>
                  <label className="gen-row-label">
                    Work:{' '}
                    <select value={rowIndex} onChange={(e) => setRowIndex(Number(e.target.value))}>
                      {table.rows.map((row, i) => (
                        <option value={i} key={i}>
                          {rowLabel(row, table.headers, i)}
                        </option>
                      ))}
                    </select>
                  </label>

                  {isPartyDoc(expandedDoc) &&
                    (() => {
                      const p = expandedDoc.id === PARTY_3RD_ID ? qcParties?.third : qcParties?.fourth
                      const which = expandedDoc.id === PARTY_3RD_ID ? '3rd-party QC agency' : '4th-party testing agency'
                      return (
                        <div className="gen-party-fields">
                          {p?.name?.trim() ? (
                            <p className="gen-party-hint">
                              {which}: <strong>{p.name}</strong> — from the Works List. Change it there if the agency
                              changed.
                            </p>
                          ) : (
                            <p className="gen-party-hint">
                              No {which} set yet. Add it once on the <strong>Works List</strong> page (below the office
                              details) and it fills here automatically.
                            </p>
                          )}
                        </div>
                      )
                    })()}

                  <div className="gen-output-modes" role="group" aria-label="Output">
                    <button
                      type="button"
                      className={outputMode === 'word' ? 'seg active' : 'seg'}
                      onClick={() => setOutputMode('word')}
                    >
                      <IconDownload /> Word
                    </button>
                    <button
                      type="button"
                      className={outputMode === 'pdf' ? 'seg active' : 'seg'}
                      onClick={() => setOutputMode('pdf')}
                    >
                      <IconDownload /> PDF
                    </button>
                    <button
                      type="button"
                      className={outputMode === 'print' ? 'seg active' : 'seg'}
                      onClick={() => setOutputMode('print')}
                    >
                      <IconPrint /> Print
                    </button>
                  </div>

                  <div className="gen-actions">
                    <button className="ghost" disabled={genBusy} onClick={() => handlePreview(expandedDoc)}>
                      <IconEye /> Preview
                    </button>
                    <button className="primary" disabled={genBusy} onClick={() => handleIssue(expandedDoc)}>
                      <IconDownload /> {genBusy ? 'Working…' : `Issue ${expandedDoc.name}`}
                    </button>
                  </div>

                  {genError && <div className="notice error">{genError}</div>}
                  {genNotice && <div className="notice">{genNotice}</div>}
                </>
              )}
            </div>
            <div className="confirm-actions">
              <button className="ghost" onClick={() => toggleExpand(expandedDoc)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {preview && expandedDoc && (
        <div className="editor-overlay" onClick={() => setPreview(null)}>
          <div className="editor-modal" onClick={(e) => e.stopPropagation()}>
            <div className="editor-head">
              <span className="editor-title">{expandedDoc.name} — preview</span>
              <div className="editor-head-actions">
                <button className="ghost" onClick={() => setPreview(null)}>
                  Close
                </button>
              </div>
            </div>
            <div className="doc-desk">
              <div className="doc-editor-wrap">
                <div ref={previewRef} className="docx-editor-canvas" />
                {previewPages > 1 && <span className="doc-page-badge">{previewPages} pages</span>}
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
