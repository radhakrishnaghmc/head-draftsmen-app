import { useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import type { PlaceholderMatch } from '@core/createDocument'
import { withComputedAmounts } from '@core/worksAmounts'
import type { CreatedDocument, ExcelTable } from '@core/types'
import type { Office } from '../office'
import { IconDoc, IconTrash, IconEye, IconPrint, IconDownload, IconCheck, IconWarn, IconPlus, IconSearch } from './Icons'
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

export default function PrintDocumentTab({ tables, documents, onChange, onGoToWorksList, office }: Props) {
  const table = tables[0] ?? null

  // Reorder/delete still act on the full synced list; only display is filtered.
  const visibleCount = useMemo(() => documents.filter((d) => isDocForOffice(d, office)).length, [documents, office])

  const [pendingDelete, setPendingDelete] = useState<CreatedDocument | null>(null)
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
  const [wantDocx, setWantDocx] = useState(true)
  const [wantPdf, setWantPdf] = useState(false)
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

  function confirmDelete() {
    if (!pendingDelete) return
    onChange(documents.filter((d) => d.id !== pendingDelete.id))
    if (expandedId === pendingDelete.id) setExpandedId(null)
    setPendingDelete(null)
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
    const row = withComputedAmounts(table?.rows[rowIndex] ?? {})
    const columns = table?.headers ?? []
    if (labels.length === 0 || columns.length === 0) {
      const resolved = labels.map((label) => ({ label, column: null, score: 0 }))
      return { docx: await api.fillPlaceholdersInDocument(doc.docx, resolved, row), resolved }
    }
    let embeddings: { labelVectors: number[][]; columnVectors: number[][] } | undefined
    try {
      const [labelVectors, columnVectors] = await Promise.all([
        api.embedTexts(labels),
        api.embedTexts(columns)
      ])
      embeddings = { labelVectors, columnVectors }
    } catch {
      // Neural matching unavailable — matchPlaceholdersToColumns falls back
      // to plain token overlap automatically when no embeddings are passed.
      embeddings = undefined
    }
    const resolved = matchPlaceholdersToColumns(labels, columns, embeddings)
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

  async function handleCreate(doc: CreatedDocument) {
    const formats: ('docx' | 'pdf')[] = [...(wantDocx ? (['docx'] as const) : []), ...(wantPdf ? (['pdf'] as const) : [])]
    if (formats.length === 0) {
      setGenError('Pick Word and/or PDF before creating the document.')
      return
    }
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
                  TILE_TONES[i % TILE_TONES.length],
                  dragIndex === i ? 'dragging' : '',
                  overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
                ]
                  .filter(Boolean)
                  .join(' ')}
                key={doc.id}
                draggable
                onDragStart={(e) => handleDragStart(e, i)}
                onDragOver={(e) => handleDragOver(e, i)}
                onDrop={(e) => handleDrop(e, i)}
                onDragEnd={handleDragEnd}
              >
                <DocThumbnail docx={doc.docx} />
                <div className="doc-tile-card-name" title={doc.name}>
                  {doc.name}
                </div>
                <div className="doc-tile-card-meta">Added {doc.createdDate}</div>
                <div className="doc-tile-card-actions">
                  <button
                    className="doc-tile-issue-btn"
                    title={`Issue ${doc.name}`}
                    onClick={() => toggleExpand(doc)}
                  >
                    Issue {doc.name}
                  </button>
                  <button className="danger-ghost" title="Remove" onClick={() => setPendingDelete(doc)}>
                    <IconTrash />
                  </button>
                </div>
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

                  <div className="gen-formats">
                    <label>
                      <input type="checkbox" checked={wantDocx} onChange={(e) => setWantDocx(e.target.checked)} />{' '}
                      Word (.docx)
                    </label>
                    <label>
                      <input type="checkbox" checked={wantPdf} onChange={(e) => setWantPdf(e.target.checked)} />{' '}
                      PDF
                    </label>
                  </div>

                  <div className="gen-actions">
                    <button className="ghost" disabled={genBusy} onClick={() => handlePreview(expandedDoc)}>
                      <IconEye /> Preview
                    </button>
                    <button className="ghost" disabled={genBusy} onClick={() => handlePrint(expandedDoc)}>
                      <IconPrint /> Print
                    </button>
                    <button className="primary" disabled={genBusy} onClick={() => handleCreate(expandedDoc)}>
                      <IconDownload /> {genBusy ? 'Working…' : 'Create Document'}
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
            <div className="gen-match-summary">
              {preview.resolved.map((r) => (
                <span className={`tag ${r.column ? 'green' : 'rose'}`} key={r.label}>
                  {r.column ? <IconCheck /> : <IconWarn />} {'{{' + r.label + '}}'}
                  {r.column ? ` → ${r.column}` : ' (unresolved)'}
                </span>
              ))}
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

      {pendingDelete && (
        <div className="editor-overlay" onClick={() => setPendingDelete(null)}>
          <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <div className="confirm-ic">
              <IconTrash />
            </div>
            <h3>Delete this document?</h3>
            <p className="confirm-warn">
              You're about to permanently remove <strong>{pendingDelete.name}</strong>.
            </p>
            <p className="confirm-hint">Once deleted, this cannot be recovered.</p>
            <div className="confirm-actions">
              <button className="ghost" onClick={() => setPendingDelete(null)}>
                Cancel
              </button>
              <button className="danger" onClick={confirmDelete}>
                Delete
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
