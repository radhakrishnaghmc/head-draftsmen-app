import { useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { findPlaceholders, matchPlaceholdersToColumns, fillDocumentHtml } from '@core/createDocument'
import type { PlaceholderMatch } from '@core/createDocument'
import { withComputedAmounts } from '@core/worksAmounts'
import type { CreatedDocument, ExcelTable } from '@core/types'
import { IconDoc, IconTrash, IconEye, IconPrint, IconDownload, IconCheck, IconWarn, IconPlus } from './Icons'
import { pageShellStyle, usePageScrollTracker, PAGE_WIDTH } from './docPage'
import DocThumbnail from './DocThumbnail'

interface Props {
  tables: ExcelTable[]
  documents: CreatedDocument[]
  onChange: (docs: CreatedDocument[]) => void
  /** Sends a document to Create New Document for edits. */
  onEdit: (doc: CreatedDocument) => void
  /** Switches to the Works List tab — offered when there are no rows to pick from. */
  onGoToWorksList: () => void
}

function rowLabel(row: Record<string, string>, headers: string[], index: number): string {
  const nameCol = headers.find((h) => /name of (the )?work/i.test(h)) ?? headers[0]
  const v = nameCol ? (row[nameCol] ?? '').trim() : ''
  return v ? `${index + 1}. ${v}` : `Row ${index + 1}`
}

const TILE_TONES = ['tone-indigo', 'tone-sky', 'tone-rose', 'tone-amber', 'tone-teal', 'tone-green']

export default function PrintDocumentTab({ tables, documents, onChange, onEdit, onGoToWorksList }: Props) {
  const table = tables[0] ?? null

  const [pendingDelete, setPendingDelete] = useState<CreatedDocument | null>(null)
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)

  const [expandedId, setExpandedId] = useState<string | null>(null)
  const [rowIndex, setRowIndex] = useState(0)
  const [wantDocx, setWantDocx] = useState(true)
  const [wantPdf, setWantPdf] = useState(false)
  const [genBusy, setGenBusy] = useState(false)
  const [genError, setGenError] = useState<string | null>(null)
  const [genNotice, setGenNotice] = useState<string | null>(null)
  const [preview, setPreview] = useState<{ html: string; resolved: PlaceholderMatch[] } | null>(null)
  const previewFrameRef = useRef<HTMLIFrameElement>(null)
  const { current: previewPage, total: previewPages } = usePageScrollTracker(previewFrameRef, preview?.html)

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
    setRowIndex(0)
  }

  async function resolveForRow(doc: CreatedDocument): Promise<{ html: string; resolved: PlaceholderMatch[] }> {
    const labels = findPlaceholders(doc.html)
    // Amount-bearing columns (Amount of estimate, Estimate Amount ECV, EMD @
    // 1%/1.5%, ASD, Contract Amount) resolve to their computed, Indian-
    // formatted "Rs 1,00,000/-" value rather than the raw spreadsheet figure
    // — see core/worksAmounts.ts.
    const row = withComputedAmounts(table?.rows[rowIndex] ?? {})
    const columns = table?.headers ?? []
    if (labels.length === 0 || columns.length === 0) {
      const resolved = labels.map((label) => ({ label, column: null, score: 0 }))
      return { html: fillDocumentHtml(doc.html, resolved, row), resolved }
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
    return { html: fillDocumentHtml(doc.html, resolved, row), resolved }
  }

  async function handlePreview(doc: CreatedDocument) {
    setGenBusy(true)
    setGenError(null)
    try {
      const result = await resolveForRow(doc)
      setPreview(result)
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
      const { html } = await resolveForRow(doc)
      await api.printCreatedDocument(html)
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
      const { html } = await resolveForRow(doc)
      const res = await api.exportCreatedDocument(html, doc.name, formats)
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

        {documents.length === 0 ? (
          <div className="empty">
            <IconDoc />
            <p>Nothing added yet — paste a document on the Create New Document tab first.</p>
          </div>
        ) : (
          <div className="doc-tile-grid">
            {documents.map((doc, i) => (
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
                <button className="doc-tile-edit-btn" title="Edit Document" onClick={() => onEdit(doc)}>
                  EDIT
                </button>
                <DocThumbnail html={doc.html} />
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
              <div className="doc-editor-wrap" style={{ width: PAGE_WIDTH }}>
                <iframe
                  ref={previewFrameRef}
                  className="doc-editor"
                  title="Filled document preview"
                  srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>${pageShellStyle()}</style></head><body>${preview.html}</body></html>`}
                />
                {previewPages > 1 && (
                  <span className="doc-page-badge">
                    Page {previewPage} of {previewPages}
                  </span>
                )}
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
