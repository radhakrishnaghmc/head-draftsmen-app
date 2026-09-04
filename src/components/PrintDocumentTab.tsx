import { useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import type { PlaceholderMatch } from '@core/createDocument'
import { withComputedAmounts } from '@core/worksAmounts'
import type { CreatedDocument, ExcelTable, QcOfficeParties } from '@core/types'
import type { Office } from '../office'
import { corporationByName } from '../zoneCircleDirectory'
import { IconDoc, IconEye, IconPrint, IconDownload, IconSearch } from './Icons'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview } from './docPage'
import { closeOnBackdropMouseDown } from '../overlayClose'
import DocThumbnail from './DocThumbnail'
import type { VerifyDocItem } from '../../electron/ipc-contract'
import { VerifyButton } from './VerifyButton'

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
  'estimate lakhs',
  'ts no and date',
  'corporation',
  'corp full',
  'corp full caps'
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

  // -1 = no work chosen: documents fill with the office details only (the work's
  // Works-List columns stay blank). A work is picked explicitly from the dropdown.
  const [rowIndex, setRowIndex] = useState(-1)
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
    // Leave the selection on "no work" until the user actively picks one — don't
    // auto-select a row just because they typed in the search box.
    if (rowIndex < 0) return
    const q = v.trim().toLowerCase()
    const matches = table.rows.map((row, i) => ({ row, i })).filter(({ row }) => !q || workNameOf(row).toLowerCase().includes(q))
    if (matches.length > 0 && !matches.some((m) => m.i === rowIndex)) setRowIndex(matches[0].i)
  }
  // Batch actions — issue several documents at once against the selected work.
  // Which documents: the ticked ones, or (when none are ticked) all shown for
  // this office. Output: Word/PDF (saved together into one folder) or Print.
  const [batchFormat, setBatchFormat] = useState<'word' | 'pdf' | 'print'>('word')
  const [batchBusy, setBatchBusy] = useState(false)
  const [batchError, setBatchError] = useState<string | null>(null)
  const [batchNotice, setBatchNotice] = useState<string | null>(null)
  // Per-document selection (tile checkboxes) for the batch action.
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const toggleSelected = (id: string) =>
    setSelectedIds((prev) => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  const selectedCount = useMemo(
    () => documents.filter((d) => isDocForOffice(d, office) && selectedIds.has(d.id)).length,
    [documents, office, selectedIds]
  )
  const allSelected = visibleCount > 0 && selectedCount === visibleCount
  function toggleSelectAll() {
    setSelectedIds(allSelected ? new Set() : new Set(documents.filter((d) => isDocForOffice(d, office)).map((d) => d.id)))
  }
  const [preview, setPreview] = useState<{ docx: string; resolved: PlaceholderMatch[] } | null>(null)
  const [previewPages, setPreviewPages] = useState(0)
  // Title shown on the preview modal — one document's name, or "N documents".
  const [previewTitle, setPreviewTitle] = useState('')
  const previewRef = useRef<HTMLDivElement>(null)
  // Works List column headers don't change between one preview and the
  // next — without this, resolveForRow re-embedded the same ~20-40 headers
  // on every single Preview click, on top of the doc's own (much shorter)
  // placeholder labels, roughly doubling per-click latency for no reason.
  const columnEmbedCacheRef = useRef<{ key: string; vectors: number[][] } | null>(null)
  // Off-screen render target used to turn filled .docx files into HTML — for
  // batch printing (handleBatch 'print') and for building the combined preview
  // (handleBatchPreview). printCreatedDocument needs plain HTML, not a docx
  // buffer, so printing goes through the OS print dialog against a temp HTML
  // file rather than requiring LibreOffice the way PDF export does.
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

  // useAi: whether to call out to the local embedding model for placeholder↔
  // column matching. Preview (handlePreview/handleBatchPreview) passes
  // false — Agreement/Intimation's previews are instant because their
  // placeholders are matched against a fixed, developer-known mapping, never
  // touching the AI model; Issue Documents' templates are arbitrary
  // user-uploaded .docx files with unknown placeholder names, so token-
  // overlap (matchPlaceholdersToColumns's built-in fallback with no
  // embeddings passed) is the only fast, synchronous option — same one this
  // function already falls back to on any embedding error. The real
  // Download/Print output (handleBatch) still opts into the AI match for
  // better accuracy, since that's a deliberate, less latency-sensitive action.
  async function resolveForRow(
    doc: CreatedDocument,
    useAi = true
  ): Promise<{ docx: string; resolved: PlaceholderMatch[]; values: Record<string, string> }> {
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
    // Technical Sanction No & Date combined into one placeholder so the " dt. "
    // separator only appears when there's something to separate — a blank work
    // (no selection) leaves the whole field blank instead of a stray "dt.".
    const tsNo = (rawRow['Technical Sanc No'] ?? '').trim()
    const tsDate = (rawRow['TS date'] ?? '').trim()
    // Corporation abbreviation/full name — several of these bundled templates
    // (e.g. Action Taken Report, Completion Report) had "CYBERABAD MUNICIPAL
    // CORPORATION" hard-coded into their letterhead; now the office's actual
    // corporation, same as the Work Order/Agreement tab's placeholders.
    const corpFull = corporationByName(office.corporation)?.fullName ?? ''
    const manualValues: Record<string, string> = {
      'Party Name': party?.name ?? '',
      'Party Address': party?.address ?? '',
      'Party Phone': party?.phone ?? '',
      'Test Type': '',
      'AE Name Phone': '',
      'Estimate Lakhs': rawEstimate && !Number.isNaN(Number(rawEstimate)) ? Number(rawEstimate).toFixed(2) : rawEstimate,
      'TS No and Date': tsNo && tsDate ? `${tsNo} dt. ${tsDate}` : tsNo || tsDate,
      Corporation: office.corporation ?? '',
      'Corp Full': corpFull,
      'Corp Full Caps': corpFull.toUpperCase()
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
      return { docx: await api.fillPlaceholdersInDocument(doc.docx, resolved, row), resolved, values: row }
    }
    let embeddings: { labelVectors: number[][]; columnVectors: number[][] } | undefined
    if (useAi) {
      try {
        const columnsKey = columns.join('␟')
        const cached = columnEmbedCacheRef.current
        const [labelVectors, columnVectors] =
          cached?.key === columnsKey
            ? [await api.embedTexts(otherLabels), cached.vectors]
            : await Promise.all([api.embedTexts(otherLabels), api.embedTexts(columns)])
        if (cached?.key !== columnsKey) columnEmbedCacheRef.current = { key: columnsKey, vectors: columnVectors }
        embeddings = { labelVectors, columnVectors }
      } catch {
        // Neural matching unavailable — matchPlaceholdersToColumns falls back
        // to plain token overlap automatically when no embeddings are passed.
        embeddings = undefined
      }
    }
    const resolved = [...matchPlaceholdersToColumns(otherLabels, columns, embeddings), ...manualResolved]
    return { docx: await api.fillPlaceholdersInDocument(doc.docx, resolved, row), resolved, values: row }
  }

  // Preview one document (tile click) — fills it against the selected work
  // (office details only when no work is chosen) and renders it read-only.
  async function handlePreview(doc: CreatedDocument) {
    setBatchBusy(true)
    setBatchError(null)
    setPreviewTitle(doc.name)
    try {
      const result = await resolveForRow(doc, false)
      setPreview(result)
      requestAnimationFrame(() => {
        void (async () => {
          const container = previewRef.current
          if (!container) return
          const { pageCount } = await renderDocPreview(base64ToUint8(result.docx), container)
          setPreviewPages(pageCount)
        })()
      })
    } catch (e) {
      setBatchError(e instanceof Error ? e.message : String(e))
    } finally {
      setBatchBusy(false)
    }
  }

  // Preview the batch target (ticked documents, or all) together in one modal,
  // each rendered with a page break between — same set the Download/Print acts on.
  async function handleBatchPreview() {
    const docs = batchTargets()
    if (docs.length === 0) {
      setBatchError('No documents to preview for this office.')
      return
    }
    setBatchBusy(true)
    setBatchError(null)
    setBatchNotice(null)
    try {
      const scratch = printScratchRef.current
      if (!scratch) throw new Error('Preview failed to initialize.')
      const parts: string[] = []
      for (const doc of docs) {
        const { docx } = await resolveForRow(doc, false)
        await renderDocPreview(base64ToUint8(docx), scratch)
        parts.push(scratch.innerHTML)
      }
      const combined = parts.join('<div style="page-break-after:always"></div>')
      setPreviewTitle(docs.length === 1 ? docs[0].name : `${docs.length} documents`)
      setPreview({ docx: '', resolved: [] })
      requestAnimationFrame(() => {
        const c = previewRef.current
        if (!c) return
        c.innerHTML = combined
        setPreviewPages(c.querySelectorAll('section.docx').length)
      })
    } catch (e) {
      setBatchError(e instanceof Error ? e.message : String(e))
    } finally {
      setBatchBusy(false)
    }
  }

  // Which documents the batch acts on: the ticked ones, or — when none are
  // ticked — every document shown for this office.
  function batchTargets(): CreatedDocument[] {
    const visible = documents.filter((d) => isDocForOffice(d, office))
    const picked = visible.filter((d) => selectedIds.has(d.id))
    return picked.length > 0 ? picked : visible
  }

  // Batch action: fill each target document against the selected work (office
  // details only when no work is chosen), then either save them all into ONE
  // folder (Word / PDF) or send them to the printer as a single job (Print).
  async function handleBatch(format: 'word' | 'pdf' | 'print') {
    const docs = batchTargets()
    if (docs.length === 0) {
      setBatchError('No documents to issue for this office.')
      return
    }
    setBatchBusy(true)
    setBatchError(null)
    setBatchNotice(null)
    try {
      if (format === 'print') {
        // Render every selected document to HTML and print them as one job, with
        // a hard page break between documents.
        const container = printScratchRef.current
        if (!container) throw new Error('Print failed to initialize.')
        const parts: string[] = []
        for (const doc of docs) {
          const { docx } = await resolveForRow(doc)
          await renderDocPreview(base64ToUint8(docx), container)
          parts.push(container.innerHTML)
        }
        const combined = parts.join('<div style="page-break-after:always"></div>')
        await api.printCreatedDocument(combined)
        setBatchNotice(`Sent ${docs.length} document(s) to the printer.`)
        return
      }
      const files: { name: string; bytes: Uint8Array }[] = []
      for (const doc of docs) {
        const { docx } = await resolveForRow(doc)
        if (format === 'word') {
          files.push({ name: doc.name, bytes: base64ToUint8(docx) })
        } else {
          const pdf = await api.docxToPdf(base64ToUint8(docx))
          files.push({ name: doc.name, bytes: pdf })
        }
      }
      const res = format === 'word' ? await api.saveDocxsToFolder(files) : await api.savePdfsToFolder(files)
      setBatchNotice(res ? `Saved ${res.files.length} document(s) to ${res.dir}` : 'Cancelled.')
    } catch (e) {
      setBatchError(e instanceof Error ? e.message : String(e))
    } finally {
      setBatchBusy(false)
    }
  }

  // Builds one VerifyDocItem per target document — these are arbitrary
  // user-uploaded templates with unknown placeholder names, so unlike the
  // fixed Intimation/Agreement documents there's no known EMD/ASD/corporation
  // wiring to cross-check; only placeholder coverage is checked (every
  // matched placeholder resolved to a non-blank value that actually made it
  // into the document text).
  async function getVerifyItems(): Promise<VerifyDocItem[]> {
    const items: VerifyDocItem[] = []
    for (const doc of batchTargets()) {
      const { docx, resolved, values } = await resolveForRow(doc)
      const labelValues: Record<string, string> = {}
      const requiredLabels: string[] = []
      for (const match of resolved) {
        if (!match.column) continue
        labelValues[match.label] = values[match.column] ?? ''
        requiredLabels.push(match.label)
      }
      items.push({ name: doc.name, docxBase64: docx, values: labelValues, requiredLabels })
    }
    return items
  }

  return (
    <>
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <section className="card">
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
              <option value={-1}>— No work (office details only) —</option>
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

        <div className="card-head">
          {visibleCount > 0 && (
            <div className="doc-batch-actions">
              <div className="gen-output-modes" role="group" aria-label="Batch output format">
                <button
                  type="button"
                  className={batchFormat === 'word' ? 'seg active' : 'seg'}
                  onClick={() => setBatchFormat('word')}
                >
                  <IconDownload /> Word
                </button>
                <button
                  type="button"
                  className={batchFormat === 'pdf' ? 'seg active' : 'seg'}
                  onClick={() => setBatchFormat('pdf')}
                >
                  <IconDownload /> PDF
                </button>
                <button
                  type="button"
                  className={batchFormat === 'print' ? 'seg active' : 'seg'}
                  onClick={() => setBatchFormat('print')}
                >
                  <IconPrint /> Print
                </button>
              </div>
              <button className="ghost" disabled={batchBusy} onClick={handleBatchPreview}>
                <IconEye /> Preview
              </button>
              <button className="primary" disabled={batchBusy} onClick={() => handleBatch(batchFormat)}>
                {batchFormat === 'print' ? <IconPrint /> : <IconDownload />}{' '}
                {batchBusy
                  ? 'Working…'
                  : `${batchFormat === 'print' ? 'Print' : 'Download'} ${
                      selectedCount > 0 ? `selected (${selectedCount})` : 'all'
                    }`}
              </button>
            </div>
          )}
        </div>

        {(batchError || batchNotice) && (
          <div className={`notice ${batchError ? 'error' : 'ok'}`}>{batchError ?? batchNotice}</div>
        )}

        {visibleCount > 0 && (
          <div className="doc-select-bar">
            <label className="doc-select-all">
              <input type="checkbox" checked={allSelected} onChange={toggleSelectAll} />
              Select all
            </label>
            <span className="doc-select-hint">
              {selectedCount > 0
                ? `${selectedCount} selected — the button above prints/downloads only these`
                : 'Tick documents to print or download only those (none ticked = all)'}
            </span>
            <VerifyButton getItems={getVerifyItems} />
            {selectedCount > 0 && (
              <button className="ghost" onClick={() => setSelectedIds(new Set())}>
                Clear
              </button>
            )}
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
                  selectedIds.has(doc.id) ? 'selected' : '',
                  dragIndex === i ? 'dragging' : '',
                  overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
                ]
                  .filter(Boolean)
                  .join(' ')}
                key={doc.id}
                role="button"
                aria-pressed={selectedIds.has(doc.id)}
                tabIndex={0}
                title={`Select ${doc.name}`}
                onClick={() => toggleSelected(doc.id)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault()
                    toggleSelected(doc.id)
                  }
                }}
                draggable
                onDragStart={(e) => handleDragStart(e, i)}
                onDragOver={(e) => handleDragOver(e, i)}
                onDrop={(e) => handleDrop(e, i)}
                onDragEnd={handleDragEnd}
              >
                <label
                  className="doc-tile-check"
                  title="Select for batch print / download"
                  onClick={(e) => e.stopPropagation()}
                >
                  <input
                    type="checkbox"
                    checked={selectedIds.has(doc.id)}
                    onChange={() => toggleSelected(doc.id)}
                  />
                </label>
                <DocThumbnail docx={doc.docx} />
                <span className="doc-tile-card-name">{doc.name}</span>
                <span className="doc-tile-card-meta">Added {doc.createdDate}</span>
                <span
                  className="tool-card-cta doc-preview-cta"
                  title={`Preview ${doc.name}`}
                  onClick={(e) => {
                    e.stopPropagation()
                    void handlePreview(doc)
                  }}
                >
                  <IconEye /> Preview
                </span>
              </div>
            ))}
          </div>
        )}
      </section>

      {preview &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setPreview(null))}>
            <div className="editor-modal" onClick={(e) => e.stopPropagation()}>
              <div className="editor-head">
                <span className="editor-title">{previewTitle} — preview</span>
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
          </div>,
          document.body
        )}
    </>
  )
}
