import { Fragment, memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { IconPlus, IconTrash, IconSearch, IconLink, IconChevronRight } from './Icons'
import type { ExcelTable } from '@core/types'

// Fixed column widths (px) — the sheet uses table-layout:fixed so column widths
// never depend on which rows happen to be rendered/scrolled. The table
// overflows and scrolls horizontally when the columns don't fit.
const COL_W = 160
const ROWNUM_W = 44
const ROWDEL_W = 44
// The work-name column gets a wider fixed width and wraps its text across
// lines (via a textarea, matched to the Monitoring Format list's read-only
// table) instead of scrolling off in a single-line input like every other
// column.
const NAME_COL_W = 340
const NAME_RE = /name\s*of\s*(the\s*)?work/i

// One Works List row, memoised so editing a single cell re-renders ONLY that
// row — not all 200+ rows × ~28 inputs. setCell replaces just the edited row's
// array (others keep their reference), and the callbacks below are stable, so
// React.memo's shallow prop compare skips every unchanged row.
interface SheetRowProps {
  ri: number
  row: string[]
  cols: number[]
  nameIndex: number
  onCell: (ri: number, ci: number, value: string) => void
  onDelete: (ri: number) => void
  flashed: boolean
  flashMessage?: string
  isFirstFlash: boolean
  firstFlashRef: React.RefObject<HTMLTableRowElement>
}
const SheetRow = memo(function SheetRow({
  ri,
  row,
  cols,
  nameIndex,
  onCell,
  onDelete,
  flashed,
  flashMessage,
  isFirstFlash,
  firstFlashRef
}: SheetRowProps) {
  return (
    <Fragment>
      <tr ref={isFirstFlash ? firstFlashRef : undefined} className={flashed ? 'row-flash' : ''}>
        <td className="rownum">{ri + 1}</td>
        {cols.map((ci) =>
          ci === nameIndex ? (
            <td key={ci}>
              <textarea className="cell-wrap" value={row[ci] ?? ''} onChange={(e) => onCell(ri, ci, e.target.value)} />
            </td>
          ) : (
            <td key={ci}>
              <input value={row[ci] ?? ''} onChange={(e) => onCell(ri, ci, e.target.value)} />
            </td>
          )
        )}
        <td className="rowdel">
          <button className="danger-ghost" title="Delete row" onClick={() => onDelete(ri)}>
            <IconTrash />
          </button>
        </td>
      </tr>
      {flashed && flashMessage && (
        <tr className="row-flash-msg">
          <td colSpan={cols.length + 2}>✓ {flashMessage}</td>
        </tr>
      )}
    </Fragment>
  )
})

interface Props {
  table: ExcelTable
  onChange: (updated: ExcelTable) => void
  /**
   * Optional per-row auto-fill run after a cell edit (Works List only): derives
   * blank Zone / Circle / Circle number from the row's own values. Given the
   * edited row as an object, returns it with those blanks filled.
   */
  autofillRow?: (row: Record<string, string>) => Record<string, string>
  /** Row indices (in table order) to briefly blink after an external update, e.g. "Update from L1". */
  flashRows?: number[]
  /** Message shown under each flashed row so the user can verify the right row was updated. */
  flashMessage?: string
  /** Rendered above the search bar, inside this same card — e.g. the Works List link-import row. */
  header?: React.ReactNode
  /** Always-visible buttons in the toolbar, after the works count — e.g. Clear / Delete database. */
  toolbarActions?: React.ReactNode
  /** Downloads exactly the rows the column filters leave (and without the deleted columns) as a table. Resolves when saved. */
  onDownloadFiltered?: (table: ExcelTable) => Promise<unknown> | void
}

const WIN_RE = /win/
const SERIAL_RE = /^(sl\.?\s*no|s\.?\s*no|sr\.?\s*no|serial)/i

function winIndex(hs: string[]): number {
  return hs.findIndex((h) => WIN_RE.test(h.toLowerCase().replace(/[^a-z]/g, '')))
}

function serialIndex(hs: string[]): number {
  const i = hs.findIndex((h) => SERIAL_RE.test(h.trim()))
  return i >= 0 ? i : -1
}

// Serial first, WIN CODE second — reorder the WIN CODE column to index 1.
function orderHeaders(hs: string[]): string[] {
  const order = [...hs]
  const wi = winIndex(order)
  if (wi > 1) {
    const [w] = order.splice(wi, 1)
    order.splice(1, 0, w)
  }
  return order
}

/**
 * The current Excel shown directly on the Data tab as an editable spreadsheet.
 * Rename headers, edit cells, and add/delete rows & columns — every valid edit
 * is committed straight back to the workspace (and persisted).
 */
export default function ExcelInline({
  table,
  onChange,
  autofillRow,
  flashRows,
  flashMessage,
  header,
  toolbarActions,
  onDownloadFiltered
}: Props) {
  const orderedHeaders = orderHeaders(table.headers)
  const [headers, setHeaders] = useState<string[]>(() => orderedHeaders)
  const [matrix, setMatrix] = useState<string[][]>(() => table.rows.map((r) => orderedHeaders.map((h) => r[h] ?? '')))
  const [query, setQuery] = useState('')
  // The link-import row (`header`) is rare and destructive, so it stays folded away
  // behind a toolbar button instead of pushing the table down the page.
  const [importOpen, setImportOpen] = useState(false)
  const nameIndex = useMemo(() => headers.findIndex((h) => NAME_RE.test(h)), [headers])

  // Column filters (one per column, opened from the funnel on the heading) and "delete columns" — the same toolbar
  // behaviour as the Monitoring Format table. Deleting a column here only hides it from this view and from the
  // filtered download; the database itself keeps it (use the trash button on a heading to really delete it).
  const [colMode, setColMode] = useState<'filter' | 'delete' | null>(null)
  const [colFilters, setColFilters] = useState<Record<number, string>>({})
  const [hiddenCols, setHiddenCols] = useState<string[]>([])
  const [filterCol, setFilterCol] = useState<{ col: number; x: number; y: number } | null>(null)
  const [downloading, setDownloading] = useState(false)
  const shownCols = useMemo(
    () => headers.map((_, i) => i).filter((i) => !hiddenCols.includes(headers[i])),
    [headers, hiddenCols]
  )
  const filtersOn = Object.values(colFilters).some((v) => v.trim() !== '')
  const isFilterable = (h: string) => !/^(sl\.?\s*no|s\.?\s*no|sr\.?\s*no)/i.test(h.trim())
  // A column with few distinct values filters from a list; any other one with a "contains" box.
  const colOptions = useMemo(
    () =>
      headers.map((h) => {
        if (NAME_RE.test(h) || /estimate|amount|ecv|emd|length|remark|address|details/i.test(h)) return null
        const values = new Set<string>()
        const ci = headers.indexOf(h)
        for (const row of matrix) {
          const v = (row[ci] ?? '').trim()
          if (v) values.add(v)
          if (values.size > 40) return null
        }
        return [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      }),
    [headers, matrix]
  )
  // Filter / Delete-columns mode switches itself off on a click anywhere else; the toggles, the heading buttons and
  // the filter pop-up (all data-mf-mode-ui) are left alone so several columns can be handled in a row.
  useEffect(() => {
    if (!colMode) return
    const off = (e: MouseEvent) => {
      if (!(e.target as Element | null)?.closest?.('[data-mf-mode-ui]')) {
        setColMode(null)
        setFilterCol(null)
      }
    }
    document.addEventListener('mousedown', off)
    return () => document.removeEventListener('mousedown', off)
  }, [colMode])
  useEffect(() => {
    if (!filterCol) return
    const close = () => setFilterCol(null)
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && close()
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
    }
  }, [filterCol])

  // Blink the just-updated rows, then settle: the "Updated from L1" message and
  // highlight clear after a few seconds so they don't linger over later edits.
  const flashSet = useMemo(() => new Set(flashRows ?? []), [flashRows])
  // Scroll to the LAST updated row, not the topmost one — flashRows is in the
  // order rows were actually matched/updated, so its last element is the most
  // recently updated row regardless of its position in the sheet.
  const firstFlashIndex = flashRows && flashRows.length > 0 ? flashRows[flashRows.length - 1] : -1
  const firstFlashRef = useRef<HTMLTableRowElement | null>(null)
  const [flashing, setFlashing] = useState(flashSet.size > 0)
  useEffect(() => {
    if (flashSet.size === 0) return
    setFlashing(true)
    firstFlashRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' })
    const t = setTimeout(() => setFlashing(false), 12000)
    return () => clearTimeout(t)
  }, [flashSet])

  // Persist the reordered column layout (WIN CODE → second) once on mount.
  useEffect(() => {
    if (orderedHeaders.join('\u0001') !== table.headers.join('\u0001')) {
      commit(orderedHeaders, matrix)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Filter rows across every cell, but keep each row's real matrix index so
  // edits and deletes still target the correct underlying row.
  const visible = useMemo(() => {
    const q = query.trim().toLowerCase()
    const active = Object.entries(colFilters).filter(([, v]) => v.trim() !== '')
    return matrix
      .map((row, ri) => ({ row, ri }))
      .filter(({ row }) => {
        for (const [c, v] of active) {
          const text = (row[Number(c)] ?? '').trim()
          const ok = colOptions[Number(c)] ? text === v : text.toLowerCase().includes(v.trim().toLowerCase())
          if (!ok) return false
        }
        return !q || row.some((c) => c.toLowerCase().includes(q))
      })
  }, [matrix, query, colFilters, colOptions])

  async function downloadFiltered() {
    if (!onDownloadFiltered || downloading) return
    const keep = shownCols
    const outHeaders = keep.map((i) => headers[i].trim())
    setDownloading(true)
    try {
      await onDownloadFiltered({
        ...table,
        headers: outHeaders,
        sourceHeaders: undefined,
        rows: visible.map(({ row }) => {
          const obj: Record<string, string> = {}
          keep.forEach((i, k) => {
            obj[outHeaders[k]] = row[i] ?? ''
          })
          return obj
        })
      })
    } finally {
      setDownloading(false)
    }
  }

  const error = useMemo(() => {
    const trimmed = headers.map((h) => h.trim())
    if (headers.length === 0) return 'Add at least one column.'
    if (trimmed.some((h) => h === '')) return 'Column names cannot be empty.'
    const seen = new Set<string>()
    for (const h of trimmed) {
      if (seen.has(h)) return `Duplicate column name: “${h}”.`
      seen.add(h)
    }
    return null
  }, [headers])

  // Push a valid edit back up to the workspace so it persists.
  function commit(hs: string[], m: string[][]) {
    const trimmed = hs.map((h) => h.trim())
    if (hs.length === 0 || trimmed.some((h) => h === '')) return
    const seen = new Set<string>()
    for (const h of trimmed) {
      if (seen.has(h)) return
      seen.add(h)
    }
    const rows = m.map((row) => {
      const obj: Record<string, string> = {}
      trimmed.forEach((h, ci) => {
        obj[h] = row[ci] ?? ''
      })
      return obj
    })
    onChange({ ...table, headers: trimmed, rows })
  }

  // Same validation/shape as commit, but for a single cell's edit — the by
  // far most frequent case, firing on every keystroke while typing into a
  // Works List cell. commit() rebuilds every row's Record<string,string>
  // from scratch every time it's called; for a few-hundred-row list that's
  // thousands of object writes PER KEYSTROKE just to change one row, and it
  // discards every OTHER row's previous object identity in the process —
  // this instead reuses table.rows' existing objects for every row except
  // the one that actually changed, rebuilding only that row. Falls back to
  // the general commit() if the row index isn't in the current table prop
  // yet (e.g. right after addRow, before the parent's state catches up) —
  // correctness over speed in that rare edge case.
  function commitRow(hs: string[], m: string[][], changedRi: number) {
    const trimmed = hs.map((h) => h.trim())
    if (hs.length === 0 || trimmed.some((h) => h === '')) return
    const seen = new Set<string>()
    for (const h of trimmed) {
      if (seen.has(h)) return
      seen.add(h)
    }
    if (changedRi < 0 || changedRi >= table.rows.length || changedRi >= m.length) {
      commit(hs, m)
      return
    }
    const changedRow: Record<string, string> = {}
    trimmed.forEach((h, ci) => {
      changedRow[h] = m[changedRi]?.[ci] ?? ''
    })
    const rows = table.rows.map((row, i) => (i === changedRi ? changedRow : row))
    onChange({ ...table, headers: trimmed, rows })
  }

  function setHeader(ci: number, value: string) {
    const next = headers.map((h, i) => (i === ci ? value : h))
    setHeaders(next)
    commit(next, matrix)
  }

  function setCell(ri: number, ci: number, value: string) {
    let next = matrix.map((row, i) => (i === ri ? row.map((c, j) => (j === ci ? value : c)) : row))
    // Live auto-fill for the Works List: derive blank Zone/Circle/Circle number
    // from the row the user just edited, and show the result in the grid at once
    // (this component holds its own matrix state, so the fill has to land here).
    if (autofillRow) {
      const trimmed = headers.map((h) => h.trim())
      const rowObj: Record<string, string> = {}
      trimmed.forEach((h, j) => {
        rowObj[h] = next[ri][j] ?? ''
      })
      const filled = autofillRow(rowObj)
      const filledRow = trimmed.map((h, j) => filled[h] ?? next[ri][j] ?? '')
      next = next.map((row, i) => (i === ri ? filledRow : row))
    }
    setMatrix(next)
    commitRow(headers, next, ri)
  }

  function addColumn() {
    let n = headers.length + 1
    const existing = new Set(headers)
    while (existing.has(`Column ${n}`)) n += 1
    const nextH = [...headers, `Column ${n}`]
    const nextM = matrix.length === 0 ? [['']] : matrix.map((row) => [...row, ''])
    setHeaders(nextH)
    setMatrix(nextM)
    commit(nextH, nextM)
  }

  function deleteColumn(ci: number) {
    const nextH = headers.filter((_, i) => i !== ci)
    const nextM = matrix.map((row) => row.filter((_, j) => j !== ci))
    setHeaders(nextH)
    setMatrix(nextM)
    commit(nextH, nextM)
  }

  function addRow() {
    const newRow = headers.map(() => '')
    const si = serialIndex(headers)
    if (si >= 0) {
      let max = 0
      for (const row of matrix) {
        const n = parseInt(String(row[si] ?? '').replace(/[^0-9]/g, ''), 10)
        if (!Number.isNaN(n) && n > max) max = n
      }
      newRow[si] = String(max + 1)
    }
    const nextM = [...matrix, newRow]
    setMatrix(nextM)
    commit(headers, nextM)
  }

  function deleteRow(ri: number) {
    const nextM = matrix.filter((_, i) => i !== ri)
    setMatrix(nextM)
    commit(headers, nextM)
  }

  // Stable handlers passed to the memoised rows: they always call the latest
  // setCell/deleteRow (via the ref) but keep the same identity across renders,
  // so a memoised row isn't forced to re-render just because its callback prop
  // changed. Without this, every edit would re-render all rows.
  const handlersRef = useRef({ setCell, deleteRow })
  handlersRef.current.setCell = setCell
  handlersRef.current.deleteRow = deleteRow
  const onCell = useCallback((ri: number, ci: number, value: string) => handlersRef.current.setCell(ri, ci, value), [])
  const onDelete = useCallback((ri: number) => handlersRef.current.deleteRow(ri), [])

  return (
    <section className="card sheet-inline">
      {error && <div className="notice warn editor-warn">{error}</div>}

      <div className="sheet-toolbar">
        <div className="sheet-search">
          <IconSearch />
          <input
            value={query}
            placeholder="Search the works database…"
            aria-label="Search the works database"
            onChange={(e) => setQuery(e.target.value)}
          />
          {query && (
            <button className="link-btn" onClick={() => setQuery('')}>
              Clear
            </button>
          )}
        </div>
        <span className="sheet-count" aria-live="polite">
          {query ? `${visible.length} of ${matrix.length} works` : `${matrix.length} works`}
        </span>
        <button
          type="button"
          data-mf-mode-ui
          className={`ghost${colMode === 'filter' ? ' mf-mode-on' : ''}${filtersOn ? ' mf-filter-applied' : ''}`}
          aria-pressed={colMode === 'filter' || filtersOn}
          onClick={() => {
            setFilterCol(null)
            // Clicking Filter again (while it is on, or while filters are applied) removes the filters.
            if (colMode === 'filter' || filtersOn) {
              setColFilters({})
              setColMode(null)
            } else {
              setColMode('filter')
            }
          }}
          title={
            colMode === 'filter' || filtersOn ? 'Remove the applied filters' : 'Show a filter on each column heading'
          }
        >
          Filter
        </button>
        <button
          type="button"
          data-mf-mode-ui
          className={`ghost${colMode === 'delete' ? ' mf-mode-on' : ''}`}
          aria-pressed={colMode === 'delete'}
          onClick={() => {
            setFilterCol(null)
            setColMode((m) => (m === 'delete' ? null : 'delete'))
          }}
          title="Show a delete × on each column heading (the column is left out of this view and of the filtered download)"
        >
          Delete columns
        </button>
        {onDownloadFiltered && (
          <button
            type="button"
            className="ghost"
            onClick={() => void downloadFiltered()}
            disabled={!filtersOn || downloading || visible.length === 0}
            title={filtersOn ? 'Download the works left after filtering' : 'Apply a column filter first'}
          >
            {downloading ? 'Preparing…' : 'Download Filtered list'}
          </button>
        )}
        {hiddenCols.length > 0 && (
          <button
            type="button"
            className="ghost"
            onClick={() => setHiddenCols([])}
            title="Bring back the deleted columns"
          >
            Restore columns ({hiddenCols.length})
          </button>
        )}
        {toolbarActions}
        {header && (
          <button
            type="button"
            className={`sheet-import-toggle${importOpen ? ' open' : ''}`}
            aria-expanded={importOpen}
            onClick={() => setImportOpen((o) => !o)}
          >
            <IconLink /> Import from link
            <IconChevronRight className="sheet-import-chev" />
          </button>
        )}
      </div>
      {header && importOpen && <div className="sheet-inline-header">{header}</div>}

      {filterCol && (
        <div
          data-mf-mode-ui
          className="mf-filter-pop"
          style={{ left: filterCol.x, top: filterCol.y }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          <div className="mf-filter-pop-title">{headers[filterCol.col]}</div>
          {colOptions[filterCol.col] ? (
            <div className="mf-filter-pop-list">
              {['', ...colOptions[filterCol.col]!].map((o) => (
                <button
                  type="button"
                  key={o || '__all'}
                  className={(colFilters[filterCol.col] ?? '') === o ? 'on' : ''}
                  onClick={() => {
                    setColFilters((p) => ({ ...p, [filterCol.col]: o }))
                    setFilterCol(null)
                  }}
                >
                  {o || 'All'}
                </button>
              ))}
            </div>
          ) : (
            <input
              autoFocus
              value={colFilters[filterCol.col] ?? ''}
              placeholder="Contains…"
              onChange={(e) => setColFilters((p) => ({ ...p, [filterCol.col]: e.target.value }))}
              onKeyDown={(e) => e.key === 'Enter' && setFilterCol(null)}
            />
          )}
        </div>
      )}

      <div className="sheet-wrap">
        <table
          className="sheet"
          style={{
            tableLayout: 'fixed',
            width:
              Math.max(
                ROWNUM_W +
                  shownCols.length * COL_W +
                  (shownCols.includes(nameIndex) ? NAME_COL_W - COL_W : 0) +
                  ROWDEL_W,
                0
              ) || undefined
          }}
        >
          <colgroup>
            <col style={{ width: ROWNUM_W }} />
            {shownCols.map((ci) => (
              <col key={ci} style={{ width: ci === nameIndex ? NAME_COL_W : COL_W }} />
            ))}
            <col style={{ width: ROWDEL_W }} />
          </colgroup>
          <thead>
            <tr>
              <th className="rownum">#</th>
              {shownCols.map((ci) => {
                const h = headers[ci]
                return (
                  <th key={ci} className={`mf-th${(colFilters[ci] ?? '').trim() ? ' mf-th-active' : ''}`}>
                    <div className="col-head">
                      <input
                        className="col-name"
                        value={h}
                        placeholder={`Column ${ci + 1}`}
                        onChange={(e) => setHeader(ci, e.target.value)}
                      />
                      <span className="mf-th-tools">
                        {colMode === 'delete' && (
                          <button
                            type="button"
                            data-mf-mode-ui
                            className="mf-th-delete"
                            title={`Delete the ${h} column (left out of this view and the filtered download)`}
                            aria-label={`Delete the ${h} column`}
                            onMouseDown={(e) => e.stopPropagation()}
                            onClick={() => {
                              setColFilters({})
                              setFilterCol(null)
                              setHiddenCols((p) => [...p, h])
                            }}
                          >
                            ×
                          </button>
                        )}
                        {colMode === 'filter' && isFilterable(h) && (
                          <button
                            type="button"
                            data-mf-mode-ui
                            className={`mf-th-filter${(colFilters[ci] ?? '').trim() ? ' on' : ''}${filterCol?.col === ci ? ' open' : ''}`}
                            title={`Filter ${h}`}
                            onMouseDown={(e) => e.stopPropagation()}
                            onClick={(e) => {
                              const r = e.currentTarget.getBoundingClientRect()
                              setFilterCol(
                                filterCol?.col === ci
                                  ? null
                                  : { col: ci, x: Math.min(r.left, window.innerWidth - 260), y: r.bottom + 4 }
                              )
                            }}
                          >
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                              <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z" />
                            </svg>
                          </button>
                        )}
                      </span>
                      <button className="col-del" title="Delete column" onClick={() => deleteColumn(ci)}>
                        <IconTrash />
                      </button>
                    </div>
                  </th>
                )
              })}
              <th className="rowdel">
                <button className="col-add" title="Add column" onClick={addColumn}>
                  <IconPlus />
                </button>
              </th>
            </tr>
          </thead>
          <tbody>
            {visible.map(({ row, ri }) => (
              <SheetRow
                key={ri}
                ri={ri}
                row={row}
                cols={shownCols}
                nameIndex={nameIndex}
                onCell={onCell}
                onDelete={onDelete}
                flashed={flashing && flashSet.has(ri)}
                flashMessage={flashMessage}
                isFirstFlash={ri === firstFlashIndex}
                firstFlashRef={firstFlashRef}
              />
            ))}
            {visible.length === 0 && (
              <tr>
                <td className="sheet-empty" colSpan={shownCols.length + 2}>
                  No rows match the search / filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="sheet-foot">
        <button onClick={addRow}>
          <IconPlus /> Add row
        </button>
        <button onClick={addColumn}>
          <IconPlus /> Add column
        </button>
        <span className="hint">Rename headers to change field names. Rows here drive how many documents generate.</span>
      </div>
    </section>
  )
}
