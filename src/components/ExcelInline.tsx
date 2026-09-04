import { Fragment, memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { IconPlus, IconTrash, IconSearch } from './Icons'
import type { ExcelTable } from '@core/types'

// Fixed column widths (px) — the sheet uses table-layout:fixed so column widths
// never depend on which rows happen to be rendered/scrolled. The table
// overflows and scrolls horizontally when the columns don't fit.
const COL_W = 160
const ROWNUM_W = 44
const ROWDEL_W = 44

// One Works List row, memoised so editing a single cell re-renders ONLY that
// row — not all 200+ rows × ~28 inputs. setCell replaces just the edited row's
// array (others keep their reference), and the callbacks below are stable, so
// React.memo's shallow prop compare skips every unchanged row.
interface SheetRowProps {
  ri: number
  row: string[]
  headerCount: number
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
  headerCount,
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
        {Array.from({ length: headerCount }, (_, ci) => (
          <td key={ci}>
            <input value={row[ci] ?? ''} onChange={(e) => onCell(ri, ci, e.target.value)} />
          </td>
        ))}
        <td className="rowdel">
          <button className="danger-ghost" title="Delete row" onClick={() => onDelete(ri)}>
            <IconTrash />
          </button>
        </td>
      </tr>
      {flashed && flashMessage && (
        <tr className="row-flash-msg">
          <td colSpan={headerCount + 2}>✓ {flashMessage}</td>
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
export default function ExcelInline({ table, onChange, autofillRow, flashRows, flashMessage, header }: Props) {
  const orderedHeaders = orderHeaders(table.headers)
  const [headers, setHeaders] = useState<string[]>(() => orderedHeaders)
  const [matrix, setMatrix] = useState<string[][]>(() =>
    table.rows.map((r) => orderedHeaders.map((h) => r[h] ?? ''))
  )
  const [query, setQuery] = useState('')

  // Blink the just-updated rows, then settle: the "Updated from L1" message and
  // highlight clear after a few seconds so they don't linger over later edits.
  const flashSet = useMemo(() => new Set(flashRows ?? []), [flashRows])
  const firstFlashIndex = flashSet.size > 0 ? Math.min(...flashSet) : -1
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
    const indexed = matrix.map((row, ri) => ({ row, ri }))
    if (!q) return indexed
    return indexed.filter(({ row }) => row.some((c) => c.toLowerCase().includes(q)))
  }, [matrix, query])

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
      {header && <div className="sheet-inline-header">{header}</div>}
      {error && <div className="notice warn editor-warn">{error}</div>}

      <div className="sheet-search">
        <IconSearch />
        <input
          value={query}
          placeholder="Search the works database…"
          onChange={(e) => setQuery(e.target.value)}
        />
        {query && (
          <>
            <span className="sheet-search-count">
              {visible.length} of {matrix.length}
            </span>
            <button className="link-btn" onClick={() => setQuery('')}>
              Clear
            </button>
          </>
        )}
      </div>

      <div className="sheet-wrap">
        <table
          className="sheet"
          style={{ tableLayout: 'fixed', width: Math.max(ROWNUM_W + headers.length * COL_W + ROWDEL_W, 0) || undefined }}
        >
          <colgroup>
            <col style={{ width: ROWNUM_W }} />
            {headers.map((_, ci) => (
              <col key={ci} style={{ width: COL_W }} />
            ))}
            <col style={{ width: ROWDEL_W }} />
          </colgroup>
          <thead>
            <tr>
              <th className="rownum">#</th>
              {headers.map((h, ci) => (
                <th key={ci}>
                  <div className="col-head">
                    <input
                      className="col-name"
                      value={h}
                      placeholder={`Column ${ci + 1}`}
                      onChange={(e) => setHeader(ci, e.target.value)}
                    />
                    <button className="col-del" title="Delete column" onClick={() => deleteColumn(ci)}>
                      <IconTrash />
                    </button>
                  </div>
                </th>
              ))}
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
                headerCount={headers.length}
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
                <td className="sheet-empty" colSpan={headers.length + 2}>
                  No rows match “{query}”.
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
        <span className="hint">
          Rename headers to change field names. Rows here drive how many documents generate.
        </span>
      </div>
    </section>
  )
}
