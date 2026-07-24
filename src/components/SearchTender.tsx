import { useCallback, useEffect, useState } from 'react'
import { fetchTenders } from '../tenderCache'
import { IconSearch, IconTable, IconPlus } from './Icons'

const PAGE_SIZES = [10, 25, 50]

const COLUMNS = [
  { key: 1, label: 'Tender ID', cls: 'tk-id' },
  { key: 2, label: 'NIT / Tender No', cls: 'tk-nit' },
  { key: 0, label: 'Organisation', cls: 'tk-org' },
  { key: 4, label: 'Work', cls: 'tk-work' },
  { key: 5, label: 'Value', cls: 'tk-val' },
  { key: 7, label: 'Bid closing', cls: 'tk-date' },
  { key: 8, label: 'Opens', cls: 'tk-date' }
] as const

const inr = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  maximumFractionDigits: 0
})

function formatValue(raw: string): string {
  const n = Number(raw)
  if (!raw || Number.isNaN(n)) return raw || '—'
  return inr.format(n)
}

interface Props {
  /** Called once per selected row when "Add to Reminders" is clicked. */
  onAddReminder?: (row: string[]) => void
}

export default function SearchTender({ onAddReminder }: Props) {
  const [type] = useState('current')
  const [length, setLength] = useState(25)
  const [start, setStart] = useState(0)
  const [rows, setRows] = useState<string[][]>([])
  const [total, setTotal] = useState(-1)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  // `search` is the input box; `submitted` is the keyword sent to the portal.
  const [search, setSearch] = useState('')
  const [submitted, setSubmitted] = useState('')
  // Indices (within the current page of `rows`) the user has checked.
  const [selected, setSelected] = useState<Set<number>>(new Set())

  const load = useCallback(
    async (force = false) => {
      if (!submitted) return
      setLoading(true)
      setError(null)
      try {
        const res = await fetchTenders({ start, length, type, search: submitted }, force)
        setRows(res.data)
        setTotal(res.total)
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e))
        setRows([])
      } finally {
        setLoading(false)
      }
    },
    [start, length, type, submitted]
  )

  // Only fetch once the user has actually searched for something — no
  // unprompted preview of the whole live tender list.
  useEffect(() => {
    if (submitted) {
      load()
    } else {
      setRows([])
      setTotal(-1)
      setError(null)
    }
  }, [load, submitted])

  // The row list just changed (new search, page, refresh) — any prior
  // checkbox selection no longer lines up with what's on screen.
  useEffect(() => {
    setSelected(new Set())
  }, [rows])

  function toggleRow(i: number) {
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(i)) next.delete(i)
      else next.add(i)
      return next
    })
  }

  const allSelected = rows.length > 0 && selected.size === rows.length

  function toggleAll() {
    setSelected(allSelected ? new Set() : new Set(rows.map((_, i) => i)))
  }

  function addSelectedToReminders() {
    rows.forEach((r, i) => {
      if (selected.has(i)) onAddReminder?.(r)
    })
    setSelected(new Set())
  }

  // Debounce typing: fire a fresh server-side search ~600ms after the user
  // stops typing (from page 1) — no Search button needed.
  useEffect(() => {
    const term = search.trim()
    if (term === submitted) return
    const t = setTimeout(() => {
      setStart(0)
      setSubmitted(term)
    }, 600)
    return () => clearTimeout(t)
  }, [search, submitted])

  const clearSearch = useCallback(() => {
    setSearch('')
    setStart(0)
    setSubmitted('')
  }, [])

  const hasNext = total >= 0 ? start + length < total : rows.length === length
  const hasPrev = start > 0
  const rangeEnd = start + rows.length

  return (
    <div className="tender">
      <section className="card tender-controls">
        <div className="tsearch">
          <IconSearch />
          <input
            type="text"
            placeholder="Search by Tender ID, IFB / NIT no, or name of work…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {loading && <span className="tsearch-spin">Searching…</span>}
          {search && (
            <button className="tsearch-clear" onClick={clearSearch}>
              Clear
            </button>
          )}
        </div>
        <div className="tender-tools">
          <label className="tender-size">
            Rows
            <select
              value={length}
              onChange={(e) => {
                setStart(0)
                setLength(Number(e.target.value))
              }}
            >
              {PAGE_SIZES.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </label>
          <button className="ghost" onClick={() => load(true)} disabled={loading || !submitted}>
            {loading ? 'Loading…' : 'Refresh'}
          </button>
          <button
            className="ghost"
            onClick={addSelectedToReminders}
            disabled={loading || selected.size === 0 || !onAddReminder}
            title="Add the checked tenders to Tender Reminders"
          >
            <IconPlus /> Add to Reminders{selected.size > 0 ? ` (${selected.size})` : ''}
          </button>
        </div>
      </section>

      {error ? (
        <div className="card">
          <div className="empty">
            <IconTable />
            <p>{error}</p>
            <button className="primary" onClick={() => load()}>
              Try again
            </button>
          </div>
        </div>
      ) : !submitted ? null : (
        <section className="card tender-results">
          <div className="tender-meta">
            <span>
              {loading
                ? 'Fetching live tenders…'
                : `Showing ${rows.length === 0 ? 0 : start + 1}–${rangeEnd}${
                    total >= 0 ? ` of ${total}` : ''
                  }`}
              {submitted && !loading ? ` · matching “${submitted}”` : ''}
            </span>
            <div className="tender-page">
              <button
                className="ghost"
                disabled={!hasPrev || loading}
                onClick={() => setStart(Math.max(0, start - length))}
              >
                ‹ Prev
              </button>
              <button
                className="ghost"
                disabled={!hasNext || loading}
                onClick={() => setStart(start + length)}
              >
                Next ›
              </button>
            </div>
          </div>

          <div className="tender-wrap">
            <table className="tender-table">
              <thead>
                <tr>
                  <th className="tk-select">
                    <input
                      type="checkbox"
                      checked={allSelected}
                      onChange={toggleAll}
                      aria-label="Select all"
                    />
                  </th>
                  {COLUMNS.map((c) => (
                    <th key={c.key} className={c.cls}>
                      {c.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.length === 0 && !loading ? (
                  <tr>
                    <td className="tender-none" colSpan={COLUMNS.length + 1}>
                      No tenders match “{submitted}”.
                    </td>
                  </tr>
                ) : (
                  rows.map((r, i) => (
                    <tr key={`${r[1]}-${i}`} className={selected.has(i) ? 'tk-row-selected' : ''}>
                      <td className="tk-select">
                        <input
                          type="checkbox"
                          checked={selected.has(i)}
                          onChange={() => toggleRow(i)}
                          aria-label={`Select ${r[2] || r[1] || 'row'}`}
                        />
                      </td>
                      {COLUMNS.map((c) => (
                        <td key={c.key} className={c.cls} title={r[c.key]}>
                          {c.key === 5 ? formatValue(r[5]) : r[c.key] || '—'}
                        </td>
                      ))}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </div>
  )
}
