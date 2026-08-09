import { useEffect, useState } from 'react'
import { api } from '../ipc'
import type { SheetPreview } from '@core/excel'
import { IconChevronLeft, IconTable, IconFolder, IconOpen, IconWarn, IconCheck } from './Icons'

// Live per-sheet progress pushed from the main process while a split runs.
interface SplitProgress {
  done: number
  total: number
  sheet: string
}

interface Pick {
  path: string
  name: string
  sheets: SheetPreview[]
}

interface Props {
  /** Return to the Tools grid. */
  onBack: () => void
}

/**
 * The Excel Sheet Separator's full page: pick a workbook, see every sheet it
 * contains listed with a checkbox, then separate just the ticked sheets — or all
 * of them at once — into one .xlsx per sheet (named after the tab) in a folder
 * the user picks. The split itself runs in the main process (api.splitWorkbook),
 * so the UI stays responsive and shows live per-sheet progress. Mirrors the PDF
 * workspace's layout: previews/list on the left, an action rail on the right.
 */
export default function ExcelSeparatorPage({ onBack }: Props) {
  const [pick, setPick] = useState<Pick | null>(null)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState<SplitProgress | null>(null)
  const [result, setResult] = useState<{ dir: string; files: string[] } | null>(null)
  const [error, setError] = useState<string | null>(null)

  // Subscribe once to the main process's per-sheet progress events.
  useEffect(() => api.onSplitProgress(setProgress), [])

  async function chooseWorkbook() {
    if (busy) return
    setError(null)
    setResult(null)
    setProgress(null)
    try {
      const picked = await api.pickWorkbookForSplit() // null when the user cancels
      if (!picked) return
      setPick(picked)
      setSelected(new Set())
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  function toggle(sheet: string) {
    setSelected((cur) => {
      const next = new Set(cur)
      if (next.has(sheet)) next.delete(sheet)
      else next.add(sheet)
      return next
    })
  }

  function selectAll() {
    if (pick) setSelected(new Set(pick.sheets.map((s) => s.name)))
  }

  // Separate the given sheets (null = every sheet) into a folder the user picks.
  async function doSplit(sheetNames: string[] | null) {
    if (!pick || busy) return
    setBusy(true)
    setError(null)
    setResult(null)
    setProgress(null)
    try {
      const res = await api.splitWorkbook(pick.path, sheetNames && sheetNames.length ? sheetNames : null)
      setResult(res) // null when the folder dialog is cancelled
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      setProgress(null)
    }
  }

  const total = pick?.sheets.length ?? 0
  const allOn = total > 0 && selected.size === total
  const pct = progress && progress.total > 0 ? (progress.done / progress.total) * 100 : 8

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconTable /> <span className="pdf-ws-title-label">Tool:</span> Excel Sheet Separator
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={busy}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <div className="card pdf-workspace">
      <div className="pdf-ws-body">
        <div className="pdf-ws-main">
          {error && (
            <div className="notice error tool-outcome">
              <IconWarn /> {error}
            </div>
          )}

          {!pick ? (
            <button className="pdf-ws-dropzone" onClick={chooseWorkbook}>
              <IconTable />
              <span className="pdf-ws-dropzone-title">Choose workbook</span>
              <span className="pdf-ws-dropzone-sub">Pick a .xlsx / .xls to split into one file per sheet</span>
            </button>
          ) : (
            <section className="pdf-ws-doc">
              <header className="pdf-ws-doc-head">
                <span className="pdf-ws-doc-name" title={pick.name}>
                  {pick.name}
                </span>
                <span className="pdf-ws-doc-meta">
                  {total} sheet{total === 1 ? '' : 's'}
                  {selected.size > 0 && ` · ${selected.size} selected`}
                </span>
                <span className="pdf-ws-doc-ctl">
                  <button className="ghost" onClick={() => (allOn ? setSelected(new Set()) : selectAll())}>
                    {allOn ? 'Deselect all' : 'Select all'}
                  </button>
                </span>
              </header>
              <div className="pdf-ws-pages excel-sheet-grid">
                {pick.sheets.map((s) => {
                  const on = selected.has(s.name)
                  const writing = busy && progress?.sheet === s.name
                  return (
                    <button
                      key={s.name}
                      type="button"
                      className={`pdf-ws-page excel-sheet-tile${on ? ' on' : ''}${s.hidden ? ' hidden' : ''}`}
                      onClick={() => toggle(s.name)}
                      aria-pressed={on}
                      disabled={busy}
                    >
                      <span className="excel-sheet-preview">
                        <table>
                          <tbody>
                            {(s.rows.length ? s.rows : [['(empty)']]).map((r, ri) => (
                              <tr key={ri}>
                                {r.map((c, ci) => (
                                  <td key={ci}>{c}</td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <span className="pdf-ws-page-check" aria-hidden>
                          {on && <IconCheck />}
                        </span>
                        {writing && <span className="excel-sheet-writing">writing…</span>}
                      </span>
                      <span className="pdf-ws-page-no" title={s.name}>
                        {s.name}
                        {s.hidden && <span className="excel-sheet-hidden-tag"> · hidden</span>}
                      </span>
                    </button>
                  )
                })}
              </div>
            </section>
          )}
        </div>

        <aside className="pdf-ws-rail">
          <button className="primary pdf-ws-upload" onClick={chooseWorkbook} disabled={busy}>
            <IconFolder /> {pick ? 'Choose another workbook' : 'Choose workbook'}
          </button>

          {pick && (
            <>
              <div className="pdf-ws-rail-count">
                <strong>{selected.size}</strong> of {total} sheet{total === 1 ? '' : 's'} selected
              </div>

              <div className="pdf-ws-rail-selrow">
                <button className="pdf-ws-selbtn" onClick={selectAll} disabled={allOn || busy}>
                  Select all
                </button>
                <button className="pdf-ws-selbtn" onClick={() => setSelected(new Set())} disabled={selected.size === 0 || busy}>
                  Clear
                </button>
              </div>

              <div className="pdf-ws-rail-actions">
                <button className="primary" onClick={() => doSplit([...selected])} disabled={selected.size === 0 || busy}>
                  <IconFolder /> {busy ? 'Separating…' : 'Separate selected sheets'}
                </button>
                <button className="pdf-ws-railbtn" onClick={() => doSplit(null)} disabled={busy}>
                  <IconFolder /> Separate all sheets
                </button>
              </div>

              {busy && (
                <div className="excel-sep-progress">
                  <span className="excel-sep-progress-label">
                    {progress ? `Writing ${progress.done} / ${progress.total} sheets…` : 'Reading workbook…'}
                  </span>
                  <span className="excel-sep-progress-track" aria-hidden>
                    <span className="excel-sep-progress-bar" style={{ width: `${pct}%` }} />
                  </span>
                </div>
              )}
            </>
          )}

          {result && (
            <div className="notice ok tool-result pdf-ws-rail-result">
              <span>
                <IconFolder /> Saved {result.files.length} sheet{result.files.length === 1 ? '' : 's'} to {result.dir}
              </span>
              <button className="ghost" onClick={() => api.openPath(result.dir)}>
                <IconOpen /> Open folder
              </button>
            </div>
          )}

          <p className="pdf-ws-rail-hint">
            Each selected sheet is saved as its own .xlsx, named after the tab, in the folder you pick.
          </p>
        </aside>
      </div>
      </div>
    </>
  )
}
