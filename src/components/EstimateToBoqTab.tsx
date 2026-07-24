import { useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractWorkName } from '@core/estimateExtract'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { applyEcvFromBoq, formatRupees } from '@core/worksAmounts'
import { mismatchHint } from '../docClassify'
import { buildBoqFromEstimate, computeEcvFromItems } from '../boqTransform'
import { IconFolder, IconDownload, IconWarn, IconCheck, IconTrash, IconTable } from './Icons'
import type { ExcelTable } from '@core/types'

type EcvStatus = 'matched' | 'not-found' | 'no-works-list' | 'no-name' | null

interface Entry {
  id: string
  fileName: string
  itemCount: number
  output: ExcelTable | null
  /** The work name detected in the estimate's own title block, if any. */
  workName?: string
  /** The BOQ's own Amount total — the same figure the estimate's item amounts sum to, i.e. the work's ECV. */
  ecvRupees: number
  ecvStatus: EcvStatus
  /** Column labels (Quantity/Rate/Unit/Serial Number) only resolved via semantic matching, not a plain header match — worth a "please double-check" flag. */
  aiAssisted: string[]
  error: string | null
  saved: string | null
  busy: boolean
}

interface Props {
  /** The Works List database — tables[0], by the app's own convention. */
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function nextId(): string {
  return `est-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * Download BOQ: upload one or more estimates — each one's work items
 * (extracted from its multi-row description/measurement/summary blocks) are
 * mapped onto the standard BOQ column layout. Every estimate gets its own
 * row with its own download, suggested as "<estimate name> BOQ".
 *
 * Each estimate's own title block is also scanned for its "Name of Work" —
 * when found, the BOQ's own Amount total (which the estimate's own item
 * amounts sum to as well) is treated as that work's ECV and saved straight
 * to the matching Works List row, along with EMD @ 1%/1.5% computed from it.
 */
export default function EstimateToBoqTab({ tables, onChange }: Props) {
  const [entries, setEntries] = useState<Entry[]>([])
  const [pickError, setPickError] = useState<string | null>(null)
  const [batchBusy, setBatchBusy] = useState(false)
  const [batchSaved, setBatchSaved] = useState<string | null>(null)

  async function uploadEstimates() {
    setPickError(null)
    const grids = await api.pickExcelGrids()
    if (grids.length === 0) return

    // Applied in sequence against a running copy of the Works List so several
    // estimates in the same batch that match different rows all land — a
    // single onChange at the end, not one per estimate.
    let worksTable = tables[0]
    const added: Entry[] = []

    for (const g of grids) {
      try {
        const headerRow = guessHeaderRow(g.grid)
        const { items, aiAssisted } = await extractEstimateItemsWithAi(g.grid, headerRow)
        if (items.length === 0) {
          throw new Error('No work items with a quantity, rate, and unit were found in that estimate.')
        }
        const workName = extractWorkName(g.grid, headerRow)
        const ecvRupees = computeEcvFromItems(items)

        let ecvStatus: EcvStatus = null
        if (workName) {
          if (!worksTable) {
            ecvStatus = 'no-works-list'
          } else {
            const result = applyEcvFromBoq(worksTable, workName, ecvRupees)
            if (result.matched) {
              worksTable = result.table
              ecvStatus = 'matched'
            } else {
              ecvStatus = 'not-found'
            }
          }
        } else {
          ecvStatus = 'no-name'
        }

        added.push({
          id: nextId(),
          fileName: g.name,
          itemCount: items.length,
          output: buildBoqFromEstimate(items),
          workName,
          ecvRupees,
          ecvStatus,
          aiAssisted,
          error: null,
          saved: null,
          busy: false
        })
      } catch (e) {
        const headerRow = guessHeaderRow(g.grid)
        const hint = mismatchHint(g.grid[headerRow] ?? [], 'estimate')
        const message = (e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : '')
        added.push({
          id: nextId(),
          fileName: g.name,
          itemCount: 0,
          output: null,
          ecvRupees: 0,
          ecvStatus: null,
          aiAssisted: [],
          error: message,
          saved: null,
          busy: false
        })
      }
    }

    if (worksTable && worksTable !== tables[0]) onChange(worksTable)
    setEntries((prev) => [...prev, ...added])
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  async function download(id: string) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, busy: true, error: null } : e)))
    const entry = entries.find((e) => e.id === id)
    if (!entry?.output) return
    try {
      const suggestedName = `${stripExt(entry.fileName)} BOQ`
      const path = await api.exportBoq(entry.output, suggestedName, entry.workName)
      setEntries((prev) =>
        prev.map((e) => (e.id === id ? { ...e, busy: false, saved: path ?? e.saved } : e))
      )
    } catch (e) {
      setEntries((prev) =>
        prev.map((x) =>
          x.id === id ? { ...x, busy: false, error: e instanceof Error ? e.message : String(e) } : x
        )
      )
    }
  }

  async function downloadAll() {
    const ready = entries.filter((e): e is Entry & { output: ExcelTable } => !!e.output)
    if (ready.length === 0) return
    setBatchBusy(true)
    setBatchSaved(null)
    setPickError(null)
    try {
      const paths = await api.exportBoqBatch(
        ready.map((e) => ({
          table: e.output,
          suggestedName: `${stripExt(e.fileName)} BOQ`,
          workName: e.workName
        }))
      )
      if (paths) {
        const savedById = new Map(ready.map((e, i) => [e.id, paths[i]]))
        setEntries((prev) =>
          prev.map((e) => (savedById.has(e.id) ? { ...e, saved: savedById.get(e.id) ?? e.saved } : e))
        )
        setBatchSaved(paths.length > 0 ? paths[0].replace(/[^/\\]+$/, '') : null)
      }
    } catch (e) {
      setPickError(e instanceof Error ? e.message : String(e))
    } finally {
      setBatchBusy(false)
    }
  }

  const readyCount = entries.filter((e) => e.output).length

  function ecvNote(e: Entry): string | null {
    switch (e.ecvStatus) {
      case 'matched':
        return `ECV ${formatRupees(e.ecvRupees)} saved to Works List for "${e.workName}"`
      case 'not-found':
        return `No matching "${e.workName}" row found in the Works List — ECV not saved`
      case 'no-works-list':
        return `ECV ${formatRupees(e.ecvRupees)} — add "${e.workName}" to the Works List to save it there`
      case 'no-name':
        return `Could not detect a "Name of Work" in this estimate — ECV ${formatRupees(e.ecvRupees)} not saved`
      default:
        return null
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconTable />
        <p>
          {entries.length === 0
            ? 'Upload one or more estimates to generate their BOQs.'
            : `${entries.length} estimate${entries.length === 1 ? '' : 's'} loaded.`}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={uploadEstimates}>
            <IconFolder /> Add Estimate(s)
          </button>
          {readyCount > 0 && (
            <button className="primary" onClick={downloadAll} disabled={batchBusy}>
              <IconDownload /> {batchBusy ? 'Saving…' : `Download All (${readyCount})`}
            </button>
          )}
        </div>
      </div>

      {pickError && (
        <div className="notice error">
          <IconWarn />
          {pickError}
        </div>
      )}

      {batchSaved && (
        <div className="notice ok">
          <IconCheck />
          Saved {readyCount} BOQ{readyCount === 1 ? '' : 's'} to {batchSaved}
        </div>
      )}

      {entries.length > 0 && (
        <ul className="todo-list boq-entries">
          {entries.map((e) => (
            <li key={e.id} className="todo-item">
              <div className="todo-body">
                <span className="todo-text">{e.fileName}</span>
                <span className="todo-dates">
                  {e.error ? e.error : `${e.itemCount} item${e.itemCount === 1 ? '' : 's'} extracted`}
                  {e.aiAssisted.length > 0 &&
                    ` · ${e.aiAssisted.join(', ')} column${e.aiAssisted.length === 1 ? '' : 's'} matched by AI — please double-check`}
                  {!e.error && ecvNote(e) && ` · ${ecvNote(e)}`}
                  {e.saved && ` · Saved to ${e.saved}`}
                </span>
              </div>
              {e.output && (
                <button className="primary" onClick={() => download(e.id)} disabled={e.busy}>
                  <IconDownload /> {e.busy ? 'Saving…' : 'Download BOQ'}
                </button>
              )}
              <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                <IconTrash />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
