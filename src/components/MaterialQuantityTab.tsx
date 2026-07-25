import { useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractWorkName } from '@core/estimateExtract'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { computeEcvFromItems } from '../boqTransform'
import { computeMaterialTotals } from '@core/materialEstimate'
import type { MaterialTotals } from '@core/materialEstimate'
import { mismatchHint } from '../docClassify'
import { IconFolder, IconDownload, IconWarn, IconTrash, IconTable } from './Icons'

interface Entry {
  id: string
  fileName: string
  itemCount: number
  matchedItemCount: number
  totals: MaterialTotals | null
  /** The work name detected in the estimate's own title block, if any. */
  workName?: string
  ecvRupees: number
  departmentName: string
  district: string
  aiAssisted: string[]
  error: string | null
  saved: string | null
  busy: boolean
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function nextId(): string {
  return `mat-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * Download Material Quantity: upload one or more estimates — each one's work
 * items are turned into the 7 material totals the bundled Material
 * Estimation Template asks for (Stone Aggregates/Sand/Gravel/Granite Slabs/
 * Napa Slabs/Cement/Steel), using standard SSR concrete/masonry/plaster
 * coefficients (see core/materialEstimate.ts).
 */
export default function MaterialQuantityTab() {
  const [entries, setEntries] = useState<Entry[]>([])
  const [pickError, setPickError] = useState<string | null>(null)

  async function uploadEstimates() {
    setPickError(null)
    const grids = await api.pickExcelGrids()
    if (grids.length === 0) return

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
        const { totals, matchedItemCount } = computeMaterialTotals(items)

        added.push({
          id: nextId(),
          fileName: g.name,
          itemCount: items.length,
          matchedItemCount,
          totals,
          workName,
          ecvRupees,
          departmentName: '',
          district: '',
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
          matchedItemCount: 0,
          totals: null,
          ecvRupees: 0,
          departmentName: '',
          district: '',
          aiAssisted: [],
          error: message,
          saved: null,
          busy: false
        })
      }
    }
    setEntries((prev) => [...prev, ...added])
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  function setField(id: string, field: 'departmentName' | 'district', value: string) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, [field]: value } : e)))
  }

  async function download(id: string) {
    const entry = entries.find((e) => e.id === id)
    if (!entry?.totals) return
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, busy: true, error: null } : e)))
    try {
      const suggestedName = `${stripExt(entry.fileName)} Material Quantity`
      const path = await api.exportMaterialEstimate(
        entry.totals,
        {
          workName: entry.workName,
          departmentName: entry.departmentName.trim() || undefined,
          district: entry.district.trim() || undefined,
          ecvRupees: entry.ecvRupees
        },
        suggestedName
      )
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, busy: false, saved: path ?? e.saved } : e)))
    } catch (e) {
      setEntries((prev) =>
        prev.map((x) => (x.id === id ? { ...x, busy: false, error: e instanceof Error ? e.message : String(e) } : x))
      )
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconTable />
        <p>
          {entries.length === 0
            ? 'Upload one or more estimates to compute their material quantities.'
            : `${entries.length} estimate${entries.length === 1 ? '' : 's'} loaded.`}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={uploadEstimates}>
            <IconFolder /> Add Estimate(s)
          </button>
        </div>
      </div>

      {pickError && (
        <div className="notice error">
          <IconWarn />
          {pickError}
        </div>
      )}

      {entries.length > 0 && (
        <ul className="todo-list boq-entries">
          {entries.map((e) => (
            <li key={e.id} className="todo-item">
              <div className="todo-body">
                <span className="todo-text">{e.fileName}</span>
                <span className="todo-dates">
                  {e.error
                    ? e.error
                    : `${e.matchedItemCount} of ${e.itemCount} items contributed material quantities` +
                      (e.workName ? ` · "${e.workName}"` : '')}
                  {e.aiAssisted.length > 0 &&
                    ` · ${e.aiAssisted.join(', ')} column${e.aiAssisted.length === 1 ? '' : 's'} matched by AI — please double-check`}
                  {e.saved && ` · Saved to ${e.saved}`}
                </span>
              </div>
              {e.totals && (
                <input
                  className="editor-name"
                  style={{ maxWidth: 160 }}
                  placeholder="Department Name"
                  value={e.departmentName}
                  onChange={(ev) => setField(e.id, 'departmentName', ev.target.value)}
                />
              )}
              {e.totals && (
                <input
                  className="editor-name"
                  style={{ maxWidth: 130 }}
                  placeholder="District"
                  value={e.district}
                  onChange={(ev) => setField(e.id, 'district', ev.target.value)}
                />
              )}
              {e.totals && (
                <button className="primary" onClick={() => download(e.id)} disabled={e.busy}>
                  <IconDownload /> {e.busy ? 'Saving…' : 'Download'}
                </button>
              )}
              <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                <IconTrash />
              </button>
            </li>
          ))}
        </ul>
      )}

      <p className="hint" style={{ margin: '10px 18px 18px' }}>
        Computed from standard SSR material constants — review against the department's own analysis of rates before
        using this for procurement.
      </p>
    </div>
  )
}
