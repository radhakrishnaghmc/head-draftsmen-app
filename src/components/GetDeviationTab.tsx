import { useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractWorkName } from '@core/estimateExtract'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { extractEstimateAmountLakhs } from '@core/deviation'
import { findWorksRowByName } from '@core/scheduleA'
import { mismatchHint } from '../docClassify'
import { IconFolder, IconDownload, IconWarn, IconTrash, IconTable } from './Icons'
import type { ExcelTable } from '@core/types'
import type { DeviationItem } from '../../electron/ipc-contract'

interface Entry {
  id: string
  fileName: string
  items: DeviationItem[]
  /** The work name detected in the estimate's own title block, if any. */
  workName?: string
  /** Matched from the Works List by workName, if found. */
  circle?: string
  estimateAmountLakhs: number
  agencyName: string
  /** Column labels (Quantity/Rate/Unit/Serial Number) only resolved via semantic matching, not a plain header match — worth a "please double-check" flag. */
  aiAssisted: string[]
  error: string | null
  saved: string | null
  busy: boolean
}

interface Props {
  /** The Works List database — tables[0], by the app's own convention — used to look up the matching Circle. */
  tables: ExcelTable[]
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function nextId(): string {
  return `dev-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * Get Deviation Statement: upload one or more estimates — each one's work
 * items fill the "As per sanctioned estimate" side of the bundled Deviation
 * Statement template (Description/Unit/Qty/Rate, Amount as a live formula).
 * The "As per work done" side, and everything below it (Sub Total, Labour
 * Cess, GST, Seigniorage, ...), is left as live formulas so it all computes
 * once the actual work-done figures are filled in later.
 *
 * Circle and Name of the work come from the estimate's own title block and
 * a Works List lookup by that name; Estimate Amount from a labeled
 * "Estimate Amount"/"ECV" cell in the estimate, or computed from the items
 * when neither is present. Name of the Agency has no reliable source, so
 * it's asked for directly, same as the Bid Document flow.
 */
export default function GetDeviationTab({ tables }: Props) {
  const [entries, setEntries] = useState<Entry[]>([])
  const [pickError, setPickError] = useState<string | null>(null)

  async function uploadEstimates() {
    setPickError(null)
    const grids = await api.pickExcelGrids()
    if (grids.length === 0) return

    const worksTable = tables[0]
    const added: Entry[] = []

    for (const g of grids) {
      try {
        const headerRow = guessHeaderRow(g.grid)
        const { items: estimateItems, aiAssisted } = await extractEstimateItemsWithAi(g.grid, headerRow)
        if (estimateItems.length === 0) {
          throw new Error('No work items with a quantity, rate, and unit were found in that estimate.')
        }
        const items: DeviationItem[] = estimateItems.map((it) => ({
          description: it.description,
          unit: it.unit,
          quantity: it.quantity,
          rate: it.rate
        }))
        const workName = extractWorkName(g.grid, headerRow)
        const estimateAmountLakhs = extractEstimateAmountLakhs(g.grid, headerRow, estimateItems)
        const circle = workName && worksTable ? findWorksRowByName(worksTable, workName)?.['Circle'] : undefined

        added.push({
          id: nextId(),
          fileName: g.name,
          items,
          workName,
          circle,
          estimateAmountLakhs,
          agencyName: '',
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
          items: [],
          estimateAmountLakhs: 0,
          agencyName: '',
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

  function setAgencyName(id: string, agencyName: string) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, agencyName } : e)))
  }

  async function generate(id: string) {
    const entry = entries.find((e) => e.id === id)
    if (!entry || entry.items.length === 0) return
    if (!entry.agencyName.trim()) {
      setEntries((prev) =>
        prev.map((e) => (e.id === id ? { ...e, error: "Enter the agency's name before generating." } : e))
      )
      return
    }
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, busy: true, error: null } : e)))
    try {
      const suggestedName = `${stripExt(entry.fileName)} Deviation`
      const path = await api.exportDeviation(
        entry.items,
        {
          circle: entry.circle,
          nameOfWork: entry.workName,
          agencyName: entry.agencyName.trim(),
          estimateAmountLakhs: entry.estimateAmountLakhs
        },
        suggestedName
      )
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, busy: false, saved: path ?? e.saved } : e)))
    } catch (e) {
      setEntries((prev) =>
        prev.map((x) =>
          x.id === id ? { ...x, busy: false, error: e instanceof Error ? e.message : String(e) } : x
        )
      )
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconTable />
        <p>
          {entries.length === 0
            ? 'Upload one or more estimates to generate their Deviation Statements.'
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
                    : `${e.items.length} item${e.items.length === 1 ? '' : 's'} extracted` +
                      (e.workName ? ` · "${e.workName}"` : '') +
                      (e.circle ? ` · ${e.circle}` : '') +
                      ` · Estimate Amount ${e.estimateAmountLakhs} Lakhs`}
                  {e.aiAssisted.length > 0 &&
                    ` · ${e.aiAssisted.join(', ')} column${e.aiAssisted.length === 1 ? '' : 's'} matched by AI — please double-check`}
                  {e.saved && ` · Saved to ${e.saved}`}
                </span>
              </div>
              {e.items.length > 0 && (
                <input
                  className="editor-name"
                  style={{ maxWidth: 220 }}
                  placeholder="Name of the Agency"
                  value={e.agencyName}
                  onChange={(ev) => setAgencyName(e.id, ev.target.value)}
                />
              )}
              {e.items.length > 0 && (
                <button className="primary" onClick={() => generate(e.id)} disabled={e.busy}>
                  <IconDownload /> {e.busy ? 'Saving…' : 'Generate Deviation'}
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
