import { useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractWorkName } from '@core/estimateExtract'
import type { EstimateWorkItem } from '@core/estimateExtract'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { extractEstimateAmountLakhs } from '@core/deviation'
import { computeEcvFromItems } from '../boqTransform'
import { formatRupees } from '@core/worksAmounts'
import {
  matchWorksRow,
  saveEcvToWorksList,
  downloadBoqFromItems,
  downloadScheduleAFromItems,
  downloadDeviationFromItems,
  downloadMaterialFromItems
} from '../estimateDownloads'
import { mismatchHint } from '../docClassify'
import { IconFolder, IconDownload, IconWarn, IconTrash, IconTable } from './Icons'
import type { ExcelTable } from '@core/types'

type ActionKey = 'boq' | 'scheduleA' | 'deviation' | 'material'

interface Entry {
  id: string
  fileName: string
  items: EstimateWorkItem[]
  workName?: string
  ecvRupees: number
  estimateAmountLakhs: number
  agencyName: string
  departmentName: string
  district: string
  aiAssisted: string[]
  busyAction: ActionKey | null
  error: string | null
  saved: string | null
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
 * Upload one or more estimates — each one's work items (extracted once) can
 * then generate its BOQ, Schedule A, Deviation Statement, and Material
 * Quantity independently, without re-uploading the same estimate for each.
 * Replaces the previously-separate Download BOQ / Get Deviation Statement /
 * Download Material Quantity tabs, and Schedule A's own upload-a-BOQ step
 * for the common case of a BOQ this app itself would generate anyway.
 */
interface EcvConfirmState {
  entryId: string
  workName: string
  ecvRupees: number
  matchedName: string
}

export default function EstimateUploadTab({ tables, onChange }: Props) {
  const [entries, setEntries] = useState<Entry[]>([])
  const [pickError, setPickError] = useState<string | null>(null)
  const [ecvConfirm, setEcvConfirm] = useState<EcvConfirmState | null>(null)

  function updateEntry(id: string, patch: Partial<Entry>) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)))
  }

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
        const estimateAmountLakhs = extractEstimateAmountLakhs(g.grid, headerRow, items)

        added.push({
          id: nextId(),
          fileName: g.name,
          items,
          workName,
          ecvRupees,
          estimateAmountLakhs,
          agencyName: '',
          departmentName: '',
          district: '',
          aiAssisted,
          busyAction: null,
          error: null,
          saved: null
        })
      } catch (e) {
        const headerRow = guessHeaderRow(g.grid)
        const hint = mismatchHint(g.grid[headerRow] ?? [], 'estimate')
        const message = (e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : '')
        added.push({
          id: nextId(),
          fileName: g.name,
          items: [],
          ecvRupees: 0,
          estimateAmountLakhs: 0,
          agencyName: '',
          departmentName: '',
          district: '',
          aiAssisted: [],
          busyAction: null,
          error: message,
          saved: null
        })
      }
    }

    setEntries((prev) => [...prev, ...added])
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  async function run(entry: Entry, action: ActionKey, fn: () => Promise<string | null>): Promise<boolean> {
    updateEntry(entry.id, { busyAction: action, error: null })
    try {
      const path = await fn()
      updateEntry(entry.id, { busyAction: null, saved: path ?? entry.saved })
      return true
    } catch (e) {
      updateEntry(entry.id, { busyAction: null, error: e instanceof Error ? e.message : String(e) })
      return false
    }
  }

  async function downloadBoq(entry: Entry) {
    const ok = await run(entry, 'boq', () => downloadBoqFromItems(entry.items, entry.workName, stripExt(entry.fileName)))
    const worksTable = tables[0]
    if (!ok || !entry.workName || !worksTable) return
    const match = await matchWorksRow(entry.workName, worksTable)
    if (match) {
      setEcvConfirm({
        entryId: entry.id,
        workName: entry.workName,
        ecvRupees: entry.ecvRupees,
        matchedName: match.row['Name of the work'] || entry.workName
      })
    }
  }

  async function confirmEcvUpdate() {
    const pending = ecvConfirm
    setEcvConfirm(null)
    const worksTable = tables[0]
    if (!pending || !worksTable) return
    const updated = await saveEcvToWorksList(pending.workName, pending.ecvRupees, worksTable)
    if (updated) onChange(updated)
  }

  function downloadScheduleA(entry: Entry) {
    return run(entry, 'scheduleA', () =>
      downloadScheduleAFromItems(entry.items, entry.workName, tables[0] ?? null, stripExt(entry.fileName))
    )
  }

  function downloadDeviation(entry: Entry) {
    if (!entry.agencyName.trim()) {
      updateEntry(entry.id, { error: "Enter the agency's name before generating the Deviation Statement." })
      return
    }
    return run(entry, 'deviation', () =>
      downloadDeviationFromItems(
        entry.items,
        entry.workName,
        entry.agencyName,
        entry.estimateAmountLakhs,
        tables[0] ?? null,
        stripExt(entry.fileName)
      )
    )
  }

  function downloadMaterial(entry: Entry) {
    return run(entry, 'material', () =>
      downloadMaterialFromItems(
        entry.items,
        entry.workName,
        entry.ecvRupees,
        entry.departmentName,
        entry.district,
        stripExt(entry.fileName)
      )
    )
  }

  return (
    <div className="card">
      <div className="empty">
        <IconTable />
        <p>
          {entries.length === 0
            ? 'Upload one or more estimates to download their BOQ, Schedule A, Deviation Statement, or Material Quantity.'
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
        <ul className="estimate-list">
          {entries.map((e) => (
            <li key={e.id} className="estimate-entry">
              <div className="estimate-entry-head">
                <span className="estimate-entry-name">{e.fileName}</span>
                <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                  <IconTrash />
                </button>
              </div>

              {e.error ? (
                <p className="estimate-entry-error">
                  <IconWarn /> {e.error}
                </p>
              ) : (
                <div className="live-tiles">
                  <div className="live-tile">
                    <span className="live-tile-label">Items</span>
                    <span className="live-tile-value">{e.items.length}</span>
                  </div>
                  <div className="live-tile">
                    <span className="live-tile-label">ECV</span>
                    <span className="live-tile-value">{formatRupees(e.ecvRupees)}</span>
                  </div>
                  {e.workName && (
                    <div className="live-tile live-tile-wide">
                      <span className="live-tile-label">Name of the work</span>
                      <span className="live-tile-value live-tile-value-text">{e.workName}</span>
                    </div>
                  )}
                </div>
              )}

              {e.aiAssisted.length > 0 && (
                <p className="estimate-hint">
                  {e.aiAssisted.join(', ')} column{e.aiAssisted.length === 1 ? '' : 's'} matched by AI — please
                  double-check.
                </p>
              )}
              {e.saved && <p className="estimate-hint">Saved to {e.saved}</p>}

              {e.items.length > 0 && (
                <ul className="estimate-actions">
                  <li className="estimate-action-row">
                    <button className="primary" onClick={() => downloadBoq(e)} disabled={e.busyAction !== null}>
                      <IconDownload /> {e.busyAction === 'boq' ? 'Saving…' : 'BOQ'}
                    </button>
                    <button
                      className="primary"
                      onClick={() => downloadScheduleA(e)}
                      disabled={e.busyAction !== null}
                    >
                      <IconDownload /> {e.busyAction === 'scheduleA' ? 'Saving…' : 'Schedule A'}
                    </button>
                  </li>
                  <li className="estimate-action-row">
                    <input
                      className="editor-name"
                      placeholder="Name of the Agency"
                      value={e.agencyName}
                      onChange={(ev) => updateEntry(e.id, { agencyName: ev.target.value })}
                    />
                    <button className="primary" onClick={() => downloadDeviation(e)} disabled={e.busyAction !== null}>
                      <IconDownload /> {e.busyAction === 'deviation' ? 'Saving…' : 'Deviation'}
                    </button>
                  </li>
                  <li className="estimate-action-row">
                    <input
                      className="editor-name"
                      placeholder="Department Name"
                      value={e.departmentName}
                      onChange={(ev) => updateEntry(e.id, { departmentName: ev.target.value })}
                    />
                    <input
                      className="editor-name"
                      placeholder="District"
                      value={e.district}
                      onChange={(ev) => updateEntry(e.id, { district: ev.target.value })}
                    />
                    <button className="primary" onClick={() => downloadMaterial(e)} disabled={e.busyAction !== null}>
                      <IconDownload /> {e.busyAction === 'material' ? 'Saving…' : 'Material Quantity'}
                    </button>
                  </li>
                </ul>
              )}
            </li>
          ))}
        </ul>
      )}

      {ecvConfirm && (
        <div className="editor-overlay" onClick={() => setEcvConfirm(null)}>
          <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(ev) => ev.stopPropagation()}>
            <h3>Update ECV in the Works List?</h3>
            <p className="confirm-warn">
              "{ecvConfirm.matchedName}" is already in the Works List. Update its ECV to{' '}
              <strong>{formatRupees(ecvConfirm.ecvRupees)}</strong>, computed from this BOQ?
            </p>
            <p className="confirm-hint">Every other field on that row stays untouched.</p>
            <div className="confirm-actions">
              <button className="ghost" onClick={() => setEcvConfirm(null)}>
                No, keep existing
              </button>
              <button className="primary" onClick={confirmEcvUpdate}>
                Update ECV
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
