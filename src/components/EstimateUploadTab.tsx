import { useMemo, useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow, buildTableFromGrid } from '@core/sheet'
import { indianDigitGroups } from '@core/worksAmounts'
import { extractWorkName, splitEstimateBlocks, extractGrandTotalLakhs, itemsMissingEstimateAmount } from '@core/estimateExtract'
import type { EstimateWorkItem } from '@core/estimateExtract'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { extractEstimateAmountLakhs } from '@core/deviation'
import { computeEcvFromItems, buildBoqFromEstimate, boqToItems, extractWorkNameFromBoq } from '../boqTransform'
import { formatRupees } from '@core/worksAmounts'
import { metaFromWorksRow } from '@core/scheduleA'
import type { WorksRowMatch } from '@core/scheduleA'
import { computeMaterialTotals } from '@core/materialEstimate'
import type { MaterialTotals } from '@core/materialEstimate'
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
import type { Office } from '../office'
import { corporationByName } from '../zoneCircleDirectory'
import { closeOnBackdropMouseDown } from '../overlayClose'

type ActionKey = 'boq' | 'scheduleA' | 'deviation' | 'material'

const DOC_TABS: { key: ActionKey; label: string }[] = [
  { key: 'boq', label: 'BOQ' },
  { key: 'scheduleA', label: 'Schedule A' },
  { key: 'deviation', label: 'Deviation' },
  { key: 'material', label: 'Material Quantity' }
]

function itemAmount(it: EstimateWorkItem): number {
  return (Number(it.quantity) || 0) * (Number(it.rate) || 0)
}

function itemsTotal(items: EstimateWorkItem[]): number {
  return items.reduce((sum, it) => sum + itemAmount(it), 0)
}

const MATERIAL_ROWS: { key: keyof MaterialTotals; label: string; unit: string }[] = [
  { key: 'stoneAggregatesMt', label: 'Stone Aggregates', unit: 'MT' },
  { key: 'sandMt', label: 'Sand', unit: 'MT' },
  { key: 'gravelMt', label: 'Gravel', unit: 'MT' },
  { key: 'graniteSqft', label: 'Granite Slabs', unit: 'Sqft' },
  { key: 'napaSqft', label: 'Napa Slabs', unit: 'Sqft' },
  { key: 'cementMt', label: 'Cement', unit: 'MT' },
  { key: 'steelMt', label: 'Steel', unit: 'MT' }
]

interface Entry {
  id: string
  fileName: string
  /** The workbook sheet/tab this estimate came from — the BOQ's name descriptor. */
  sheetName?: string
  items: EstimateWorkItem[]
  /** True when the upload was a flat BOQ, not a detailed estimate — a BOQ has no
   * sanctioned estimate amount of its own, so its total is the ECV alone. */
  isBoq?: boolean
  workName?: string
  /** The Works List row this work matched (exact name, else AI/embedding),
   * resolved once at upload so the preview's estimate/ECV/agency come from the
   * Works List rather than being recomputed. null = searched but not found. */
  worksMatch: WorksRowMatch | null
  /** True when a Works List lookup was actually attempted (a work name and a
   * loaded Works List both existed) — drives the "not found" note. */
  worksSearched: boolean
  ecvRupees: number
  estimateAmountLakhs: number
  /** Sanctioned Grand Total in lakhs (bottom of the estimate) — used in the BOQ file name. */
  grandTotalLakhs?: number
  /** Items the estimate left un-costed (a quantity and rate, but a blank/zero Amount) — the BOQ counts them, so its total exceeds the estimate's own. */
  uncostedItems: EstimateWorkItem[]
  agencyName: string
  departmentName: string
  district: string
  aiAssisted: string[]
  busyAction: ActionKey | null
  error: string | null
  saved: string | null
  /** Which document's live preview is showing right now. */
  previewDoc: ActionKey
}

interface Props {
  /** The Works List database — tables[0], by the app's own convention. */
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
  /** The chosen office — its corporation is the default Department name. */
  office: Office
}

/**
 * The default "Department" for a new estimate — the chosen corporation's full
 * name (e.g. "Cyberabad Municipal Corporation"), falling back to its short code.
 * Blank when no office/corporation is set. Stays editable per entry.
 */
function defaultDepartmentName(office: Office): string {
  if (!office.corporation) return ''
  return corporationByName(office.corporation)?.fullName ?? office.corporation
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function nextId(): string {
  return `est-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

// The contractor/agency named on a matched Works List row — pre-fills the
// Deviation's "Name of the Agency" (stays editable). Empty when nothing matched.
function agencyFromMatch(match: WorksRowMatch | null): string {
  return (match?.row['Name of the Agency'] ?? '').trim()
}

// A blank sheet — no cell holds anything but whitespace. A workbook often
// carries such empty tabs alongside its real estimates; they'd otherwise land
// in the list as "no work items found" errors, so they're skipped up front.
function isSheetEmpty(grid: string[][]): boolean {
  return grid.every((row) => row.every((cell) => cell.trim() === ''))
}

// Total Qty × Rate of the un-costed items — how far the BOQ's total runs above
// the estimate's own on their account.
function uncostedWorth(items: EstimateWorkItem[]): number {
  const num = (s: string) => Number(String(s).replace(/,/g, '')) || 0
  return items.reduce((sum, it) => sum + Math.round(num(it.quantity) * num(it.rate)), 0)
}

// A description clipped for an inline notice — enough to identify the item.
function shortDesc(description: string): string {
  const d = description.trim()
  return d.length > 60 ? `${d.slice(0, 57)}…` : d
}

// Strip characters a file name can't contain, and collapse whitespace.
function sanitizeFileName(s: string): string {
  return s
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

/**
 * The BOQ file name for an estimate: "Boq <sheet/tab name> <estimate amount in
 * lakhs>" (e.g. "Boq peddamma 112.00"). The sheet name is the descriptor
 * because a workbook's estimates are near-identically worded and only their
 * tab names (the localities) tell them apart; the amount is the estimate's
 * Grand Total in lakhs. Falls back to the file name / ECV-based lakhs when a
 * sheet name or Grand Total isn't available. exportBoqBatch de-duplicates any
 * that still collide.
 */
function boqBaseName(e: Entry): string {
  const descriptor = sanitizeFileName(e.sheetName || stripExt(e.fileName)) || 'estimate'
  const lakhs = e.grandTotalLakhs ?? e.estimateAmountLakhs
  const amount = lakhs ? ` ${lakhs.toFixed(2)}` : ''
  return sanitizeFileName(`Boq ${descriptor}${amount}`)
}

/**
 * A live, spreadsheet-styled preview of exactly one of the four documents
 * this estimate can produce — the same underlying line items, laid out the
 * way that document's own template presents them (Schedule A's meta block,
 * Deviation's Circle/Agency header, Material Quantity's computed totals),
 * so switching tabs shows what that specific download will actually look
 * like rather than one generic items grid.
 */
function DocumentPreview({ entry }: { entry: Entry }) {
  const matchedRow = entry.worksMatch?.row
  const scheduleAMeta = useMemo(() => (matchedRow ? metaFromWorksRow(matchedRow) : undefined), [matchedRow])
  const materialTotals = useMemo(() => computeMaterialTotals(entry.items).totals, [entry.items])
  const total = itemsTotal(entry.items)

  // Estimate Amount fallback when this work isn't on the Works List: the
  // estimate's own sanctioned Grand Total (Lakhs → rupees), NOT the bare
  // Σ(qty×rate) item-sum — that sum is the ECV, and showing it here made the
  // two fields identical. Only drops back to the item-sum if no Grand Total
  // was found at all.
  const estimateLakhs = entry.grandTotalLakhs ?? entry.estimateAmountLakhs
  const estimateFallback = estimateLakhs
    ? `${indianDigitGroups(Math.round(estimateLakhs * 100000))}/-`
    : `${total.toFixed(2)}/-`

  const itemTable = (
    <table className="doc-sheet-table">
      <thead>
        <tr>
          <th>Item No.</th>
          <th>Description</th>
          <th>Unit</th>
          <th>Qty</th>
          <th>Rate</th>
          <th>Amount</th>
        </tr>
      </thead>
      <tbody>
        {entry.items.map((it, i) => (
          <tr key={i}>
            <td>{i + 1}</td>
            <td>{it.description}</td>
            <td>{it.unit}</td>
            <td className="num">{it.quantity}</td>
            <td className="num">{Number(it.rate) ? Number(it.rate).toFixed(2) : it.rate}</td>
            <td className="num">{itemAmount(it).toFixed(2)}</td>
          </tr>
        ))}
      </tbody>
      <tfoot>
        <tr>
          <td colSpan={5}>Total</td>
          <td className="num">{total.toFixed(2)}</td>
        </tr>
      </tfoot>
    </table>
  )

  // Item count + ECV shown right in the sheet's title band (rather than as
  // separate chips above the preview).
  const titleStats = (
    <span className="doc-sheet-title-stats">
      {entry.items.length} items · ECV {formatRupees(entry.ecvRupees)}
    </span>
  )

  if (entry.previewDoc === 'scheduleA') {
    return (
      <div className="doc-sheet">
        <div className="doc-sheet-title">
          <span>SCHEDULE-A · BILL OF QUANTITIES</span>
          {titleStats}
        </div>
        <div className="doc-sheet-meta">
          <div>
            <span>Name of work</span>
            <strong className="doc-sheet-work-name">{entry.workName || '—'}</strong>
          </div>
          <div>
            <span>Estimate Amount</span>
            {/* A BOQ has no sanctioned estimate amount of its own — only the ECV. */}
            <strong>{entry.isBoq ? '—' : `Rs. ${scheduleAMeta?.estimateAmount ?? estimateFallback}`}</strong>
          </div>
          <div>
            <span>ECV Amount</span>
            <strong>
              {scheduleAMeta?.ecvAmount
                ? `Rs. ${scheduleAMeta.ecvAmount}`
                : entry.ecvRupees
                  ? `Rs. ${indianDigitGroups(entry.ecvRupees)}/- (computed)`
                  : 'Not on the Works List yet'}
            </strong>
          </div>
          <div>
            <span>Contract Amount</span>
            <strong>{scheduleAMeta?.contractAmount ? `Rs. ${scheduleAMeta.contractAmount}` : '—'}</strong>
          </div>
          <div>
            <span>Name of the Contractor</span>
            <strong>{matchedRow?.['Name of the Agency'] || '—'}</strong>
          </div>
        </div>
        {itemTable}
      </div>
    )
  }

  if (entry.previewDoc === 'deviation') {
    return (
      <div className="doc-sheet">
        <div className="doc-sheet-title">
          <span>DEVIATION STATEMENT</span>
          {titleStats}
        </div>
        <div className="doc-sheet-meta">
          <div>
            <span>Circle</span>
            <strong>{matchedRow?.['Circle'] || '—'}</strong>
          </div>
          <div>
            <span>Name of Work</span>
            <strong className="doc-sheet-work-name">{entry.workName || '—'}</strong>
          </div>
          <div>
            <span>Agency</span>
            <strong>{entry.agencyName || 'Not entered yet'}</strong>
          </div>
          <div>
            <span>Estimate Amount</span>
            <strong>{entry.estimateAmountLakhs ? `${entry.estimateAmountLakhs} Lakhs` : '—'}</strong>
          </div>
        </div>
        {itemTable}
      </div>
    )
  }

  if (entry.previewDoc === 'material') {
    return (
      <div className="doc-sheet">
        <div className="doc-sheet-title">
          <span>MATERIAL ESTIMATION</span>
          {titleStats}
        </div>
        <div className="doc-sheet-meta">
          <div>
            <span>Name of work</span>
            <strong className="doc-sheet-work-name">{entry.workName || '—'}</strong>
          </div>
          <div>
            <span>Department</span>
            <strong>{entry.departmentName || 'Not entered yet'}</strong>
          </div>
          <div>
            <span>District</span>
            <strong>{entry.district || 'Not entered yet'}</strong>
          </div>
          <div>
            <span>ECV</span>
            <strong>{formatRupees(entry.ecvRupees)}</strong>
          </div>
        </div>
        <table className="doc-sheet-table">
          <thead>
            <tr>
              <th>Material</th>
              <th>Quantity</th>
              <th>Unit</th>
            </tr>
          </thead>
          <tbody>
            {MATERIAL_ROWS.map((r) => (
              <tr key={r.key}>
                <td>{r.label}</td>
                <td className="num">{materialTotals[r.key].toFixed(2)}</td>
                <td>{r.unit}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  return (
    <div className="doc-sheet">
      <div className="doc-sheet-title">
        <span>BILL OF QUANTITIES</span>
        {titleStats}
      </div>
      {itemTable}
    </div>
  )
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

export default function EstimateUploadTab({ tables, onChange, office }: Props) {
  const [entries, setEntries] = useState<Entry[]>([])
  const deptDefault = defaultDepartmentName(office)
  const [pickError, setPickError] = useState<string | null>(null)
  const [ecvConfirm, setEcvConfirm] = useState<EcvConfirmState | null>(null)
  const [downloadingAllBoqs, setDownloadingAllBoqs] = useState(false)
  const [allBoqsSavedDir, setAllBoqsSavedDir] = useState<string | null>(null)

  function updateEntry(id: string, patch: Partial<Entry>) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)))
  }

  // Every loaded estimate that parsed into items — the ones a BOQ can be built for.
  const boqReadyEntries = entries.filter((e) => !e.error && e.items.length > 0)

  // Save one BOQ per loaded estimate into a single chosen folder, in one go.
  async function downloadAllBoqs() {
    if (boqReadyEntries.length === 0) return
    setPickError(null)
    setAllBoqsSavedDir(null)
    setDownloadingAllBoqs(true)
    try {
      const paths = await api.exportBoqBatch(
        boqReadyEntries.map((e) => ({
          table: buildBoqFromEstimate(e.items),
          suggestedName: boqBaseName(e),
          workName: e.workName
        }))
      )
      if (paths && paths.length > 0) setAllBoqsSavedDir(paths[0].replace(/[^/\\]+$/, ''))
    } catch (e) {
      setPickError(e instanceof Error ? e.message : String(e))
    } finally {
      setDownloadingAllBoqs(false)
    }
  }

  async function uploadEstimates() {
    setPickError(null)
    let grids = await api.pickExcelGrids()
    if (grids.length === 0) return

    // Ignore blank sheets entirely — don't generate a BOQ (or a "no items"
    // error) for a tab that has no data at all.
    grids = grids.filter((g) => !isSheetEmpty(g.grid))
    if (grids.length === 0) {
      setPickError('No data found — every sheet in the selected workbook is empty.')
      return
    }

    // A workbook may keep extra estimates on hidden sheets. Only process those
    // when the user opts in — otherwise generate for the visible sheet(s) only.
    // (Skip the prompt when every sheet is hidden: there'd be nothing to fall
    // back to, so just use them all.)
    const hiddenGrids = grids.filter((g) => g.hidden)
    const visibleGrids = grids.filter((g) => !g.hidden)
    if (hiddenGrids.length > 0 && visibleGrids.length > 0) {
      const names = hiddenGrids.map((g) => g.sheetName || g.name).join(', ')
      const includeHidden = window.confirm(
        `This workbook has ${hiddenGrids.length} hidden estimate sheet${hiddenGrids.length === 1 ? '' : 's'} (${names}).\n\n` +
          `Click OK to generate the BOQ and documents for ALL sheets, including the hidden ones.\n` +
          `Click Cancel to use only the visible sheet${visibleGrids.length === 1 ? '' : 's'}.`
      )
      grids = includeHidden ? grids : visibleGrids
    }

    const added: Entry[] = []
    const worksTable = tables[0] ?? null

    // Resolve a work to its Works List row once, up front — exact name match
    // first, then the same AI/embedding fallback the Schedule A tab uses, so a
    // slightly-reworded name still finds its row. The estimate/ECV amounts and
    // the agency in the preview all read from this match.
    async function resolveWorksMatch(
      workName: string | undefined
    ): Promise<{ match: WorksRowMatch | null; searched: boolean }> {
      if (!workName || !worksTable) return { match: null, searched: false }
      const match = (await matchWorksRow(workName, worksTable)) ?? null
      return { match, searched: true }
    }

    // Each grid is one sheet; a sheet may itself stack several complete
    // estimates, so split it into per-estimate blocks and process each.
    for (const g of grids) {
      const blocks = splitEstimateBlocks(g.grid)
      for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
        const grid = blocks[blockIndex]
        // Distinguish stacked estimates within one sheet by a #n suffix.
        const sheetName =
          blocks.length > 1 ? `${g.sheetName || g.name} #${blockIndex + 1}` : g.sheetName || g.name
        const headerRow = guessHeaderRow(grid)
        try {
          const { items, aiAssisted } = await extractEstimateItemsWithAi(grid, headerRow)
          if (items.length > 0) {
            const workName = extractWorkName(grid, headerRow)
            const { match, searched } = await resolveWorksMatch(workName)
            added.push({
              id: nextId(),
              fileName: g.name,
              sheetName,
              items,
              workName,
              worksMatch: match,
              worksSearched: searched,
              ecvRupees: computeEcvFromItems(items),
              estimateAmountLakhs: extractEstimateAmountLakhs(grid, headerRow, items),
              grandTotalLakhs: extractGrandTotalLakhs(grid, headerRow),
              uncostedItems: itemsMissingEstimateAmount(items),
              agencyName: agencyFromMatch(match),
              departmentName: deptDefault,
              district: '',
              aiAssisted,
              busyAction: null,
              error: null,
              saved: null,
              previewDoc: 'boq'
            })
          } else {
            // Not a detailed estimate — try it as a flat BOQ, so an existing BOQ
            // (from outside the app) can drive the same BOQ / Schedule A / Material
            // outputs. A BOQ has no sanctioned estimate amount of its own, so its
            // items total is the ECV alone (estimateAmountLakhs stays 0).
            const boqTable = buildTableFromGrid(grid, headerRow, { id: `boq-${nextId()}`, name: g.name, path: g.path })
            const boqItems = boqToItems(boqTable)
            if (boqItems.length === 0) {
              throw new Error('No work items with a quantity, rate, and unit were found in that estimate or BOQ.')
            }
            const boqWorkName = extractWorkNameFromBoq(boqTable)
            const { match, searched } = await resolveWorksMatch(boqWorkName)
            added.push({
              id: nextId(),
              fileName: g.name,
              sheetName,
              items: boqItems,
              isBoq: true,
              workName: boqWorkName,
              worksMatch: match,
              worksSearched: searched,
              ecvRupees: computeEcvFromItems(boqItems),
              estimateAmountLakhs: 0,
              uncostedItems: [],
              agencyName: agencyFromMatch(match),
              departmentName: deptDefault,
              district: '',
              aiAssisted: [],
              busyAction: null,
              error: null,
              saved: null,
              previewDoc: 'boq'
            })
          }
        } catch (e) {
          const hint = mismatchHint(grid[headerRow] ?? [], 'estimate')
          const message = (e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : '')
          added.push({
            id: nextId(),
            fileName: g.name,
            sheetName,
            items: [],
            worksMatch: null,
            worksSearched: false,
            ecvRupees: 0,
            estimateAmountLakhs: 0,
            uncostedItems: [],
            agencyName: '',
            departmentName: '',
            district: '',
            aiAssisted: [],
            busyAction: null,
            error: message,
            saved: null,
            previewDoc: 'boq'
          })
        }
      }
    }

    setEntries((prev) => [...prev, ...added])

    // Once a BOQ is uploaded here, switch the whole UI over to Verdana by
    // overriding the --font design token on the document root (body and the
    // rest of the chrome all read font-family from var(--font)).
    if (added.some((e) => e.isBoq && !e.error)) {
      document.documentElement.style.setProperty('--font', 'Verdana, Geneva, Tahoma, sans-serif')
      document.body.style.fontSize = '16px'
    }
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
    const ok = await run(entry, 'boq', () =>
      downloadBoqFromItems(entry.items, entry.workName, stripExt(entry.fileName), boqBaseName(entry))
    )
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

  function downloadActive(entry: Entry) {
    if (entry.previewDoc === 'scheduleA') return downloadScheduleA(entry)
    if (entry.previewDoc === 'deviation') return downloadDeviation(entry)
    if (entry.previewDoc === 'material') return downloadMaterial(entry)
    return downloadBoq(entry)
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
          <button className="primary upload-btn" onClick={uploadEstimates}>
            <IconFolder /> Add estimates/BOQs
          </button>
          {boqReadyEntries.length > 1 && (
            <button className="ghost upload-btn" onClick={downloadAllBoqs} disabled={downloadingAllBoqs}>
              <IconDownload /> {downloadingAllBoqs ? 'Saving…' : `Download all BOQs (${boqReadyEntries.length})`}
            </button>
          )}
        </div>
        {allBoqsSavedDir && (
          <div className="notice ok">
            <IconTable /> Saved {boqReadyEntries.length} BOQs to {allBoqsSavedDir}
          </div>
        )}
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
                <div className="estimate-entry-titles">
                  <span className="estimate-entry-file">
                    {e.sheetName ? `${e.fileName} · ${e.sheetName}` : e.fileName}
                  </span>
                  {e.workName && <h3 className="estimate-work-name">{e.workName}</h3>}
                </div>
                <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                  <IconTrash />
                </button>
              </div>

              {e.error && (
                <p className="estimate-entry-error">
                  <IconWarn /> {e.error}
                </p>
              )}

              {e.aiAssisted.length > 0 && (
                <p className="estimate-hint">
                  {e.aiAssisted.join(', ')} column{e.aiAssisted.length === 1 ? '' : 's'} matched by AI — please
                  double-check.
                </p>
              )}
              {e.worksSearched && !e.worksMatch && (
                <p className="estimate-hint estimate-warning">
                  "{e.workName}" wasn't found on the Works List — Estimate Amount, ECV and Name of the Agency were
                  computed from the file, not read from the Works List.
                </p>
              )}
              {e.worksMatch?.matchedViaAi && (
                <p className="estimate-hint">
                  Matched to the Works List by AI (the name differs) — please verify the agency and amounts.
                </p>
              )}
              {e.uncostedItems.length > 0 && (
                <p className="estimate-hint estimate-warning">
                  {e.uncostedItems.length} item{e.uncostedItems.length === 1 ? '' : 's'} ha
                  {e.uncostedItems.length === 1 ? 's' : 've'} a quantity and rate but a blank/zero Amount in the
                  estimate ({formatRupees(uncostedWorth(e.uncostedItems))} of work) — the BOQ counts{' '}
                  {e.uncostedItems.length === 1 ? 'it' : 'them'}, so its total is that much above the estimate's own.
                  Check: {e.uncostedItems.map((it) => shortDesc(it.description)).join('; ')}.
                </p>
              )}
              {e.saved && <p className="estimate-hint">Saved to {e.saved}</p>}

              {e.items.length > 0 && (
                <div className="estimate-workspace">
                  <div className="doc-tabs">
                    {DOC_TABS.map((t) => (
                      <button
                        key={t.key}
                        className={`doc-tab${e.previewDoc === t.key ? ' active' : ''}`}
                        onClick={() => updateEntry(e.id, { previewDoc: t.key })}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>

                  <div className="estimate-fields">
                    <label className="estimate-details-field">
                      Name of the Agency <span className="estimate-hint">— Deviation</span>
                      <input
                        className="editor-name"
                        placeholder="Name of the Agency"
                        value={e.agencyName}
                        onChange={(ev) => updateEntry(e.id, { agencyName: ev.target.value })}
                      />
                    </label>
                    <label className="estimate-details-field">
                      Department Name <span className="estimate-hint">— Material Quantity</span>
                      <input
                        className="editor-name"
                        placeholder="Department Name"
                        value={e.departmentName}
                        onChange={(ev) => updateEntry(e.id, { departmentName: ev.target.value })}
                      />
                    </label>
                    <label className="estimate-details-field">
                      District <span className="estimate-hint">— Material Quantity</span>
                      <input
                        className="editor-name"
                        placeholder="District"
                        value={e.district}
                        onChange={(ev) => updateEntry(e.id, { district: ev.target.value })}
                      />
                    </label>
                  </div>

                  <div className="estimate-preview-scroll">
                    <DocumentPreview entry={e} />
                  </div>

                  <div className="doc-sheet-footer">
                    <span className="estimate-hint">
                      Live preview of the {DOC_TABS.find((t) => t.key === e.previewDoc)?.label} — updates as you fill
                      in the details.
                    </span>
                    <button className="primary" onClick={() => downloadActive(e)} disabled={e.busyAction !== null}>
                      <IconDownload />{' '}
                      {e.busyAction === e.previewDoc
                        ? 'Saving…'
                        : `Download ${DOC_TABS.find((t) => t.key === e.previewDoc)?.label}`}
                    </button>
                  </div>
                </div>
              )}
            </li>
          ))}
        </ul>
      )}

      {ecvConfirm && (
        <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setEcvConfirm(null))}>
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
