import { useRef, useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { mismatchHint } from '../docClassify'
import { buildRateIndex } from '@core/rateDatabase'
import type { RateEntry } from '@core/rateDatabase'
import { buildSorRateIndex } from '@core/sorPdfExtract'
import { pdfPagesToTextLinesFromData } from '../pdfToText'
import { pdfPagesToDataUrlsFromData } from '../pdfToImages'
import {
  matchEstimateItems,
  buildCellEdits,
  buildRateAnalysisSheet,
  isAutoResolved,
  needsReview,
  matchByExactDescription,
  buildReplicateIndex,
  buildReplicateCellEdits
} from '@core/technicalSanction'
import type { ItemMatch, ResolvedItem, ReplicateMatch, CellEdit } from '@core/technicalSanction'
import type { EstimateWorkItem } from '@core/estimateExtract'
import { parseCementSteelRateLines, findCementSteelRateCells } from '@core/cementSteelRate'
import type { CementSteelRateCellRef, CementSteelRates } from '@core/cementSteelRate'
import { CEMENT_STEEL_MONTHS, cementSteelRatePeriodKey } from '@core/cementSteelRates'
import { IconFolder, IconDownload, IconWarn, IconCheck, IconTable, IconRefresh, IconChevronLeft, IconTrash } from './Icons'
import type { SheetGrid } from '../../electron/ipc-contract'

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

const inr = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  maximumFractionDigits: 2
})

function formatRate(rate: string): string {
  const n = Number(rate)
  return Number.isFinite(n) ? inr.format(n) : rate
}

/**
 * Give Technical Sanction: upload the Data Sheet (rates database — uploaded
 * fresh each time since its rates change monthly) and, optionally, the SoR
 * (Schedule of Rates) PDF as a fallback rate source, then an Estimate. Every
 * estimate line's description is retrieved against the combined Data
 * Sheet + SoR rate index (hybrid TF-IDF/embedding search, core/technicalSanction.ts)
 * — the Data Sheet is the locally lead-adjusted, authoritative source, the
 * SoR fills in items the Data Sheet doesn't carry. A clean single match
 * auto-fills the rate and replaces the description text (green), a
 * description with several different rates attached to it (a common
 * pattern — e.g. manual vs. mechanical means × depth range, all sharing one
 * paragraph) is left for the user to pick from, and anything below the
 * match threshold is left red for manual entry. "Generate" writes these
 * edits into a copy of the original estimate file.
 */
export default function GiveTechnicalSanctionTab() {
  // Which tile's workspace (if any) has taken over the tab: the Data Sheet/SoR
  // rate lookup above, or Replicate Technical Sanction below. null = the tile
  // grid picker.
  const [mode, setMode] = useState<'lookup' | 'replicate' | null>(null)

  const [dataSheetName, setDataSheetName] = useState<string | null>(null)
  const [dataRateIndex, setDataRateIndex] = useState<RateEntry[] | null>(null)
  const [rateSheetCount, setRateSheetCount] = useState(0)
  const [dataError, setDataError] = useState<string | null>(null)
  const [dataBusy, setDataBusy] = useState(false)

  const [sorFileName, setSorFileName] = useState<string | null>(null)
  const [sorRateIndex, setSorRateIndex] = useState<RateEntry[] | null>(null)
  const [sorError, setSorError] = useState<string | null>(null)
  const [sorBusy, setSorBusy] = useState(false)
  const [sorProgress, setSorProgress] = useState<{ page: number; total: number } | null>(null)
  const sorFileInputRef = useRef<HTMLInputElement>(null)

  const [combinedEmbeddings, setCombinedEmbeddings] = useState<number[][] | null>(null)

  // The estimate workbook's sheets (usually one — a workbook can also spread
  // one estimate's items across several, each read and matched in turn).
  const [estimate, setEstimate] = useState<SheetGrid[] | null>(null)
  const [matches, setMatches] = useState<ItemMatch[] | null>(null)
  const [selections, setSelections] = useState<(number | null)[]>([])
  const [estimateError, setEstimateError] = useState<string | null>(null)
  const [estimateBusy, setEstimateBusy] = useState(false)
  const [embedNotice, setEmbedNotice] = useState<string | null>(null)
  // Column labels (Quantity/Rate/Unit/Serial Number) only resolved via semantic matching, not a plain header match.
  const [columnsAiAssisted, setColumnsAiAssisted] = useState<string[]>([])

  const [saved, setSaved] = useState<string | null>(null)
  const [generating, setGenerating] = useState(false)
  const [generateError, setGenerateError] = useState<string | null>(null)

  // Replicate Technical Sanction: a separate workflow (see
  // core/technicalSanction.ts's matchByExactDescription) — exact-description
  // reuse of an already-sanctioned estimate's rates, not the Data Sheet/SoR
  // keyword/semantic lookup above.
  const [oldItems, setOldItems] = useState<EstimateWorkItem[]>([])
  const [oldEstimateNames, setOldEstimateNames] = useState<string[]>([])
  const [oldEstimateError, setOldEstimateError] = useState<string | null>(null)
  const [oldEstimateBusy, setOldEstimateBusy] = useState(false)

  // Each uploaded "new estimate" file is an independent work needing its own
  // Technical Sanction output — never merged with another file's items —
  // so these all live as one entry per file rather than singular state.
  interface NewEstimateEntry {
    id: string
    grids: SheetGrid[]
    matches: ReplicateMatch[] | null
    columnsAiAssisted: string[]
    cementSteelCellEdits: CellEdit[]
    cementSteelFillInfo: { sheetName: string; month: string; year: number; cementRate?: string; steelRate?: string }[]
    cementSteelFillError: string | null
    matchError: string | null
    saved: string | null
    generating: boolean
    generateError: string | null
  }
  const [newEstimates, setNewEstimates] = useState<NewEstimateEntry[]>([])
  const [newEstimateError, setNewEstimateError] = useState<string | null>(null)
  const [newEstimateBusy, setNewEstimateBusy] = useState(false)

  // Auto-fills a new estimate's own "Cement (<Month> <Year>)=Rs…" / "Steel
  // HYSD(<Month> <Year>) =Rs…" header cells (see core/cementSteelRate.ts's
  // findCementSteelRateCells) straight from that month's departmental
  // circular, so the replicated Technical Sanction carries the current rate
  // rather than whatever the old sanctioned estimate cited. Tracks whichever
  // file/sheet is currently being processed while uploadNewEstimates loops
  // through every selected file.
  const [cementSteelFillPhase, setCementSteelFillPhase] = useState<'finding' | 'downloading' | 'reading' | 'updating' | null>(
    null
  )
  const [cementSteelFillTarget, setCementSteelFillTarget] = useState<{
    fileIndex: number
    fileTotal: number
    month: string
    year: number
    sheetIndex: number
    sheetTotal: number
  } | null>(null)

  // Semantic matching is an enhancement, not a hard requirement — if the
  // bundled model can't be loaded for any reason, fall back to keyword
  // (TF-IDF) matching alone rather than blocking the whole feature.
  async function embedSafely(texts: string[]): Promise<number[][] | null> {
    try {
      return await api.embedTexts(texts)
    } catch (e) {
      setEmbedNotice(
        `Semantic matching unavailable (${e instanceof Error ? e.message : String(e)}) — using keyword matching only.`
      )
      return null
    }
  }

  async function runMatching(grids: SheetGrid[], index: RateEntry[], entryVectors: number[][] | null) {
    const items: EstimateWorkItem[] = []
    const aiAssisted = new Set<string>()
    for (const grid of grids) {
      // A multi-sheet workbook can carry sheets that aren't an estimate table
      // at all (an empty sheet, a summary table with different columns) —
      // skip those rather than letting one bad sheet abort the whole upload.
      try {
        const extracted = await extractEstimateItemsWithAi(grid.grid, guessHeaderRow(grid.grid))
        extracted.aiAssisted.forEach((a) => aiAssisted.add(a))
        items.push(...extracted.items.map((it) => ({ ...it, sheetName: grid.sheetName })))
      } catch {
        continue
      }
    }
    setColumnsAiAssisted([...aiAssisted])
    if (items.length === 0) {
      throw new Error('No work items with a quantity, rate, and unit were found in that estimate.')
    }
    let embeddings: { itemVectors: number[][]; entryVectors: number[][] } | undefined
    if (entryVectors) {
      const itemVectors = await embedSafely(items.map((it) => it.description))
      if (itemVectors) embeddings = { itemVectors, entryVectors }
    }
    const m = matchEstimateItems(items, index, embeddings)
    setMatches(m)
    // Only auto-select a match confident enough to fill without review (exact
    // code/grade, or a strong-similarity single candidate). A weak/semantic-only
    // guess — including a medium-confidence lookalike drawn by an item whose true
    // rate isn't in the Data Sheet at all — is left for the user to confirm.
    setSelections(m.map((x) => (isAutoResolved(x) ? 0 : null)))
  }

  // Re-embeds the combined Data Sheet + SoR rate index and re-runs matching
  // against the currently loaded estimate (if any) — called after either
  // source changes, since a new source shifts the whole retrieval corpus.
  async function refreshCombined(dataIdx: RateEntry[] | null, sorIdx: RateEntry[] | null) {
    const combined = [...(dataIdx ?? []), ...(sorIdx ?? [])]
    setSaved(null)
    const entryVectors = combined.length > 0 ? await embedSafely(combined.map((e) => e.description)) : null
    setCombinedEmbeddings(entryVectors)
    if (estimate) await runMatching(estimate, combined, entryVectors)
  }

  async function uploadDataSheet() {
    setDataError(null)
    setEmbedNotice(null)
    setDataBusy(true)
    try {
      const sheets = await api.pickDataSheet()
      if (!sheets) return
      const index = buildRateIndex(sheets.map((s) => ({ name: s.sheetName, grid: s.grid })))
      setDataRateIndex(index)
      setRateSheetCount(sheets.length)
      setDataSheetName(sheets[0]?.name ?? null)
      await refreshCombined(index, sorRateIndex)
    } catch (e) {
      setDataError(e instanceof Error ? e.message : String(e))
    } finally {
      setDataBusy(false)
    }
  }

  async function uploadSor(file: File) {
    setSorError(null)
    setEmbedNotice(null)
    setSorBusy(true)
    setSorProgress({ page: 0, total: 0 })
    try {
      const pages = await pdfPagesToTextLinesFromData(await file.arrayBuffer(), (page, total) => setSorProgress({ page, total }))
      setSorProgress(null) // page extraction done — the rest (parsing/embedding) has no per-step progress to show
      const index = buildSorRateIndex(pages, file.name)
      if (index.length === 0) {
        throw new Error(`No rate items recognized in ${file.name} — is this the SoR item-table PDF?`)
      }
      setSorRateIndex(index)
      setSorFileName(file.name)
      await refreshCombined(dataRateIndex, index)
    } catch (e) {
      setSorError(e instanceof Error ? e.message : String(e))
    } finally {
      setSorBusy(false)
      setSorProgress(null)
    }
  }

  async function uploadEstimate() {
    setEstimateError(null)
    setSaved(null)
    setEstimateBusy(true)
    setColumnsAiAssisted([])
    let grids: SheetGrid[] | null = null
    try {
      grids = await api.pickEstimateGrids()
      if (!grids || grids.length === 0) return
      setEstimate(grids)
      setMatches(null)
      const combined = [...(dataRateIndex ?? []), ...(sorRateIndex ?? [])]
      if (combined.length > 0) await runMatching(grids, combined, combinedEmbeddings)
    } catch (e) {
      const hint = grids ? mismatchHint(grids[0].grid[guessHeaderRow(grids[0].grid)] ?? [], 'estimate') : null
      setEstimateError((e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : ''))
    } finally {
      setEstimateBusy(false)
    }
  }

  function pick(itemIndex: number, candidateIndex: number | null) {
    setSelections((prev) => {
      const next = [...prev]
      next[itemIndex] = candidateIndex
      return next
    })
  }

  async function generate() {
    if (!estimate || !matches) return
    setGenerating(true)
    setGenerateError(null)
    try {
      const resolved: ResolvedItem[] = matches.map((m, i) => ({
        item: m.item,
        chosen: selections[i] != null ? m.candidates[selections[i] as number] : null
      }))
      const edits = buildCellEdits(resolved)
      const rateAnalysisRows = buildRateAnalysisSheet(resolved)
      const suggestedName = `${stripExt(estimate[0].name)} Technical Sanction`
      const path = await api.generateTechnicalSanction(
        estimate[0].path,
        estimate[0].sheetName,
        edits,
        suggestedName,
        rateAnalysisRows
      )
      if (path) setSaved(path)
    } catch (e) {
      setGenerateError(e instanceof Error ? e.message : String(e))
    } finally {
      setGenerating(false)
    }
  }

  async function matchAgainstOldItems(
    grids: SheetGrid[],
    items: EstimateWorkItem[]
  ): Promise<{ matches: ReplicateMatch[]; columnsAiAssisted: string[] }> {
    const index = buildReplicateIndex(items)
    const newItems: EstimateWorkItem[] = []
    const aiAssisted = new Set<string>()
    for (const grid of grids) {
      try {
        const extracted = await extractEstimateItemsWithAi(grid.grid, guessHeaderRow(grid.grid))
        extracted.aiAssisted.forEach((a) => aiAssisted.add(a))
        newItems.push(...extracted.items.map((it) => ({ ...it, sheetName: grid.sheetName })))
      } catch {
        continue
      }
    }
    if (newItems.length === 0) {
      throw new Error('No work items with a quantity, rate, and unit were found in that estimate.')
    }
    return { matches: matchByExactDescription(newItems, index), columnsAiAssisted: [...aiAssisted] }
  }

  /** Re-runs matching for one already-uploaded new-estimate entry (e.g. after more sanctioned estimates are added to the rate pool) without re-reading its file. */
  async function runReplicateMatchingFor(id: string, items: EstimateWorkItem[]) {
    const entry = newEstimates.find((e) => e.id === id)
    if (!entry) return
    try {
      const { matches, columnsAiAssisted } = await matchAgainstOldItems(entry.grids, items)
      setNewEstimates((prev) =>
        prev.map((e) => (e.id === id ? { ...e, matches, columnsAiAssisted, matchError: null } : e))
      )
    } catch (e) {
      setNewEstimates((prev) =>
        prev.map((e) => (e.id === id ? { ...e, matchError: e instanceof Error ? e.message : String(e) } : e))
      )
    }
  }

  /** Finds any "Cement (<Month> <Year>)=Rs…" / "Steel HYSD(<Month> <Year>)
   * =Rs…" cells in the uploaded estimate and, if there's internet, fetches
   * each cited month's circular from the department site (same list the
   * sidebar's Cement & Steel Rates tab browses), OCRs it, and refreshes
   * those cells' rate in place. Never blocks the rest of Replicate Technical
   * Sanction — on any failure that sheet's original rate is left untouched.
   *
   * Grouped per sheet (not treated as one workbook-wide target), because an
   * uploaded workbook can bundle several unrelated estimates each citing a
   * different rate month (e.g. TS FOR KOMPALLY.xlsx's two sheets) — using
   * only the first match found would silently apply its month's rate to
   * every other sheet too, or skip them if the first sheet's circular
   * wasn't found. Circulars already fetched for a period are cached so two
   * sheets citing the same month only trigger one download.
   *
   * `onProgress` reports which sheet (of this file) is being worked on —
   * the caller additionally knows which file (of a multi-file upload) this
   * is, so file-level position is layered on by the caller, not here. */
  async function autoFillCementSteelRates(
    grids: SheetGrid[],
    onProgress: (phase: 'finding' | 'downloading' | 'reading' | 'updating', month: string, year: number, sheetIndex: number, sheetTotal: number) => void
  ): Promise<{
    edits: CellEdit[]
    infos: { sheetName: string; month: string; year: number; cementRate?: string; steelRate?: string }[]
    error: string | null
  }> {
    const cells = findCementSteelRateCells(grids)
    if (cells.length === 0) return { edits: [], infos: [], error: null }

    const bySheet = new Map<string, CementSteelRateCellRef[]>()
    for (const cell of cells) {
      const arr = bySheet.get(cell.sheetName)
      if (arr) arr.push(cell)
      else bySheet.set(cell.sheetName, [cell])
    }
    const sheetGroups = [...bySheet.entries()]

    const rateCache = new Map<number, CementSteelRates>()
    const failCache = new Map<number, string>()
    let sheetIndexForProgress = 0

    async function ratesForPeriod(month: string, year: number): Promise<CementSteelRates> {
      const monthIndex = CEMENT_STEEL_MONTHS.indexOf(month)
      const target = year * 12 + monthIndex
      if (failCache.has(target)) throw new Error(failCache.get(target)!)
      const cached = rateCache.get(target)
      if (cached) return cached
      try {
        onProgress('finding', month, year, sheetIndexForProgress, sheetGroups.length)
        const rates = await api.fetchCementSteelRates()
        const match = rates.find((r) => cementSteelRatePeriodKey(r) === target)
        if (!match) {
          throw new Error(`No Cement & Steel Rates circular found for ${month} ${year} on the department site yet.`)
        }

        onProgress('downloading', month, year, sheetIndexForProgress, sheetGroups.length)
        const bytes = await api.fetchCementSteelRateBytes(match.token)
        const dataUrls = await pdfPagesToDataUrlsFromData(bytes)

        onProgress('reading', month, year, sheetIndexForProgress, sheetGroups.length)
        const sheet = await api.ocrEstimatePhotos(dataUrls)
        const lines = sheet.grid.map((row) => row[0] ?? '')
        const parsed = parseCementSteelRateLines(lines)
        rateCache.set(target, parsed)
        return parsed
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e)
        failCache.set(target, msg)
        throw new Error(msg)
      }
    }

    const edits: CellEdit[] = []
    const infos: { sheetName: string; month: string; year: number; cementRate?: string; steelRate?: string }[] = []
    const errors: string[] = []

    for (let i = 0; i < sheetGroups.length; i++) {
      const [sheetName, sheetCells] = sheetGroups[i]
      const { month, year } = sheetCells[0]
      sheetIndexForProgress = i + 1
      try {
        const parsed = await ratesForPeriod(month, year)
        onProgress('updating', month, year, sheetIndexForProgress, sheetGroups.length)
        for (const cell of sheetCells) {
          const newRate = cell.kind === 'cement' ? parsed.cementRate : parsed.steelRate
          if (!newRate) continue
          const value = cell.raw.replace(cell.oldRateText, newRate.replace(/,/g, ''))
          edits.push({ row: cell.row, col: cell.col, sheetName: cell.sheetName, value })
        }
        infos.push({ sheetName, month, year, cementRate: parsed.cementRate, steelRate: parsed.steelRate })
      } catch (e) {
        errors.push(
          `Sheet "${sheetName}" (${month} ${year}): ${e instanceof Error ? e.message : String(e)} — its own original rate was left unchanged.`
        )
      }
    }

    return { edits, infos, error: errors.length > 0 ? errors.join(' ') : null }
  }

  async function uploadOldEstimate() {
    setOldEstimateError(null)
    setOldEstimateBusy(true)
    try {
      const files = await api.pickEstimateGridsMulti()
      if (!files || files.length === 0) return
      const items: EstimateWorkItem[] = []
      const names: string[] = []
      for (const grids of files) {
        names.push(grids[0]?.name ?? 'Estimate')
        for (const grid of grids) {
          try {
            const extracted = await extractEstimateItemsWithAi(grid.grid, guessHeaderRow(grid.grid))
            items.push(...extracted.items)
          } catch {
            continue
          }
        }
      }
      if (items.length === 0) {
        throw new Error('No work items with a quantity, rate, and unit were found in those estimate(s).')
      }
      const merged = [...oldItems, ...items]
      setOldItems(merged)
      setOldEstimateNames((prev) => [...prev, ...names])
      setNewEstimates((prev) => prev.map((e) => ({ ...e, saved: null })))
      for (const entry of newEstimates) {
        await runReplicateMatchingFor(entry.id, merged)
      }
    } catch (e) {
      setOldEstimateError(e instanceof Error ? e.message : String(e))
    } finally {
      setOldEstimateBusy(false)
    }
  }

  async function uploadNewEstimates() {
    setNewEstimateError(null)
    setNewEstimateBusy(true)
    try {
      const files = await api.pickEstimateGridsMulti()
      if (!files || files.length === 0) return
      for (let i = 0; i < files.length; i++) {
        const grids = files[i]
        const id = `${Date.now()}-${i}`
        const entry: NewEstimateEntry = {
          id,
          grids,
          matches: null,
          columnsAiAssisted: [],
          cementSteelCellEdits: [],
          cementSteelFillInfo: [],
          cementSteelFillError: null,
          matchError: null,
          saved: null,
          generating: false,
          generateError: null
        }
        try {
          if (oldItems.length > 0) {
            const { matches, columnsAiAssisted } = await matchAgainstOldItems(grids, oldItems)
            entry.matches = matches
            entry.columnsAiAssisted = columnsAiAssisted
          }
        } catch (e) {
          const hint = mismatchHint(grids[0].grid[guessHeaderRow(grids[0].grid)] ?? [], 'estimate')
          entry.matchError = `${e instanceof Error ? e.message : String(e)}${hint ? ` ${hint}` : ''}`
        }
        const fill = await autoFillCementSteelRates(grids, (phase, month, year, sheetIndex, sheetTotal) => {
          setCementSteelFillPhase(phase)
          setCementSteelFillTarget({ fileIndex: i + 1, fileTotal: files.length, month, year, sheetIndex, sheetTotal })
        })
        entry.cementSteelCellEdits = fill.edits
        entry.cementSteelFillInfo = fill.infos
        entry.cementSteelFillError = fill.error
        setNewEstimates((prev) => [...prev, entry])
      }
    } catch (e) {
      setNewEstimateError(e instanceof Error ? e.message : String(e))
    } finally {
      setNewEstimateBusy(false)
      setCementSteelFillPhase(null)
      setCementSteelFillTarget(null)
    }
  }

  async function generateReplicateFor(id: string) {
    const entry = newEstimates.find((e) => e.id === id)
    if (!entry || !entry.matches) return
    setNewEstimates((prev) => prev.map((e) => (e.id === id ? { ...e, generating: true, generateError: null } : e)))
    try {
      const edits = [...buildReplicateCellEdits(entry.matches), ...entry.cementSteelCellEdits]
      const suggestedName = `${stripExt(entry.grids[0].name)} Technical Sanction`
      const path = await api.generateTechnicalSanction(entry.grids[0].path, entry.grids[0].sheetName, edits, suggestedName)
      setNewEstimates((prev) =>
        prev.map((e) => (e.id === id ? { ...e, saved: path ?? e.saved, generating: false } : e))
      )
    } catch (e) {
      setNewEstimates((prev) =>
        prev.map((e) => (e.id === id ? { ...e, generateError: e instanceof Error ? e.message : String(e), generating: false } : e))
      )
    }
  }

  function clearLookup() {
    setDataSheetName(null)
    setDataRateIndex(null)
    setRateSheetCount(0)
    setDataError(null)
    setSorFileName(null)
    setSorRateIndex(null)
    setSorError(null)
    setSorProgress(null)
    setCombinedEmbeddings(null)
    setEmbedNotice(null)
    setEstimate(null)
    setMatches(null)
    setSelections([])
    setEstimateError(null)
    setColumnsAiAssisted([])
    setSaved(null)
    setGenerateError(null)
  }

  function clearReplicate() {
    setOldItems([])
    setOldEstimateNames([])
    setOldEstimateError(null)
    setNewEstimates([])
    setNewEstimateError(null)
    setCementSteelFillPhase(null)
    setCementSteelFillTarget(null)
  }

  const CEMENT_STEEL_FILL_STEPS = {
    finding: { pct: 25, label: 'Finding circular…' },
    downloading: { pct: 50, label: 'Downloading circular…' },
    reading: { pct: 75, label: 'Reading cement & steel rates…' },
    updating: { pct: 100, label: 'Updating estimate…' }
  } as const

  const combinedCount = (dataRateIndex?.length ?? 0) + (sorRateIndex?.length ?? 0)

  // Resolved = anything now carrying a chosen rate, auto-filled or user-picked.
  // Needs-pick = reviewable matches still awaiting a choice. Not-found = red.
  const foundCount = matches ? selections.filter((s) => s != null).length : 0
  const ambiguousCount = matches ? matches.filter((m, i) => needsReview(m) && selections[i] == null).length : 0
  const notFoundCount = matches ? matches.filter((m) => m.candidates.length === 0).length : 0

  if (mode === null) {
    return (
      <div className="card">
        <div className="doc-tile-grid tools-grid ts-mode-grid">
          <button className="doc-tile-card tone-green tool-card" onClick={() => setMode('lookup')}>
            <span className="tool-card-ic">
              <IconTable />
            </span>
            <span className="doc-tile-card-name">Rate Lookup</span>
            <span className="doc-tile-card-meta">Data Sheet / SoR keyword &amp; semantic rate lookup</span>
            <span className="tool-card-cta">
              <IconFolder /> Open workspace
            </span>
          </button>

          <button className="doc-tile-card tone-green tool-card" onClick={() => setMode('replicate')}>
            <span className="tool-card-ic">
              <IconRefresh />
            </span>
            <span className="doc-tile-card-name">Replicate Technical Sanction</span>
            <span className="doc-tile-card-meta">Reuse rates from an already-sanctioned estimate, by matching description</span>
            <span className="tool-card-cta">
              <IconFolder /> Open workspace
            </span>
          </button>
        </div>
      </div>
    )
  }

  if (mode === 'replicate') {
    return (
      <>
        <div className="pdf-ws-topbar">
          <div className="pdf-ws-title">
            <IconRefresh /> <span className="pdf-ws-title-label">Tool:</span> Replicate Technical Sanction
          </div>
          <button className="ghost" onClick={clearReplicate}>
            <IconTrash /> Clear
          </button>
          <button className="ghost pdf-ws-back" onClick={() => setMode(null)}>
            <IconChevronLeft /> Back
          </button>
        </div>
        <div className="card pdf-workspace">
          <div className="empty">
            <IconRefresh />
            <p>
              {oldEstimateNames.length > 0
                ? `${oldItems.length} items loaded across ${oldEstimateNames.length} already-sanctioned estimate${oldEstimateNames.length === 1 ? '' : 's'} (“${oldEstimateNames.join('”, “')}”).`
                : 'Upload one or more already Technical Sanctioned estimates, as the source of rates to reuse.'}
            </p>
            <div className="boq-actions">
              <button className="primary upload-btn" onClick={uploadOldEstimate} disabled={oldEstimateBusy}>
                <IconFolder /> {oldEstimateBusy ? 'Loading…' : oldEstimateNames.length > 0 ? 'Add Another Sanctioned Estimate' : 'Upload Sanctioned Estimate'}
              </button>
            </div>
            {oldEstimateError && (
              <div className="notice error">
                <IconWarn />
                {oldEstimateError}
              </div>
            )}

            <p style={{ marginTop: 18 }}>
              {newEstimates.length > 0
                ? `${newEstimates.length} new estimate${newEstimates.length === 1 ? '' : 's'} loaded.`
                : 'Upload the new estimate(s) whose rates should be filled in from the sanctioned estimate(s) above.'}
            </p>
            <div className="boq-actions">
              <button className="primary upload-btn" onClick={uploadNewEstimates} disabled={newEstimateBusy}>
                <IconFolder />{' '}
                {newEstimateBusy
                  ? 'Loading & computing matches…'
                  : newEstimates.length > 0
                    ? 'Add More New Estimates'
                    : 'Upload New Estimate(s)'}
              </button>
            </div>
            {newEstimateError && (
              <div className="notice error">
                <IconWarn />
                {newEstimateError}
              </div>
            )}
            {cementSteelFillPhase && cementSteelFillTarget && (
              <div className="progress" style={{ margin: '10px 0' }}>
                <div className="progress-bar">
                  <div style={{ width: `${CEMENT_STEEL_FILL_STEPS[cementSteelFillPhase].pct}%` }} />
                </div>
                <div className="progress-label">
                  File {cementSteelFillTarget.fileIndex} of {cementSteelFillTarget.fileTotal} — Sheet{' '}
                  {cementSteelFillTarget.sheetIndex} of {cementSteelFillTarget.sheetTotal} —{' '}
                  {CEMENT_STEEL_FILL_STEPS[cementSteelFillPhase].label} ({cementSteelFillTarget.month}{' '}
                  {cementSteelFillTarget.year} Cement &amp; Steel rate)
                </div>
              </div>
            )}
          </div>

          {newEstimates.map((entry) => {
            const foundCount = entry.matches ? entry.matches.filter((m) => m.rate !== undefined).length : 0
            const notFoundCount = entry.matches ? entry.matches.filter((m) => m.rate === undefined).length : 0
            return (
              <div key={entry.id} className="card ts-replicate-card">
                <div className="boq-actions" style={{ justifyContent: 'space-between' }}>
                  <strong>
                    {entry.grids[0].name}
                    {entry.grids.length > 1 ? ` (${entry.grids.length} sheets)` : ''}
                  </strong>
                  {entry.matches && (
                    <button className="primary" onClick={() => generateReplicateFor(entry.id)} disabled={entry.generating}>
                      <IconDownload /> {entry.generating ? 'Saving…' : 'Generate Technical Sanction'}
                    </button>
                  )}
                </div>
                {entry.matchError && (
                  <div className="notice error">
                    <IconWarn />
                    {entry.matchError}
                  </div>
                )}
                {entry.cementSteelFillError && (
                  <div className="notice error">
                    <IconWarn />
                    {entry.cementSteelFillError}
                  </div>
                )}
                {entry.cementSteelFillInfo.length > 0 && (
                  <div className="notice ok">
                    <IconCheck />
                    {entry.cementSteelFillInfo.map((info, i) => (
                      <div key={i}>
                        “{info.sheetName}” — {info.month} {info.year} rates filled —{' '}
                        {info.cementRate ? `Cement ${formatRate(info.cementRate)}` : 'Cement not found'}
                        {', '}
                        {info.steelRate ? `Steel ${formatRate(info.steelRate)}` : 'Steel not found'}
                      </div>
                    ))}
                  </div>
                )}
                {entry.columnsAiAssisted.length > 0 && (
                  <div className="notice warn">
                    <IconWarn />
                    {entry.columnsAiAssisted.join(', ')} column{entry.columnsAiAssisted.length === 1 ? '' : 's'} matched by AI —
                    please double-check the estimate's items below.
                  </div>
                )}
                {entry.generateError && (
                  <div className="notice error">
                    <IconWarn />
                    {entry.generateError}
                  </div>
                )}
                {entry.saved && (
                  <div className="notice ok">
                    <IconCheck />
                    Saved to {entry.saved}
                  </div>
                )}
                {entry.matches && (
                  <>
                    <p className="hint" style={{ margin: '10px 0 0' }}>
                      {foundCount} rate{foundCount === 1 ? '' : 's'} found · {notFoundCount} not found
                    </p>
                    <ul className="todo-list boq-entries ts-replicate-items-scroll" style={{ margin: '10px 0 0' }}>
                      {entry.matches.map((m, i) => (
                        <li key={i} className="todo-item">
                          <div className="todo-body">
                            <span
                              className="todo-text"
                              title={m.item.description}
                              style={{
                                background: m.rate !== undefined ? '#c6efce' : '#ffc7ce',
                                borderRadius: 4,
                                padding: '1px 4px'
                              }}
                            >
                              {m.item.description.length > 100 ? `${m.item.description.slice(0, 100)}…` : m.item.description}
                            </span>
                            <span className="todo-dates" style={{ color: m.rate !== undefined ? '#008000' : '#cc0000' }}>
                              {m.rate !== undefined ? `Rate found: ${formatRate(m.rate)}` : 'Not found — left for manual entry'}
                            </span>
                          </div>
                        </li>
                      ))}
                    </ul>
                  </>
                )}
              </div>
            )
          })}
        </div>
      </>
    )
  }

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconTable /> <span className="pdf-ws-title-label">Tool:</span> Rate Lookup
        </div>
        <button className="ghost" onClick={clearLookup}>
          <IconTrash /> Clear
        </button>
        <button className="ghost pdf-ws-back" onClick={() => setMode(null)}>
          <IconChevronLeft /> Back
        </button>
      </div>
      <div className="card pdf-workspace">
      <div className="empty">
        <IconTable />
        <p>
          {dataRateIndex
            ? `Data Sheet loaded — ${dataRateIndex.length} rate entries across ${rateSheetCount} sheets (“${dataSheetName}”).`
            : 'Upload the Data Sheet (rates database) — since rates update monthly, upload the latest copy each time.'}
        </p>
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={uploadDataSheet} disabled={dataBusy}>
            <IconFolder /> {dataBusy ? 'Loading & computing matches…' : dataRateIndex ? 'Replace Data Sheet' : 'Upload Data Sheet'}
          </button>
        </div>
        {dataError && (
          <div className="notice error">
            <IconWarn />
            {dataError}
          </div>
        )}

        <p style={{ marginTop: 18 }}>
          {sorRateIndex
            ? `SoR PDF loaded — ${sorRateIndex.length} rate entries (“${sorFileName}”).`
            : 'Optionally also upload the SoR (Schedule of Rates) PDF, as a fallback source for items the Data Sheet doesn’t carry.'}
        </p>
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={() => sorFileInputRef.current?.click()} disabled={sorBusy}>
            <IconFolder />{' '}
            {sorBusy
              ? sorProgress && sorProgress.total > 0
                ? `Reading page ${sorProgress.page}/${sorProgress.total} (${Math.round((sorProgress.page / sorProgress.total) * 100)}%)…`
                : 'Computing matches…'
              : sorRateIndex
                ? 'Replace SoR PDF'
                : 'Upload SoR PDF'}
          </button>
          <input
            ref={sorFileInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void uploadSor(file)
              e.target.value = ''
            }}
          />
        </div>
        {sorError && (
          <div className="notice error">
            <IconWarn />
            {sorError}
          </div>
        )}
        {combinedEmbeddings && (
          <p className="hint" style={{ marginTop: 6 }}>
            {combinedCount} combined rate entries, with semantic matching enabled.
          </p>
        )}
        {embedNotice && (
          <div className="notice error">
            <IconWarn />
            {embedNotice}
          </div>
        )}

        <p style={{ marginTop: 18 }}>
          {estimate
            ? `Loaded “${estimate[0].name}”${estimate.length > 1 ? ` (${estimate.length} sheets)` : ''}.`
            : 'Upload the Estimate to match its items against the Data Sheet.'}
        </p>
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={uploadEstimate} disabled={estimateBusy}>
            <IconFolder /> {estimateBusy ? 'Loading & computing matches…' : 'Upload Estimate'}
          </button>
          {matches && (
            <button className="primary" onClick={generate} disabled={generating}>
              <IconDownload /> {generating ? 'Saving…' : 'Generate Technical Sanction'}
            </button>
          )}
        </div>
        {estimateError && (
          <div className="notice error">
            <IconWarn />
            {estimateError}
          </div>
        )}
        {columnsAiAssisted.length > 0 && (
          <div className="notice warn">
            <IconWarn />
            {columnsAiAssisted.join(', ')} column{columnsAiAssisted.length === 1 ? '' : 's'} matched by AI — please
            double-check the estimate's items below.
          </div>
        )}
        {generateError && (
          <div className="notice error">
            <IconWarn />
            {generateError}
          </div>
        )}
        {saved && (
          <div className="notice ok">
            <IconCheck />
            Saved to {saved}
          </div>
        )}
      </div>

      {matches && (
        <>
          <p className="hint" style={{ margin: '14px 18px 0' }}>
            {foundCount} rate{foundCount === 1 ? '' : 's'} resolved · {ambiguousCount} need your pick ·{' '}
            {notFoundCount} not found
          </p>
          <ul className="todo-list boq-entries" style={{ margin: '10px 18px 18px' }}>
            {matches.map((m, i) => (
              <li key={i} className="todo-item">
                <div className="todo-body">
                  <span className="todo-text" title={m.item.description}>
                    {m.item.description.length > 100 ? `${m.item.description.slice(0, 100)}…` : m.item.description}
                  </span>
                  {m.candidates.length === 0 && (
                    <span className="todo-dates" style={{ color: '#cc0000' }}>
                      Not found — left for manual entry
                    </span>
                  )}
                  {isAutoResolved(m) && (
                    <span className="todo-dates" style={{ color: '#008000' }}>
                      Rate found: {formatRate(m.candidates[0].rate)}
                    </span>
                  )}
                  {needsReview(m) && (
                    <label className="todo-dates" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ color: '#b8860b' }}>
                        {m.candidates.length > 1
                          ? `${m.candidates.length} possible rates — pick one:`
                          : m.semanticOnly
                            ? 'Matched by meaning only, not exact wording — please verify:'
                            : 'Low-confidence match — please verify:'}
                      </span>
                      <select
                        value={selections[i] ?? ''}
                        onChange={(e) => pick(i, e.target.value === '' ? null : Number(e.target.value))}
                      >
                        <option value="">Not resolved</option>
                        {m.candidates.map((c, ci) => (
                          <option key={ci} value={ci}>
                            {formatRate(c.rate)}
                          </option>
                        ))}
                      </select>
                    </label>
                  )}
                </div>
              </li>
            ))}
          </ul>
        </>
      )}
      </div>
    </>
  )
}
