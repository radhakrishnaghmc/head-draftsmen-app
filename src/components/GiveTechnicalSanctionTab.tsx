import { useRef, useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractEstimateItemsWithAi } from '../aiEstimateColumns'
import { mismatchHint } from '../docClassify'
import { buildRateIndex } from '@core/rateDatabase'
import type { RateEntry } from '@core/rateDatabase'
import { buildSorRateIndex } from '@core/sorPdfExtract'
import { pdfPagesToTextLinesFromData } from '../pdfToText'
import { matchEstimateItems, buildCellEdits, buildRateAnalysisSheet, isAutoResolved, needsReview } from '@core/technicalSanction'
import type { ItemMatch, ResolvedItem } from '@core/technicalSanction'
import { IconFolder, IconDownload, IconWarn, IconCheck, IconTable } from './Icons'
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

  const [estimate, setEstimate] = useState<SheetGrid | null>(null)
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

  async function runMatching(grid: SheetGrid, index: RateEntry[], entryVectors: number[][] | null) {
    const { items, aiAssisted } = await extractEstimateItemsWithAi(grid.grid, guessHeaderRow(grid.grid))
    setColumnsAiAssisted(aiAssisted)
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
    let grid: SheetGrid | null = null
    try {
      grid = await api.pickEstimateGrid()
      if (!grid) return
      setEstimate(grid)
      setMatches(null)
      const combined = [...(dataRateIndex ?? []), ...(sorRateIndex ?? [])]
      if (combined.length > 0) await runMatching(grid, combined, combinedEmbeddings)
    } catch (e) {
      const hint = grid ? mismatchHint(grid.grid[guessHeaderRow(grid.grid)] ?? [], 'estimate') : null
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
      const suggestedName = `${stripExt(estimate.name)} Technical Sanction`
      const path = await api.generateTechnicalSanction(
        estimate.path,
        estimate.sheetName,
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

  const combinedCount = (dataRateIndex?.length ?? 0) + (sorRateIndex?.length ?? 0)

  // Resolved = anything now carrying a chosen rate, auto-filled or user-picked.
  // Needs-pick = reviewable matches still awaiting a choice. Not-found = red.
  const foundCount = matches ? selections.filter((s) => s != null).length : 0
  const ambiguousCount = matches ? matches.filter((m, i) => needsReview(m) && selections[i] == null).length : 0
  const notFoundCount = matches ? matches.filter((m) => m.candidates.length === 0).length : 0

  return (
    <div className="card">
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
          {estimate ? `Loaded “${estimate.name}”.` : 'Upload the Estimate to match its items against the Data Sheet.'}
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
  )
}
