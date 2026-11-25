import { useState } from 'react'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { stripItemNoTag } from '@core/bidDocument'
import { compareWorkNames } from '@core/workNameMatch'
import { circleFromWorkName } from '../zoneCircleCheck'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractEstimateItems, extractWorkName } from '@core/estimateExtract'
import { computeWorkAmounts, indianDigitGroups, rupeesFromCell, tenderPercentFromRow } from '@core/worksAmounts'
import { buildWorksRowEmbeddingCache, matchWorksRowCached } from '../estimateDownloads'
import { inferTypeOfWork } from '@core/workTypeClassifier'
import type { DeviationItem, DeviationMeta } from '@core/deviationTemplate'
import type { ExcelTable, StoredLetter } from '@core/types'
import type { PickedTenderDocument } from '../../electron/ipc-contract'
import { IconChevronLeft, IconCheck, IconDoc, IconFolder, IconWarn } from './Icons'

interface Props {
  /** The office's Works List — each estimate's work is looked up there for its agency, circle, NIT and estimate amount. */
  worksTable?: ExcelTable | null
  /** Receives every Deviation Statement made — they land in Issued Documents → Deviation Statements. */
  onAddLetters?: (letters: StoredLetter[]) => void
  onBack: () => void
}

interface Result {
  file: string
  ok: boolean
  note: string
}

/** One estimate workbook read: the first sheet that reads as an estimate. */
interface ReadEstimate {
  file: string
  workName: string
  items: DeviationItem[]
}

/** The estimate's own contract value: the sum of its items' quantity × rate (what the L-1 sheet's ECV is the tender's figure for). */
const estimateEcv = (items: DeviationItem[]) =>
  items.reduce((t, it) => t + Math.round((Number(String(it.quantity).replace(/,/g, '')) || 0) * (Number(String(it.rate).replace(/,/g, '')) || 0)), 0)

/** Rupees of slack allowed between the estimate's ECV and the L-1 sheet's (rounding). */
const ECV_TOLERANCE = 2

const pdfName = (n: string) => n.replace(/\.[^.]+$/, '')

/** Bulk Deviation Statements. Estimates alone: the Works List supplies each work's details. With L-1 sheets as well: every L-1 work is
 * paired with the estimate of the same work name and filled from the L-1 (agency, tender %, NIT, ECV); an L-1 with no matching estimate
 * gets no Deviation Statement. Statements land in Issued Documents → Deviation Statements. */
export default function BulkDeviationsWorkspace({ worksTable, onAddLetters, onBack }: Props) {
  // Adding folders only records their paths; nothing is read until Generate Deviations is clicked.
  const [l1Paths, setL1Paths] = useState<string[]>([])
  const [estimatePaths, setEstimatePaths] = useState<string[]>([])
  const [progress, setProgress] = useState<{ label: string; done: number; total: number } | null>(null)
  const [results, setResults] = useState<Result[]>([])
  const [error, setError] = useState<string | null>(null)
  const busy = progress !== null
  const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p

  async function addPaths(kind: 'l1' | 'estimates') {
    setError(null)
    try {
      const picked = await api.pickBulkDeviationPaths(kind)
      if (!picked) return
      ;(kind === 'l1' ? setL1Paths : setEstimatePaths)(picked)
      setResults([])
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  async function generate() {
    setError(null)
    setResults([])
    try {
      // 1) Read every L-1 sheet.
      const l1s: { file: string; ev: TenderEvaluation }[] = []
      let pickedL1: PickedTenderDocument[] = []
      if (l1Paths.length > 0) {
        setProgress({ label: 'Finding L-1 sheets', done: 0, total: 0 })
        const stop = api.onTenderScanProgress((p) => setProgress({ label: p.phase === 'scanning' ? 'Finding L-1 sheets' : 'Opening L-1 sheets', done: p.done, total: p.total }))
        try {
          pickedL1 = (await api.readTenderPaths(l1Paths)).filter((d) => /\.pdf$/i.test(d.name))
        } finally {
          stop()
        }
      }
      for (const [i, p] of pickedL1.entries()) {
        setProgress({ label: 'Reading L-1 sheets', done: i, total: pickedL1.length })
        try {
          const ev = parseTenderEvaluation(await pdfToTextLinesOrOcr(new File([new Uint8Array(p.bytes)], p.name)))
          if (ev.nameOfWork && ev.l1AgencyName) l1s.push({ file: p.name, ev })
        } catch {
          // Not a readable L-1 page (an LOA, a scan that failed) — it just isn't one of the works.
        }
      }
      // 2) Read every estimate.
      const estimates: ReadEstimate[] = []
      const unreadEstimates: Result[] = []
      setProgress({ label: 'Finding estimates', done: 0, total: 0 })
      const workbooks = await api.listEstimateWorkbooks(estimatePaths)
      for (const [i, wb] of workbooks.entries()) {
        setProgress({ label: 'Reading estimates', done: i, total: workbooks.length })
        const sheets = await api.readEstimateWorkbook(wb)
        const file = baseName(wb)
        // A workbook often holds many estimates, one per sheet — every estimate sheet is a work of its own.
        const found: ReadEstimate[] = []
        let unnamed = 0
        for (const sh of sheets) {
          if (sh.hidden) continue
          const headerRow = guessHeaderRow(sh.grid)
          let items: ReturnType<typeof extractEstimateItems>
          try {
            items = extractEstimateItems(sh.grid, headerRow)
          } catch {
            continue // not an estimate sheet (cover page, abstract…) — try the workbook's next sheet
          }
          if (items.length === 0) continue
          const workName = stripItemNoTag(extractWorkName(sh.grid, headerRow) ?? '')
          if (!workName) {
            unnamed++
            continue
          }
          found.push({
            file: `${file} · ${sh.sheetName}`,
            workName,
            items: items.map((it) => ({ description: it.description, unit: it.unit, quantity: it.quantity, rate: it.rate }))
          })
        }
        if (found.length > 0) estimates.push(...found)
        else unreadEstimates.push({ file, ok: false, note: unnamed > 0 ? 'The work name could not be read from the estimate.' : 'No work items with a quantity, rate and unit were found.' })
      }
      // 3) Pair and generate.
      await build(l1s, estimates, unreadEstimates)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setProgress(null)
    }
  }

  async function build(l1s: { file: string; ev: TenderEvaluation }[], estimates: ReadEstimate[], unreadEstimates: Result[]) {
    setProgress({ label: 'Looking up the Works List', done: 0, total: 0 })
    const cache = worksTable ? await buildWorksRowEmbeddingCache(worksTable) : null
    const made: StoredLetter[] = []
    const out: Result[] = [...unreadEstimates]
    const today = new Date().toISOString().slice(0, 10)
    const stamp = Date.now()
    const letter = (n: number, e: ReadEstimate, d: { workName: string; agency: string; nitNo: string; nitDate: string; lakhs: number; circle?: string; tenderPercent?: number; extra?: Record<string, string> }) => {
      const meta: DeviationMeta = { circle: d.circle, nameOfWork: d.workName, agencyName: d.agency, estimateAmountLakhs: d.lakhs, tenderPercent: d.tenderPercent }
      made.push({
        id: `sl-${stamp}-dv${n}`,
        kind: 'deviation',
        nitNo: d.nitNo,
        nitDate: d.nitDate,
        agency: d.agency,
        workName: d.workName,
        typeOfWork: inferTypeOfWork(d.workName) ?? '',
        estimateRupees: d.lakhs > 0 ? Math.round(d.lakhs * 100000) : null,
        variant: 'default',
        values: { items: JSON.stringify(e.items), meta: JSON.stringify(meta), ...(d.extra ?? {}) },
        fileName: `Deviation Statement${d.agency ? ` - ${d.agency}` : ''}${d.lakhs > 0 ? ` - ${d.lakhs.toFixed(2)} Lakhs` : ''}`,
        createdDate: today
      })
    }
    // Exact name first, then the Works List's closest name (an L-1 / estimate title often words the work slightly differently).
    const ecvNote = (e: ReadEstimate, ecv?: number) => {
      if (ecv == null) return ' — ECV not on the L-1, not checked'
      const est = estimateEcv(e.items)
      return Math.abs(est - ecv) <= ECV_TOLERANCE
        ? ' — ECV matches the L-1'
        : ` — ECV MISMATCH: estimate ${indianDigitGroups(est)}/- vs L-1 ${indianDigitGroups(ecv)}/-`
    }
    const worksRow = async (name: string) => (worksTable ? (await matchWorksRowCached(name, worksTable, cache))?.row : undefined)
    const estimateLakhs = (row?: Record<string, string>) => (row ? computeWorkAmounts(row).estimate / 100000 : 0)

    if (l1s.length > 0) {
      // Every L-1 work gets its estimate: the one whose work name matches. No match → no Deviation Statement.
      const claimed = new Set<ReadEstimate>()
      for (const [i, { file, ev }] of l1s.entries()) {
        setProgress({ label: 'Generating deviations', done: i, total: l1s.length })
        await new Promise((r) => setTimeout(r, 0))
        const name = ev.nameOfWork ?? ''
        // Same work name; if several estimates share it, the one whose ECV equals the L-1's wins.
        const named = estimates.filter((e) => !claimed.has(e) && compareWorkNames(name, e.workName).status === 'match')
        const ecvOk = (e: ReadEstimate) => ev.ecvRupees != null && Math.abs(estimateEcv(e.items) - ev.ecvRupees) <= ECV_TOLERANCE
        // An L-1 ECV that no same-named estimate matches (within ₹2) means it's the wrong estimate: no Deviation Statement.
        const hit = ev.ecvRupees != null ? named.find(ecvOk) : named[0]
        if (!hit) {
          out.push({
            file: pdfName(file),
            ok: false,
            note: named.length
              ? `ECV mismatch: estimate ${named.map((e) => indianDigitGroups(estimateEcv(e.items))).join(' / ')}/- vs L-1 ${indianDigitGroups(ev.ecvRupees ?? 0)}/- — no Deviation Statement made.`
              : 'No estimate with the same work name — no Deviation Statement made.'
          })
          continue
        }
        claimed.add(hit)
        const row = await worksRow(name)
        const lakhs = estimateLakhs(row)
        letter(i, hit, {
          workName: stripItemNoTag(name),
          agency: ev.l1AgencyName ?? '',
          nitNo: ev.noticeNo ?? '',
          nitDate: ev.noticeDate ?? '',
          lakhs,
          circle: row?.['Circle'] || circleFromWorkName(name)?.circle,
          tenderPercent: ev.tenderPercentage,
          extra: { ...(ev.tenderId ? { tenderId: ev.tenderId } : {}), ...(ev.ecvRupees != null ? { ecvRupees: String(ev.ecvRupees) } : {}), ...(ev.contractRupees != null ? { contractRupees: String(ev.contractRupees) } : {}) }
        })
        out.push({
          file: pdfName(file),
          ok: true,
          note: `${hit.items.length} items from ${hit.file}${ecvNote(hit, ev.ecvRupees)}${ev.tenderPercentage == null ? ' — no tender percentage on the L-1, template 25% kept' : ''}`
        })
      }
    } else {
      // The Works List is this office's own database: an estimate whose work name isn't on it isn't this office's, so no Deviation Statement.
      const done = new Map<string, string>() // Works List work name -> estimate file that already made its deviation
      for (const [i, e] of estimates.entries()) {
        setProgress({ label: 'Generating deviations', done: i, total: estimates.length })
        await new Promise((r) => setTimeout(r, 0))
        const row = await worksRow(e.workName)
        if (!row) {
          out.push({ file: e.file, ok: false, note: 'Work not found in the Works List — no Deviation Statement made.' })
          continue
        }
        // One deviation per work: a repeated estimate, or two estimates closest to the same Works List row, would otherwise each make one.
        const key = (row['Name of the work'] ?? e.workName).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim()
        const first = done.get(key)
        if (first) {
          out.push({ file: e.file, ok: false, note: `Same work as ${first} — already has a Deviation Statement, skipped.` })
          continue
        }
        done.set(key, e.file)
        const pct = parseFloat(tenderPercentFromRow(row))
        const ecv = rupeesFromCell(row['ECV'] ?? '')
        const contract = rupeesFromCell(row['Contract Amount'] ?? '')
        const tenderId = (row['Tender ID'] ?? '').trim()
        letter(i, e, {
          workName: (row['Name of the work'] ?? '').trim() || e.workName,
          agency: (row['Name of the Agency'] ?? '').trim(),
          nitNo: (row['Tender Notice No'] ?? '').trim(),
          nitDate: (row['Tender notice Date'] ?? '').trim(),
          lakhs: estimateLakhs(row),
          circle: row['Circle'] || undefined,
          tenderPercent: Number.isFinite(pct) ? pct : undefined,
          extra: { ...(tenderId ? { tenderId } : {}), ...(ecv ? { ecvRupees: String(ecv) } : {}), ...(contract ? { contractRupees: String(contract) } : {}) }
        })
        out.push({ file: e.file, ok: true, note: `${e.items.length} items — details from the Works List` })
      }
    }
    if (made.length > 0) onAddLetters?.(made)
    setResults(out)
  }

  const madeCount = results.filter((r) => r.ok).length
  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">Bulk Deviations</div>
        <button className="ghost pdf-ws-back" onClick={onBack}>
          <IconChevronLeft /> Back
        </button>
      </div>
      <div className="card">
        <div className="card-head">
          <div className="head-ic">
            <IconDoc />
          </div>
          <div className="titles">
            <h2>Bulk Deviations</h2>
            <p className="sub">
              Upload the L-1 sheets and the technical-sanctioned estimates. Each L-1 work is matched to the estimate with the same work name and its Deviation Statement is added to Issued Documents → Deviation Statements; an L-1 with no matching estimate gets none.
            </p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="primary" onClick={() => void addPaths('l1')} disabled={busy}>
            <IconFolder /> Add L-1 sheets folder
          </button>
          <button className="primary" onClick={() => void addPaths('estimates')} disabled={busy}>
            <IconFolder /> Add estimates folder
          </button>
          <button className="primary" onClick={() => void generate()} disabled={busy || estimatePaths.length === 0}>
            <IconDoc /> Generate Deviations
          </button>
        </div>
        <p className="estimate-hint">
          {l1Paths.length > 0 ? `L-1 sheets: ${l1Paths.map(baseName).join(', ')}` : 'No L-1 folder added'} · {estimatePaths.length > 0 ? `Estimates: ${estimatePaths.map(baseName).join(', ')}` : 'No estimates folder added'}
          {l1Paths.length === 0 && estimatePaths.length > 0 ? ' — with no L-1 sheets, details come from the Works List.' : ''}
        </p>
        {progress && (
          <div className="bulk-progress" role="progressbar" aria-valuemin={0} aria-valuemax={progress.total} aria-valuenow={progress.done}>
            <span className="estimate-hint">
              {progress.label}
              {progress.total > 0 ? ` ${progress.done} of ${progress.total}` : '…'}
            </span>
            <progress value={progress.total > 0 ? progress.done : undefined} max={progress.total || undefined} style={{ width: '100%' }} />
          </div>
        )}
        {error && (
          <div className="notice error">
            <IconWarn />
            {error}
          </div>
        )}
        {results.length > 0 && (
          <>
            <p className="estimate-hint">
              {madeCount} of {results.length} Deviation Statement{madeCount === 1 ? '' : 's'} added to Issued Documents.
            </p>
            <ul className="bulk-deviation-results">
              {results.map((r, i) => (
                <li key={i}>
                  {r.ok && !r.note.includes("MISMATCH") ? <IconCheck /> : <IconWarn />} <strong>{r.file}</strong> — {r.note}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </>
  )
}
