import { useEffect, useRef, useState } from 'react'
import { api } from '../ipc'
import type { TenderScanProgress } from '../../electron/ipc-contract'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import type { ExcelTable } from '@core/types'
import { IconFolder, IconWarn } from './Icons'

interface Props {
  table: ExcelTable
  onChange: (table: ExcelTable) => void
  /** Report a successful update so the host can flash the matched rows and show the message under them. */
  onUpdated: (rowIndices: number[], message: string) => void
}

/**
 * "Update from L1": pick one or more L-1 selection sheets and fold their
 * tender details into the Works List by matching each sheet's Name of Work
 * to a row — Tender ID, Notice No/Date, ECV (with EMD 1%/1.5% and ASD
 * auto-computed), Name of the Agency, Tender %, Contract Amount and the
 * Reservation flag. Scoped to the one Name-of-Work row. Unmatched sheets
 * are reported. Flashes the affected row(s) via onUpdated so the user can
 * verify them.
 */
export default function WorksListL1Update({ table, onChange, onUpdated }: Props) {
  const [busy, setBusy] = useState<null | 'l1'>(null)
  // Only errors / no-match feedback shows beside the button; a successful update
  // is surfaced under the flashed row(s) in the table instead (see onUpdated).
  const [result, setResult] = useState<{ message: string; unmatched: string[] } | null>(null)
  // Progress shown in the button while a picked folder is processed: 'scanning'
  // and 'reading' come from the main process's disk walk/read-back
  // (electron/tenderDocumentScan.ts); 'parsing' is set locally, one tick per
  // file, while this component's own loop below text-extracts/OCRs and
  // matches each file — a real office's Tender Evaluations folder can have
  // 500+ relevant PDFs, and that phase used to run with zero feedback at all
  // (the button just said "Updating…" for minutes), which read as the whole
  // app having frozen even though it hadn't.
  const [scanProgress, setScanProgress] = useState<TenderScanProgress | null>(null)
  const busyRef = useRef(busy)
  busyRef.current = busy

  useEffect(() => {
    return api.onTenderScanProgress((p) => {
      // Ignore a straggling event from a scan that's no longer the active one
      // (e.g. the previous pick's last tick arriving after busy was cleared).
      if (busyRef.current) setScanProgress(p)
    })
  }, [])

  async function handleFiles(files: File[]) {
    setBusy('l1')
    setResult(null)
    try {
      // Row name vectors only depend on the Works List's own "Name of the Work"
      // column, which this update never touches — computed once up front so
      // each file below only has to embed its own (single) title, and can be
      // matched and applied to the table the moment it's parsed instead of
      // waiting for the whole folder to finish.
      const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
      let rowNameVectors: number[][] | undefined
      if (nameHeader) {
        try {
          rowNameVectors = await api.embedTexts(table.rows.map((r) => r[nameHeader] ?? ''))
        } catch {
          rowNameVectors = undefined
        }
      }

      let current = table
      let matchedCount = 0
      let evaluationCount = 0
      const unmatched: string[] = []
      // Row(s) actually changed, in the order they were matched — flashed/
      // scrolled to ONCE after the whole batch, not per file: remounting the
      // sheet (and re-running scrollIntoView) on every single match froze the
      // app partway through a large folder scan, and flashed rows the user
      // had no time to read before the next match replaced them.
      const allMatchedIndices: number[] = []

      for (const file of files) {
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
        let ev: TenderEvaluation | null = null
        if (!isPdf) {
          continue // HTML Online Intimation alone carries no L-1 fields to fold in here.
        }
        const lines = await pdfToTextLinesOrOcr(file)
        const parsed = parseTenderEvaluation(lines)
        if (parsed.nameOfWork || parsed.tenderId) ev = parsed
        if (!ev) continue // a printed Intimation / LOA PDF, not an L-1 evaluation.

        evaluationCount++
        let evalNameVectors: number[][] | undefined
        if (rowNameVectors) {
          try {
            evalNameVectors = await api.embedTexts([ev.nameOfWork ?? ''])
          } catch {
            evalNameVectors = undefined
          }
        }
        const embeddings = rowNameVectors && evalNameVectors ? { rowNameVectors, evalNameVectors } : undefined
        const result = updateWorksListFromEvaluations(current, [ev], embeddings)
        if (result.matchedCount > 0) {
          current = result.table
          matchedCount++
          allMatchedIndices.push(...result.matchedRowIndices)
          // Progress only advances on an actual Works List row update, not on
          // every PDF read — a big unmatched batch would otherwise look like
          // it's racing through updates that never happen.
          setScanProgress({ phase: 'parsing', done: matchedCount, total: files.length })
          onChange(current) // live table data — applied immediately; the flash/scroll below waits for the batch to finish.
        } else {
          unmatched.push(...result.unmatched)
        }
      }

      if (allMatchedIndices.length > 0) {
        onUpdated(
          allMatchedIndices,
          `Updated from L1 — Tender ID, Notice No/Date, ECV, EMD 1%/1.5%, ASD, Agency, Tender %, Contract Amount, Reservation.`
        )
      }

      if (evaluationCount === 0) {
        setResult({
          message: "Couldn't read any L-1 selection sheet from the files. Add the Commercial Evaluation / Stage Selected PDF(s).",
          unmatched: []
        })
        return
      }
      // Only report problems here; each success was already flashed live above.
      if (matchedCount === 0 || unmatched.length > 0) {
        setResult({
          message:
            matchedCount === 0
              ? 'Read the L-1 sheet(s), but no Works List row matched by Name of Work — nothing was changed.'
              : `Updated ${matchedCount} work${matchedCount === 1 ? '' : 's'}; some sheets matched no row.`,
          unmatched
        })
      } else {
        setResult(null)
      }
    } catch (e) {
      setResult({ message: e instanceof Error ? e.message : String(e), unmatched: [] })
    } finally {
      setBusy(null)
    }
  }

  // The native dialog lets the user pick loose files, whole folders, or both
  // together (see electron/tenderDocumentScan.ts) — a folder is scanned
  // recursively for anything named like an L1 sheet or Online Intimation, so
  // an entire circle's Tender Evaluations tree can be picked once instead of
  // hunting down each work's own file individually. Files picked directly by
  // name are never filtered out, so a single sheet still works exactly as
  // before — this replaces the old file-only <input>, it doesn't add a
  // second, separate "folder mode". busy is set BEFORE the dialog/scan even
  // starts (not just once handleFiles begins) — that scan used to be the
  // part with zero visual feedback while it silently froze the whole app;
  // now it's both non-blocking and visible.
  async function pickAndRun(kind: 'l1', run: (files: File[]) => Promise<void>) {
    setBusy(kind)
    setScanProgress(null)
    setResult(null)
    try {
      const picked = await api.pickTenderDocuments()
      if (picked.length === 0) return
      // Scan/read is done — handleFiles' own PDF parsing and row-matching
      // has no progress signal of its own, fall back to a plain busy label
      // for that part (same as before this change).
      setScanProgress(null)
      await run(picked.map((f) => new File([new Uint8Array(f.bytes)], f.name)))
    } finally {
      setScanProgress(null)
      setBusy(null)
    }
  }

  // Percentage fill shown inside the button while reading files back (the
  // phase with a real, known total) — 'scanning' has no fixed total yet, so
  // it gets a plain running count instead of a fill.
  const readPct =
    (scanProgress?.phase === 'reading' || scanProgress?.phase === 'parsing') && scanProgress.total > 0
      ? Math.round((scanProgress.done / scanProgress.total) * 100)
      : null

  function buttonLabel(kind: 'l1', idleLabel: string): React.ReactNode {
    if (busy !== kind) return idleLabel
    if (!scanProgress) return 'Opening…'
    if (scanProgress.phase === 'scanning') return `Scanning… (${scanProgress.done} found)`
    if (scanProgress.phase === 'parsing') return `Updated ${scanProgress.done}/${scanProgress.total} (${readPct}%)`
    if (readPct !== null) return `Reading ${scanProgress.done}/${scanProgress.total} (${readPct}%)`
    return 'Updating…'
  }

  return (
    <>
      <button
        className="ghost scan-progress-btn"
        onClick={() => void pickAndRun('l1', handleFiles)}
        disabled={busy !== null}
        title="Pick the L-1 selection sheet(s) and/or the LOA / Online Intimation — files, a whole folder, or both — to fold tender details into the matching Works List row(s) by Name of Work"
      >
        {busy === 'l1' && readPct !== null && (
          <span className="scan-progress-fill" style={{ width: `${readPct}%` }} />
        )}
        <span className="scan-progress-label">
          <IconFolder /> {buttonLabel('l1', 'Update from L1/LOA')}
        </span>
      </button>
      {result && (
        <div className="notice warn" style={{ flexBasis: '100%' }}>
          <IconWarn /> {result.message}
          {result.unmatched.length > 0 && (
            <div className="estimate-hint">
              No match for: {result.unmatched.map((n) => `“${n}”`).join(', ')} — add these to the Works List first.
            </div>
          )}
        </div>
      )}
    </>
  )
}
