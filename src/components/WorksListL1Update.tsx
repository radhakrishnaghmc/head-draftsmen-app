import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import {
  parseIntimationNotice,
  parseIntimationNoticeText,
  type IntimationNotice
} from '@core/intimationNotice'
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
 * Works List action: pick one or more L-1 selection sheets (and, optionally, the
 * Online Intimation) and fold their tender details into the Works List by
 * matching each sheet's Name of Work to a row. Reuses the same
 * updateWorksListFromEvaluations the Give Intimation / Work Order flows use, so
 * matched rows get Tender ID, Tender Notice No/Date, ECV (with EMD 1% / 1.5% and
 * ASD auto-computed), Name of the Agency, Tender %, Contract Amount and the
 * Reservation flag. A sheet whose work name matches nothing is reported, not
 * guessed at.
 */
export default function WorksListL1Update({ table, onChange, onUpdated }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  // Only errors / no-match feedback shows beside the button; a successful update
  // is surfaced under the flashed row(s) in the table instead (see onUpdated).
  const [result, setResult] = useState<{ message: string; unmatched: string[] } | null>(null)

  async function handleFiles(files: File[]) {
    setBusy(true)
    setResult(null)
    try {
      const evaluations: TenderEvaluation[] = []
      const notices: IntimationNotice[] = []
      for (const file of files) {
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
        if (!isPdf) {
          // HTML "View Intimation Notice" page → an Online Intimation.
          try {
            notices.push(parseIntimationNotice(await file.text()))
          } catch {
            /* ignore unreadable file */
          }
          continue
        }
        const lines = await pdfToTextLines(file)
        const ev = parseTenderEvaluation(lines)
        if (ev.nameOfWork || ev.tenderId) {
          evaluations.push(ev)
        } else {
          // A printed Intimation / LOA PDF rather than an L-1 evaluation.
          const notice = parseIntimationNoticeText(lines)
          if (notice.agencyName || notice.address || notice.nitNo) notices.push(notice)
        }
      }

      if (evaluations.length === 0) {
        setResult({
          message: "Couldn't read any L-1 selection sheet from the files. Add the Commercial Evaluation / Stage Selected PDF(s).",
          unmatched: []
        })
        return
      }

      // Embedding vectors let the matcher bridge wording drift between a tender's
      // title and the Works List's own entry (abbreviations, punctuation).
      let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
      const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
      if (nameHeader) {
        try {
          const [rowNameVectors, evalNameVectors] = await Promise.all([
            api.embedTexts(table.rows.map((r) => r[nameHeader] ?? '')),
            api.embedTexts(evaluations.map((e) => e.nameOfWork ?? ''))
          ])
          embeddings = { rowNameVectors, evalNameVectors }
        } catch {
          embeddings = undefined
        }
      }

      const { table: updated, matchedCount, unmatched, matchedRowIndices } = updateWorksListFromEvaluations(
        table,
        evaluations,
        embeddings,
        notices.length === 1 ? notices[0] : undefined
      )
      if (matchedCount > 0) {
        onChange(updated)
        onUpdated(
          matchedRowIndices,
          `Updated from L1 — Tender ID, Notice No/Date, ECV, EMD 1%/1.5%, ASD, Agency, Tender %, Contract Amount, Reservation.`
        )
      }
      // Only report problems here; the success is shown under the flashed row(s).
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
      setBusy(false)
    }
  }

  return (
    <>
      <button
        className="ghost"
        onClick={() => inputRef.current?.click()}
        disabled={busy}
        title="Pick the L-1 selection sheet(s) (and, optionally, the Online Intimation) to fold tender details into the Works List by Name of Work"
      >
        <IconFolder /> {busy ? 'Updating…' : 'Update from L1'}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept="application/pdf,.pdf,.html,.htm,text/html"
        multiple
        style={{ display: 'none' }}
        onChange={(e) => {
          const files = Array.from(e.target.files ?? [])
          if (files.length > 0) void handleFiles(files)
          e.target.value = ''
        }}
      />
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
