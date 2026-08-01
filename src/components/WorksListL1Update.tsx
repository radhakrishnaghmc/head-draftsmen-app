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
 * Two Works List update actions:
 *  - "Update from L1": pick one or more L-1 selection sheets (and, optionally, the
 *    Online Intimation) and fold their tender details into the Works List by
 *    matching each sheet's Name of Work to a row — Tender ID, Notice No/Date, ECV
 *    (with EMD 1%/1.5% and ASD auto-computed), Name of the Agency, Address (when
 *    the Intimation is included), Tender %, Contract Amount and the Reservation
 *    flag. Scoped to the one Name-of-Work row. Unmatched sheets are reported.
 *  - "Address from Intimation": upload the Online Intimation (HTML or PDF) alone
 *    to update the agency's Address of the agency on EVERY Works List row that
 *    carries that agency (matched by Name of the Agency).
 * Both flash the affected row(s) via onUpdated so the user can verify them.
 */
/** Agency-name normaliser for matching the Intimation's agency to Works List rows (case / punctuation / spacing insensitive). */
function normAgency(s: string): string {
  return (s ?? '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

export default function WorksListL1Update({ table, onChange, onUpdated }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const intimationRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState<null | 'l1' | 'intimation'>(null)
  // Only errors / no-match feedback shows beside the button; a successful update
  // is surfaced under the flashed row(s) in the table instead (see onUpdated).
  const [result, setResult] = useState<{ message: string; unmatched: string[] } | null>(null)

  // Online Intimation (HTML or PDF) alone → update the agency's ADDRESS on every
  // Works List row carrying that agency (matched by Name of the Agency). This is
  // distinct from "Update from L1", which scopes to the one Name-of-Work row.
  async function handleIntimation(file: File) {
    setBusy('intimation')
    setResult(null)
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const notice: IntimationNotice = isPdf
        ? parseIntimationNoticeText(await pdfToTextLines(file))
        : parseIntimationNotice(await file.text())
      const agency = (notice.agencyName ?? '').trim()
      const address = (notice.address ?? '').trim()
      const agencyHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the agency')
      const addrHeader = table.headers.find((h) => h.trim().toLowerCase() === 'address of the agency')
      if (!agency) {
        setResult({ message: "Couldn't read the agency name from that Online Intimation.", unmatched: [] })
        return
      }
      if (!agencyHeader || !addrHeader) {
        setResult({ message: 'The Works List has no “Name of the Agency” / “Address of the agency” column.', unmatched: [] })
        return
      }
      const target = normAgency(agency)
      // Match every row for this agency: exact normalised name first; if none,
      // fall back to a contains match (handles an "M/s" prefix on one side).
      let matched = table.rows
        .map((r, i) => ({ i, n: normAgency(r[agencyHeader] ?? '') }))
        .filter(({ n }) => n && n === target)
      if (matched.length === 0 && target.length >= 4) {
        matched = table.rows
          .map((r, i) => ({ i, n: normAgency(r[agencyHeader] ?? '') }))
          .filter(({ n }) => n && (n.includes(target) || target.includes(n)))
      }
      if (matched.length === 0) {
        setResult({ message: `No Works List row has the agency “${agency}”.`, unmatched: [] })
        return
      }
      const matchedSet = new Set(matched.map((m) => m.i))
      const rows = table.rows.map((r, i) => (matchedSet.has(i) && address ? { ...r, [addrHeader]: address } : r))
      onChange({ ...table, rows })
      onUpdated(
        matched.map((m) => m.i),
        `Updated the address for “${agency}” from the Online Intimation on ${matched.length} row${matched.length === 1 ? '' : 's'}.`
      )
    } catch (e) {
      setResult({ message: e instanceof Error ? e.message : String(e), unmatched: [] })
    } finally {
      setBusy(null)
    }
  }

  async function handleFiles(files: File[]) {
    setBusy('l1')
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
      setBusy(null)
    }
  }

  return (
    <>
      <button
        className="ghost"
        onClick={() => inputRef.current?.click()}
        disabled={busy !== null}
        title="Pick the L-1 selection sheet(s) together with the LOA / Online Intimation, to fold tender details (and the matched work's address & agency from the LOA) into the Works List by Name of Work"
      >
        <IconFolder /> {busy === 'l1' ? 'Updating…' : 'Update from L1/LOA'}
      </button>
      <button
        className="ghost"
        onClick={() => intimationRef.current?.click()}
        disabled={busy !== null}
        title="Upload the Online Intimation (HTML or PDF) to update this agency's address on every Works List row that carries it"
      >
        <IconFolder /> {busy === 'intimation' ? 'Updating…' : 'Address from Intimation'}
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
      <input
        ref={intimationRef}
        type="file"
        accept="application/pdf,.pdf,.html,.htm,text/html"
        style={{ display: 'none' }}
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) void handleIntimation(file)
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
