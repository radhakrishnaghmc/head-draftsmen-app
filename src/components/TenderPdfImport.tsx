import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { IconFolder, IconWarn, IconCheck } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  /** The Works List database — tables[0], by the app's own convention. */
  table: ExcelTable | null
  onChange: (table: ExcelTable) => void
}

/**
 * Reads one or more tender-evaluation PDFs (the portal's "Commercial
 * Evaluation" / "Stage Selected" pages) and updates the Works List against
 * each PDF's Name of Work — filling Tender ID, Tender Notice No, ECV, the
 * L-1 agency's name, Tender Percentage, and Contract Amount. See
 * core/tenderEvaluationPdf.ts (parse) and core/worksTenderUpdate.ts (match
 * + write).
 */
export default function TenderPdfImport({ table, onChange }: Props) {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  async function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    if (!table) {
      setError('Add works to the Works List first — a PDF updates an existing work row by its name.')
      return
    }
    setLoading(true)
    setError(null)
    setDone(null)
    try {
      const evaluations: TenderEvaluation[] = []
      const skipped: string[] = []
      for (const file of Array.from(fileList)) {
        const lines = await pdfToTextLines(file)
        const ev = parseTenderEvaluation(lines)
        if (ev.nameOfWork) evaluations.push(ev)
        else skipped.push(file.name)
      }

      if (evaluations.length === 0) {
        setError(
          `Couldn't read a Name of Work from ${skipped.length === 1 ? 'that PDF' : 'those PDFs'} — make sure it's the tender's Commercial Evaluation / Stage Selected page.`
        )
        return
      }

      // Embeddings let a PDF's work title match a Works List row worded a
      // little differently; fall back to exact-name matching if unavailable.
      const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
      let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
      if (nameHeader) {
        try {
          const rowNames = table.rows.map((r) => r[nameHeader] ?? '')
          const evalNames = evaluations.map((e) => e.nameOfWork!)
          const [rowNameVectors, evalNameVectors] = await Promise.all([
            api.embedTexts(rowNames),
            api.embedTexts(evalNames)
          ])
          embeddings = { rowNameVectors, evalNameVectors }
        } catch {
          embeddings = undefined
        }
      }

      const { table: updated, matchedCount, unmatched } = updateWorksListFromEvaluations(
        table,
        evaluations,
        embeddings
      )
      if (matchedCount > 0) onChange(updated)

      const parts = [
        `Updated ${matchedCount} work${matchedCount === 1 ? '' : 's'} from ${evaluations.length} PDF${evaluations.length === 1 ? '' : 's'}.`
      ]
      if (unmatched.length > 0) parts.push(`No matching work found for: ${unmatched.join('; ')}.`)
      if (skipped.length > 0) parts.push(`Skipped (no Name of Work): ${skipped.join(', ')}.`)
      setDone(parts.join(' '))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="card link-import">
      <div className="link-import-row">
        <IconFolder />
        <span className="tender-pdf-label">
          Update from tender evaluation PDF — fills Tender ID, Tender Notice No, ECV, Agency (L-1), Tender %, and
          Contract Amount against each work.
        </span>
        <button className="primary" onClick={() => inputRef.current?.click()} disabled={loading}>
          {loading ? 'Reading…' : 'Upload Tender PDF(s)'}
        </button>
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          multiple
          style={{ display: 'none' }}
          onChange={(e) => {
            void handleFiles(e.target.files)
            e.target.value = ''
          }}
        />
      </div>
      {error && (
        <div className="notice error">
          <IconWarn />
          {error}
        </div>
      )}
      {done && (
        <div className="notice ok">
          <IconCheck />
          {done}
        </div>
      )}
    </div>
  )
}
