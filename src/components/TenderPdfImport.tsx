import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines, pdfToTextLinesFromData } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { base64ToUint8 } from './docPage'
import { IconFolder, IconWarn, IconCheck } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  /** The Works List database — tables[0], by the app's own convention. */
  table: ExcelTable | null
  onChange: (table: ExcelTable) => void
}

/** One PDF to process, named for reporting, with a lazy text extractor (a picked File, or a folder path read via the main process). */
interface PdfSource {
  name: string
  getLines: () => Promise<string[]>
}

/**
 * Updates the Works List from tender-evaluation PDFs (the portal's
 * "Commercial Evaluation" / "Stage Selected" pages) — either a few picked
 * files or a whole folder from the e-procurement platform, with a progress
 * bar for the folder case. Each PDF's Name of Work is matched to a Works
 * List row, filling Tender ID, Tender Notice No/Date, ECV, the L-1 agency,
 * Tender Percentage, and Contract Amount. See core/tenderEvaluationPdf.ts
 * (parse) and core/worksTenderUpdate.ts (match + write).
 */
export default function TenderPdfImport({ table, onChange }: Props) {
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState<string | null>(null)
  const [driveLink, setDriveLink] = useState('')
  const inputRef = useRef<HTMLInputElement>(null)

  async function process(sources: PdfSource[]) {
    if (sources.length === 0) return
    if (!table) {
      setError('Add works to the Works List first — a PDF updates an existing work row by its name.')
      return
    }
    setLoading(true)
    setError(null)
    setDone(null)
    setProgress({ done: 0, total: sources.length })
    try {
      const evaluations: TenderEvaluation[] = []
      const skipped: string[] = []
      for (let i = 0; i < sources.length; i++) {
        try {
          const ev = parseTenderEvaluation(await sources[i].getLines())
          if (ev.nameOfWork) evaluations.push(ev)
          else skipped.push(sources[i].name)
        } catch {
          skipped.push(sources[i].name)
        }
        setProgress({ done: i + 1, total: sources.length })
      }

      if (evaluations.length === 0) {
        setError(
          `Couldn't read a Name of Work from ${sources.length === 1 ? 'that PDF' : 'any of those PDFs'} — make sure they're the tender's Commercial Evaluation / Stage Selected pages.`
        )
        return
      }

      // Embeddings let a PDF's work title match a Works List row worded a
      // little differently; fall back to exact-name matching if unavailable.
      const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
      let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
      if (nameHeader) {
        try {
          const [rowNameVectors, evalNameVectors] = await Promise.all([
            api.embedTexts(table.rows.map((r) => r[nameHeader] ?? '')),
            api.embedTexts(evaluations.map((e) => e.nameOfWork!))
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
      if (unmatched.length > 0) parts.push(`No matching work for: ${unmatched.join('; ')}.`)
      if (skipped.length > 0) parts.push(`Skipped (no tender details): ${skipped.join(', ')}.`)
      setDone(parts.join(' '))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setLoading(false)
      setProgress(null)
    }
  }

  function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    void process(Array.from(fileList).map((file) => ({ name: file.name, getLines: () => pdfToTextLines(file) })))
  }

  async function handleFolder() {
    const picked = await api.pickTenderEvalFolder()
    if (!picked) return
    if (picked.files.length === 0) {
      setError('No PDF files found in that folder.')
      return
    }
    void process(
      picked.files.map((filePath) => ({
        name: filePath.split(/[/\\]/).pop() ?? filePath,
        getLines: async () => pdfToTextLinesFromData(base64ToUint8(await api.readFileBase64(filePath)))
      }))
    )
  }

  async function handleDriveLink() {
    const link = driveLink.trim()
    if (!link || loading) return
    setLoading(true)
    setError(null)
    setDone(null)
    let files: { id: string; name: string }[]
    try {
      files = await api.listDriveFolderPdfs(link)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
      setLoading(false)
      return
    }
    // process() manages its own loading/progress from here.
    setLoading(false)
    void process(
      files.map((f) => ({
        name: f.name,
        getLines: async () => pdfToTextLinesFromData(base64ToUint8(await api.downloadDriveFile(f.id)))
      }))
    )
  }

  const pct = progress ? Math.round((progress.done / progress.total) * 100) : 0

  return (
    <div className="card link-import">
      <div className="link-import-row">
        <IconFolder />
        <span className="tender-pdf-label">
          Update from tender evaluation PDFs — fills Tender ID, Tender Notice No/Date, ECV, Agency (L-1), Tender %, and
          Contract Amount against each work.
        </span>
        <button className="primary" onClick={() => inputRef.current?.click()} disabled={loading}>
          {loading ? 'Reading…' : 'Upload PDF(s)'}
        </button>
        <button className="primary" onClick={handleFolder} disabled={loading}>
          {loading ? 'Reading…' : 'Select Folder'}
        </button>
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          multiple
          style={{ display: 'none' }}
          onChange={(e) => {
            handleFiles(e.target.files)
            e.target.value = ''
          }}
        />
      </div>
      <div className="link-import-row" style={{ marginTop: 8 }}>
        <input
          value={driveLink}
          placeholder="…or paste a Google Drive folder link of tender PDFs"
          onChange={(e) => {
            setDriveLink(e.target.value)
            setError(null)
            setDone(null)
          }}
          onKeyDown={(e) => e.key === 'Enter' && handleDriveLink()}
          disabled={loading}
        />
        <button className="primary" onClick={handleDriveLink} disabled={loading || !driveLink.trim()}>
          {loading ? 'Reading…' : 'From Drive Link'}
        </button>
      </div>
      {progress && (
        <div className="tender-progress">
          <div className="tender-progress-bar" style={{ width: `${pct}%` }} />
          <span className="tender-progress-text">
            Reading PDFs… {progress.done} / {progress.total}
          </span>
        </div>
      )}
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
