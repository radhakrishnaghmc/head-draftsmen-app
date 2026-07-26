import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines, pdfToTextLinesFromData } from '../pdfToText'
import { parseTenderEvaluation, mergeEvaluationsByWork, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { parseIntimationNotice, parseIntimationNoticeLines } from '@core/intimationNotice'
import { updateWorksListFromEvaluations, applyAgencyAddresses } from '@core/worksTenderUpdate'
import { base64ToUint8 } from './docPage'
import { IconFolder, IconWarn, IconCheck } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  /** The Works List database — tables[0], by the app's own convention. */
  table: ExcelTable | null
  onChange: (table: ExcelTable) => void
}

/** A tender-evaluation PDF to parse (a picked File, a folder path, or a Drive id — all resolve to text lines). */
interface EvalSource {
  name: string
  getLines: () => Promise<string[]>
}
/** An intimation notice (.html or .html.pdf) to parse for the agency's name + address. */
interface IntimationSource {
  name: string
  parse: () => Promise<{ agencyName?: string; address?: string }>
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')
const isIntimationName = (n: string) => /viewintimationnotice/i.test(n)
const isHtmlName = (n: string) => /\.html?$/i.test(n)
const htmlFromB64 = (b64: string) => new TextDecoder().decode(base64ToUint8(b64))

/** Run `fn` over `items` with bounded concurrency, calling `onEach` after each completes (for the progress bar). */
async function pool<T>(items: T[], concurrency: number, fn: (item: T) => Promise<void>): Promise<void> {
  let next = 0
  const worker = async () => {
    for (;;) {
      const i = next++
      if (i >= items.length) return
      await fn(items[i])
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, worker))
}

/**
 * Updates the Works List from a batch of tender-evaluation documents — from a
 * few picked files, a local folder, or a Google Drive folder link (with a
 * progress bar). Two kinds of document are read:
 *  - the "Commercial Evaluation" / "Stage Selected" PDFs → each work's row is
 *    filled (matched by Name of Work) with Tender ID, Tender Notice No/Date,
 *    ECV, the L-1 agency, Tender %, and Contract Amount.
 *  - the intimation notices (.html / .html.pdf) → the agency's postal address,
 *    written to every row of that agency (matched by agency name, since one
 *    agency has a single address).
 * See core/tenderEvaluationPdf.ts, core/intimationNotice.ts, core/worksTenderUpdate.ts.
 */
export default function TenderPdfImport({ table, onChange }: Props) {
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState<string | null>(null)
  const [driveLink, setDriveLink] = useState('')
  const inputRef = useRef<HTMLInputElement>(null)

  async function process(evalSources: EvalSource[], intimationSources: IntimationSource[]) {
    if (evalSources.length + intimationSources.length === 0) return
    if (!table) {
      setError('Add works to the Works List first — these documents update existing work rows.')
      return
    }
    setLoading(true)
    setError(null)
    setDone(null)
    const total = evalSources.length + intimationSources.length
    setProgress({ done: 0, total })
    try {
      let completed = 0
      const tick = () => setProgress({ done: ++completed, total })

      // Phase 1 — read the evaluation PDFs.
      const parsed: TenderEvaluation[] = []
      const skipped: string[] = []
      await pool(evalSources, 5, async (s) => {
        try {
          const ev = parseTenderEvaluation(await s.getLines())
          if (ev.nameOfWork) parsed.push(ev)
          else skipped.push(s.name)
        } catch {
          skipped.push(s.name)
        }
        tick()
      })

      // Phase 2 — read the intimation notices for agency addresses.
      const addressByAgency = new Map<string, string>()
      await pool(intimationSources, 5, async (s) => {
        try {
          const { agencyName, address } = await s.parse()
          if (agencyName && address) addressByAgency.set(norm(agencyName), address)
        } catch {
          /* skip an unreadable notice */
        }
        tick()
      })

      // Several pages per work (responsiveness + commercial "L1") collapse
      // into one complete record per work.
      const evaluations = mergeEvaluationsByWork(parsed)
      if (evaluations.length === 0 && addressByAgency.size === 0) {
        setError(
          "Couldn't read tender or intimation details from those files — make sure they're the Commercial Evaluation / Stage Selected pages and the intimation notices."
        )
        return
      }

      let working = table
      let matchedCount = 0
      let unmatched: string[] = []

      if (evaluations.length > 0) {
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
        const res = updateWorksListFromEvaluations(working, evaluations, embeddings)
        working = res.table
        matchedCount = res.matchedCount
        unmatched = res.unmatched
      }

      // Addresses apply after the evaluations, so a work whose agency was just
      // set in this same batch also gets its address filled.
      let addressFilled = 0
      if (addressByAgency.size > 0) {
        const res = applyAgencyAddresses(working, addressByAgency)
        working = res.table
        addressFilled = res.filledCount
      }

      if (working !== table) onChange(working)

      const parts: string[] = []
      if (evaluations.length > 0) {
        parts.push(`Updated ${matchedCount} work${matchedCount === 1 ? '' : 's'} from ${evaluations.length} tender${evaluations.length === 1 ? '' : 's'}.`)
      }
      if (addressByAgency.size > 0) {
        parts.push(`Filled ${addressFilled} agency address${addressFilled === 1 ? '' : 'es'} from ${addressByAgency.size} intimation notice${addressByAgency.size === 1 ? '' : 's'}.`)
      }
      if (unmatched.length > 0) parts.push(`No matching work for: ${unmatched.join('; ')}.`)
      if (skipped.length > 0) parts.push(`Ignored ${skipped.length} unreadable file${skipped.length === 1 ? '' : 's'}.`)
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
    const evalSources: EvalSource[] = []
    const intimationSources: IntimationSource[] = []
    for (const file of Array.from(fileList)) {
      if (isIntimationName(file.name)) {
        intimationSources.push({
          name: file.name,
          parse: isHtmlName(file.name)
            ? async () => parseIntimationNotice(await file.text())
            : async () => parseIntimationNoticeLines(await pdfToTextLines(file))
        })
      } else if (isHtmlName(file.name)) {
        intimationSources.push({ name: file.name, parse: async () => parseIntimationNotice(await file.text()) })
      } else {
        evalSources.push({ name: file.name, getLines: () => pdfToTextLines(file) })
      }
    }
    void process(evalSources, intimationSources)
  }

  async function handleFolder() {
    const picked = await api.pickTenderEvalFolder()
    if (!picked) return
    if (picked.tenderPdfs.length === 0 && picked.intimationFiles.length === 0) {
      setError('No tender-evaluation PDFs or intimation notices found in that folder.')
      return
    }
    const evalSources: EvalSource[] = picked.tenderPdfs.map((p) => ({
      name: p.split(/[/\\]/).pop() ?? p,
      getLines: async () => pdfToTextLinesFromData(base64ToUint8(await api.readFileBase64(p)))
    }))
    const intimationSources: IntimationSource[] = picked.intimationFiles.map((p) => {
      const name = p.split(/[/\\]/).pop() ?? p
      return {
        name,
        parse: isHtmlName(p)
          ? async () => parseIntimationNotice(htmlFromB64(await api.readFileBase64(p)))
          : async () => parseIntimationNoticeLines(await pdfToTextLinesFromData(base64ToUint8(await api.readFileBase64(p))))
      }
    })
    void process(evalSources, intimationSources)
  }

  async function handleDriveLink() {
    const link = driveLink.trim()
    if (!link || loading) return
    setLoading(true)
    setError(null)
    setDone(null)
    let files: { tenderPdfs: { id: string; name: string }[]; intimationFiles: { id: string; name: string }[] }
    try {
      files = await api.listDriveFolderTenderFiles(link)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
      setLoading(false)
      return
    }
    setLoading(false) // process() manages loading/progress from here
    const evalSources: EvalSource[] = files.tenderPdfs.map((f) => ({
      name: f.name,
      getLines: async () => pdfToTextLinesFromData(base64ToUint8(await api.downloadDriveFile(f.id)))
    }))
    const intimationSources: IntimationSource[] = files.intimationFiles.map((f) => ({
      name: f.name,
      parse: isHtmlName(f.name)
        ? async () => parseIntimationNotice(htmlFromB64(await api.downloadDriveFile(f.id)))
        : async () => parseIntimationNoticeLines(await pdfToTextLinesFromData(base64ToUint8(await api.downloadDriveFile(f.id))))
    }))
    void process(evalSources, intimationSources)
  }

  const pct = progress ? Math.round((progress.done / progress.total) * 100) : 0

  return (
    <div className="card link-import">
      <div className="link-import-row">
        <IconFolder />
        <span className="tender-pdf-label">
          Update from tender evaluation documents — fills Tender ID, Tender Notice No/Date, ECV, Agency (L-1), Tender %,
          and Contract Amount per work, and each agency's address from its intimation notice.
        </span>
        <button className="primary" onClick={() => inputRef.current?.click()} disabled={loading}>
          {loading ? 'Reading…' : 'Upload File(s)'}
        </button>
        <button className="primary" onClick={handleFolder} disabled={loading}>
          {loading ? 'Reading…' : 'Select Folder'}
        </button>
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf,.html,.htm"
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
          placeholder="…or paste a Google Drive folder link of tender documents"
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
            Reading files… {progress.done} / {progress.total}
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
