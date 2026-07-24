import { useMemo, useState } from 'react'
import { api } from '../ipc'
import {
  boqToScheduleA,
  extractWorkNameFromBoq,
  boqColumnsMatchedViaEmbedding,
  BOQ_SCHEDULE_A_COLUMN_SPECS
} from '../boqTransform'
import { mismatchHint } from '../docClassify'
import { IconFolder, IconDownload, IconWarn, IconCheck, IconTable } from './Icons'
import type { ExcelTable } from '@core/types'
import type { ColumnEmbeddings } from '@core/columnMatch'
import { buildScheduleARows, rowsToScheduleAItems, metaFromWorksRow, findWorksRowByName } from '@core/scheduleA'
import type { ScheduleAMeta } from '../../electron/ipc-contract'

/**
 * Only called when the fast, synchronous regex-based column match already
 * failed — computes embeddings for this BOQ's headers plus the standard
 * column labels (Estimate Quantity/Description/Rate/UOM/Amount) so
 * boqToScheduleA gets one more chance via semantic similarity (e.g. a
 * hand-built BOQ with unusual column names). Local model only, no network —
 * see electron/embeddingsWorker.ts. Returns null (rather than throwing) if
 * the model itself is unavailable, so the caller's original error stands.
 */
async function embedBoqColumns(headers: string[]): Promise<ColumnEmbeddings | null> {
  try {
    const labels = BOQ_SCHEDULE_A_COLUMN_SPECS.map((s) => s.label)
    const vectors = await api.embedTexts([...headers, ...labels])
    return { headerVectors: vectors.slice(0, headers.length), labelVectors: vectors.slice(headers.length) }
  } catch {
    return null
  }
}

interface Props {
  tables: ExcelTable[]
}

function findHeader(headers: string[], name: string): string | undefined {
  return headers.find((h) => h.trim().toLowerCase() === name.toLowerCase())
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

// The uploaded file is a BOQ, but the saved document is Schedule A — drop
// any "BOQ" wording from its name so the two don't end up side by side
// (e.g. "7-1 BOQ.xls" → "7-1 Schedule A.xlsx", not "7-1 BOQ Schedule A.xlsx").
function stripBoqWord(name: string): string {
  return name
    .replace(/\bboq\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
}

/**
 * Schedule A download: pick a Win Code from the Works List (its work details
 * feed the name-of-work / estimate / contractor fields), upload the matching
 * BOQ, and save the combined Schedule A document.
 */
export default function ScheduleATab({ tables }: Props) {
  const worksTable = tables[0] ?? null

  const winCodeHeader = worksTable ? findHeader(worksTable.headers, 'Wincode') : undefined

  const winCodes = useMemo(() => {
    if (!worksTable || !winCodeHeader) return []
    const seen = new Set<string>()
    for (const row of worksTable.rows) {
      const v = (row[winCodeHeader] ?? '').trim()
      if (v) seen.add(v)
    }
    return Array.from(seen)
  }, [worksTable, winCodeHeader])

  const [winCode, setWinCode] = useState('')
  const selectedRow = useMemo(() => {
    if (!worksTable || !winCodeHeader || !winCode) return null
    return worksTable.rows.find((r) => (r[winCodeHeader] ?? '').trim() === winCode) ?? null
  }, [worksTable, winCodeHeader, winCode])

  const [source, setSource] = useState<ExcelTable | null>(null)
  const [output, setOutput] = useState<ExcelTable | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)

  // The uploaded BOQ's own embedded "Name of Work" (written there by the
  // Download BOQ tab), and the Works List row it matched — when present,
  // this drives the meta fields instead of the manually picked Win Code, so
  // uploading a BOQ generated in-app auto-fills Schedule A without the user
  // having to separately hunt down and pick its Win Code.
  const [detectedName, setDetectedName] = useState<string | null>(null)
  const [autoRow, setAutoRow] = useState<Record<string, string> | null>(null)
  // True when autoRow only matched via semantic similarity (embeddings),
  // not an exact "Name of the work" match — worth a "please verify" note.
  const [autoRowMatchedViaAi, setAutoRowMatchedViaAi] = useState(false)
  // Column labels only resolved via semantic matching, not a plain header match.
  const [aiAssisted, setAiAssisted] = useState<string[]>([])

  const effectiveRow = autoRow ?? selectedRow

  const meta: ScheduleAMeta | undefined = useMemo(
    () => (effectiveRow ? metaFromWorksRow(effectiveRow) : undefined),
    [effectiveRow]
  )

  const previewRows = useMemo(
    () => (output ? buildScheduleARows(rowsToScheduleAItems(output), meta) : null),
    [output, meta]
  )

  async function upload() {
    setError(null)
    setSaved(null)
    setAiAssisted([])
    const uploaded = await api.pickExcels()
    if (uploaded.length === 0) return
    const t = uploaded[0]
    setSource(t)
    try {
      // Fast path first (plain regex, no model) — only if that fails does
      // this reach for the local embedding model.
      try {
        setOutput(boqToScheduleA(t))
      } catch (syncError) {
        const embeddings = await embedBoqColumns(t.headers)
        if (!embeddings) throw syncError
        setOutput(boqToScheduleA(t, embeddings))
        setAiAssisted(boqColumnsMatchedViaEmbedding(t.headers, embeddings))
      }
      const detected = extractWorkNameFromBoq(t)
      setDetectedName(detected ?? null)
      if (detected && worksTable) {
        let match = findWorksRowByName(worksTable, detected)
        // Fast exact match failed — a BOQ's title-block work name often
        // differs slightly in wording from the Works List's own entry, so
        // give the local embedding model one more chance before giving up.
        if (!match) {
          const nameHeader = worksTable.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
          if (nameHeader) {
            try {
              const rowNames = worksTable.rows.map((r) => r[nameHeader] ?? '')
              const vectors = await api.embedTexts([detected, ...rowNames])
              match = findWorksRowByName(worksTable, detected, {
                workNameVector: vectors[0],
                rowNameVectors: vectors.slice(1)
              })
            } catch {
              // Embeddings unavailable — leave unmatched, same as before.
            }
          }
        }
        setAutoRow(match?.row ?? null)
        setAutoRowMatchedViaAi(match?.matchedViaAi ?? false)
      } else {
        setAutoRow(null)
        setAutoRowMatchedViaAi(false)
      }
    } catch (e) {
      setOutput(null)
      setDetectedName(null)
      setAutoRow(null)
      setAutoRowMatchedViaAi(false)
      const hint = mismatchHint(t.headers, 'boq')
      setError((e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : ''))
    }
  }

  function autoDetectNote(): string | null {
    if (!detectedName) return null
    if (!autoRow) {
      return `Detected "Name of Work" in the BOQ — "${detectedName}" — but no matching row was found in the Works List.`
    }
    const base = `Detected "Name of Work" in the BOQ — "${detectedName}". Estimate/ECV/Contract Amount, Contractor and Tender Percentage auto-filled from the Works List.`
    return autoRowMatchedViaAi ? `${base} Matched via AI, not an exact name — please verify.` : base
  }

  async function download() {
    if (!output) return
    setBusy(true)
    setError(null)
    try {
      const base = source ? stripBoqWord(stripExt(source.name)) : ''
      const suggestedName = base ? `${base} Schedule A` : 'Schedule A'
      const savedPath = await api.exportScheduleA(output, suggestedName, meta)
      if (savedPath) setSaved(savedPath)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconTable />
        {winCodes.length > 0 ? (
          <>
            <p>
              Select the Win Code for this work, then upload its BOQ — or just upload a BOQ generated
              in-app, since its own "Name of Work" auto-fills these details.
            </p>
            <select
              className="win-code-select"
              value={winCode}
              onChange={(e) => {
                setWinCode(e.target.value)
                setSaved(null)
              }}
            >
              <option value="">Select a Win Code…</option>
              {winCodes.map((w) => (
                <option key={w} value={w}>
                  {w}
                </option>
              ))}
            </select>
            {selectedRow && (
              <p className="hint">{selectedRow['Name of the work'] || 'No work name on file for this Win Code.'}</p>
            )}
          </>
        ) : (
          <p>No Win Codes found — add works to the Works List to select one here.</p>
        )}
        <p>
          {source
            ? `Loaded “${source.name}” — ${source.rows.length} row${source.rows.length === 1 ? '' : 's'}.`
            : 'Upload BOQ to generate the Schedule A.'}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={upload}>
            <IconFolder /> Upload BOQ
          </button>
          {output && (
            <button className="primary" onClick={download} disabled={busy}>
              <IconDownload /> {busy ? 'Saving…' : 'Download Schedule A'}
            </button>
          )}
        </div>
      </div>
      {aiAssisted.length > 0 && (
        <div className="notice warn">
          <IconWarn />
          {aiAssisted.join(', ')} column{aiAssisted.length === 1 ? '' : 's'} matched by AI — please double-check
          the preview below.
        </div>
      )}
      {autoDetectNote() && (
        <div className={`notice ${autoRow ? 'ok' : 'warn'}`}>
          {autoRow ? <IconCheck /> : <IconWarn />}
          {autoDetectNote()}
        </div>
      )}
      {error && (
        <div className="notice error">
          <IconWarn />
          {error}
        </div>
      )}
      {saved && (
        <div className="notice ok">
          <IconCheck />
          Saved to {saved}
        </div>
      )}
      {previewRows && (
        <div className="sa-preview">
          <div className="sa-preview-head">Preview — this is exactly what gets saved</div>
          <div className="sheet-wrap">
            <table className="sa-preview-table">
              <tbody>
                {previewRows.map((row, ri) => {
                  const cells = Array.from({ length: 6 }, (_, ci) => row[ci] ?? '')
                  const wide = cells.slice(1).every((c) => c === '')
                  return (
                    <tr key={ri}>
                      {wide ? (
                        <td colSpan={6} className="sa-wide">
                          {String(cells[0])}
                        </td>
                      ) : (
                        cells.map((c, ci) => <td key={ci}>{String(c)}</td>)
                      )}
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
