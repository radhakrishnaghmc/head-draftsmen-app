import { useEffect, useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { findHtmlPlaceholders, fillHtmlPlaceholders } from '@core/htmlPlaceholders'
import { extractLabeledLine } from '@core/ocrLabels'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import { pdfPagesToDataUrls } from '../pdfToImages'
import { IconFolder, IconImage, IconTrash, IconPrint, IconDownload, IconWarn, IconBell } from './Icons'
import type { ExcelTable } from '@core/types'

interface Photo {
  id: string
  name: string
  dataUrl: string
}

interface Props {
  tables: ExcelTable[]
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

function readAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsText(file)
  })
}

function nextId(): string {
  return `intim-photo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function rowLabel(row: Record<string, string>, headers: string[], index: number): string {
  const nameCol = headers.find((h) => /name of (the )?work/i.test(h)) ?? headers[0]
  const v = nameCol ? (row[nameCol] ?? '').trim() : ''
  return v ? `${index + 1}. ${v}` : `Row ${index + 1}`
}

/**
 * Reads a PDF (or a series of photos) of a site document plus an existing
 * Intimation format (an .html file the user already has — this app's own
 * document templates moved to .docx, but this stays HTML, e.g. an older
 * format kept as-is) and fills the format's {{placeholders}} straight from
 * whatever's printed/handwritten on it (OCR, see core/ocrLabels.ts) —
 * falling back to a picked Works List row for standard fields the source
 * document doesn't carry (Name of the work, Circle, Agency, ...), the same
 * semantic placeholder-to-column matching Create/Issue Document already
 * uses. A PDF is split into one page-image per photo first (src/pdfToImages.ts),
 * the same conversion Upload Photos to Get Estimate already uses.
 */
export default function GiveIntimationTab({ tables }: Props) {
  const table = tables[0] ?? null

  const [templateHtml, setTemplateHtml] = useState<string | null>(null)
  const [templateName, setTemplateName] = useState('')
  const [templateError, setTemplateError] = useState<string | null>(null)

  const [photos, setPhotos] = useState<Photo[]>([])
  const [ocrLines, setOcrLines] = useState<string[] | null>(null)
  const [ocrBusy, setOcrBusy] = useState(false)
  const [ocrError, setOcrError] = useState<string | null>(null)

  const [rowIndex, setRowIndex] = useState(0)
  const [values, setValues] = useState<Record<string, string>>({})

  const [printBusy, setPrintBusy] = useState(false)
  const [saveBusy, setSaveBusy] = useState(false)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)

  const templateInputRef = useRef<HTMLInputElement>(null)
  const photoInputRef = useRef<HTMLInputElement>(null)

  const placeholders = useMemo(() => (templateHtml ? findHtmlPlaceholders(templateHtml) : []), [templateHtml])
  const filledHtml = useMemo(() => (templateHtml ? fillHtmlPlaceholders(templateHtml, values) : ''), [
    templateHtml,
    values
  ])
  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null

  // Auto-fill only the placeholders still blank — from OCR'd photo text
  // first (the more specific, just-read source), then from the selected
  // Works List row — never overwriting a value the user already typed or
  // corrected by hand.
  function autoFill(source: 'ocr' | 'row') {
    setValues((prev) => {
      const next = { ...prev }
      const missing = placeholders.filter((label) => !(next[label] ?? '').trim())
      if (missing.length === 0) return prev

      if (source === 'ocr' && ocrLines) {
        for (const label of missing) {
          const found = extractLabeledLine(ocrLines, label)
          if (found) next[label] = found
        }
      } else if (source === 'row' && selectedRow && table) {
        const stillMissing = missing.filter((label) => !(next[label] ?? '').trim())
        const mapping = matchPlaceholdersToColumns(stillMissing, table.headers)
        for (const m of mapping) {
          if (m.column) {
            const v = (selectedRow[m.column] ?? '').trim()
            if (v) next[m.label] = v
          }
        }
      }
      return next
    })
  }

  useEffect(() => {
    if (placeholders.length === 0) return
    if (ocrLines) autoFill('ocr')
    // autoFill intentionally omitted: it closes over `values` on every
    // render, and including it here would re-run this effect (and re-fill
    // already-blank placeholders redundantly) on every keystroke.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [placeholders, ocrLines])

  useEffect(() => {
    if (placeholders.length === 0) return
    if (selectedRow) autoFill('row')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [placeholders, selectedRow])

  async function pickTemplate(file: File) {
    setTemplateError(null)
    try {
      const text = await readAsText(file)
      setTemplateHtml(text)
      setTemplateName(stripExt(file.name))
      setValues({})
      setActionSaved(null)
    } catch (e) {
      setTemplateError(e instanceof Error ? e.message : String(e))
    }
  }

  async function handlePhotoFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    setOcrError(null)
    try {
      const added: Photo[] = []
      for (const file of Array.from(fileList)) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // A PDF here is a scanned/photographed multi-page document saved
          // as one file — split it into one page-image per photo so it
          // feeds the exact same per-photo OCR pipeline as a direct photo
          // upload.
          const pages = await pdfPagesToDataUrls(file)
          const base = stripExt(file.name)
          pages.forEach((dataUrl, i) => added.push({ id: nextId(), name: `${base} (page ${i + 1})`, dataUrl }))
        } else if (file.type.startsWith('image/')) {
          added.push({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) })
        }
      }
      setPhotos((prev) => [...prev, ...added])
      setOcrLines(null)
    } catch (e) {
      setOcrError(e instanceof Error ? e.message : String(e))
    }
  }

  function removePhoto(id: string) {
    setPhotos((prev) => prev.filter((p) => p.id !== id))
    setOcrLines(null)
  }

  async function readPhotos() {
    if (photos.length === 0) return
    setOcrBusy(true)
    setOcrError(null)
    try {
      const grid = await api.ocrEstimatePhotos(photos.map((p) => p.dataUrl))
      setOcrLines(grid.grid.map((row) => row[0] ?? '').filter((l) => l.trim()))
    } catch (e) {
      setOcrError(e instanceof Error ? e.message : String(e))
    } finally {
      setOcrBusy(false)
    }
  }

  function updateValue(label: string, value: string) {
    setValues((prev) => ({ ...prev, [label]: value }))
  }

  async function printIntimation() {
    setPrintBusy(true)
    setActionError(null)
    try {
      await api.printCreatedDocument(filledHtml)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setPrintBusy(false)
    }
  }

  async function saveIntimation() {
    setSaveBusy(true)
    setActionError(null)
    try {
      const path = await api.exportIntimationHtml(filledHtml, templateName || 'Intimation')
      if (path) setActionSaved(path)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setSaveBusy(false)
    }
  }

  return (
    <div className="card">
      <div className="empty">
        <IconBell />
        <p>
          Upload the Intimation format (.html) and a PDF (or photos) of the site document — the app reads
          whatever's printed/written on it and fills the format's placeholders automatically.
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={() => templateInputRef.current?.click()}>
            <IconFolder /> {templateHtml ? 'Change Intimation Format' : 'Upload Intimation Format (.html)'}
          </button>
          <button className="primary" onClick={() => photoInputRef.current?.click()}>
            <IconImage /> Upload PDF or Photos
          </button>
          <input
            ref={templateInputRef}
            type="file"
            accept=".html,.htm,text/html"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void pickTemplate(file)
              e.target.value = ''
            }}
          />
          <input
            ref={photoInputRef}
            type="file"
            accept="image/*,application/pdf"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              void handlePhotoFiles(e.target.files)
              e.target.value = ''
            }}
          />
        </div>
      </div>

      {templateError && (
        <div className="notice error">
          <IconWarn /> {templateError}
        </div>
      )}

      {photos.length > 0 && (
        <>
          <div className="photo-tile-grid">
            {photos.map((photo, i) => (
              <div key={photo.id} className="photo-tile-card">
                <span className="photo-tile-number">{i + 1}</span>
                <button className="photo-tile-remove" title="Remove" onClick={() => removePhoto(photo.id)}>
                  <IconTrash />
                </button>
                <img className="photo-tile-img" src={photo.dataUrl} alt={photo.name} />
                <div className="photo-tile-name" title={photo.name}>
                  {photo.name}
                </div>
              </div>
            ))}
          </div>
          <div className="boq-actions" style={{ marginTop: 10 }}>
            <button className="primary" onClick={readPhotos} disabled={ocrBusy}>
              <IconImage /> {ocrBusy ? 'Reading photos…' : 'Read Photos'}
            </button>
            {ocrLines && <span className="estimate-hint">{ocrLines.length} line(s) read from {photos.length} photo(s)</span>}
          </div>
        </>
      )}

      {ocrError && (
        <div className="notice error">
          <IconWarn /> {ocrError}
        </div>
      )}

      {templateHtml && (
        <div className="estimate-body">
          <div className="estimate-preview">
            <span className="estimate-preview-title">Live Preview — {templateName}</span>
            <div className="estimate-preview-scroll intimation-frame-wrap">
              <iframe className="intimation-frame" srcDoc={filledHtml} title="Intimation preview" />
            </div>
            <div className="doc-sheet-footer">
              <span className="estimate-hint">Updates live as details are filled in.</span>
              <button className="primary" onClick={printIntimation} disabled={printBusy}>
                <IconPrint /> {printBusy ? 'Opening…' : 'Print / Save as PDF'}
              </button>
              <button className="primary" onClick={saveIntimation} disabled={saveBusy}>
                <IconDownload /> {saveBusy ? 'Saving…' : 'Save as HTML'}
              </button>
            </div>
            {actionError && (
              <div className="notice error">
                <IconWarn /> {actionError}
              </div>
            )}
            {actionSaved && <p className="estimate-hint">Saved to {actionSaved}</p>}
          </div>

          <div className="estimate-details">
            <span className="estimate-preview-title">Details needed</span>
            {table && table.rows.length > 0 && (
              <label className="estimate-details-field">
                Work (fills standard fields from the Works List)
                <select
                  className="editor-name"
                  value={rowIndex}
                  onChange={(e) => setRowIndex(Number(e.target.value))}
                >
                  {table.rows.map((row, i) => (
                    <option value={i} key={i}>
                      {rowLabel(row, table.headers, i)}
                    </option>
                  ))}
                </select>
              </label>
            )}
            {placeholders.length === 0 ? (
              <p className="estimate-hint">No {'{{placeholders}}'} found in this format.</p>
            ) : (
              placeholders.map((label) => (
                <label className="estimate-details-field" key={label}>
                  {label}
                  <input
                    className="editor-name"
                    placeholder={label}
                    value={values[label] ?? ''}
                    onChange={(e) => updateValue(label, e.target.value)}
                  />
                </label>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  )
}
