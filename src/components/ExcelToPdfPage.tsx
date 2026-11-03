import { useEffect, useRef, useState } from 'react'
import type { ExcelPrintSheetInfo, PrintOrientation, PrintPaper } from '@core/excelPrintHtml'
import { api } from '../ipc'
import { pdfPagesToDataUrlsFromData } from '../pdfToImages'
import { IconChevronLeft, IconTable, IconFolder, IconDownload, IconPrint, IconWarn, IconCheck, IconOpen } from './Icons'

interface Props {
  /** Return to the Tools grid. */
  onBack: () => void
}

/**
 * The Excel → Printable PDF tool's full page: pick a workbook, tick the sheets to print, and the page margins, print
 * area, orientation and scale-to-width are set automatically (paper size and orientation can be overridden). The
 * result shows as live page previews and can be saved as a PDF or sent straight to the printer. Layout runs in the
 * main process (api.excelToPrintPdf / api.excelPrint); the previews are that same PDF rasterized.
 */
export default function ExcelToPdfPage({ onBack }: Props) {
  const [fileName, setFileName] = useState('')
  const [bytes, setBytes] = useState<Uint8Array | null>(null)
  const [sheets, setSheets] = useState<ExcelPrintSheetInfo[]>([])
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [paper, setPaper] = useState<PrintPaper>('a4')
  const [orientation, setOrientation] = useState<PrintOrientation>('auto')
  const [pages, setPages] = useState<string[]>([])
  const [pdf, setPdf] = useState<Uint8Array | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const fileInput = useRef<HTMLInputElement>(null)
  const runId = useRef(0)

  async function chooseWorkbook(file: File | undefined) {
    if (!file) return
    setError(null)
    setSaved(null)
    setPages([])
    setPdf(null)
    setBusy(true)
    try {
      const data = new Uint8Array(await file.arrayBuffer())
      const found = await api.excelPrintInspect(data, paper)
      if (found.length === 0) throw new Error('No sheet in this workbook has anything to print.')
      setFileName(file.name)
      setBytes(data)
      setSheets(found)
      setSelected(new Set(found.map((s) => s.name)))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      if (fileInput.current) fileInput.current.value = ''
    }
  }

  // Re-lay-out the preview whenever the workbook, the ticked sheets or an option changes.
  useEffect(() => {
    if (!bytes || selected.size === 0) {
      setPages([])
      setPdf(null)
      return
    }
    const id = ++runId.current
    setBusy(true)
    setError(null)
    setSaved(null)
    const t = setTimeout(async () => {
      try {
        const out = await api.excelToPrintPdf(bytes, { sheets: [...selected], paper, orientation })
        const imgs = await pdfPagesToDataUrlsFromData(out.slice(), 1.2)
        if (id !== runId.current) return
        setPdf(out)
        setPages(imgs)
      } catch (e) {
        if (id === runId.current) setError(e instanceof Error ? e.message : String(e))
      } finally {
        if (id === runId.current) setBusy(false)
      }
    }, 250)
    return () => clearTimeout(t)
  }, [bytes, selected, paper, orientation])

  function toggle(name: string) {
    setSelected((cur) => {
      const next = new Set(cur)
      if (next.has(name)) next.delete(name)
      else next.add(name)
      return next
    })
  }

  async function savePdf() {
    if (!pdf) return
    try {
      const res = await api.savePdf(pdf, `${fileName.replace(/\.[^.]+$/, '') || 'Workbook'}.pdf`)
      if (res) setSaved(res.file)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  async function printNow() {
    if (!bytes) return
    try {
      await api.excelPrint(bytes, { sheets: [...selected], paper, orientation })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  const allOn = sheets.length > 0 && selected.size === sheets.length

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconTable /> <span className="pdf-ws-title-label">Tool:</span> Excel → Printable PDF
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <input
        ref={fileInput}
        type="file"
        accept=".xlsx,.xlsm"
        hidden
        onChange={(e) => chooseWorkbook(e.target.files?.[0])}
      />

      <div className="card pdf-workspace">
        <div className="pdf-ws-body">
          <div className="pdf-ws-main">
            {error && (
              <div className="notice error tool-outcome">
                <IconWarn /> {error}
              </div>
            )}

            {!bytes ? (
              <button className="pdf-ws-dropzone" onClick={() => fileInput.current?.click()} disabled={busy}>
                <IconTable />
                <span className="pdf-ws-dropzone-title">{busy ? 'Reading workbook…' : 'Choose workbook'}</span>
                <span className="pdf-ws-dropzone-sub">
                  Pick an .xlsx — margins, print area and page fit are set automatically
                </span>
              </button>
            ) : (
              <section className="pdf-ws-doc">
                <header className="pdf-ws-doc-head">
                  <span className="pdf-ws-doc-name" title={fileName}>
                    {fileName}
                  </span>
                  <span className="pdf-ws-doc-meta">
                    {busy ? 'Laying out…' : `${pages.length} page${pages.length === 1 ? '' : 's'}`}
                  </span>
                </header>
                <div className="pdf-ws-pages">
                  {pages.map((src, i) => (
                    <figure key={i} className="pdf-ws-page">
                      <img
                        src={src}
                        alt={`Page ${i + 1}`}
                        style={{ width: '100%', background: '#fff', boxShadow: '0 0 0 1px rgba(16,24,40,.15)' }}
                      />
                      <figcaption className="pdf-ws-page-no">{i + 1}</figcaption>
                    </figure>
                  ))}
                </div>
              </section>
            )}
          </div>

          <aside className="pdf-ws-rail">
            <button className="primary pdf-ws-upload" onClick={() => fileInput.current?.click()} disabled={busy}>
              <IconFolder /> {bytes ? 'Choose another workbook' : 'Choose workbook'}
            </button>

            {bytes && (
              <>
                <div className="pdf-ws-rail-count">
                  <strong>{selected.size}</strong> of {sheets.length} sheet{sheets.length === 1 ? '' : 's'} selected
                </div>
                <div className="pdf-ws-rail-selrow">
                  <button
                    className="pdf-ws-selbtn"
                    onClick={() => setSelected(new Set(sheets.map((s) => s.name)))}
                    disabled={allOn}
                  >
                    Select all
                  </button>
                  <button
                    className="pdf-ws-selbtn"
                    onClick={() => setSelected(new Set())}
                    disabled={selected.size === 0}
                  >
                    Clear
                  </button>
                </div>
                <div style={{ display: 'grid', gap: 4, maxHeight: 220, overflow: 'auto' }}>
                  {sheets.map((s) => (
                    <label key={s.name} style={{ display: 'flex', gap: 6, alignItems: 'center', fontSize: 13 }}>
                      <input type="checkbox" checked={selected.has(s.name)} onChange={() => toggle(s.name)} />
                      <span
                        style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                        title={s.name}
                      >
                        {s.name}
                      </span>
                      <span style={{ opacity: 0.6 }}>
                        {s.rows}×{s.cols}
                      </span>
                    </label>
                  ))}
                </div>

                <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
                  Paper
                  <select value={paper} onChange={(e) => setPaper(e.target.value as PrintPaper)}>
                    <option value="a4">A4</option>
                    <option value="a3">A3</option>
                    <option value="letter">Letter</option>
                    <option value="legal">Legal</option>
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
                  Orientation
                  <select value={orientation} onChange={(e) => setOrientation(e.target.value as PrintOrientation)}>
                    <option value="auto">Auto (best fit per sheet)</option>
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                  </select>
                </label>

                <div className="pdf-ws-rail-actions">
                  <button className="primary" onClick={savePdf} disabled={!pdf || busy}>
                    <IconDownload /> Download PDF
                  </button>
                  <button className="pdf-ws-railbtn" onClick={printNow} disabled={selected.size === 0 || busy}>
                    <IconPrint /> Print
                  </button>
                </div>

                {saved && (
                  <div className="notice tool-outcome">
                    <IconCheck /> Saved: {saved}
                    <button className="ghost" onClick={() => api.openPath(saved)}>
                      <IconOpen /> Open
                    </button>
                  </div>
                )}
              </>
            )}
          </aside>
        </div>
      </div>
    </>
  )
}
