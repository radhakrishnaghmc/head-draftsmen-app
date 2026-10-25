import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { pdfPagesToDataUrlsFromData } from '../pdfToImages'
import { IconClose, IconDoc, IconDownload, IconPrint } from './Icons'

interface Props {
  bytes: Uint8Array
  /** Suggested file name for Save PDF. */
  fileName: string
  /** When given, an Excel button saves the same list as a spreadsheet. */
  onExcel?: () => Promise<unknown>
  onClose: () => void
}

/** Preview of a generated PDF (Chromium's viewer in an iframe) with Print and Save PDF. */
export default function PdfPreviewModal({ bytes, fileName, onExcel, onClose }: Props) {
  const url = useMemo(() => URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type: 'application/pdf' })), [bytes])
  useEffect(() => () => URL.revokeObjectURL(url), [url])
  const [busy, setBusy] = useState<'print' | 'save' | 'excel' | null>(null)
  const [error, setError] = useState('')

  async function save() {
    setBusy('save')
    setError('')
    try {
      await api.savePdf(bytes, fileName)
    } catch (e) {
      setError(`Couldn't save the PDF: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setBusy(null)
    }
  }

  async function excel() {
    if (!onExcel) return
    setBusy('excel')
    setError('')
    try {
      await onExcel()
    } catch (e) {
      setError(`Couldn't save the Excel: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setBusy(null)
    }
  }

  // Print goes through the app's own print window (same path as every created document):
  // the PDF's pages are rasterised into one full-bleed image per page, sized to the page's own orientation.
  async function print() {
    setBusy('print')
    setError('')
    try {
      const pages = await pdfPagesToDataUrlsFromData(bytes.slice(), 2.5)
      const img = new Image()
      img.src = pages[0]
      await img.decode()
      const landscape = img.naturalWidth > img.naturalHeight
      const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
        @page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 0 }
        html, body { margin: 0; padding: 0; background: #fff }
        img { display: block; width: 100%; height: ${landscape ? '210mm' : '297mm'}; object-fit: contain; page-break-after: always; break-after: page }
        img:last-child { page-break-after: auto; break-after: auto }
        @media screen { html { background: #d8d8dc } img { margin: 12px auto; width: ${landscape ? '297mm' : '210mm'}; background: #fff } }
      </style></head><body>${pages.map((p) => `<img src="${p}">`).join('')}</body></html>`
      await api.printCreatedDocument(html)
    } catch (e) {
      setError(`Couldn't print the PDF: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setBusy(null)
    }
  }

  return createPortal(
    <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(onClose)}>
      <div className="pdf-preview-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-preview-head">
          <span className="pdf-preview-title">
            <IconDoc /> {fileName.replace(/\.pdf$/i, '')}
          </span>
          <div className="pdf-preview-head-actions">
            {error && <span className="mf-combine-count">{error}</span>}
            <button className="pdf-ws-railbtn" onClick={print} disabled={busy !== null}>
              <IconPrint /> {busy === 'print' ? 'Preparing…' : 'Print'}
            </button>
            <button className="pdf-ws-railbtn" onClick={save} disabled={busy !== null}>
              <IconDownload /> {busy === 'save' ? 'Saving…' : 'Save PDF'}
            </button>
            {onExcel && (
              <button className="pdf-ws-railbtn" onClick={excel} disabled={busy !== null}>
                <IconDownload /> {busy === 'excel' ? 'Saving…' : 'Excel'}
              </button>
            )}
            <button className="pdf-ws-railbtn" onClick={onClose}>
              <IconClose /> Close
            </button>
          </div>
        </div>
        <iframe className="pdf-preview-frame" src={url} title="PDF preview" />
      </div>
    </div>,
    document.body
  )
}
