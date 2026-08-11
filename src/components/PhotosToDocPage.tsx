import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfPagesToDataUrls } from '../pdfToImages'
import { extractPdfGeometry } from '../pdfReconstruct'
import { geometryToBlocks, geometryHasText } from '@core/reconstructDoc'
import { layoutToBlocks } from '@core/ocrReconstruct'
import { blocksToRows, type DocBlock } from '@core/docxBuilder'
import {
  IconChevronLeft,
  IconImage,
  IconFolder,
  IconDoc,
  IconTable,
  IconDownload,
  IconTrash,
  IconWarn,
  IconOpen
} from './Icons'

// One picked page image, held as a data URL (in reading order — drag to reorder).
interface Pic {
  id: string
  name: string
  dataUrl: string
}

function nextId(): string {
  return `pic-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(file)
  })
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

interface Props {
  onBack: () => void
}

type Busy = null | 'load' | 'word' | 'excel'

// One uploaded PDF, kept as raw bytes so it can be converted to Word with its
// layout preserved (LibreOffice), separate from the page-images used for OCR.
interface PdfFile {
  name: string
  bytes: Uint8Array
}

/**
 * Tool: convert photos (or the pages of an uploaded PDF) into a Word (.docx) or
 * Excel (.xlsx) document. Each page is OCR'd with the app's own local engine
 * (the same one behind Estimate-from-Photos), the recognised text is shown in
 * one editable box for review/correction, and then exported — Word gets one
 * paragraph per line, Excel one row per line (split into columns on wide gaps).
 * All the heavy work (OCR + file building) is in the main process; this page
 * just gathers the images, shows the editable text, and triggers the save.
 */
export default function PhotosToDocPage({ onBack }: Props) {
  const [pics, setPics] = useState<Pic[]>([])
  const [pdfFiles, setPdfFiles] = useState<PdfFile[]>([])
  const [busy, setBusy] = useState<Busy>(null)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const hasPics = pics.length > 0
  const working = busy !== null
  const suggestedName = pics[0] ? stripExt(pics[0].name).replace(/\s*\(page \d+\)$/i, '') : 'Document'

  async function addFiles(files: File[]) {
    setError(null)
    setSaved(null)
    setBusy('load')
    try {
      const added: Pic[] = []
      const addedPdfs: PdfFile[] = []
      for (const file of files) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // Keep the raw PDF bytes (for layout-preserving Word) AND image its
          // pages (for the OCR/text path + thumbnails).
          addedPdfs.push({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) })
          const pages = await pdfPagesToDataUrls(file)
          const base = stripExt(file.name)
          pages.forEach((dataUrl, i) => added.push({ id: nextId(), name: `${base} (page ${i + 1})`, dataUrl }))
        } else if (file.type.startsWith('image/') || /\.(jpe?g|png|gif|bmp|webp)$/i.test(file.name)) {
          added.push({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) })
        }
      }
      if (added.length === 0) throw new Error('No photos or PDFs were found in the selection.')
      setPics((cur) => [...cur, ...added])
      if (addedPdfs.length > 0) setPdfFiles((cur) => [...cur, ...addedPdfs])
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  function onPick(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files ? Array.from(e.target.files) : []
    e.target.value = ''
    if (files.length) void addFiles(files)
  }

  function removePic(id: string) {
    setPics((cur) => cur.filter((p) => p.id !== id))
  }

  function clearAll() {
    setPics([])
    setPdfFiles([])
    setError(null)
    setSaved(null)
  }

  function onDragStart(e: React.DragEvent, i: number) {
    setDragIndex(i)
    e.dataTransfer.effectAllowed = 'move'
  }
  function onDragOver(e: React.DragEvent, i: number) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    if (i !== overIndex) setOverIndex(i)
  }
  function onDrop(e: React.DragEvent, i: number) {
    e.preventDefault()
    if (dragIndex !== null && dragIndex !== i) {
      setPics((cur) => {
        const next = [...cur]
        const [moved] = next.splice(dragIndex, 1)
        next.splice(i, 0, moved)
        return next
      })
    }
    setDragIndex(null)
    setOverIndex(null)
  }
  function onDragEnd() {
    setDragIndex(null)
    setOverIndex(null)
  }

  // Download Word: for uploaded PDF(s), rebuild them as an editable .docx that
  // preserves the layout (real Word tables/paragraphs reconstructed from the
  // PDF's own text + ruling lines); a scanned PDF with no text falls back to the
  // Build the doc-model blocks for the current upload:
  //  • a PDF with a real text layer → precise reconstruction from that text
  //    (exact characters, no OCR errors) — best for the office's PDF documents;
  //  • photos, or scanned/no-text PDFs → lightweight image reconstruction from
  //    the OCR detector's own text boxes (offline, no extra model).
  async function getBlocks(): Promise<DocBlock[]> {
    if (pdfFiles.length > 0) {
      const acc: DocBlock[] = []
      let anyText = false
      for (const pdf of pdfFiles) {
        const pages = await extractPdfGeometry(pdf.bytes)
        if (geometryHasText(pages)) {
          anyText = true
          if (acc.length > 0) acc.push({ kind: 'paragraph', runs: [], pageBreak: true })
          acc.push(...geometryToBlocks(pages))
        }
      }
      if (anyText) return acc
    }
    return layoutToBlocks(await api.ocrPhotosToLayout(pics.map((p) => p.dataUrl)))
  }

  async function downloadWord() {
    if (!hasPics || working) return
    setBusy('word')
    setError(null)
    setSaved(null)
    try {
      const path = await api.saveWordDoc(await getBlocks(), suggestedName)
      if (path) setSaved(path)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function downloadExcel() {
    if (!hasPics || working) return
    setBusy('excel')
    setError(null)
    setSaved(null)
    try {
      const path = await api.saveRowsAsExcel(blocksToRows(await getBlocks()), suggestedName)
      if (path) setSaved(path)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconDoc /> <span className="pdf-ws-title-label">Tool:</span> Photos / PDF → Word / Excel
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={working}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <input ref={fileInputRef} type="file" accept="image/*,application/pdf,.pdf" multiple hidden onChange={onPick} />

      <div className="card pdf-workspace">
        <div className="pdf-ws-body">
          <div className="pdf-ws-main">
            {error && (
              <div className="notice error tool-outcome">
                <IconWarn /> {error}
              </div>
            )}

            {!hasPics ? (
              <button className="pdf-ws-dropzone" onClick={() => fileInputRef.current?.click()} disabled={working}>
                <IconImage />
                <span className="pdf-ws-dropzone-title">{working ? 'Reading…' : 'Upload photos or PDF'}</span>
                <span className="pdf-ws-dropzone-sub">
                  Each photo (and each PDF page) is read with OCR, then exported to Word or Excel
                </span>
              </button>
            ) : (
              <>
                <div className="photo-tile-grid">
                  {pics.map((pic, i) => (
                    <div
                      key={pic.id}
                      className={[
                        'photo-tile-card',
                        dragIndex === i ? 'dragging' : '',
                        overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      draggable
                      onDragStart={(e) => onDragStart(e, i)}
                      onDragOver={(e) => onDragOver(e, i)}
                      onDrop={(e) => onDrop(e, i)}
                      onDragEnd={onDragEnd}
                    >
                      <span className="photo-tile-number">{i + 1}</span>
                      <button className="photo-tile-remove" title="Remove" onClick={() => removePic(pic.id)}>
                        <IconTrash />
                      </button>
                      <img className="photo-tile-img" src={pic.dataUrl} alt={pic.name} />
                      <div className="photo-tile-name" title={pic.name}>
                        {pic.name}
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>

          <aside className="pdf-ws-rail">
            <button className="primary pdf-ws-upload" onClick={() => fileInputRef.current?.click()} disabled={working}>
              <IconFolder /> {working ? 'Working…' : hasPics ? 'Add more' : 'Upload photos or PDF'}
            </button>

            {hasPics && (
              <>
                <div className="pdf-ws-rail-count">
                  <strong>{pics.length}</strong> page{pics.length === 1 ? '' : 's'} — drag to reorder
                </div>

                <div className="pdf-ws-rail-actions">
                  <button className="primary" onClick={downloadWord} disabled={working}>
                    <IconDoc /> {busy === 'word' ? 'Working…' : 'Download Word'}
                  </button>
                  <button className="pdf-ws-railbtn" onClick={downloadExcel} disabled={working}>
                    <IconTable /> {busy === 'excel' ? 'Working…' : 'Download Excel'}
                  </button>
                </div>

                <button className="pdf-ws-clearall" onClick={clearAll} disabled={working}>
                  <IconTrash /> Remove all
                </button>
              </>
            )}

            {saved && (
              <div className="notice ok tool-result pdf-ws-rail-result">
                <span>
                  <IconDoc /> Saved: {saved}
                </span>
                <button className="ghost" onClick={() => api.openPath(saved)}>
                  <IconOpen /> Open
                </button>
              </div>
            )}

            <p className="pdf-ws-rail-hint">
              A <strong>PDF with real text</strong> is rebuilt exactly from that text (no OCR errors) — best for office
              documents. <strong>Photos and scans</strong> use offline table-AI (cell detection + OCR), which is
              best-effort — review the result. Download Word gives an editable document; Download Excel a spreadsheet.
            </p>
          </aside>
        </div>
      </div>
    </>
  )
}
