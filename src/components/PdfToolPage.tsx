import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { PDFDocument } from 'pdf-lib'
import type { PDFDocumentProxy, PDFDocumentLoadingTask } from 'pdfjs-dist'
import { api } from '../ipc'
import { IconChevronLeft, IconDownload, IconDoc, IconFolder, IconTrash, IconWarn, IconOpen } from './Icons'

// A PDF the user uploaded: its raw bytes (kept for pdf-lib to copy pages from)
// and how many pages it has. The pdf.js document used to draw thumbnails is
// held in a ref (docsRef), not in state — it isn't serialisable and each page is
// only rendered lazily, when it scrolls into view.
interface PdfSource {
  id: string
  name: string
  base: string
  bytes: Uint8Array
  pageCount: number
}

// One page, addressed as `${sourceId}:${pageIndex}` (0-based page).
const keyOf = (id: string, i: number) => `${id}:${i}`

// Thumbnails only need to be legible, not print-sharp — half a page's 72dpi
// keeps them small so even a 40-page upload stays light.
const THUMB_SCALE = 0.5

type Busy = 'loading' | 'merge' | 'individual' | 'all' | null
type Result =
  | { kind: 'merge'; file: string }
  | { kind: 'folder'; dir: string; count: number }
  | null

interface Props {
  /** Return to the Tools grid. */
  onBack: () => void
}

/**
 * The PDF workspace behind the single "PDF Merge / Separator" tile: a full page
 * where the user uploads one or more PDFs, sees every page as a thumbnail (each
 * file stacked one below the other), ticks the pages they want — freely across
 * files — then, from the action rail on the right, either merges those pages into
 * one new PDF or saves each selected page as its own PDF. Thumbnails render
 * lazily (only as they scroll into view) so a large document — e.g. a 41-page
 * tender — loads and stays responsive. All page-copying is done here with
 * pdf-lib; the main process only writes the finished bytes to disk.
 */
export default function PdfToolPage({ onBack }: Props) {
  const [sources, setSources] = useState<PdfSource[]>([])
  // Selected pages in the order they were ticked — the serial the merged output
  // follows (and the number shown on each page). A derived Set/Map give O(1)
  // membership and serial lookups during render.
  const [order, setOrder] = useState<string[]>([])
  const selectedSet = useMemo(() => new Set(order), [order])
  const serialOf = useMemo(() => {
    const m = new Map<string, number>()
    order.forEach((k, i) => m.set(k, i + 1))
    return m
  }, [order])
  const [busy, setBusy] = useState<Busy>(null)
  const [error, setError] = useState<string | null>(null)
  const [result, setResult] = useState<Result>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  // Live pdf.js documents (for drawing thumbnails on demand) and a cache of the
  // thumbnails already drawn — both keyed by source id / page key, kept out of
  // React state since they're non-serialisable and change without a re-render.
  // The loading task is held alongside the doc because releasing the underlying
  // worker document (destroy) is a method of the task, not of the doc proxy.
  const docsRef = useRef<Map<string, { doc: PDFDocumentProxy; task: PDFDocumentLoadingTask }>>(new Map())
  const thumbCacheRef = useRef<Map<string, string>>(new Map())

  const totalPages = sources.reduce((n, s) => n + s.pageCount, 0)
  const selectedCount = order.length
  const working = busy === 'merge' || busy === 'individual' || busy === 'all'

  useEffect(() => {
    // Release every pdf.js document when the workspace unmounts.
    const docs = docsRef.current
    return () => {
      for (const d of docs.values()) void d.task.destroy()
      docs.clear()
    }
  }, [])

  // Draw one page to a JPEG data URL, on demand. Cached so scrolling a page out
  // and back doesn't redraw it; a failing page returns null (its tile shows a
  // small warning) instead of taking down the whole upload.
  const renderThumb = useCallback(async (id: string, pageIndex: number): Promise<string | null> => {
    const key = keyOf(id, pageIndex)
    const cached = thumbCacheRef.current.get(key)
    if (cached) return cached
    const entry = docsRef.current.get(id)
    if (!entry) return null
    try {
      const page = await entry.doc.getPage(pageIndex + 1)
      const viewport = page.getViewport({ scale: THUMB_SCALE })
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.floor(viewport.width))
      canvas.height = Math.max(1, Math.floor(viewport.height))
      const ctx = canvas.getContext('2d')
      if (!ctx) return null
      await page.render({ canvas, canvasContext: ctx, viewport }).promise
      const url = canvas.toDataURL('image/jpeg', 0.82)
      thumbCacheRef.current.set(key, url)
      return url
    } catch {
      return null
    }
  }, [])

  // Read the picked files and append each as a source. Only its page count is
  // needed up front — the pages themselves draw lazily — so even a big PDF is
  // ready the moment its page count is read.
  async function addFiles(files: File[]) {
    setError(null)
    setResult(null)
    setBusy('loading')
    try {
      const { pdfjsLib } = await import('../pdfjsSetup')
      for (const file of files) {
        if (!file.name.toLowerCase().endsWith('.pdf')) continue
        const bytes = new Uint8Array(await file.arrayBuffer())
        const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`
        let doc: PDFDocumentProxy
        let task: PDFDocumentLoadingTask
        try {
          // pdf.js may transfer its data buffer to the worker, so hand it a copy
          // and keep `bytes` intact for pdf-lib to read from at save time.
          task = pdfjsLib.getDocument({ data: bytes.slice() })
          doc = await task.promise
        } catch (e) {
          setError(`Couldn't read "${file.name}"${e instanceof Error ? ` — ${e.message}` : ''}.`)
          continue
        }
        docsRef.current.set(id, { doc, task })
        const base = file.name.replace(/\.pdf$/i, '')
        setSources((cur) => [...cur, { id, name: file.name, base, bytes, pageCount: doc.numPages }])
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  function onPick(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files ? Array.from(e.target.files) : []
    e.target.value = '' // let the same file be picked again after removal
    if (files.length) void addFiles(files)
  }

  // Toggle a page: appending it to the end of the order (so it takes the next
  // serial) or dropping it and letting the pages after it shift up a number.
  function toggle(id: string, i: number) {
    const k = keyOf(id, i)
    setOrder((cur) => (cur.includes(k) ? cur.filter((x) => x !== k) : [...cur, k]))
  }

  // Select / clear every page of one source. Selecting appends the not-yet-picked
  // pages in page order, so they keep reading order within their serials.
  function toggleAllIn(src: PdfSource) {
    const keys = Array.from({ length: src.pageCount }, (_, i) => keyOf(src.id, i))
    const allOn = keys.every((k) => selectedSet.has(k))
    setOrder((cur) => {
      if (allOn) return cur.filter((k) => !keys.includes(k))
      const missing = keys.filter((k) => !cur.includes(k))
      return [...cur, ...missing]
    })
  }

  // Select every page, in document (reading) order.
  function selectAll() {
    const all: string[] = []
    for (const s of sources) for (let i = 0; i < s.pageCount; i++) all.push(keyOf(s.id, i))
    setOrder(all)
  }

  function removeSource(id: string) {
    void docsRef.current.get(id)?.task.destroy()
    docsRef.current.delete(id)
    for (const k of Array.from(thumbCacheRef.current.keys())) if (k.startsWith(`${id}:`)) thumbCacheRef.current.delete(k)
    setSources((cur) => cur.filter((s) => s.id !== id))
    setOrder((cur) => cur.filter((k) => !k.startsWith(`${id}:`)))
  }

  function clearAll() {
    for (const d of docsRef.current.values()) void d.task.destroy()
    docsRef.current.clear()
    thumbCacheRef.current.clear()
    setSources([])
    setOrder([])
    setError(null)
    setResult(null)
  }

  // The pages to act on. For a selection, they come out in the serial order the
  // user ticked them (that's what the merged file follows). `all` ignores the
  // selection and takes every page in document (reading) order.
  function orderedPages(all = false): { source: PdfSource; pageIndex: number }[] {
    const byKey = new Map<string, { source: PdfSource; pageIndex: number }>()
    for (const source of sources) {
      for (let i = 0; i < source.pageCount; i++) byKey.set(keyOf(source.id, i), { source, pageIndex: i })
    }
    if (all) return [...byKey.values()]
    const out: { source: PdfSource; pageIndex: number }[] = []
    for (const k of order) {
      const v = byKey.get(k)
      if (v) out.push(v)
    }
    return out
  }

  async function mergePages(all: boolean) {
    const picks = orderedPages(all)
    if (picks.length === 0 || working) return
    setBusy(all ? 'all' : 'merge')
    setError(null)
    setResult(null)
    try {
      const out = await PDFDocument.create()
      const cache = new Map<string, PDFDocument>()
      for (const { source, pageIndex } of picks) {
        let doc = cache.get(source.id)
        if (!doc) {
          doc = await PDFDocument.load(source.bytes, { ignoreEncryption: true })
          cache.set(source.id, doc)
        }
        const [pg] = await out.copyPages(doc, [pageIndex])
        out.addPage(pg)
      }
      const bytes = await out.save()
      const res = await api.savePdf(bytes, 'Merged.pdf')
      if (res) setResult({ kind: 'merge', file: res.file })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function saveSeparately() {
    const picks = orderedPages(false)
    if (picks.length === 0 || working) return
    setBusy('individual')
    setError(null)
    setResult(null)
    try {
      const cache = new Map<string, PDFDocument>()
      const files: { name: string; bytes: Uint8Array }[] = []
      const used = new Set<string>()
      for (const { source, pageIndex } of picks) {
        let doc = cache.get(source.id)
        if (!doc) {
          doc = await PDFDocument.load(source.bytes, { ignoreEncryption: true })
          cache.set(source.id, doc)
        }
        const out = await PDFDocument.create()
        const [pg] = await out.copyPages(doc, [pageIndex])
        out.addPage(pg)
        const bytes = await out.save()
        let name = `${source.base} p${pageIndex + 1}.pdf`
        let dupe = 2
        while (used.has(name.toLowerCase())) name = `${source.base} p${pageIndex + 1} (${dupe++}).pdf`
        used.add(name.toLowerCase())
        files.push({ name, bytes })
      }
      const res = await api.savePdfsToFolder(files)
      if (res) setResult({ kind: 'folder', dir: res.dir, count: res.files.length })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const hasSources = sources.length > 0

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconDoc /> <span className="pdf-ws-title-label">Tool:</span> PDF Merge / Separator
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={working}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <div className="card pdf-workspace">
      <input ref={fileInputRef} type="file" accept="application/pdf,.pdf" multiple hidden onChange={onPick} />

      <div className="pdf-ws-body">
        <div className="pdf-ws-main">
          {error && (
            <div className="notice error tool-outcome">
              <IconWarn /> {error}
            </div>
          )}

          {!hasSources ? (
            <button
              className="pdf-ws-dropzone"
              onClick={() => fileInputRef.current?.click()}
              disabled={busy === 'loading'}
            >
              <IconDoc />
              <span className="pdf-ws-dropzone-title">{busy === 'loading' ? 'Loading…' : 'Upload PDF(s)'}</span>
              <span className="pdf-ws-dropzone-sub">Pick one file or several at once</span>
            </button>
          ) : (
            <div className="pdf-ws-docs">
              {sources.map((src) => {
                const keys = Array.from({ length: src.pageCount }, (_, i) => keyOf(src.id, i))
                const allOn = keys.length > 0 && keys.every((k) => selectedSet.has(k))
                const selInDoc = keys.filter((k) => selectedSet.has(k)).length
                return (
                  <section key={src.id} className="pdf-ws-doc">
                    <header className="pdf-ws-doc-head">
                      <span className="pdf-ws-doc-name" title={src.name}>
                        {src.name}
                      </span>
                      <span className="pdf-ws-doc-meta">
                        {src.pageCount} page{src.pageCount === 1 ? '' : 's'}
                        {selInDoc > 0 && ` · ${selInDoc} selected`}
                      </span>
                      <span className="pdf-ws-doc-ctl">
                        <button className="ghost" onClick={() => toggleAllIn(src)}>
                          {allOn ? 'Deselect all' : 'Select all'}
                        </button>
                        <button
                          className="ghost"
                          onClick={() => removeSource(src.id)}
                          title="Remove this PDF"
                          aria-label="Remove this PDF"
                        >
                          <IconTrash />
                        </button>
                      </span>
                    </header>
                    <div className="pdf-ws-pages">
                      {keys.map((k, i) => (
                        <PdfPageThumb
                          key={k}
                          id={src.id}
                          pageIndex={i}
                          serial={serialOf.get(k) ?? 0}
                          onToggle={() => toggle(src.id, i)}
                          render={renderThumb}
                        />
                      ))}
                    </div>
                  </section>
                )
              })}
            </div>
          )}
        </div>

        <aside className="pdf-ws-rail">
          <button
            className="primary pdf-ws-upload"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy === 'loading'}
          >
            <IconFolder /> {busy === 'loading' ? 'Loading…' : hasSources ? 'Add more PDFs' : 'Upload PDF(s)'}
          </button>

          {hasSources && (
            <>
              <div className="pdf-ws-rail-count">
                <strong>{selectedCount}</strong> of {totalPages} page{totalPages === 1 ? '' : 's'} selected
              </div>

              <div className="pdf-ws-rail-selrow">
                <button className="pdf-ws-selbtn" onClick={selectAll} disabled={selectedCount === totalPages}>
                  Select all
                </button>
                <button className="pdf-ws-selbtn" onClick={() => setOrder([])} disabled={selectedCount === 0}>
                  Clear
                </button>
              </div>

              <div className="pdf-ws-rail-actions">
                <button className="primary" onClick={() => mergePages(false)} disabled={selectedCount === 0 || working}>
                  <IconDoc /> {busy === 'merge' ? 'Merging…' : 'Merge selected → one PDF'}
                </button>
                <button className="pdf-ws-railbtn" onClick={saveSeparately} disabled={selectedCount === 0 || working}>
                  <IconDownload /> {busy === 'individual' ? 'Saving…' : 'Save each selected page'}
                </button>
                <button className="pdf-ws-railbtn" onClick={() => mergePages(true)} disabled={totalPages === 0 || working}>
                  <IconDoc /> {busy === 'all' ? 'Merging…' : 'Merge all pages → one PDF'}
                </button>
              </div>

              <button className="pdf-ws-clearall" onClick={clearAll} disabled={working}>
                <IconTrash /> Remove all PDFs
              </button>
            </>
          )}

          {result && (
            <div className="notice ok tool-result pdf-ws-rail-result">
              <span>
                {result.kind === 'merge' ? (
                  <>
                    <IconDoc /> Saved: {result.file}
                  </>
                ) : (
                  <>
                    <IconFolder /> Saved {result.count} PDF{result.count === 1 ? '' : 's'} to {result.dir}
                  </>
                )}
              </span>
              <button className="ghost" onClick={() => api.openPath(result.kind === 'merge' ? result.file : result.dir)}>
                <IconOpen /> Open
              </button>
            </div>
          )}

          <p className="pdf-ws-rail-hint">
            Tick pages across any of the uploaded PDFs, then merge them into one file or save each on its own.
          </p>
        </aside>
      </div>
      </div>
    </>
  )
}

// One page tile. Draws its thumbnail only once it scrolls near the viewport
// (IntersectionObserver) so a big PDF doesn't render every page up front. Shared
// with the Word workspace, whose pages are the same PDF pages (from a docx→PDF
// conversion).
export function PdfPageThumb({
  id,
  pageIndex,
  serial,
  onToggle,
  render
}: {
  id: string
  pageIndex: number
  /** 1-based position in the pick order (the serial shown); 0 when not selected. */
  serial: number
  onToggle: () => void
  render: (id: string, pageIndex: number) => Promise<string | null>
}) {
  const selected = serial > 0
  const [url, setUrl] = useState<string | null>(null)
  const [failed, setFailed] = useState(false)
  const ref = useRef<HTMLButtonElement>(null)

  useEffect(() => {
    if (url) return
    const el = ref.current
    if (!el) return
    let cancelled = false
    const io = new IntersectionObserver(
      (entries) => {
        if (entries.some((e) => e.isIntersecting)) {
          io.disconnect()
          void render(id, pageIndex).then((u) => {
            if (cancelled) return
            if (u) setUrl(u)
            else setFailed(true)
          })
        }
      },
      { rootMargin: '300px' }
    )
    io.observe(el)
    return () => {
      cancelled = true
      io.disconnect()
    }
  }, [id, pageIndex, url, render])

  return (
    <button
      ref={ref}
      type="button"
      className={`pdf-ws-page${selected ? ' on' : ''}`}
      onClick={onToggle}
      aria-pressed={selected}
    >
      <span className="pdf-ws-page-thumb">
        {url ? (
          <img src={url} alt={`Page ${pageIndex + 1}`} />
        ) : failed ? (
          <span className="pdf-ws-page-fail" title="This page couldn't be previewed">
            <IconWarn />
          </span>
        ) : (
          <span className="pdf-ws-page-spin">…</span>
        )}
        <span className="pdf-ws-page-check" aria-hidden>
          {selected && serial}
        </span>
      </span>
      <span className="pdf-ws-page-no">Page {pageIndex + 1}</span>
    </button>
  )
}
