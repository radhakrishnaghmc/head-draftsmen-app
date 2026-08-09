import { useCallback, useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { IconChevronLeft, IconDownload, IconDoc, IconFolder, IconTrash, IconWarn, IconOpen } from './Icons'
import { PdfPageThumb } from './PdfToolPage'

// A Word document the user uploaded, already split into one .docx per page (at
// its page breaks). Each section is a full-formatting .docx; a PDF rendition of
// its first page is drawn lazily as the tile thumbnail. The section bytes are
// also kept in sectionsRef (outside state) so the lazy thumbnail renderer can
// reach them without a stale closure.
interface WordSource {
  id: string
  name: string
  base: string
  docxBytes: Uint8Array
  sections: Uint8Array[]
}

const keyOf = (id: string, i: number) => `${id}:${i}`
const THUMB_SCALE = 0.5

type Busy = 'loading' | 'merge' | 'save' | 'files' | null
type Result = { kind: 'docx'; file: string } | { kind: 'folder'; dir: string; count: number } | null

interface Props {
  onBack: () => void
}

/**
 * The Word Merge / Separator page. Upload one or more .docx files; each is split
 * at its page breaks into one full-formatting .docx per page, shown as tiles
 * (previewed via a PDF conversion). Pick pages across files, in order, then merge
 * the picked pages into one .docx or save each on its own — or merge the whole
 * uploaded files into one .docx. The split/merge/save are pure .docx operations
 * (they keep styles, tables and images and don't need LibreOffice); only the tile
 * previews use the docx→PDF conversion.
 */
export default function WordToolPage({ onBack }: Props) {
  const [sources, setSources] = useState<WordSource[]>([])
  // Picked pages in the order ticked — the serial the merged file follows.
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
  const sectionsRef = useRef<Map<string, Uint8Array[]>>(new Map())
  const thumbCacheRef = useRef<Map<string, string>>(new Map())

  const totalPages = sources.reduce((n, s) => n + s.sections.length, 0)
  const selectedCount = order.length
  const working = busy === 'merge' || busy === 'save' || busy === 'files'

  // Draw one page's thumbnail on demand: convert that section .docx to PDF
  // (LibreOffice) and render its first page. Cached; a failure (e.g. no
  // LibreOffice) returns null and the tile shows a small warning — the actual
  // merge/split still work without it.
  const renderThumb = useCallback(async (id: string, sectionIndex: number): Promise<string | null> => {
    const key = keyOf(id, sectionIndex)
    const cached = thumbCacheRef.current.get(key)
    if (cached) return cached
    const bytes = sectionsRef.current.get(id)?.[sectionIndex]
    if (!bytes) return null
    try {
      const pdfBytes = await api.docxToPdf(bytes)
      const { pdfjsLib } = await import('../pdfjsSetup')
      const task = pdfjsLib.getDocument({ data: pdfBytes.slice() })
      const doc = await task.promise
      const page = await doc.getPage(1)
      const viewport = page.getViewport({ scale: THUMB_SCALE })
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.floor(viewport.width))
      canvas.height = Math.max(1, Math.floor(viewport.height))
      const ctx = canvas.getContext('2d')
      if (!ctx) {
        void task.destroy()
        return null
      }
      await page.render({ canvas, canvasContext: ctx, viewport }).promise
      const url = canvas.toDataURL('image/jpeg', 0.82)
      void task.destroy()
      thumbCacheRef.current.set(key, url)
      return url
    } catch {
      return null
    }
  }, [])

  // Read each picked .docx and split it into per-page sections up front (fast —
  // pure XML). Thumbnails render lazily afterwards.
  async function addFiles(files: File[]) {
    setError(null)
    setResult(null)
    setBusy('loading')
    try {
      for (const file of files) {
        if (!/\.docx?$/i.test(file.name)) continue
        const docxBytes = new Uint8Array(await file.arrayBuffer())
        let sections: Uint8Array[]
        try {
          sections = await api.splitDocxSections(docxBytes)
        } catch (e) {
          setError(`Couldn't read "${file.name}"${e instanceof Error ? ` — ${e.message}` : ''}.`)
          continue
        }
        if (sections.length === 0) sections = [docxBytes]
        const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`
        sectionsRef.current.set(id, sections)
        const base = file.name.replace(/\.docx?$/i, '')
        setSources((cur) => [...cur, { id, name: file.name, base, docxBytes, sections }])
      }
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

  function toggle(id: string, i: number) {
    const k = keyOf(id, i)
    setOrder((cur) => (cur.includes(k) ? cur.filter((x) => x !== k) : [...cur, k]))
  }

  function toggleAllIn(src: WordSource) {
    const keys = src.sections.map((_, i) => keyOf(src.id, i))
    const allOn = keys.every((k) => selectedSet.has(k))
    setOrder((cur) => {
      if (allOn) return cur.filter((k) => !keys.includes(k))
      const missing = keys.filter((k) => !cur.includes(k))
      return [...cur, ...missing]
    })
  }

  function selectAll() {
    const all: string[] = []
    for (const s of sources) for (let i = 0; i < s.sections.length; i++) all.push(keyOf(s.id, i))
    setOrder(all)
  }

  function removeSource(id: string) {
    sectionsRef.current.delete(id)
    for (const k of Array.from(thumbCacheRef.current.keys())) if (k.startsWith(`${id}:`)) thumbCacheRef.current.delete(k)
    setSources((cur) => cur.filter((s) => s.id !== id))
    setOrder((cur) => cur.filter((k) => !k.startsWith(`${id}:`)))
  }

  function clearAll() {
    sectionsRef.current.clear()
    thumbCacheRef.current.clear()
    setSources([])
    setOrder([])
    setError(null)
    setResult(null)
  }

  // Picked pages, in serial (pick) order, resolved to their section bytes + a name.
  function orderedSelection(): { bytes: Uint8Array; base: string; pageNo: number }[] {
    const byKey = new Map<string, { bytes: Uint8Array; base: string; pageNo: number }>()
    for (const s of sources) {
      s.sections.forEach((bytes, i) => byKey.set(keyOf(s.id, i), { bytes, base: s.base, pageNo: i + 1 }))
    }
    const out: { bytes: Uint8Array; base: string; pageNo: number }[] = []
    for (const k of order) {
      const v = byKey.get(k)
      if (v) out.push(v)
    }
    return out
  }

  async function mergeSelected() {
    const picks = orderedSelection()
    if (picks.length === 0 || working) return
    setBusy('merge')
    setError(null)
    setResult(null)
    try {
      const res = await api.mergeDocx(picks.map((p) => p.bytes))
      if (res) setResult({ kind: 'docx', file: res.file })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function saveEachSelected() {
    const picks = orderedSelection()
    if (picks.length === 0 || working) return
    setBusy('save')
    setError(null)
    setResult(null)
    try {
      const used = new Set<string>()
      const files = picks.map((p) => {
        let name = `${p.base} p${p.pageNo}.docx`
        let dupe = 2
        while (used.has(name.toLowerCase())) name = `${p.base} p${p.pageNo} (${dupe++}).docx`
        used.add(name.toLowerCase())
        return { name, bytes: p.bytes }
      })
      const res = await api.saveDocxsToFolder(files)
      if (res) setResult({ kind: 'folder', dir: res.dir, count: res.files.length })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function mergeWholeFiles() {
    if (sources.length < 2 || working) return
    setBusy('files')
    setError(null)
    setResult(null)
    try {
      const res = await api.mergeDocx(sources.map((s) => s.docxBytes))
      if (res) setResult({ kind: 'docx', file: res.file })
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
          <IconDoc /> <span className="pdf-ws-title-label">Tool:</span> Word Merge / Separator
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack} disabled={working}>
          <IconChevronLeft /> Back to Tools
        </button>
      </div>

      <div className="card pdf-workspace">
      <input ref={fileInputRef} type="file" accept=".docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden onChange={onPick} />

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
              <span className="pdf-ws-dropzone-title">{busy === 'loading' ? 'Reading…' : 'Upload Word file(s)'}</span>
              <span className="pdf-ws-dropzone-sub">Pick one .docx or several at once</span>
            </button>
          ) : (
            <div className="pdf-ws-docs">
              {sources.map((src) => {
                const keys = src.sections.map((_, i) => keyOf(src.id, i))
                const allOn = keys.length > 0 && keys.every((k) => selectedSet.has(k))
                const selInDoc = keys.filter((k) => selectedSet.has(k)).length
                return (
                  <section key={src.id} className="pdf-ws-doc">
                    <header className="pdf-ws-doc-head">
                      <span className="pdf-ws-doc-name" title={src.name}>
                        {src.name}
                      </span>
                      <span className="pdf-ws-doc-meta">
                        {src.sections.length} page{src.sections.length === 1 ? '' : 's'}
                        {selInDoc > 0 && ` · ${selInDoc} selected`}
                      </span>
                      <span className="pdf-ws-doc-ctl">
                        <button className="ghost" onClick={() => toggleAllIn(src)}>
                          {allOn ? 'Deselect all' : 'Select all'}
                        </button>
                        <button
                          className="ghost"
                          onClick={() => removeSource(src.id)}
                          title="Remove this document"
                          aria-label="Remove this document"
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
            <IconFolder /> {busy === 'loading' ? 'Reading…' : hasSources ? 'Add more Word files' : 'Upload Word file(s)'}
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
                <div className="pdf-ws-rail-grouplabel">Selected pages → Word</div>
                <button className="primary" onClick={mergeSelected} disabled={selectedCount === 0 || working}>
                  <IconDoc /> {busy === 'merge' ? 'Merging…' : 'Merge selected → one .docx'}
                </button>
                <button className="pdf-ws-railbtn" onClick={saveEachSelected} disabled={selectedCount === 0 || working}>
                  <IconDownload /> {busy === 'save' ? 'Saving…' : 'Save each selected page'}
                </button>
              </div>

              <div className="pdf-ws-rail-actions">
                <div className="pdf-ws-rail-grouplabel">Whole files → Word</div>
                <button className="pdf-ws-railbtn" onClick={mergeWholeFiles} disabled={sources.length < 2 || working}>
                  <IconDoc /> {busy === 'files' ? 'Merging…' : 'Merge files → one .docx'}
                </button>
                {sources.length < 2 && <p className="pdf-ws-rail-hint">Add at least two Word files to merge whole files.</p>}
              </div>

              <button className="pdf-ws-clearall" onClick={clearAll} disabled={working}>
                <IconTrash /> Remove all files
              </button>
            </>
          )}

          {result && (
            <div className="notice ok tool-result pdf-ws-rail-result">
              <span>
                {result.kind === 'folder' ? (
                  <>
                    <IconFolder /> Saved {result.count} file{result.count === 1 ? '' : 's'} to {result.dir}
                  </>
                ) : (
                  <>
                    <IconDoc /> Saved: {result.file}
                  </>
                )}
              </span>
              <button className="ghost" onClick={() => api.openPath(result.kind === 'folder' ? result.dir : result.file)}>
                <IconOpen /> Open
              </button>
            </div>
          )}

          <p className="pdf-ws-rail-hint">
            Pages are split at the document's page breaks, and every output stays in Word (.docx) with its formatting,
            tables and images kept. Previews need LibreOffice, but the merge/split do not.
          </p>
        </aside>
      </div>
      </div>
    </>
  )
}
