import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfPagesToDataUrlsFromData } from '../pdfToImages'
import { IconDoc, IconFolder, IconTable, IconTrash, IconWarn } from './Icons'
import { emptyMbModel, mergeMbPage, type MbModel } from '@core/measurementBook'

const avg = (v: (number | null)[]) => {
  const a = v.filter((x): x is number => x != null)
  return a.length ? a.reduce((x, y) => x + y, 0) / a.length : null
}

interface PagePreview {
  url: string
  ratio: number
}

const loadRatio = (url: string) =>
  new Promise<number>((resolve) => {
    const img = new Image()
    img.onload = () => resolve(img.naturalWidth / img.naturalHeight || 1)
    img.onerror = () => resolve(1)
    img.src = url
  })

export default function ExtractMeasurementPage() {
  const [busy, setBusy] = useState(false)
  const [step, setStep] = useState('')
  const [pages, setPages] = useState<PagePreview[]>([])
  const [current, setCurrent] = useState(0) // 1-based page being read; 0 = none
  const [done, setDone] = useState(0)
  const [model, setModel] = useState<MbModel>(emptyMbModel)
  const [fileName, setFileName] = useState('')
  const [saved, setSaved] = useState<string | null>(null)
  const [error, setError] = useState('')
  const fileInput = useRef<HTMLInputElement>(null)
  const stopped = useRef(false)

  async function stop() {
    stopped.current = true
    setStep('Stopping…')
    await api.cancelClaude()
  }

  // Stops any scan in flight and removes the loaded PDF with its results.
  async function clear() {
    if (busy) await stop()
    setPages([])
    setModel(emptyMbModel())
    setDone(0)
    setCurrent(0)
    setError('')
    setSaved(null)
  }

  async function choose(file?: File) {
    if (!file) return
    stopped.current = false
    setBusy(true)
    setError('')
    setModel(emptyMbModel())
    setSaved(null)
    setFileName(file.name.replace(/\.pdf$/i, ''))
    setPages([])
    setDone(0)
    setCurrent(0)
    try {
      const bytes = new Uint8Array(await file.arrayBuffer())
      const urls = await pdfPagesToDataUrlsFromData(bytes.slice(), 0.8)
      const previews = await Promise.all(urls.map(async (url) => ({ url, ratio: await loadRatio(url) })))
      if (stopped.current) return
      setPages(previews)
      if (!(await api.claudeAuthStatus())) {
        setStep('Sign in to your Claude account in the browser window that opened…')
        await api.claudeLogin()
      }
      for (let p = 1; p <= previews.length; p++) {
        if (stopped.current) break
        setCurrent(p)
        setStep(`Reading page ${p} of ${previews.length}`)
        const out = await api.extractMeasurements(bytes, p)
        if (stopped.current) break
        setModel((cur) => mergeMbPage(cur, p, out) ?? cur) // an unreadable reply just skips that page
        setDone(p)
      }
    } catch (e) {
      if (!stopped.current) setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      setCurrent(0)
      setStep('')
    }
  }

  const measured = model.items.reduce((n, i) => n + i.lines.length, 0)

  async function exportExcel() {
    setError('')
    try {
      const file = await api.exportMeasurementBook(model, fileName || 'MB')
      if (file) setSaved(file)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    }
  }

  return (
    <div className="card pdf-workspace">
      <input
        ref={fileInput}
        type="file"
        accept=".pdf"
        hidden
        onChange={(e) => {
          choose(e.target.files?.[0])
          e.target.value = ''
        }}
      />
      <div className="gps-toolbar">
        {/* With nothing loaded, the dropzone below is the upload control. */}
        {(busy || pages.length > 0) && (
          <button className="primary" onClick={() => fileInput.current?.click()} disabled={busy}>
            <IconFolder /> {busy ? 'Reading…' : 'Choose another PDF'}
          </button>
        )}
        {busy && (
          <button className="pdf-ws-clearbtn" onClick={stop}>
            Stop
          </button>
        )}
        {pages.length > 0 && (
          <button className="pdf-ws-clearbtn" onClick={clear}>
            <IconTrash /> Clear
          </button>
        )}
        {measured + model.abstract.length > 0 && (
          <button className="primary" onClick={exportExcel} disabled={busy}>
            <IconTable /> Export to Excel
          </button>
        )}
        {pages.length > 0 && (
          <span className="gps-counts">
            <strong>{pages.length}</strong> page{pages.length === 1 ? '' : 's'} · <strong>{model.items.length}</strong> item
            {model.items.length === 1 ? '' : 's'} · <strong>{measured}</strong> measurement line{measured === 1 ? '' : 's'}
          </span>
        )}
      </div>

      {busy && pages.length > 0 && current > 0 && (() => {
        // Same windowed coverflow as GPS Photos: the page being read in the centre,
        // a few neighbours either side (indices wrap, so no card ever jumps across).
        const N = pages.length
        const center = current - 1
        const W = Math.max(0, Math.min(3, Math.floor(N / 2) - 1))
        const STEP = 30
        const RADIUS = 360
        const MAXW = 380
        const MAXH = 300
        const cards = []
        for (let off = -W; off <= W; off++) cards.push({ off, idx: (((center + off) % N) + N) % N })
        return (
          <div className="gps-coverflow">
            <div className="gps-coverflow-stage">
              {cards.map(({ off, idx }) => {
                const p = pages[idx]
                let w = MAXW
                let h = MAXW / p.ratio
                if (h > MAXH) {
                  h = MAXH
                  w = MAXH * p.ratio
                }
                return (
                  <div
                    key={idx}
                    className={`gps-cf-card${off === 0 ? ' active' : ''}`}
                    style={{
                      width: `${Math.round(w)}px`,
                      height: `${Math.round(h)}px`,
                      marginLeft: `${-Math.round(w) / 2}px`,
                      marginTop: `${-Math.round(h) / 2}px`,
                      transform: `rotateY(${off * STEP}deg) translateZ(${RADIUS}px)`,
                      zIndex: 100 - Math.abs(off),
                      opacity: Math.abs(off) >= 3 ? 0.55 : 1
                    }}
                  >
                    <img src={p.url} alt={`Page ${idx + 1}`} />
                  </div>
                )
              })}
            </div>
            <div className="gps-coverflow-name">Reading page {current}</div>
          </div>
        )
      })()}

      {busy && (
        <div className="gps-progress">
          <span className="gps-progress-label">
            {step || 'Working…'}
            {pages.length > 0 && ` (${Math.round((done / pages.length) * 100)}%)`}
          </span>
          <span className="gps-progress-track" aria-hidden>
            <span className="gps-progress-bar" style={{ width: `${pages.length ? (done / pages.length) * 100 : 0}%` }} />
          </span>
        </div>
      )}

      {error && (
        <div className="notice error tool-outcome">
          <IconWarn /> {error}
        </div>
      )}

      {pages.length === 0 && !busy ? (
        <button className="pdf-ws-dropzone" onClick={() => fileInput.current?.click()}>
          <IconDoc />
          <span className="pdf-ws-dropzone-title">Upload a measurement book PDF</span>
          <span className="pdf-ws-dropzone-sub">
            Sign in to your Claude account if asked — Claude then reads the PDF page by page
          </span>
        </button>
      ) : (
        measured > 0 && (
          <div className="gps-table-wrap">
            <table className="gps-table">
              <thead>
                <tr>
                  {['Item / location', 'No.', 'L', 'B', 'D', 'Written'].map((h) => (
                    <th key={h}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {model.items.flatMap((it, i) => [
                  <tr key={`h${i}`}>
                    <td colSpan={6}>
                      <strong>
                        Item {i + 1} (page {it.mbPage}):
                      </strong>{' '}
                      {it.description}
                    </td>
                  </tr>,
                  ...it.lines.map((l, k) => (
                    <tr key={`${i}-${k}`}>
                      <td>{l.location}</td>
                      <td>{l.no ?? ''}</td>
                      <td>{l.l ?? ''}</td>
                      <td>{avg(l.b)?.toFixed(2) ?? ''}</td>
                      <td>{avg(l.d)?.toFixed(2) ?? ''}</td>
                      <td>{l.written ?? ''}</td>
                    </tr>
                  ))
                ])}
              </tbody>
            </table>
          </div>
        )
      )}
      {saved && (
        <div className="notice ok tool-result tool-outcome">
          <span>
            <IconTable /> Saved Excel: {saved}
          </span>
          <button className="ghost" onClick={() => api.openPath(saved)}>
            Open
          </button>
        </div>
      )}
    </div>
  )
}
