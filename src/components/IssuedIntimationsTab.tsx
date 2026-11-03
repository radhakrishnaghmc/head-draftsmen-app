import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import type { IssuedIntimation } from '@core/types'
import type { PlaceholderMatch } from '@core/createDocument'
import { matchesBoardSearch } from '@core/boardSearch'
import { intimationLetterDate } from '@core/intimationFill'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { base64ToUint8, renderDocPreview } from './docPage'
import BoardSearch from './BoardSearch'
import { IconCheck, IconDoc, IconDownload, IconTrash, IconWarn } from './Icons'

interface Template {
  b64: string
  labels: string[]
}

const formatKey = (f: IssuedIntimation['format']) =>
  f.kind === 'loa' ? `loa:${f.reserved ? 'reserved' : 'open'}:${f.corporation}` : `intimation:${f.variant}`

/** Loads (and caches) the bundled format each issued letter was made with. */
function useTemplates(items: IssuedIntimation[]) {
  const [templates, setTemplates] = useState<Record<string, Template>>({})
  const [error, setError] = useState<string | null>(null)
  const wanted = [...new Map(items.map((i) => [formatKey(i.format), i.format])).values()]
  const wantedKey = wanted.map(formatKey).sort().join('|')
  useEffect(() => {
    let cancelled = false
    void (async () => {
      for (const f of wanted) {
        const key = formatKey(f)
        if (templates[key]) continue
        try {
          const b64 =
            f.kind === 'loa'
              ? await api.loaSeTemplate(f.reserved, f.corporation)
              : await api.intimationTemplate(f.variant)
          const labels = await api.findPlaceholdersInDocument(b64)
          if (!cancelled) setTemplates((t) => ({ ...t, [key]: { b64, labels } }))
        } catch (e) {
          if (!cancelled) setError(e instanceof Error ? e.message : String(e))
        }
      }
    })()
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wantedKey])
  return { templates, error }
}

async function fillIssued(item: IssuedIntimation, tpl: Template): Promise<string> {
  const resolved: PlaceholderMatch[] = tpl.labels.map((label) => ({ label, column: label, score: 1 }))
  return api.fillPlaceholdersInDocument(tpl.b64, resolved, item.values)
}

/** One issued letter's page, rendered only once it scrolls into view. */
function IssuedPage({
  item,
  tpl,
  full = false
}: {
  item: IssuedIntimation
  tpl: Template | undefined
  full?: boolean
}) {
  const ref = useRef<HTMLDivElement>(null)
  const [visible, setVisible] = useState(full)
  const [failed, setFailed] = useState(false)
  useEffect(() => {
    if (full) return
    const el = ref.current
    if (!el || typeof IntersectionObserver === 'undefined') return setVisible(true)
    const io = new IntersectionObserver((entries) => entries.some((e) => e.isIntersecting) && setVisible(true), {
      rootMargin: '300px'
    })
    io.observe(el)
    return () => io.disconnect()
  }, [full])
  useEffect(() => {
    const el = ref.current
    if (!visible || !tpl || !el) return
    let cancelled = false
    void fillIssued(item, tpl)
      .then((docx) => renderDocPreview(base64ToUint8(docx), el))
      .then(() => !cancelled && setFailed(false))
      .catch(() => !cancelled && setFailed(true))
    return () => {
      cancelled = true
    }
  }, [visible, tpl, item])
  if (failed)
    return (
      <div className="settings-tile-error">
        <IconWarn /> Couldn't load
      </div>
    )
  return full ? (
    <div className="ns-preview">
      <div className="ns-preview-scroll">
        <div ref={ref} className="intimation-docx-preview" />
      </div>
    </div>
  ) : (
    <div ref={ref} className="wo-tile-doc" />
  )
}

/** The letter's own date; older records saved the generation day, so it is re-read from the stored values. */
const letterDate = (i: IssuedIntimation) => intimationLetterDate(i.values) || i.issuedDate

const dmy = (iso: string) => (/^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso.split('-').reverse().join('.') : iso)

/**
 * Issued Intimations: every Intimation / Letter of Acceptance downloaded, printed or bundled from the Issue Intimation
 * tab is collected here (per office) — searchable, previewable, and downloadable again.
 */
export default function IssuedIntimationsTab({
  items,
  onChange,
  searchSlot
}: {
  /** This office's issued intimations (including binned ones). */
  items: IssuedIntimation[]
  onChange: (next: IssuedIntimation[]) => void
  /** Element in the page's sub-tab row to show the search box in (it sits in the card until that exists). */
  searchSlot?: HTMLElement | null
}) {
  const mine = useMemo(
    () => items.filter((i) => !i.deletedAt).sort((a, b) => letterDate(b).localeCompare(letterDate(a))),
    [items]
  )
  const { templates, error } = useTemplates(mine)
  const [query, setQuery] = useState('')
  const [preview, setPreview] = useState<IssuedIntimation | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [message, setMessage] = useState<{ ok: boolean; text: string } | null>(null)
  const shown = useMemo(() => mine.filter((i) => matchesBoardSearch(query, i.workName, i.agency)), [mine, query])
  // Grouped by the intimation's own date, newest first (`mine` is already sorted that way).
  const groups = useMemo(() => {
    const byDate = new Map<string, IssuedIntimation[]>()
    for (const i of shown) byDate.set(letterDate(i), [...(byDate.get(letterDate(i)) ?? []), i])
    return [...byDate.entries()]
  }, [shown])
  const searchBox = <BoardSearch value={query} onChange={setQuery} count={shown.length} total={mine.length} />
  const tplOf = (i: IssuedIntimation) => templates[formatKey(i.format)]

  function remove(id: string) {
    onChange(items.map((i) => (i.id === id ? { ...i, deletedAt: new Date().toISOString() } : i)))
  }

  async function download(item: IssuedIntimation, fmt: 'docx' | 'pdf') {
    setMessage(null)
    const tpl = tplOf(item)
    if (!tpl) return
    setBusy(fmt === 'pdf' ? 'Preparing PDF…' : 'Saving…')
    try {
      const res = await api.exportCreatedDocument(await fillIssued(item, tpl), item.fileName, [fmt])
      if (res && res.length) setMessage({ ok: true, text: `Saved: ${res.map((r) => r.file).join(', ')}` })
    } catch (e) {
      setMessage({ ok: false, text: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <div className="card pdf-workspace">
        {searchSlot ? (
          createPortal(searchBox, searchSlot)
        ) : (
          <div className="letters-toolbar">{searchBox}</div>
        )}
        {error && (
          <div className="notice error">
            <IconWarn /> {error}
          </div>
        )}
        {message && !preview && (
          <div className={`notice${message.ok ? '' : ' warn'}`}>
            {message.ok ? <IconCheck /> : <IconWarn />} {message.text}
          </div>
        )}
        {mine.length === 0 && (
          <div className="empty">
            <IconDoc />
            <p>
              No intimations issued yet. Each Intimation you download or print in the Issue Intimation tab is collected
              here.
            </p>
          </div>
        )}
        {mine.length > 0 && shown.length === 0 && (
          <div className="empty">
            <IconDoc />
            <p>No intimations match “{query.trim()}”.</p>
          </div>
        )}
        {groups.map(([date, list]) => (
          <div key={date} className="notice-nit-group">
            <div className="ntc-group-head">
              <span className="ntc-group-title">LOA Date {dmy(date)}</span>
              <span className="ntc-group-count">{list.length}</span>
            </div>
            <div className="wo-tiles">
              {list.map((i) => (
                <div
                  key={i.id}
                  className="wo-tile"
                  role="button"
                  tabIndex={0}
                  onClick={() => setPreview(i)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault()
                      setPreview(i)
                    }
                  }}
                  title={i.fileName}
                >
                  <div className="wo-tile-preview">
                    <IssuedPage item={i} tpl={tplOf(i)} />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot notice-tile-foot">
                    <div>
                      <span>Name of the work:</span> {i.workName || '—'}
                    </div>
                    <div>
                      <span>Name of the agency:</span> {i.agency || '—'}
                    </div>
                    <div>
                      <span>NIT No:</span> {i.nitNo || '—'}
                    </div>
                    <div>
                      <span>Issued:</span> {dmy(letterDate(i))}
                    </div>
                  </div>
                  <div className="wo-tile-actions" onClick={(e) => e.stopPropagation()}>
                    <button className="ghost small" onClick={() => remove(i.id)} title="Remove from this list">
                      <IconTrash />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      {preview &&
        createPortal(
          <div className="wo-modal-overlay" onMouseDown={closeOnBackdropMouseDown(() => setPreview(null))}>
            <div className="wo-modal wide" onClick={(e) => e.stopPropagation()}>
              <div className="wo-modal-head">
                <span className="wo-modal-title">{preview.fileName}</span>
                <button className="wo-modal-close" onClick={() => setPreview(null)} title="Close" aria-label="Close">
                  ×
                </button>
              </div>
              <div className="wo-modal-main">
                <div className="wo-modal-body">
                  <div className="ns-body">
                    <IssuedPage item={preview} tpl={tplOf(preview)} full />
                  </div>
                </div>
              </div>
              <div className="wo-modal-foot">
                {message && <span className="estimate-hint">{message.text}</span>}
                <button className="ghost wo-foot-close" onClick={() => setPreview(null)}>
                  Close
                </button>
                <button className="primary" onClick={() => void download(preview, 'docx')} disabled={!!busy}>
                  <IconDownload /> Word
                </button>
                <button className="primary" onClick={() => void download(preview, 'pdf')} disabled={!!busy}>
                  <IconDownload /> PDF
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
    </>
  )
}
