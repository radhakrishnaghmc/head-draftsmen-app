import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import type { StoredLetter } from '@core/types'
import type { PlaceholderMatch } from '@core/createDocument'
import { matchesBoardSearch } from '@core/boardSearch'
import { formatRupees } from '@core/worksAmounts'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { base64ToUint8, renderDocPreview } from './docPage'
import BoardSearch from './BoardSearch'
import DocFieldsEditor from './DocFieldsEditor'
import { DB_NOTE_FIELDS } from '../dbNoteFields'
import { formatDbNoteAmounts } from '@core/workOrderAgreement'
import { IconChevronLeft, IconCheck, IconDoc, IconDownload, IconFolder, IconTrash, IconWarn } from './Icons'

type LetterKind = StoredLetter['kind']

const KINDS: { key: LetterKind; name: string; meta: string; docName: string }[] = [
  {
    key: 'forwarding',
    name: 'AE/Dy.EE Forwarding Letters',
    meta: 'Forwarding letters made in bulk from Settings → Bulk Forwarding Letters',
    docName: 'AE-Dy.EE Forwarding Letters'
  },
  {
    key: 'qc',
    name: 'QC Intimation/Completion Letters',
    meta: 'QC letters made in bulk from Settings → Bulk QC Intimations',
    docName: 'QC Inti-Comp'
  },
  {
    key: 'dbnote',
    name: 'Drawing Branch Note',
    meta: 'DB notes issued in the Agreement workspace or made in bulk from Settings → Bulk Issue Drawing Branch Notes',
    docName: 'DB Note'
  }
]

interface Template {
  b64: string
  labels: string[]
}

/** Loads (and caches) the bundled template of each variant the letters were made with. */
function useTemplates(kind: LetterKind, variants: string[]) {
  const [templates, setTemplates] = useState<Record<string, Template>>({})
  const [error, setError] = useState<string | null>(null)
  const wanted = [...new Set(variants)].sort().join('|')
  useEffect(() => {
    let cancelled = false
    void (async () => {
      for (const variant of wanted ? wanted.split('|') : []) {
        const key = `${kind}:${variant}`
        if (templates[key]) continue
        try {
          const b64 =
            kind === 'qc'
              ? await api.qccIntimationTemplate(variant)
              : kind === 'dbnote'
                ? await api.workBillNoteTemplate()
                : await api.aeDyeeForwardingTemplate(variant)
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
  }, [kind, wanted])
  return { templates, error }
}

/** The letter's docx — its template filled with the stored values (plus the AE name & phone, for QC letters). */
async function fillLetter(letter: StoredLetter, tpl: Template, extra: Record<string, string>): Promise<string> {
  const resolved: PlaceholderMatch[] = tpl.labels.map((label) => ({ label, column: label, score: 1 }))
  const values = { ...letter.values, ...(letter.edits ?? {}), ...extra }
  if (letter.kind === 'dbnote') formatDbNoteAmounts(values)
  return api.fillPlaceholdersInDocument(tpl.b64, resolved, values)
}

/** A letter's page, rendered only once it scrolls into view — a board can hold hundreds. */
function LetterPage({ letter, tpl, extra, full = false }: { letter: StoredLetter; tpl: Template | undefined; extra: Record<string, string>; full?: boolean }) {
  const ref = useRef<HTMLDivElement>(null)
  const [visible, setVisible] = useState(full)
  const [failed, setFailed] = useState(false)
  const extraKey = JSON.stringify(extra)
  useEffect(() => {
    if (full) return
    const el = ref.current
    if (!el || typeof IntersectionObserver === 'undefined') return setVisible(true)
    const io = new IntersectionObserver((entries) => entries.some((e) => e.isIntersecting) && setVisible(true), { rootMargin: '300px' })
    io.observe(el)
    return () => io.disconnect()
  }, [full])
  useEffect(() => {
    const el = ref.current
    if (!visible || !tpl || !el) return
    let cancelled = false
    void fillLetter(letter, tpl, extra)
      .then((docx) => renderDocPreview(base64ToUint8(docx), el, { flowPages: letter.kind === 'dbnote' }))
      .then(() => !cancelled && setFailed(false))
      .catch(() => !cancelled && setFailed(true))
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visible, tpl, letter.values, letter.edits, extraKey])
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

function LetterBoard({
  kind,
  letters,
  onChange,
  onBack
}: {
  kind: LetterKind
  letters: StoredLetter[]
  onChange: (next: StoredLetter[]) => void
  onBack: () => void
}) {
  const meta = KINDS.find((k) => k.key === kind)!
  const mine = useMemo(() => letters.filter((l) => l.kind === kind && !l.deletedAt), [letters, kind])
  const { templates, error: templateError } = useTemplates(kind, mine.map((l) => l.variant))
  const [query, setQuery] = useState('')
  const [aeName, setAeName] = useState('')
  const [aePhone, setAePhone] = useState('')
  const [format, setFormat] = useState<'docx' | 'pdf'>('docx')
  const [preview, setPreview] = useState<StoredLetter | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [message, setMessage] = useState<{ ok: boolean; text: string } | null>(null)

  const shown = useMemo(() => mine.filter((l) => matchesBoardSearch(query, l.workName, l.agency)), [mine, query])
  // Grouped by NIT, like the bulk run that made them; groups newest-made first.
  const groups = useMemo(() => {
    const map = new Map<string, StoredLetter[]>()
    for (const l of shown) {
      const list = map.get(l.nitNo)
      if (list) list.push(l)
      else map.set(l.nitNo, [l])
    }
    for (const list of map.values()) list.sort((a, b) => a.agency.localeCompare(b.agency) || a.workName.localeCompare(b.workName))
    return [...map.entries()].sort(([, a], [, b]) => (b[0].createdDate || '').localeCompare(a[0].createdDate || ''))
  }, [shown])

  const extra: Record<string, string> =
    kind === 'qc' && (aeName.trim() || aePhone.trim()) ? { 'AE Name Phone': [aeName.trim(), aePhone.trim()].filter(Boolean).join(', ') } : {}
  const tplOf = (l: StoredLetter) => templates[`${kind}:${l.variant}`]

  // The letter open in the preview, read live from the list so a hand edit shows straight away.
  const previewLive = preview ? (letters.find((l) => l.id === preview.id) ?? preview) : null
  function setEdits(id: string, edits: Record<string, string>) {
    onChange(letters.map((l) => (l.id === id ? { ...l, edits: Object.keys(edits).length > 0 ? edits : undefined } : l)))
  }

  function remove(id: string) {
    onChange(letters.map((l) => (l.id === id ? { ...l, deletedAt: new Date().toISOString() } : l)))
  }

  async function downloadAll() {
    setMessage(null)
    if (shown.length === 0) return
    if (kind === 'qc' && (!aeName.trim() || !aePhone.trim())) {
      setMessage({ ok: false, text: 'Enter the AE name and phone number first — they fill the 4th-party letter.' })
      return
    }
    setBusy(`Preparing 0/${shown.length}`)
    const unsubscribe = api.onAgreementBundleProgress(({ done, total }) => setBusy(`Saving ${done}/${total}`))
    try {
      const files = []
      for (const l of shown) {
        const tpl = tplOf(l)
        if (!tpl) throw new Error('A letter format is still loading — try again in a moment.')
        files.push({ name: l.fileName, format, docxBase64: await fillLetter(l, tpl, extra) })
        setBusy(`Preparing ${files.length}/${shown.length}`)
      }
      const nits = [...new Set(shown.map((l) => l.nitNo.replace(/\s*\bdated?\b.*$/i, '').replace(/\s*\/\s*/g, '/').trim()).filter(Boolean))]
      const folder = `${kind === 'qc' ? 'QC intimation' : kind === 'dbnote' ? 'DB notes' : 'AE-DyEE forwarding letters'}${nits.length ? ` of ${nits.join(', ')}` : ''}`
      const res = await api.exportAgreementBundle(files, undefined, folder)
      if (!res) return
      setMessage(
        res.failed.length
          ? { ok: false, text: `Saved ${res.written.length}; ${res.failed.length} failed: ${res.failed.join(', ')}` }
          : { ok: true, text: `Saved ${shown.length} letter${shown.length === 1 ? '' : 's'}.` }
      )
    } catch (e) {
      setMessage({ ok: false, text: e instanceof Error ? e.message : String(e) })
    } finally {
      unsubscribe()
      setBusy(null)
    }
  }

  async function downloadOne(l: StoredLetter, fmt: 'docx' | 'pdf') {
    setMessage(null)
    const tpl = tplOf(l)
    if (!tpl) return
    if (kind === 'qc' && (!aeName.trim() || !aePhone.trim())) {
      setMessage({ ok: false, text: 'Enter the AE name and phone number first — they fill the 4th-party letter.' })
      return
    }
    setBusy(fmt === 'pdf' ? 'Preparing PDF…' : 'Saving…')
    try {
      const res = await api.exportCreatedDocument(await fillLetter(l, tpl, extra), l.fileName, [fmt])
      if (res && res.length) setMessage({ ok: true, text: `Saved: ${res.map((r) => r.file).join(', ')}` })
    } catch (e) {
      setMessage({ ok: false, text: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconDoc /> {meta.name}
        </div>
        <button className="ghost pdf-ws-back" onClick={onBack}>
          <IconChevronLeft /> Back
        </button>
      </div>
      <div className="card pdf-workspace">
        <div className="letters-toolbar">
          <BoardSearch value={query} onChange={setQuery} count={shown.length} total={mine.length} />
          {kind === 'qc' && (
            <>
              <input type="text" className="ae-input" placeholder="AE name *" value={aeName} onChange={(e) => setAeName(e.target.value)} aria-label="AE name" />
              <input type="text" className="ae-input" placeholder="AE phone number *" value={aePhone} onChange={(e) => setAePhone(e.target.value)} aria-label="AE phone number" />
            </>
          )}
          {mine.length > 0 && (
            <>
              <div className="format-toggle" role="group" aria-label="Download format">
                <button type="button" className={format === 'docx' ? 'on' : ''} onClick={() => setFormat('docx')} disabled={!!busy}>
                  Word
                </button>
                <button type="button" className={format === 'pdf' ? 'on' : ''} onClick={() => setFormat('pdf')} disabled={!!busy}>
                  PDF
                </button>
              </div>
              <button className="primary" onClick={() => void downloadAll()} disabled={!!busy || shown.length === 0}>
                <IconDownload /> {busy ?? `Download all (${shown.length})`}
              </button>
            </>
          )}
        </div>
        {templateError && (
          <div className="notice error">
            <IconWarn /> {templateError}
          </div>
        )}
        {message && (
          <div className={`notice${message.ok ? '' : ' warn'}`}>
            {message.ok ? <IconCheck /> : <IconWarn />} {message.text}
          </div>
        )}
        {mine.length === 0 && (
          <div className="empty">
            <IconDoc />
            <p>
              {kind === 'dbnote'
                ? 'No DB notes yet. Download a DB Note in the Agreement workspace, or upload a folder in Settings → Bulk Issue Drawing Branch Notes, and they are collected here.'
                : `No letters yet. Upload a folder in Settings → ${kind === 'qc' ? 'Bulk QC Intimations' : 'Bulk Forwarding Letters'} and they are collected here.`}
            </p>
          </div>
        )}
        {mine.length > 0 && shown.length === 0 && (
          <div className="empty">
            <IconDoc />
            <p>No letters match “{query.trim()}”.</p>
          </div>
        )}
        {groups.map(([nit, list]) => (
          <div key={nit || 'none'} className="notice-nit-group">
            <div className="ntc-group-head">
              <span className="ntc-group-title">NIT {nit || '— number not known'}</span>
              <span className="ntc-group-count">{list.length}</span>
            </div>
            <div className="wo-tiles">
              {list.map((l) => (
                <div
                  key={l.id}
                  className="wo-tile"
                  role="button"
                  tabIndex={0}
                  onClick={() => setPreview(l)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault()
                      setPreview(l)
                    }
                  }}
                  title={meta.docName}
                >
                  <div className="wo-tile-preview">
                    <LetterPage letter={l} tpl={tplOf(l)} extra={extra} />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot notice-tile-foot">
                    <div>
                      <span>Name of the work:</span> {l.workName || '—'}
                    </div>
                    <div>
                      <span>NIT date:</span> {l.nitDate || '—'}
                    </div>
                    <div>
                      <span>Name of the agency:</span> {l.agency || '—'}
                    </div>
                    <div>
                      <span>Type of work:</span> {l.typeOfWork || '—'}
                    </div>
                    <div>
                      <span>Estimate amount:</span> {l.estimateRupees != null ? formatRupees(l.estimateRupees) : '—'}
                    </div>
                  </div>
                  <div className="wo-tile-actions" onClick={(e) => e.stopPropagation()}>
                    <button className="ghost small" onClick={() => remove(l.id)} title="Remove this letter">
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
                <span className="wo-modal-title">
                  {meta.name} — {preview.agency || preview.nitNo}
                </span>
                <button className="wo-modal-close" onClick={() => setPreview(null)} title="Close" aria-label="Close">
                  ×
                </button>
              </div>
              <div className="wo-modal-main">
                <div className="wo-modal-body">
                  <div className="ns-body">
                    {previewLive?.kind === 'dbnote' && (
                      <DocFieldsEditor
                        groups={DB_NOTE_FIELDS}
                        derived={previewLive.values}
                        edits={previewLive.edits ?? {}}
                        onChange={(label, value) => setEdits(previewLive.id, { ...(previewLive.edits ?? {}), [label]: value })}
                        onRevert={(label) => {
                          const { [label]: _gone, ...rest } = previewLive.edits ?? {}
                          setEdits(previewLive.id, rest)
                        }}
                        onResetAll={() => setEdits(previewLive.id, {})}
                      />
                    )}
                    <LetterPage letter={previewLive ?? preview} tpl={tplOf(preview)} extra={extra} full />
                  </div>
                </div>
              </div>
              <div className="wo-modal-foot">
                {message && <span className="estimate-hint">{message.text}</span>}
                <button className="ghost wo-foot-close" onClick={() => setPreview(null)}>
                  Close
                </button>
                <button className="primary" onClick={() => void downloadOne(previewLive ?? preview, 'docx')} disabled={!!busy}>
                  <IconDownload /> Word
                </button>
                <button className="primary" onClick={() => void downloadOne(previewLive ?? preview, 'pdf')} disabled={!!busy}>
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

/**
 * Forwarding Slips, QC Intimations & DB Note: three tiles — the AE/Dy.EE Forwarding Letters, the QC Intimation/Completion
 * letters and the Drawing Branch Notes — each opening a board of the letters Settings' bulk tools made, with search and download.
 */
export default function LettersTab({
  letters,
  onChange,
  onWorkspaceChange
}: {
  /** This office's letters (both kinds, including binned ones). */
  letters: StoredLetter[]
  onChange: (next: StoredLetter[]) => void
  onWorkspaceChange?: (open: boolean) => void
}) {
  const [kind, setKind] = useState<LetterKind | null>(null)
  useEffect(() => {
    onWorkspaceChange?.(kind !== null)
  }, [kind, onWorkspaceChange])
  if (kind)
    return <LetterBoard kind={kind} letters={letters} onChange={onChange} onBack={() => setKind(null)} />
  return (
    <div className="card">
      <div className="doc-tile-grid tools-grid notice-kind-grid">
        {KINDS.map((k) => {
          const count = letters.filter((l) => l.kind === k.key && !l.deletedAt).length
          return (
            <button key={k.key} className="doc-tile-card tone-green tool-card" onClick={() => setKind(k.key)}>
              <span className="tool-card-ic">
                <IconDoc />
              </span>
              <span className="doc-tile-card-name">{k.name}</span>
              <span className="doc-tile-card-meta">
                {k.meta} · {count} letter{count === 1 ? '' : 's'}
              </span>
              <span className="tool-card-cta">
                <IconFolder /> Open workspace
              </span>
            </button>
          )
        })}
      </div>
    </div>
  )
}
