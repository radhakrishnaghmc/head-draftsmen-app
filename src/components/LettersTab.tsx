import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import type { AgreementBundleFile } from '../../electron/ipc-contract'
import { intimationLetterDate } from '@core/intimationFill'
import type { ExcelTable, StoredLetter } from '@core/types'
import type { ScheduleAMeta } from '@core/scheduleA'
import type { PlaceholderMatch } from '@core/createDocument'
import { matchesBoardSearch } from '@core/boardSearch'
import { formatRupees } from '@core/worksAmounts'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { base64ToUint8, renderDocPreview } from './docPage'
import BoardSearch from './BoardSearch'
import DocFieldsEditor, { groupLabels, humanizeLabel, type DocFieldGroup } from './DocFieldsEditor'
import { DOC_FIELDS } from './WorkOrderAgreementTab'
import { noteSubmittedSections, type NoteSubmittedData } from '@core/noteSubmitted'
import { noteDocxHtml, noteMarginsTwips, paginateNotes } from '../notePagination'
import { AGREEMENT_DOCS, isAgreementDoc } from '../bulkAgreementDocs'
import { DB_NOTE_FIELDS } from '../dbNoteFields'
import type { DeviationItem, DeviationMeta } from '@core/deviationTemplate'
import { formatDbNoteAmounts } from '@core/workOrderAgreement'
import DeleteAllButton from './DeleteAllButton'
import RefreshFromWorksButton from './RefreshFromWorksButton'
import { refreshAgainstWorks, refreshLetter } from '@core/refreshFromWorks'
import { IconChevronLeft, IconCheck, IconDoc, IconDownload, IconFolder, IconTrash, IconWarn } from './Icons'

type LetterKind = StoredLetter['kind']

const KINDS: { key: LetterKind; name: string; meta: string; docName: string }[] = [
  {
    key: 'intimation',
    name: 'Intimations',
    meta: 'Intimations made in bulk from Settings → Bulk Intimations',
    docName: 'Intimation'
  },
  {
    key: 'forwarding',
    name: 'AE/Dy.EE Forwarding Letters',
    meta: 'Forwarding Letters downloaded in the Agreement workspace or made in bulk from Settings → Bulk Forwarding Letters',
    docName: 'AE-Dy.EE Forwarding Letters'
  },
  {
    key: 'qc',
    name: 'QC Intimation/Completion Letters',
    meta: 'QC letters downloaded in the Agreement workspace or made in bulk from Settings → Bulk QC Intimations',
    docName: 'QC Inti-Comp'
  },
  {
    key: 'filebacker',
    name: 'File Backers',
    meta: 'File Backers downloaded in the Agreement workspace or made in bulk from Settings → Bulk File Backers',
    docName: 'File Backer'
  },
  {
    key: 'agreementbond',
    name: 'Agreement Bonds',
    meta: 'Agreement Bonds downloaded in the Agreement workspace or made in bulk from Settings → Bulk Agreement Bonds',
    docName: 'Agreement Bond'
  },
  {
    key: 'workorder',
    name: 'Work Orders',
    meta: 'Work Orders downloaded in the Agreement workspace or made in bulk from Settings → Bulk Work Orders',
    docName: 'Work Order'
  },
  {
    key: 'tenderdoc',
    name: 'Tender Documents',
    meta: 'Tender Documents downloaded in the Agreement workspace or made in bulk from Settings → Bulk Tender Documents',
    docName: 'Tender Document'
  },
  {
    key: 'schedulea',
    name: 'Schedule A',
    meta: 'Schedule A sheets downloaded in the Agreement workspace',
    docName: 'Schedule A'
  },
  {
    key: 'deviation',
    name: 'Deviation Statements',
    meta: 'Deviation Statements downloaded in the Agreement workspace',
    docName: 'Deviation Statement'
  },
  {
    key: 'notesubmitted',
    name: 'Notes Submitted',
    meta: 'Notes Submitted downloaded in the Agreement workspace',
    docName: 'Note Submitted'
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
          const b64 = kind === 'notesubmitted' || kind === 'schedulea' || kind === 'deviation'
            ? ''
            : kind === 'tenderdoc'
              ? await api.civilTenderTemplate(variant)
              : kind === 'intimation'
            ? await api.intimationTemplate(variant)
            : isAgreementDoc(kind)
            ? await AGREEMENT_DOCS[kind].template(variant)
            : kind === 'qc'
              ? await api.qccIntimationTemplate(variant)
              : kind === 'dbnote'
                ? await api.workBillNoteTemplate()
                : await api.aeDyeeForwardingTemplate(variant)
          const labels = b64 ? await api.findPlaceholdersInDocument(b64) : []
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

/** A Schedule A's estimate rows, meta (hand edits applied) and office kind, as stored when it was issued. */
function scheduleParts(letter: StoredLetter) {
  const { headers, rows } = JSON.parse(letter.values.table) as { headers: string[]; rows: Record<string, string>[] }
  const table: ExcelTable = { id: letter.id, name: letter.fileName, path: '', headers, rows }
  const meta = { ...JSON.parse(letter.values.meta), ...(letter.edits ?? {}) } as ScheduleAMeta
  return { table, meta, se: letter.values.se === '1' }
}

/** A Deviation Statement's estimate items and header details (hand edits applied), as stored when it was issued. */
function deviationParts(letter: StoredLetter) {
  const items = JSON.parse(letter.values.items) as DeviationItem[]
  const meta = { ...JSON.parse(letter.values.meta), ...(letter.edits ?? {}) } as DeviationMeta
  const pct = meta.tenderPercent == null || String(meta.tenderPercent).trim() === '' ? undefined : Number(meta.tenderPercent)
  return { items, meta: { ...meta, estimateAmountLakhs: Number(meta.estimateAmountLakhs) || 0, tenderPercent: Number.isFinite(pct) ? pct : undefined } }
}

/** The letter's docx — its template filled with the stored values (plus the AE name & phone, for QC letters). */
async function fillLetter(letter: StoredLetter, tpl: Template, extra: Record<string, string>): Promise<string> {
  if (letter.kind === 'notesubmitted') {
    // Stored as the note's data + margins; the page is rebuilt from them, hand edits overriding the data's own fields.
    const data = { ...JSON.parse(letter.values.data), ...(letter.edits ?? {}) } as NoteSubmittedData
    const margins = JSON.parse(letter.values.margins)
    const raw = noteSubmittedSections(data).map((sec) => sec.html)
    return api.noteSubmittedDocx(noteDocxHtml(paginateNotes(raw, margins)), noteMarginsTwips(margins))
  }
  const resolved: PlaceholderMatch[] = tpl.labels.map((label) => ({ label, column: label, score: 1 }))
  const values = { ...letter.values, ...(letter.edits ?? {}), ...extra }
  if (letter.kind === 'dbnote') formatDbNoteAmounts(values)
  return api.fillPlaceholdersInDocument(tpl.b64, resolved, values)
}

const DOC_KIND_OF: Partial<Record<StoredLetter['kind'], string>> = {
  forwarding: 'aeDyeeForwarding',
  qc: 'qccIntimation',
  dbnote: 'workBillNote',
  filebacker: 'fileBacker',
  agreementbond: 'agreement',
  workorder: 'workOrder',
  tenderdoc: 'civilTender'
}

/** A Note Submitted's editable fields: every plain text / number field of its data. */
const noteScalars = (letter: StoredLetter): Record<string, string> =>
  Object.fromEntries(
    Object.entries(JSON.parse(letter.values.data) as Record<string, unknown>)
      .filter(([, v]) => typeof v === 'string' || typeof v === 'number')
      .map(([k, v]) => [k, String(v)])
  )

/** What the editor panel lists for a letter — the Agreement workspace's own field list where it has one, else the template's labels. */
function editorGroups(letter: StoredLetter, tpl: Template | undefined): DocFieldGroup[] {
  if (letter.kind === 'dbnote') return DB_NOTE_FIELDS
  if (letter.kind === 'schedulea')
    return [{ title: 'Schedule A details', fields: Object.keys(scheduleParts(letter).meta).map((k) => ({ label: k, title: humanizeLabel(k), multiline: /name/i.test(k) })) }]
  if (letter.kind === 'deviation')
    return [{ title: 'Deviation Statement details', fields: Object.keys(JSON.parse(letter.values.meta)).map((k) => ({ label: k, title: humanizeLabel(k), multiline: /name/i.test(k) })) }]
  if (letter.kind === 'notesubmitted')
    return [{ title: 'The note', fields: Object.keys(noteScalars(letter)).map((k) => ({ label: k, title: humanizeLabel(k), multiline: /name|note|newspapers/i.test(k) })) }]
  const listed = DOC_FIELDS[DOC_KIND_OF[letter.kind] as keyof typeof DOC_FIELDS]
  return listed ?? groupLabels(tpl?.labels ?? [])
}

const editorDerived = (letter: StoredLetter): Record<string, string> =>
  letter.kind === 'notesubmitted'
    ? noteScalars(letter)
    : letter.kind === 'deviation'
      ? Object.fromEntries(Object.entries(JSON.parse(letter.values.meta) as Record<string, unknown>).map(([k, v]) => [k, v == null ? '' : String(v)]))
      : letter.kind === 'schedulea'
      ? (JSON.parse(letter.values.meta) as Record<string, string>)
      : letter.values

/** A letter's page, rendered only once it scrolls into view — a board can hold hundreds. */
/** A Schedule A: the printed page itself in the full preview, a small card on the board. */
function ScheduleAPage({ letter, full }: { letter: StoredLetter; full: boolean }) {
  const [html, setHtml] = useState('')
  const { table, meta, se } = scheduleParts(letter)
  const metaKey = JSON.stringify(meta)
  useEffect(() => {
    if (!full) return
    let cancelled = false
    void api.scheduleAPreviewHtml(table, meta, se).then((h) => !cancelled && setHtml(h))
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [full, letter.values.table, metaKey])
  return full ? (
    <iframe title="Schedule A" srcDoc={html} style={{ width: '100%', height: '100%', minHeight: '70vh', border: 0, background: '#d8d8dc' }} />
  ) : (
    <div className="wo-tile-doc">
      <p className="sub">Schedule A · {table.rows.length} item{table.rows.length === 1 ? '' : 's'}</p>
    </div>
  )
}

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
  if (letter.kind === 'schedulea') return <ScheduleAPage letter={letter} full={full} />
  if (letter.kind === 'deviation') {
    const { items, meta } = deviationParts(letter)
    return (
      <div className={full ? 'ns-preview' : 'wo-tile-doc'}>
        <p className="sub">
          Deviation Statement · {items.length} item{items.length === 1 ? '' : 's'}
        </p>
        {full && <p>{meta.nameOfWork}</p>}
      </div>
    )
  }
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
  loginId,
  worksTable,
  onChange,
  onBack
}: {
  kind: LetterKind
  loginId: string
  worksTable: ExcelTable | null
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
  // Grouped by the LOA date the letter is for, newest first; letters whose date can't be told come last.
  const groups = useMemo(() => {
    const map = new Map<string, StoredLetter[]>()
    for (const l of shown) {
      const date = loaDateOf(l)
      const list = map.get(date)
      if (list) list.push(l)
      else map.set(date, [l])
    }
    for (const list of map.values()) list.sort((a, b) => a.nitNo.localeCompare(b.nitNo) || a.agency.localeCompare(b.agency) || a.workName.localeCompare(b.workName))
    return [...map.entries()].sort(([a], [b]) => (a && !b ? -1 : !a && b ? 1 : b.localeCompare(a)))
  }, [shown])

  const extra: Record<string, string> =
    kind === 'qc' && (aeName.trim() || aePhone.trim()) ? { 'AE Name Phone': [aeName.trim(), aePhone.trim()].filter(Boolean).join(', ') } : {}
  const tplOf = (l: StoredLetter) => templates[`${kind}:${l.variant}`]

  // The letter open in the preview, read live from the list so a hand edit shows straight away.
  const previewLive = preview ? (letters.find((l) => l.id === preview.id) ?? preview) : null
  function setEdits(id: string, edits: Record<string, string>) {
    onChange(letters.map((l) => (l.id === id ? { ...l, edits: Object.keys(edits).length > 0 ? edits : undefined } : l)))
  }

  function removeAll() {
    const now = new Date().toISOString()
    onChange(letters.map((l) => (l.kind === kind && !l.deletedAt ? { ...l, deletedAt: now } : l)))
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
      const files: AgreementBundleFile[] = []
      for (const l of shown) {
        const tpl = tplOf(l)
        if (!tpl) throw new Error('The document format is still loading — try again in a moment.')
        if (l.kind === 'schedulea') {
          const { table, meta, se } = scheduleParts(l)
          files.push({ name: l.fileName, format: format === 'pdf' ? 'pdf' : 'xlsx', scheduleATable: table, scheduleAMeta: meta, scheduleAIsSe: se })
        } else if (l.kind === 'deviation') {
          const { items, meta } = deviationParts(l)
          files.push({ name: l.fileName, format: 'xlsx', deviationItems: items, deviationMeta: meta })
        } else files.push({ name: l.fileName, format, docxBase64: await fillLetter(l, tpl, extra) })
        setBusy(`Preparing ${files.length}/${shown.length}`)
      }
      const nits = [...new Set(shown.map((l) => l.nitNo.replace(/\s*\bdated?\b.*$/i, '').replace(/\s*\/\s*/g, '/').trim()).filter(Boolean))]
      const folder = `${kind === 'tenderdoc' || kind === 'notesubmitted' || kind === 'deviation' ? meta.docName : kind === 'intimation' ? 'Intimations' : isAgreementDoc(kind) ? AGREEMENT_DOCS[kind].docName : kind === 'qc' ? 'QC intimation' : kind === 'dbnote' ? 'DB notes' : 'AE-DyEE forwarding letters'}${nits.length ? ` of ${nits.join(', ')}` : ''}`
      const res = await api.exportAgreementBundle(files, undefined, folder)
      if (!res) return
      setMessage(
        res.failed.length
          ? { ok: false, text: `Saved ${res.written.length}; ${res.failed.length} failed: ${res.failed.join(', ')}` }
          : { ok: true, text: `Saved ${shown.length} document${shown.length === 1 ? '' : 's'}.` }
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
      if (l.kind === 'schedulea') {
        const { table, meta, se } = scheduleParts(l)
        const saved =
          fmt === 'pdf'
            ? await api.exportScheduleAPdf(table, l.fileName, meta, se)
            : se
              ? await api.exportSeScheduleA(table, l.fileName, meta)
              : await api.exportScheduleA(table, l.fileName, meta)
        if (saved) setMessage({ ok: true, text: `Saved: ${saved}` })
        return
      }
      if (l.kind === 'deviation') {
        const { items, meta } = deviationParts(l)
        const saved = await api.exportDeviation(items, meta, l.fileName)
        if (saved) setMessage({ ok: true, text: `Saved: ${saved}` })
        return
      }
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
          <DeleteAllButton count={mine.length} label={meta.name} loginId={loginId} onConfirmed={removeAll} />
          <RefreshFromWorksButton
            hasWorks={!!worksTable}
            onRefresh={() => {
              const r = refreshAgainstWorks(letters, worksTable, refreshLetter, (l) => l.kind === kind)
              onChange(r.items)
              return { found: r.found, total: mine.length }
            }}
          />
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
                  {kind === 'schedulea' || kind === 'deviation' ? 'Excel' : 'Word'}
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
                : `Nothing here yet. ${kind === 'tenderdoc' || kind === 'notesubmitted' || kind === 'deviation' ? `Download one in the Agreement workspace` : `Upload a folder in Settings → ${kind === 'intimation' ? 'Bulk Intimations' : isAgreementDoc(kind) ? AGREEMENT_DOCS[kind].title : kind === 'qc' ? 'Bulk QC Intimations' : 'Bulk Forwarding Letters'}`} and they are collected here.`}
            </p>
          </div>
        )}
        {mine.length > 0 && shown.length === 0 && (
          <div className="empty">
            <IconDoc />
            <p>No documents match “{query.trim()}”.</p>
          </div>
        )}
        {groups.map(([date, list]) => (
          <div key={date || 'none'} className="notice-nit-group">
            <div className="ntc-group-head">
              <span className="ntc-group-title">LOA Date {date ? date.split('-').reverse().join('.') : '— not known'}</span>
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
                      {l.kind === 'deviation' ? (
                        <>
                          <span>Tender ID:</span> {l.values?.tenderId || '—'}
                        </>
                      ) : (
                        <>
                          <span>NIT date:</span> {l.nitDate || '—'}
                        </>
                      )}
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
                    {previewLive && (
                      <DocFieldsEditor
                        groups={editorGroups(previewLive, tplOf(previewLive))}
                        derived={editorDerived(previewLive)}
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
                  <IconDownload /> {kind === 'schedulea' || kind === 'deviation' ? 'Excel' : 'Word'}
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
 * Issued Documents: one tile per document type, each opening a board of the documents the Agreement workspace issued or Settings' bulk tools made, with search and download.
 */
/** The LOA date a letter is for: stored when it was made, else read back from its values (older letters). */
const loaDateOf = (l: StoredLetter) => l.loaDate || intimationLetterDate({ ...l.values, ...(l.edits ?? {}) })

export default function LettersTab({
  letters,
  onChange,
  onWorkspaceChange,
  loginId,
  worksTable
}: {
  /** This office's letters (both kinds, including binned ones). */
  letters: StoredLetter[]
  onChange: (next: StoredLetter[]) => void
  onWorkspaceChange?: (open: boolean) => void
  loginId: string
  worksTable: ExcelTable | null
}) {
  const [kind, setKind] = useState<LetterKind | null>(null)
  useEffect(() => {
    onWorkspaceChange?.(kind !== null)
  }, [kind, onWorkspaceChange])
  if (kind)
    return <LetterBoard kind={kind} letters={letters} loginId={loginId} worksTable={worksTable} onChange={onChange} onBack={() => setKind(null)} />
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
                {k.meta} · {count} document{count === 1 ? '' : 's'}
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
