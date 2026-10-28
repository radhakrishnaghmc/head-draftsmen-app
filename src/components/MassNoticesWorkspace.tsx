import { useEffect, useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import {
  parseIntimationNotice,
  parseIntimationNoticeText,
  parseSupplierDetailsText,
  type IntimationNotice
} from '@core/intimationNotice'
import { detectTenderId } from '@core/tenderAgents/tenderId'
import {
  pairL1WithLoa,
  findMonitoringWork,
  decideFromMonitoringFormat,
  type ParsedDoc,
  type SupplierDetails
} from '@core/massNotices'
import { defaultLoaNo, isNoticeOverdue, EMPTY_ISSUE_NOTICE_MANUAL_FIELDS } from '@core/issueNotice'
import { nthWorkingDay, dmyToIso, type CalendarData } from '@core/calendar'
import { agencyKey } from '@core/agencyName'
import { normalizeWorkNameForMatch } from '@core/tenderAgents/nameOfWork'
import type { MonitoringFormatWorkRow } from '@core/monitoringFormat'
import type { PendingNoticeItem } from '@core/types'
import { CMC_ZONE_CIRCLES, corporationByName } from '../zoneCircleDirectory'
import { isZoneOnlyOffice, type Office } from '../office'
import { detectOffice } from './IssueNoticesTab'
import { IconFolder, IconWarn, IconCheck, IconBell, IconChevronLeft, IconRestore } from './Icons'

interface Props {
  office: Office
  /** This office's Monitoring Format "list of works" rows (the Agreement column decides whether a notice is due). */
  monitoringWorks?: MonitoringFormatWorkRow[]
  calendar: CalendarData | null
  /** This office's Notice tiles — generated ones are appended; they show up under Issue Notices → "Agreement not Concluded — Reminders". */
  items: PendingNoticeItem[]
  onChange: (next: PendingNoticeItem[]) => void
  /** Leaves the workspace — the parent unmounts it, which resets it to the Upload folder state. */
  onBack: () => void
}

/** A tender that got no notice, and why. */
interface Skipped {
  key: string
  nitNo: string
  agency: string
  workName: string
  note: string
}

type Phase = 'idle' | 'opening' | 'reading' | 'generating' | 'done'

function isoToday(): string {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

const pause = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))

/**
 * Settings → Generate Mass Notices workspace. Pick a folder of tender
 * documents: every L-1 selection sheet is paired with the LOA (Online
 * Intimation) of the SAME work (agency + NIT — core/sameWorkCheck.ts, as in the
 * Intimation and Work Order tabs), or — when there's no LOA — with the bidder's
 * Supplier Details page (same Tender ID + agency). Each pair is then looked up
 * in the office's Monitoring Format: a date in its Agreement column means the
 * agreement is done (skip); blank or words means a notice is due.
 *
 * Notices are generated one at a time — each appears as a tile and is added to
 * Issue Notices → "Agreement not Concluded — Reminders" straight away (where
 * they group by NIT No and count as Active / Overdue). Leaving the workspace
 * unmounts it, so coming back starts again at Upload folder; anything already
 * generated stays in Reminders.
 */
export default function MassNoticesWorkspace({ office, monitoringWorks, calendar, items, onChange, onBack }: Props) {
  const [phase, setPhase] = useState<Phase>('idle')
  const [readProgress, setReadProgress] = useState({ done: 0, total: 0 })
  // Main-process folder scan: 'scanning' (files found so far) then 'reading' (bytes loaded) — before any parsing starts.
  const [scan, setScan] = useState<{ phase: 'scanning' | 'reading'; done: number; total: number } | null>(null)
  const openingRef = useRef(false)
  const [genProgress, setGenProgress] = useState({ done: 0, total: 0 })
  const [error, setError] = useState<string | null>(null)
  const [tiles, setTiles] = useState<PendingNoticeItem[]>([])
  // Why each tile was generated, by notice id — e.g. the work isn't in the Monitoring Format.
  const [tileNotes, setTileNotes] = useState<Record<string, string>>({})
  const [skipped, setSkipped] = useState<Skipped[]>([])
  const [ignored, setIgnored] = useState(0)

  // The parent re-creates onChange every render (it closes over the whole list), so the
  // long-running loop below always calls the latest one.
  const onChangeRef = useRef(onChange)
  onChangeRef.current = onChange
  const cancelled = useRef(false)
  const runId = useRef(0)
  useEffect(
    () =>
      api.onTenderScanProgress((p) => {
        if (openingRef.current && p.phase !== 'parsing') setScan({ phase: p.phase, done: p.done, total: p.total })
      }),
    []
  )
  useEffect(() => {
    cancelled.current = false
    return () => {
      cancelled.current = true
    }
  }, [])

  const ee = !isZoneOnlyOffice(office)
  const corp = corporationByName(office.corporation)
  const entries = corp?.entries?.length ? corp.entries : CMC_ZONE_CIRCLES
  const corporationName = corp?.entries?.length ? office.corporation! : 'CMC'
  const busy = phase === 'opening' || phase === 'reading' || phase === 'generating'

  // Back to the Upload folder state. Stops a run in progress; notices already added to Reminders stay there.
  function clear() {
    runId.current++
    openingRef.current = false
    setPhase('idle')
    setError(null)
    setTiles([])
    setTileNotes({})
    setSkipped([])
    setIgnored(0)
    setScan(null)
    setReadProgress({ done: 0, total: 0 })
    setGenProgress({ done: 0, total: 0 })
  }

  async function run() {
    const myRun = ++runId.current
    const stale = () => cancelled.current || runId.current !== myRun
    setPhase('opening')
    setScan(null)
    openingRef.current = true
    setError(null)
    setTiles([])
    setTileNotes({})
    setSkipped([])
    setIgnored(0)
    try {
      const picked = await api.pickTenderDocuments().finally(() => {
        openingRef.current = false
      })
      setScan(null)
      if (stale()) return
      if (picked.length === 0) {
        setPhase('idle')
        return
      }

      // 1) Classify every file as an L-1 sheet, an LOA (Online Intimation) or a Supplier Details page.
      setPhase('reading')
      const l1s: ParsedDoc<TenderEvaluation>[] = []
      const loas: ParsedDoc<IntimationNotice>[] = []
      const suppliers: ParsedDoc<SupplierDetails>[] = []
      let unread = 0
      for (let i = 0; i < picked.length; i++) {
        if (stale()) return
        const f = new File([new Uint8Array(picked[i].bytes)], picked[i].name)
        setReadProgress({ done: i + 1, total: picked.length })
        try {
          if (/\.pdf$/i.test(f.name)) {
            const lines = await pdfToTextLinesOrOcr(f)
            const ev = parseTenderEvaluation(lines)
            if ((ev.nameOfWork || ev.tenderId) && ev.l1AgencyName) {
              l1s.push({ file: f.name, doc: ev })
              continue
            }
            // The eGP "Supplier Details" page: portal Tender ID + company + address/phone.
            const sup = parseSupplierDetailsText(lines)
            const supTenderId = detectTenderId(lines)
            if (supTenderId && sup.agencyName && (sup.address || sup.phone)) {
              suppliers.push({ file: f.name, doc: { ...sup, tenderId: supTenderId } })
              continue
            }
            const n = parseIntimationNoticeText(lines)
            if (n.agencyName && n.nitNo) loas.push({ file: f.name, doc: n })
            else unread++
          } else {
            const n = parseIntimationNotice(await f.text())
            if (n.agencyName && n.nitNo) loas.push({ file: f.name, doc: n })
            else unread++
          }
        } catch {
          unread++
        }
      }
      setIgnored(unread)
      if (l1s.length === 0) {
        setError("Couldn't find any L-1 selection sheet (Commercial Evaluation / Stage Selected PDF) in that folder.")
        setPhase('idle')
        return
      }

      // 2) Same-work pairing; an L-1 with no LOA is accepted on its Supplier Details page, otherwise gets no notice.
      // Strictly office-specific: only this office's tenders are considered. A tender's own circle (EE office)
      // / zone (SE office), resolved from its work name and NIT, must be the selected office's; every other
      // office's tender — and any that can't be placed — is ignored outright (not paired, not listed).
      const norm = (v?: string) => (v ?? '').trim().toLowerCase()
      const mine = ee ? norm(office.circle) : norm(office.zone)
      const isMine = (l1: TenderEvaluation, loa?: IntimationNotice) => {
        const f = detectOffice(entries, l1, loa ?? null, {})
        return (ee ? norm(f.circle) : norm(f.zone)) === mine
      }
      const ownL1s = l1s.filter((x) => isMine(x.doc))
      // Other offices' tenders are not paired or generated — they are only listed under Skipped, with the reason.
      setSkipped(
        l1s
          .filter((x) => !ownL1s.includes(x))
          .map(({ doc: l1, file }) => {
            const f = detectOffice(entries, l1, null, {})
            const where = ee ? f.circle && `${f.circle} Circle` : f.zone && `${f.zone} Zone`
            return {
              key: `other|${file}`,
              nitNo: l1.noticeNo ?? '',
              agency: l1.l1AgencyName ?? '',
              workName: l1.nameOfWork ?? '',
              note: where ? `Belongs to ${where}, not ${ee ? office.circle : office.zone}` : `Couldn’t tell which ${ee ? 'circle' : 'zone'} this work belongs to`
            }
          })
      )
      const { pairs, noLoa } = pairL1WithLoa(ownL1s, loas, suppliers)

      // 3) Monitoring Format lookup (embeddings improve the name match when available).
      const mfRows = monitoringWorks ?? []
      let rowNameVectors: number[][] | undefined
      if (mfRows.length > 0) {
        try {
          rowNameVectors = await api.embedTexts(mfRows.map((r) => r.workName))
        } catch {
          rowNameVectors = undefined
        }
      }

      // 4) Generate one notice at a time.
      setPhase('generating')
      setGenProgress({ done: 0, total: pairs.length })
      const base = items
      const existing = base.filter((n) => !n.deletedAt)
      const seen = new Set<string>()
      const made: PendingNoticeItem[] = []
      const calendarByYear = new Map(calendar ? [[Number(calendar.year), calendar]] : [])

      for (let idx = 0; idx < pairs.length; idx++) {
        if (stale()) return
        const { l1, loa } = pairs[idx]
        const nitNo = loa.nitNo ?? l1.noticeNo ?? ''
        const agency = loa.agencyName ?? l1.l1AgencyName ?? ''
        const workName = l1.nameOfWork ?? ''
        const key = `${normalizeWorkNameForMatch(nitNo)}|${agencyKey(agency)}|${normalizeWorkNameForMatch(workName)}`
        const skip = (note: string) => setSkipped((s) => [...s, { key, nitNo, agency, workName, note }])
        setGenProgress({ done: idx + 1, total: pairs.length })

        if (seen.has(key)) continue // the same tender's files present twice in the folder
        seen.add(key)

        const found = detectOffice(entries, l1, loa, {})
        const duplicate = existing.some(
          (n) =>
            (!!l1.tenderId && n.tenderId === l1.tenderId) ||
            (!n.tenderId &&
              n.nitNo === nitNo &&
              agencyKey(n.agencyName ?? '') === agencyKey(agency) &&
              normalizeWorkNameForMatch(n.workName) === normalizeWorkNameForMatch(workName))
        )
        if (duplicate) {
          skip('A notice for this tender already exists')
          continue
        }

        let embeddings: { workNameVector: number[]; rowNameVectors: number[][] } | undefined
        if (rowNameVectors) {
          try {
            const [workNameVector] = await api.embedTexts([workName])
            embeddings = { workNameVector, rowNameVectors }
          } catch {
            embeddings = undefined
          }
        }
        if (stale()) return
        const decision = decideFromMonitoringFormat(findMonitoringWork(mfRows, workName, embeddings), mfRows)
        if (decision.kind === 'skip') {
          skip(decision.reason)
          continue
        }

        const zone = !ee ? found.zone : undefined
        const circle = ee ? found.circle : undefined
        const cno = ee ? found.cno : undefined
        const ecv = loa.ecvRupees ?? l1.ecvRupees
        const contract =
          loa.contractRupees ??
          l1.contractRupees ??
          (ecv != null && l1.tenderPercentage != null ? ecv * (1 - l1.tenderPercentage / 100) : undefined)
        // Notice Date = 8 working days after the LOA's own date, as Give Intimation sets it.
        const loaDate = l1.serverDate || ''
        const intimationDate = dmyToIso(loaDate) ?? isoToday()
        const [ny, nm, nd] = nthWorkingDay(intimationDate, 8, calendarByYear).split('-')
        const item: PendingNoticeItem = {
          id: `notice-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          officeKey: undefined, // stamped with the current office when added
          seMode: !ee,
          zone,
          circle,
          cno,
          corporation: corporationName,
          agencyName: agency || undefined,
          address: loa.address,
          workName: l1.nameOfWork,
          nitNo: nitNo || undefined,
          nitDate: loa.nitDate ?? l1.noticeDate,
          tenderId: l1.tenderId,
          itemNo: l1.itemNo,
          ecvRupees: ecv,
          contractRupees: contract,
          tenderPercentage: l1.tenderPercentage,
          intimationDate,
          manual: {
            ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS,
            loaDate,
            agencyPhone: loa.phone ?? '',
            loaNo: defaultLoaNo({ zone, circle, cno, corporation: corporationName }, loaDate),
            noticeDate: `${nd}.${nm}.${ny}`
          },
          done: false
        }
        made.push(item)
        // Added to Reminders the moment its tile appears.
        onChangeRef.current([...base, ...made])
        setTiles((t) => [...t, item])
        setTileNotes((n) => ({ ...n, [item.id]: decision.mfNote }))
        await pause(120) // let each tile paint before the next one
      }
      for (const { l1 } of noLoa) {
        setSkipped((s) => [
          ...s,
          {
            key: `noloa|${l1.tenderId ?? ''}|${l1.nameOfWork ?? ''}`,
            nitNo: l1.noticeNo ?? '',
            agency: l1.l1AgencyName ?? '',
            workName: l1.nameOfWork ?? '',
            note: 'L-1 found but no LOA or matching Supplier Details (same Tender ID and agency) in the folder'
          }
        ])
      }
      setPhase('done')
    } catch (e) {
      if (stale()) return
      setError(e instanceof Error ? e.message : String(e))
      setPhase('idle')
    }
  }

  // Grouped by NIT No, like the Reminders list itself.
  const groups = useMemo(() => {
    const map = new Map<string, PendingNoticeItem[]>()
    for (const t of tiles) {
      const k = t.intimationDate || ''
      map.set(k, [...(map.get(k) ?? []), t])
    }
    return [...map.entries()].sort(([a], [b]) => b.localeCompare(a))
  }, [tiles])

  const overdue = tiles.filter((t) => isNoticeOverdue(t.manual.noticeDate, t.done)).length
  const started = phase !== 'idle'
  const progress =
    phase === 'reading'
      ? {
          label: `Reading ${readProgress.done}/${readProgress.total}`,
          pct: readProgress.total ? (readProgress.done / readProgress.total) * 100 : 0
        }
      : phase === 'generating'
        ? {
            label: `Notice ${genProgress.done}/${genProgress.total}`,
            pct: genProgress.total ? (genProgress.done / genProgress.total) * 100 : 0
          }
        : phase === 'opening'
          ? scan?.phase === 'reading'
            ? { label: `Loading files ${scan.done}/${scan.total}`, pct: scan.total ? (scan.done / scan.total) * 100 : 0 }
            : { label: scan ? `Scanning folder… ${scan.done} found` : 'Waiting for folder…', pct: 0 }
          : null

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconBell /> Mass Notices
        </div>
        {phase !== 'idle' && (
          <button className="ghost" onClick={clear} title="Clear this run and start again from Upload folder (notices already in Reminders stay)">
            <IconRestore /> Clear
          </button>
        )}
        <button className="ghost pdf-ws-back" onClick={onBack}>
          <IconChevronLeft /> Back
        </button>
      </div>

      <div className="card pdf-workspace mass-notices-body">
        {!started && (
          <div className="mass-notices-upload">
            <p className="sub">
              Pick a folder of L-1 sheets and LOAs. Notices are made where the agreement isn’t concluded and added to Reminders.
            </p>
            <button onClick={() => void run()}>
              <IconFolder /> Upload folder
            </button>
          </div>
        )}

        {!monitoringWorks?.length && !started && (
          <div className="notice warn">
            <IconWarn /> No Monitoring Format imported — agreement dates won’t be checked.
          </div>
        )}
        {error && (
          <div className="notice warn">
            <IconWarn /> {error}
            <div>
              <button className="ghost" onClick={() => void run()} disabled={busy}>
                <IconFolder /> Try another folder
              </button>
            </div>
          </div>
        )}

        {progress && (
          <div
            className="mass-notices-progress"
            role="progressbar"
            aria-valuenow={Math.round(progress.pct)}
            aria-valuemin={0}
            aria-valuemax={100}
          >
            <div className="mass-notices-progress-label">
              {progress.label} {progress.pct > 0 && `(${Math.round(progress.pct)}%)`}
            </div>
            <div className="mass-notices-progress-track">
              <div className="mass-notices-progress-fill" style={{ width: `${progress.pct}%` }} />
            </div>
          </div>
        )}

        {started && !error && (
          <div className="ntc-summary" aria-label="Notices summary">
            <div className="ntc-chip">
              <span className="ntc-chip-label">Made</span>
              <span className="ntc-chip-value">{tiles.length}</span>
            </div>
            <div className="ntc-chip">
              <span className="ntc-chip-label">Active</span>
              <span className="ntc-chip-value">{tiles.length - overdue}</span>
            </div>
            <div className={`ntc-chip${overdue > 0 ? ' ntc-chip-warn' : ''}`}>
              <span className="ntc-chip-label">Overdue</span>
              <span className="ntc-chip-value">{overdue}</span>
            </div>
            <div className="ntc-chip">
              <span className="ntc-chip-label">Skipped</span>
              <span className="ntc-chip-value">{skipped.length}</span>
            </div>
            {ignored > 0 && (
              <div className="ntc-chip">
                <span className="ntc-chip-label">Ignored</span>
                <span className="ntc-chip-value">{ignored}</span>
              </div>
            )}
          </div>
        )}

        {phase === 'done' && (
          <div className="notice">
            <IconCheck /> {tiles.length} notice{tiles.length === 1 ? '' : 's'} added to Reminders.
          </div>
        )}

        {groups.map(([issued, group]) => (
          <div key={issued} className="notice-nit-group">
            <div className="ntc-group-head">
              <span className="ntc-group-title">Intimation issued {issued ? issued.split('-').reverse().join('.') : '— date not known'}</span>
              <span className="ntc-group-count">{group.length}</span>
            </div>
            <div className="mass-notices-tiles">
              {group.map((t) => {
                const late = isNoticeOverdue(t.manual.noticeDate, t.done)
                return (
                  <div key={t.id} className="mass-notice-tile">
                    <div className="mass-notice-tile-work">{t.workName || '—'}</div>
                    <div className="estimate-hint">
                      NIT {t.nitNo || '—'} · {t.nitDate || '—'}
                    </div>
                    <div className="estimate-hint">{t.agencyName || 'Unknown agency'}</div>
                    <div className="mass-notice-tile-foot">
                      <span>Due {t.manual.noticeDate || '—'}</span>
                      <span className={`mass-notice-badge${late ? ' overdue' : ''}`}>
                        {late ? 'Overdue' : 'Active'}
                      </span>
                    </div>
                    {tileNotes[t.id] && (
                      <div className="estimate-hint">
                        <IconWarn /> {tileNotes[t.id]}
                      </div>
                    )}
                    <div className="estimate-hint">
                      <IconCheck /> In Reminders
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        ))}

        {skipped.length > 0 && (
          <details className="mass-notices-skipped">
            <summary>Skipped ({skipped.length})</summary>
            <table className="mass-notices-table">
              <tbody>
                {skipped.map((o) => (
                  <tr key={o.key}>
                    <td>
                      <div>{o.workName || '—'}</div>
                      <div className="estimate-hint">
                        {o.nitNo ? `NIT ${o.nitNo} · ` : ''}
                        {o.agency || 'Unknown agency'} · {o.note}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </details>
        )}
      </div>
    </>
  )
}
