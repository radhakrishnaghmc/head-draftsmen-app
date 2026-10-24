import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { fetchTenders } from '../tenderCache'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { readEvaluationSheetPage } from '@core/tenderAgents/evaluationSheetAgent'
import { BID_CAPACITY_THRESHOLD, defaultConditions, isBidCapacityCondition, type EvaluationSheetInput } from '@core/evaluationSheet'
import { officeKey, type Office } from '../office'
import { officeKeyFromNit } from '../tenderOffice'
import { tenderConditionsInOffice } from '@core/tenderConditions'
import type { ExcelTable, TenderConditionsEntry, ShareResult } from '@core/types'
import { matchWorksRow } from '../estimateDownloads'
import { lakhsToRupees } from '@core/worksAmounts'
import { IconFolder, IconTable, IconPlus, IconTrash, IconChecklist, IconWarn, IconDownload, IconSearch, IconEye, IconUser } from './Icons'

interface Props {
  office: Office
  /** Works List tables; the first is the office's Works List. */
  tables: ExcelTable[]
  /** Tender conditions saved in the app by tender ID (the portal drops a tender once bidding closes). */
  savedConditions: Record<string, TenderConditionsEntry>
  onSaveConditions: (tenderId: string, conditions: string[], details?: { nitNo?: string; workName?: string; ecv?: number; officeKey?: string }) => void
  /** Called once sheets are saved, so those tenders' conditions can be dropped a week later. */
  onEvaluationGenerated: (tenderIds: string[]) => void
  onDeleteConditions: (tenderId: string) => void
}

interface SheetEntry {
  id: string
  sourceName: string
  tenderId: string
  nitLine: string
  workLine: string
  ecv: string
  /** Estimate amount in rupees from the Works List ("" when the work isn't found there). */
  estimate: string
  /** Where the estimate came from / why it is missing. */
  estimateNote: string | null
  signature: string
  bidders: string[]
  /** The conditions printed on the sheet (tender's own list or the standard ones), editable. */
  conditions: string[]
  /** Why the portal conditions could not be used (shown as a hint). */
  conditionsNote: string | null
  error: string | null
  saved: string | null
  /** The ECV the small-work question was last answered for, so it isn't asked again for the same value. */
  askedEcv: string | null
}

function newId(): string {
  return `ev-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(file)
  })
}

function suggestedName(e: SheetEntry): string {
  return `Evaluation Sheet${e.tenderId ? ` ${e.tenderId}` : ''}`
}

/**
 * Evaluation Sheet — upload the portal's "View Bidders" (Supplier List) PDFs,
 * or phone photos of the same page, and issue one Bid Capacity Evaluation
 * Sheet (.xlsx) per tender: one column per participating bidder, the eleven
 * fixed tender-condition rows, and the 3AN−B bid-capacity block wired as
 * formulas. The header, bidder columns and ECV are filled from each upload;
 * the per-bidder turnover / N / B / eligibility cells are left blank for the
 * office to key in from each bidder's documents. Several tenders can be
 * loaded at once and their sheets downloaded together into one folder.
 */
// "Save all" runs live at module level, keyed by the office that started them: switching office remounts the tab
// (so nothing of one office is shown in another), but the run keeps going in the background, stamping each tender
// with its own office, and its progress / Stop reappear when you switch back.
const batchByOffice = new Map<string, { done: number; total: number }>()
const cancelBatchByOffice = new Map<string, boolean>()
const batchListeners = new Set<() => void>()
function setBatchFor(key: string, value: { done: number; total: number } | null) {
  if (value) batchByOffice.set(key, value)
  else batchByOffice.delete(key)
  batchListeners.forEach((l) => l())
}
const subscribeBatch = (l: () => void) => {
  batchListeners.add(l)
  return () => {
    batchListeners.delete(l)
  }
}

// Last user list per corporation, so the Share dialog opens with it at once and refreshes behind it.
const shareUsersCache = new Map<string, string[]>()

export default function EvaluationSheetTab({ office, tables, savedConditions, onSaveConditions, onEvaluationGenerated, onDeleteConditions }: Props) {
  const [entries, setEntries] = useState<SheetEntry[]>([])
  const [busy, setBusy] = useState<null | 'reading' | 'downloadAll'>(null)
  // Upload progress shown inside the upload button: file `index` of `total`, in `phase`.
  const [progress, setProgress] = useState<{ index: number; total: number; phase: 'reading' | 'estimate' | 'conditions' } | null>(null)
  const [pickError, setPickError] = useState<string | null>(null)
  const [savedAllDir, setSavedAllDir] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  // Manual tender ID: fetch that tender's conditions from the portal and save them by ID.
  const [manualBusy, setManualBusy] = useState(false)
  // Which tender is being saved / how its last save ended, shown on that tender's own row.
  const [savingId, setSavingId] = useState<string | null>(null)
  const [rowResult, setRowResult] = useState<{ id: string; ok: boolean; text: string } | null>(null)
  const [manualMsg, setManualMsg] = useState<{
    ok: boolean
    text: string
  } | null>(null)

  // Saved tenders are per office: an SE (zone) office's tenders stay out of an EE (circle) office and vice versa.
  const currentOfficeKey = officeKey(office) ?? ''
  const zoneLevel = !(office.circle ?? '').trim()
  const inOffice = (id: string) => !!savedConditions[id] && tenderConditionsInOffice(savedConditions[id], currentOfficeKey, zoneLevel)
  const mineIds = Object.keys(savedConditions).filter(inOffice)

  // Keyword search of live portal tenders, so the ID needn't be typed.
  const [viewId, setViewId] = useState<string | null>(null)
  // Download of an evaluation sheet straight from a saved tender: ask about agencies first.
  const [dlId, setDlId] = useState<string | null>(null)
  // Saved tender awaiting delete confirmation (a closed tender can't be re-fetched).
  const [confirmDelId, setConfirmDelId] = useState<string | null>(null)
  const [confirmDelAll, setConfirmDelAll] = useState(false)
  // "Share with other users": the corporation's other logins (from the login sheet) to copy this office's saved tenders to.
  const corporation = (office.corporation ?? '').trim()
  const [shareOpen, setShareOpen] = useState(false)
  // Which saved tenders the dialog shares: every one of this office (the bar button) or just one row's.
  const [shareIds, setShareIds] = useState<string[]>([])
  // Tenders whose work name names no circle: they keep the sender's office when shared.
  const [shareUnresolved, setShareUnresolved] = useState(0)
  const [shareUsers, setShareUsers] = useState<string[] | null>(null)
  const [shareErr, setShareErr] = useState<string | null>(null)
  const [sharePicked, setSharePicked] = useState<Set<string>>(new Set())
  const [shareBusy, setShareBusy] = useState(false)
  const [shareDone, setShareDone] = useState<ShareResult[] | null>(null)
  // Warm the list as soon as the tab is open, so the dialog doesn't wait on the login sheet.
  useEffect(() => {
    if (corporation && !shareUsersCache.has(corporation)) {
      api.listShareUsers(corporation).then((list) => shareUsersCache.set(corporation, list)).catch(() => {})
    }
  }, [corporation])
  function openShare(ids: string[]) {
    setShareIds(ids)
    setShareUnresolved(0)
    setShareOpen(true)
    setShareUsers(null)
    setShareErr(null)
    setSharePicked(new Set())
    setShareDone(null)
    const cached = shareUsersCache.get(corporation)
    if (cached) setShareUsers(cached)
    api
      .listShareUsers(corporation)
      .then((list) => {
        shareUsersCache.set(corporation, list)
        setShareUsers(list)
      })
      .catch((e) => {
        if (cached) return
        setShareErr(
          String(e).includes('NOT_SIGNED_IN')
            ? 'Sharing needs a signed-in account. This build started without the login screen (dev mode) — log in through a normal build to share.'
            : 'Could not load the users — check your internet connection.'
        )
      })
  }
  async function runShare() {
    setShareBusy(true)
    setShareErr(null)
    try {
      // The recipient sees a tender under the office its NIT number decides (EE -> that circle's office, SE -> the
      // zone-level office), not under the sender's — so each entry is re-stamped with that office key.
      const entries: Record<string, TenderConditionsEntry> = {}
      let unresolved = 0
      for (const id of shareIds) {
        const e = savedConditions[id]
        if (!e) continue
        const key = officeKeyFromNit(e.nitNo, corporation, (office.zone ?? '').trim(), e.workName)
        if (key) entries[id] = { ...e, officeKey: key }
        else {
          entries[id] = e
          unresolved++
        }
      }
      setShareUnresolved(unresolved)
      setShareDone(await api.shareTenderConditions(corporation, [...sharePicked], entries))
    } catch (e) {
      setShareErr(
        String(e).includes('NOT_SIGNED_IN')
          ? 'Sharing needs a signed-in account — log in through a normal build to share.'
          : 'Could not share — check your internet connection and try again.'
      )
    } finally {
      setShareBusy(false)
    }
  }
  const [dlAgencies, setDlAgencies] = useState<string[] | null>(null)
  const [dlMsg, setDlMsg] = useState<{ ok: boolean; text: string } | null>(null)
  // NIT / work / ECV for the sheet: prefilled from the saved tender (or the portal while it is live), editable.
  const [dlDetails, setDlDetails] = useState({ nit: '', work: '', ecv: '' })
  useEffect(() => {
    if (!dlId) return
    const saved = savedConditions[dlId]
    setDlDetails({ nit: saved?.nitNo ?? '', work: saved?.workName ?? '', ecv: saved?.ecv != null ? String(saved.ecv) : '' })
    if (saved?.nitNo && saved?.workName && saved?.ecv != null) return
    let stale = false
    void fetchTenders({ start: 0, length: 10, type: 'current', search: dlId })
      .then((res) => {
        const r = res.data.find((x) => (x[1] ?? '').trim() === dlId)
        if (!r || stale) return
        const n = Number(r[5])
        setDlDetails((d) => ({
          nit: d.nit || (r[2] ?? '').trim(),
          work: d.work || (r[4] ?? '').trim(),
          ecv: d.ecv || (Number.isFinite(n) && n > 0 ? String(n) : '')
        }))
      })
      .catch(() => {})
    return () => {
      stale = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dlId])

  async function downloadSavedSheet(id: string, agencies: string[]) {
    const saved = savedConditions[id]
    if (!saved) return
    const workName = dlDetails.work.trim()
    const nitNo = dlDetails.nit.replace(/^\s*NIT\s*No\.?:?\s*/i, '').trim()
    const ecvNum = Number(dlDetails.ecv.replace(/,/g, '')) || 0
    const workLine = `Name of the work: ${workName}${ecvNum > 0 ? `. ecv: ${ecvNum.toFixed(2)}` : ''}`
    const entry: SheetEntry = {
      id: newId(),
      sourceName: `Tender ${id}`,
      tenderId: id,
      nitLine: `NIT No: ${nitNo},Tender ID:${id}`,
      workLine,
      ecv: ecvNum > 0 ? String(ecvNum) : '',
      estimate: '',
      estimateNote: null,
      signature: defaultSignature(),
      bidders: agencies,
      conditions: defaultConditions(saved.conditions, `Name of the work: ${workName}`, null),
      conditionsNote: null,
      error: null,
      saved: null,
      askedEcv: null
    }
    // Remember what was typed so the next download is already filled in.
    onSaveConditions(id, saved.conditions, { nitNo: nitNo || undefined, workName: workName || undefined, ecv: ecvNum > 0 ? ecvNum : undefined, officeKey: currentOfficeKey })
    setDlId(null)
    setDlAgencies(null)
    const e = await confirmSmallWork(entry)
    const input = toInput(e)
    if (!input) return
    try {
      const path = await api.exportEvaluationSheet(input, suggestedName(e))
      setDlMsg(path ? { ok: true, text: `Saved to ${path}` } : null)
      if (path) onEvaluationGenerated([id])
    } catch (err) {
      setDlMsg({ ok: false, text: err instanceof Error ? err.message : String(err) })
    }
  }
  const [tq, setTq] = useState('')
  const [tResults, setTResults] = useState<string[][] | null>(null)
  const [tTotal, setTTotal] = useState(-1)
  const [tBusy, setTBusy] = useState(false)
  const [tErr, setTErr] = useState<string | null>(null)
  useEffect(() => {
    const term = tq.trim()
    if (!term) {
      setTResults(null)
      setTErr(null)
      return
    }
    let stale = false
    const t = setTimeout(async () => {
      setTBusy(true)
      setTErr(null)
      try {
        const res = await fetchTenders({
          start: 0,
          length: 100,
          type: 'current',
          search: term
        })
        if (!stale) {
          setTResults(res.data)
          setTTotal(res.total)
        }
      } catch (e) {
        if (!stale) {
          setTResults([])
          setTErr(e instanceof Error ? e.message : String(e))
        }
      } finally {
        if (!stale) setTBusy(false)
      }
    }, 600)
    return () => {
      stale = true
      clearTimeout(t)
    }
  }, [tq])

  // Tenders saved before their details were kept: fill the work name / ECV in from the portal while they are still live.
  const backfilled = useRef(new Set<string>())
  useEffect(() => {
    const missing = mineIds.filter((id) => !savedConditions[id].workName && !backfilled.current.has(id))
    if (missing.length === 0) return
    missing.forEach((id) => backfilled.current.add(id))
    void (async () => {
      for (const id of missing) {
        try {
          const res = await fetchTenders({ start: 0, length: 10, type: 'current', search: id })
          const r = res.data.find((x) => (x[1] ?? '').trim() === id)
          if (!r) continue
          const n = Number(r[5])
          onSaveConditions(id, savedConditions[id].conditions, {
            nitNo: r[2]?.trim() || undefined,
            workName: r[4]?.trim() || undefined,
            ecv: Number.isFinite(n) && n > 0 ? n : undefined
          })
        } catch {
          /* portal unreachable — try again next launch */
        }
      }
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [savedConditions])

  // "Save all" for the tenders listed by the current search: one portal fetch at a time, skipping those already saved here.
  const batch = useSyncExternalStore(subscribeBatch, () => batchByOffice.get(currentOfficeKey) ?? null)
  async function saveAllListed() {
    const todo = (tResults ?? []).filter((r) => /^\d+$/.test((r[1] ?? '').trim()) && !(inOffice((r[1] ?? '').trim()) && savedConditions[(r[1] ?? '').trim()]?.conditions.length))
    if (todo.length === 0 || batch) return
    const runOffice = currentOfficeKey
    cancelBatchByOffice.set(runOffice, false)
    let saved = 0
    let failed = 0
    setBatchFor(runOffice, { done: 0, total: todo.length })
    for (let i = 0; i < todo.length; i += 1) {
      if (cancelBatchByOffice.get(runOffice)) break
      const tid = (todo[i][1] ?? '').trim()
      // A tender the portal won't open must not block the rest: give each 45s, mark it failed and move on.
      let timer: ReturnType<typeof setTimeout> | undefined
      const ok = await Promise.race([
        saveManualConditions(tid, todo[i]),
        new Promise<boolean>((resolve) => {
          timer = setTimeout(() => {
            setRowResult({ id: tid, ok: false, text: 'Timed out waiting for the portal. Try again.' })
            resolve(false)
          }, 45000)
        })
      ])
      clearTimeout(timer)
      if (ok) saved += 1
      else failed += 1
      setBatchFor(runOffice, { done: i + 1, total: todo.length })
    }
    setBatchFor(runOffice, null)
    const left = todo.length - saved - failed
    setManualMsg({
      ok: failed === 0,
      text: `Saved ${saved} of ${todo.length} tender${todo.length === 1 ? '' : 's'}${failed ? `, ${failed} failed — use Save conditions on those rows to retry` : ''}${left > 0 ? `; stopped with ${left} not tried` : ''}.`
    })
  }

  async function saveManualConditions(idArg: string, row?: string[]): Promise<boolean> {
    const id = idArg.trim()
    if (!id) return false
    if (!/^\d+$/.test(id)) {
      setManualMsg({ ok: false, text: 'Enter the numeric tender ID.' })
      return false
    }
    if (inOffice(id) && savedConditions[id]?.conditions.length) {
      setManualMsg({
        ok: true,
        text: `Tender ${id} conditions are already saved (${savedConditions[id].conditions.length}).`
      })
      return true
    }
    let ok = false
    setManualBusy(true)
    setSavingId(id)
    setRowResult(null)
    setManualMsg(null)
    try {
      const conditions = await api.fetchTenderConditions(id)
      // Keep the portal's NIT / work / ECV with the conditions: the tender vanishes once bidding closes.
      let r = row
      if (!r) {
        try {
          r = (await fetchTenders({ start: 0, length: 10, type: 'current', search: id })).data.find((x) => (x[1] ?? '').trim() === id)
        } catch {
          r = undefined
        }
      }
      const ecvNum = Number(r?.[5])
      onSaveConditions(id, conditions, {
        nitNo: r?.[2]?.trim() || undefined,
        workName: r?.[4]?.trim() || undefined,
        ecv: Number.isFinite(ecvNum) && ecvNum > 0 ? ecvNum : undefined,
        officeKey: currentOfficeKey
      })
      setManualMsg({ ok: true, text: `Saved ${conditions.length} tender conditions for tender ${id}. Its evaluation sheet will use them.` })
      setRowResult({ id, ok: true, text: `Saved ${conditions.length} conditions.` })
      ok = true
      // Sheets already loaded for this tender switch to the saved conditions.
      setEntries((prev) =>
        prev.map((x) =>
          x.tenderId === id
            ? {
                ...x,
                conditions: defaultConditions(conditions, x.workLine, Number(x.estimate.replace(/,/g, '')) || null),
                conditionsNote: `Using the tender conditions saved earlier for tender ${id}.`
              }
            : x
        )
      )
    } catch (err) {
      const why = err instanceof Error ? err.message.replace(/^Error invoking remote method '[^']*': (Error: )?/, '') : String(err)
      setRowResult({ id, ok: false, text: `Could not save (${why}). Try again.` })
      setManualMsg({ ok: false, text: `Could not save the conditions (${why}). Standard conditions are used unless some were saved earlier.` })
    } finally {
      setManualBusy(false)
      setSavingId(null)
    }
    return ok
  }
  // Pop-up asking whether a work missing from the Works List has an estimate of 5 lakhs or below.
  const [estimateAsk, setEstimateAsk] = useState<{ work: string; why: string; answer: (below: boolean) => void } | null>(null)

  function askEstimateBelowFiveLakhs(work: string, why: string): Promise<boolean> {
    return new Promise((resolve) =>
      setEstimateAsk({
        work,
        why,
        answer: (below) => {
          setEstimateAsk(null)
          resolve(below)
        }
      })
    )
  }

  function defaultSignature(): string {
    const circle = (office.circle ?? '').trim()
    const cno = (office.circleNumber ?? '').trim()
    const corp = (office.corporation ?? '').trim() || 'CMC'
    const circlePart = circle ? `${circle} circle${cno ? `-${cno}` : ''},` : ''
    return `Executive Engineer,\n${circlePart}${corp}`
  }

  async function handleFiles(files: File[]) {
    setBusy('reading')
    setPickError(null)
    setSavedAllDir(null)
    try {
      for (const [fileIndex, file] of files.entries()) {
        setProgress({ index: fileIndex, total: files.length, phase: 'reading' })
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
        try {
          const lines = isPdf ? await pdfToTextLinesOrOcr(file) : await api.ocrPhotosToLines([await readAsDataUrl(file)])
          // Evaluation Sheet agent: header fields + participating bidders from the page.
          const { header: ev, bidders: list } = readEvaluationSheetPage(lines)
          // Estimate amount from the Works List, matched on the work's name — it
          // decides whether the sheet needs the 3AN−B rows (estimate > 5 lakhs).
          let estimate = ''
          let estimateNote: string | null = null
          const worksTable = tables[0]
          if (!ev.nameOfWork) {
            estimateNote = 'No work name found in the upload, so the Works List estimate could not be looked up.'
          } else if (!worksTable) {
            estimateNote = 'No Works List loaded, so the estimate could not be looked up.'
          } else {
            setProgress({ index: fileIndex, total: files.length, phase: 'estimate' })
            const match = await matchWorksRow(ev.nameOfWork, worksTable)
            const rupees = match ? lakhsToRupees(match.row['Amount of estimate'] ?? '') : 0
            if (rupees > 0) {
              estimate = String(rupees)
              estimateNote = `Estimate ₹${rupees.toLocaleString('en-IN')} from the Works List${match?.matchedViaAi ? ' (AI name match — please verify)' : ''}.`
            } else {
              estimateNote = `${match ? 'The Works List row has no estimate amount' : 'Work not found on the Works List'}.`
            }
          }
          // The estimate is always above the ECV, so an ECV over 5 lakhs already means
          // 3AN−B applies. A smaller ECV leaves the question open whatever the Works
          // List says, so ask: yes drops the bid-capacity conditions and rows.
          if (ev.nameOfWork && ev.ecvRupees != null && ev.ecvRupees < BID_CAPACITY_THRESHOLD) {
            const below = await askEstimateBelowFiveLakhs(ev.nameOfWork, `The ECV is ₹${ev.ecvRupees.toLocaleString('en-IN')}, below ₹5 lakhs`)
            if (below) {
              estimate = String(BID_CAPACITY_THRESHOLD)
              estimateNote = 'Estimate confirmed below ₹5 lakhs, so the 3AN−B conditions and rows are left out.'
            } else {
              estimateNote = `${estimateNote ? estimateNote + ' ' : ''}Estimate confirmed ₹5 lakhs or above, so the 3AN−B rows are kept.`
            }
          }
          // The View Bidders sheet names the tender; look it up on the portal and
          // take its listed conditions in place of the standard eleven.
          let fetched: string[] | null = null
          let conditionsNote: string | null = null
          if (ev.tenderId) {
            setProgress({ index: fileIndex, total: files.length, phase: 'conditions' })
            // Conditions saved earlier win: a closed tender is gone from the portal.
            fetched = inOffice(ev.tenderId) && savedConditions[ev.tenderId]?.conditions.length ? savedConditions[ev.tenderId].conditions : null
            if (fetched) {
              conditionsNote = `Using the tender conditions saved earlier for tender ${ev.tenderId}.`
            } else {
              try {
                fetched = await api.fetchTenderConditions(ev.tenderId)
                onSaveConditions(ev.tenderId, fetched, { officeKey: currentOfficeKey })
              } catch (err) {
                const why = err instanceof Error ? err.message.replace(/^Error invoking remote method '[^']*': (Error: )?/, '') : String(err)
                conditionsNote = `Could not fetch the tender conditions from the portal (${why}); using the standard conditions. Enter this tender ID above while the tender is still open to save its conditions.`
              }
            }
          } else {
            conditionsNote = 'No tender ID found in the upload, so the standard conditions are used.'
          }
          const entry: SheetEntry = {
            id: newId(),
            sourceName: file.name,
            tenderId: ev.tenderId ?? '',
            nitLine: `NIT No: ${ev.noticeNo ?? ''}${ev.noticeDate ? `,Dt:${ev.noticeDate}` : ''}${ev.tenderId ? `,Tender ID:${ev.tenderId}` : ''}`,
            workLine: `Name of the work: ${ev.nameOfWork ?? ''}${ev.ecvRupees != null ? `. ecv: ${ev.ecvRupees.toFixed(2)}` : ''}`,
            ecv: ev.ecvRupees != null ? String(ev.ecvRupees) : '',
            signature: defaultSignature(),
            bidders: list.map((b) => b.name),
            estimate,
            estimateNote,
            conditions: defaultConditions(fetched, `Name of the work: ${ev.nameOfWork ?? ''}`, estimate ? Number(estimate) || null : null),
            conditionsNote,
            error:
              list.length === 0
                ? `No participating bidders found in ${file.name}. Upload the portal's "View Bidders" (Supplier List) page, as a PDF or a clear photo.`
                : null,
            saved: null,
            askedEcv: ev.ecvRupees != null && ev.ecvRupees < BID_CAPACITY_THRESHOLD ? String(ev.ecvRupees) : null
          }
          setEntries((prev) => [...prev, entry])
        } catch (e) {
          setEntries((prev) => [
            ...prev,
            {
              id: newId(),
              sourceName: file.name,
              tenderId: '',
              nitLine: '',
              workLine: '',
              ecv: '',
              signature: defaultSignature(),
              bidders: [],
              estimate: '',
              estimateNote: null,
              conditions: [],
              conditionsNote: null,
              error: e instanceof Error ? e.message : String(e),
              saved: null,
              askedEcv: null
            }
          ])
        }
      }
    } finally {
      setBusy(null)
      setProgress(null)
    }
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  function updateEntry(id: string, patch: Partial<SheetEntry>) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)))
  }

  function setCondition(id: string, i: number, value: string) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: x.conditions.map((c, j) => (j === i ? value : c)) } : x)))
  }
  function addCondition(id: string) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: [...x.conditions, ''] } : x)))
  }
  function removeCondition(id: string, i: number) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: x.conditions.filter((_, j) => j !== i) } : x)))
  }

  function setBidder(id: string, i: number, value: string) {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, bidders: e.bidders.map((b, j) => (j === i ? value : b)) } : e))
    )
  }
  function addBidder(id: string) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, bidders: [...e.bidders, ''] } : e)))
  }
  function removeBidder(id: string, i: number) {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, bidders: e.bidders.filter((_, j) => j !== i) } : e))
    )
  }

  function toInput(e: SheetEntry): EvaluationSheetInput | null {
    const bidders = e.bidders.map((b) => b.trim()).filter(Boolean)
    if (bidders.length === 0) return null
    return { nitLine: e.nitLine, workLine: e.workLine, bidders, ecv: Number(e.ecv.replace(/,/g, '')) || null, estimate: Number(e.estimate.replace(/,/g, '')) || null, signature: e.signature, conditions: e.conditions, conditionsFinal: true }
  }

  /**
   * The ECV can be edited after upload. Below ₹5 lakhs, ask once per ECV value
   * whether the estimate is too; yes drops the bid-capacity conditions and rows.
   * Returns the entry to download (updated in place on screen too).
   */
  async function confirmSmallWork(e: SheetEntry): Promise<SheetEntry> {
    const ecv = Number(e.ecv.replace(/,/g, '')) || 0
    if (ecv <= 0 || ecv >= BID_CAPACITY_THRESHOLD || e.askedEcv === e.ecv) return e
    const below = await askEstimateBelowFiveLakhs(e.workLine.replace(/^Name of the work:\s*/i, ''), `The ECV is ₹${ecv.toLocaleString('en-IN')}, below ₹5 lakhs`)
    const patch: Partial<SheetEntry> = below
      ? {
          askedEcv: e.ecv,
          estimate: String(BID_CAPACITY_THRESHOLD),
          estimateNote: 'Estimate confirmed below ₹5 lakhs, so the 3AN−B conditions and rows are left out.',
          conditions: e.conditions.filter((c) => !isBidCapacityCondition(c))
        }
      : { askedEcv: e.ecv, estimateNote: 'Estimate confirmed ₹5 lakhs or above, so the 3AN−B rows are kept.' }
    updateEntry(e.id, patch)
    return { ...e, ...patch }
  }

  async function downloadOne(entry: SheetEntry) {
    const e = await confirmSmallWork(entry)
    const input = toInput(e)
    if (!input) {
      updateEntry(e.id, { error: 'Add at least one bidder before downloading.' })
      return
    }
    setBusy('reading')
    updateEntry(e.id, { error: null, saved: null })
    try {
      const path = await api.exportEvaluationSheet(input, suggestedName(e))
      updateEntry(e.id, { saved: path ? `Saved to ${path}` : 'Cancelled.' })
      if (path && e.tenderId) onEvaluationGenerated([e.tenderId])
    } catch (err) {
      updateEntry(e.id, { error: err instanceof Error ? err.message : String(err) })
    } finally {
      setBusy(null)
    }
  }

  async function downloadAll() {
    const confirmed: SheetEntry[] = []
    for (const entry of entries) confirmed.push(entry.bidders.some((b) => b.trim()) ? await confirmSmallWork(entry) : entry)
    const ready = confirmed.map((e) => ({ e, input: toInput(e) })).filter((x) => x.input != null)
    if (ready.length === 0) {
      setPickError('Add at least one bidder before downloading.')
      return
    }
    setBusy('downloadAll')
    setPickError(null)
    setSavedAllDir(null)
    try {
      const paths = await api.exportEvaluationSheetBatch(
        ready.map(({ e, input }) => ({ input: input as EvaluationSheetInput, suggestedName: suggestedName(e) }))
      )
      if (paths && paths.length > 0) {
        setSavedAllDir(paths[0].replace(/[^/\\]+$/, ''))
        onEvaluationGenerated(ready.map(({ e }) => e.tenderId).filter(Boolean))
      }
    } catch (e) {
      setPickError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Each file counts as three steps: read the PDF, look up its estimate, fetch its conditions.
  const pct = progress
    ? Math.round(((progress.index + (progress.phase === 'conditions' ? 0.66 : progress.phase === 'estimate' ? 0.33 : 0)) / progress.total) * 100)
    : 0
  const readyCount = entries.filter((e) => e.bidders.some((b) => b.trim())).length

  return (
    <div className="card">
      {dlId &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => { setDlId(null); setDlAgencies(null) }}>
            <div
              className="wo-modal tender-conditions-modal"
              role="dialog"
              aria-modal="true"
              aria-label={`Download evaluation sheet for tender ${dlId}`}
              onClick={(ev) => ev.stopPropagation()}
              onKeyDown={(ev) => ev.key === 'Escape' && (setDlId(null), setDlAgencies(null))}
            >
              <h3>Evaluation sheet — tender {dlId}</h3>
              {dlAgencies === null ? (
                <>
                  <div className="dl-fields">
                    <label>
                      NIT number
                      <input type="text" value={dlDetails.nit} onChange={(ev) => setDlDetails((d) => ({ ...d, nit: ev.target.value }))} />
                    </label>
                    <label>
                      ECV (₹)
                      <input type="text" inputMode="decimal" value={dlDetails.ecv} onChange={(ev) => setDlDetails((d) => ({ ...d, ecv: ev.target.value }))} />
                    </label>
                    <label className="dl-fields__wide">
                      Name of the work
                      <input type="text" value={dlDetails.work} onChange={(ev) => setDlDetails((d) => ({ ...d, work: ev.target.value }))} />
                    </label>
                  </div>
                  <p className="tender-conditions__hint">
                    These fill the sheet along with the tender ID and conditions. Download it without agencies (columns named Agency 1, Agency 2), or add the agency names first?
                  </p>
                  <div className="boq-actions">
                    <button className="primary" autoFocus onClick={() => void downloadSavedSheet(dlId, ['Agency 1', 'Agency 2'])}>
                      Yes, download without agencies
                    </button>
                    <button className="ghost" onClick={() => setDlAgencies([''])}>
                      Add agencies
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <ul className="todo-list tender-conditions__list">
                    {dlAgencies.map((a, i) => (
                      <li key={i} className="todo-item">
                        <input
                          type="text"
                          style={{ flex: 1 }}
                          autoFocus={i === dlAgencies.length - 1}
                          placeholder={`Agency ${i + 1} name`}
                          value={a}
                          onChange={(ev) => setDlAgencies((prev) => (prev ?? []).map((x, j) => (j === i ? ev.target.value : x)))}
                        />
                        <button className="danger-ghost" title="Remove" onClick={() => setDlAgencies((prev) => (prev ?? []).filter((_, j) => j !== i))}>
                          <IconTrash />
                        </button>
                      </li>
                    ))}
                  </ul>
                  <div className="boq-actions">
                    <button className="ghost" onClick={() => setDlAgencies((prev) => [...(prev ?? []), ''])}>
                      <IconPlus /> Add another agency
                    </button>
                    <button className="primary" disabled={!dlAgencies.some((a) => a.trim())} onClick={() => void downloadSavedSheet(dlId, dlAgencies.map((a) => a.trim()).filter(Boolean))}>
                      <IconDownload /> Download sheet
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>,
          document.body
        )}
      {shareOpen &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => !shareBusy && setShareOpen(false)}>
            <div
              className="wo-modal tender-conditions-modal"
              role="dialog"
              aria-modal="true"
              aria-label="Share saved tenders with other users"
              onClick={(ev) => ev.stopPropagation()}
              onKeyDown={(ev) => ev.key === 'Escape' && !shareBusy && setShareOpen(false)}
            >
              <h3>Share saved tenders</h3>
              <p className="tender-id-card__hint">
                Copies {shareIds.length === 1 ? `tender ${shareIds[0]}` : `${shareIds.length} saved tenders`} to the users you pick
                {corporation ? ` (${corporation} users)` : ''}. Each tender goes to the office its NIT number says (EE: the circle named in the work name, SE: the zone office), so it shows in that office’s Evaluation Sheet for them. A tender they already have is left as it is.
              </p>
              {shareDone ? (
                <ul className="todo-list">
                  {shareDone.map((r) => (
                    <li key={r.loginId} className="todo-item">
                      <div className="todo-body">
                        <span className="todo-text">{r.loginId}</span>
                        <span className="todo-dates">
                          {r.error
                            ? `Not shared — ${r.error}`
                            : `${r.added} shared${r.skipped ? `, ${r.skipped} already had` : ''}`}
                        </span>
                      </div>
                    </li>
                  ))}
                  {shareUnresolved > 0 && (
                    <li className="todo-item">
                      <span className="todo-dates">
                        {shareUnresolved} tender{shareUnresolved === 1 ? ' has' : 's have'} no EE/SE office in the NIT number (or no circle for an EE work), so {shareUnresolved === 1 ? 'it was' : 'they were'} shared under your office.
                      </span>
                    </li>
                  )}
                </ul>
              ) : shareErr ? (
                <p className="tender-id-card__hint">{shareErr}</p>
              ) : !corporation ? (
                <p className="tender-id-card__hint">Pick your office (corporation) first.</p>
              ) : shareUsers === null ? (
                <p className="tender-id-card__hint">Loading users…</p>
              ) : shareUsers.length === 0 ? (
                <p className="tender-id-card__hint">No other users are listed under {corporation} in the login sheet.</p>
              ) : (
                <>
                  <label className="share-user">
                    <input
                      type="checkbox"
                      checked={sharePicked.size === shareUsers.length}
                      onChange={(ev) => setSharePicked(ev.target.checked ? new Set(shareUsers) : new Set())}
                    />{' '}
                    Select all ({shareUsers.length})
                  </label>
                  <ul className="todo-list" style={{ maxHeight: 280, overflowY: 'auto' }}>
                    {shareUsers.map((u) => (
                      <li key={u} className="todo-item">
                        <label className="share-user">
                          <input
                            type="checkbox"
                            checked={sharePicked.has(u)}
                            onChange={() =>
                              setSharePicked((prev) => {
                                const next = new Set(prev)
                                if (next.has(u)) next.delete(u)
                                else next.add(u)
                                return next
                              })
                            }
                          />{' '}
                          {u}
                        </label>
                      </li>
                    ))}
                  </ul>
                </>
              )}
              <div className="boq-actions">
                {!shareDone && shareUsers && shareUsers.length > 0 && (
                  <button className="primary" disabled={shareBusy || sharePicked.size === 0} onClick={() => void runShare()}>
                    {shareBusy ? 'Sharing…' : `Share with ${sharePicked.size} user${sharePicked.size === 1 ? '' : 's'}`}
                  </button>
                )}
                <button className="ghost" disabled={shareBusy} onClick={() => setShareOpen(false)}>
                  {shareDone ? 'Close' : 'Cancel'}
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
      {viewId &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => setViewId(null)}>
            <div
              className="wo-modal tender-conditions-modal"
              role="dialog"
              aria-modal="true"
              aria-label={`Conditions for tender ${viewId}`}
              onClick={(ev) => ev.stopPropagation()}
              onKeyDown={(ev) => ev.key === 'Escape' && setViewId(null)}
            >
              <h3>Tender {viewId} — saved conditions</h3>
              {savedConditions[viewId]?.conditions.length ? (
                <ol className="tender-conditions-list">
                  {savedConditions[viewId].conditions.map((c, i) => (
                    <li key={i}>{c}</li>
                  ))}
                </ol>
              ) : (
                <p className="tender-id-card__hint">No conditions saved for this tender.</p>
              )}
              <div className="boq-actions">
                <button className="primary" autoFocus onClick={() => setViewId(null)}>
                  Close
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
      {estimateAsk && (
        <div className="wo-modal-overlay">
          <div className="wo-modal office-modal" role="dialog" aria-modal="true" style={{ padding: 20 }}>
            <h3 style={{ marginTop: 0 }}>Estimate amount</h3>
            <p className="estimate-hint">{estimateAsk.why}:</p>
            <p>{estimateAsk.work}</p>
            <p>Is the Estimate amount less than 5 Lakhs? (3AN−B is checked only at ₹5 lakhs and above.)</p>
            <div className="boq-actions">
              <button className="primary" onClick={() => estimateAsk.answer(true)}>
                Yes, less than 5 lakhs
              </button>
              <button className="ghost" onClick={() => estimateAsk.answer(false)}>
                No, 5 lakhs or above
              </button>
            </div>
          </div>
        </div>
      )}
      <div className="empty empty--tight">
        <IconChecklist />
        <div className="boq-actions">
          <button
            className="primary upload-btn"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy !== null}
            style={
              progress
                ? { background: `linear-gradient(to right, rgba(255,255,255,0.28) ${pct}%, transparent ${pct}%), var(--primary, #2563eb)` }
                : undefined
            }
          >
            <IconFolder />{' '}
            {progress
              ? `${progress.phase === 'reading' ? 'Reading' : progress.phase === 'estimate' ? 'Looking up estimate' : 'Fetching tender conditions'} ${progress.index + 1}/${progress.total}… ${pct}%`
              : busy === 'reading'
                ? 'Reading…'
                : entries.length > 0
                  ? 'Add more tenders'
                  : 'Upload View Bidders PDFs / photos'}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="application/pdf,.pdf,image/*"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              const files = Array.from(e.target.files ?? [])
              if (files.length > 0) void handleFiles(files)
              e.target.value = ''
            }}
          />
          {readyCount > 1 && (
            <button className="ghost upload-btn" onClick={downloadAll} disabled={busy !== null}>
              <IconDownload /> {busy === 'downloadAll' ? 'Saving…' : `Download all ${readyCount} evaluation sheets`}
            </button>
          )}
        </div>
        {pickError && (
          <div className="notice error">
            <IconWarn /> {pickError}
          </div>
        )}
        {savedAllDir && (
          <div className="notice ok">
            <IconTable /> Saved {readyCount} evaluation sheets to {savedAllDir}
          </div>
        )}
      </div>

      <section className="tender-conditions">
        <div className="card-head">
          <div className="head-ic">
            <IconChecklist />
          </div>
          <div className="titles">
            <h2>Tender Conditions</h2>
            <p className="sub">The portal drops a tender once bidding closes — save its conditions while it is live and the sheet will use them.</p>
          </div>
        </div>
        <div className="tsearch">
          <IconSearch />
          <input
            type="text"
            placeholder="Search by Tender ID or name of work…"
            value={tq}
            onChange={(ev) => setTq(ev.target.value)}
            aria-label="Search live tenders"
          />
          {(tBusy || manualBusy) && <span className="tsearch-spin">{manualBusy ? 'Saving…' : 'Searching…'}</span>}
          {tq && (
            <button className="tsearch-clear" onClick={() => setTq('')}>
              Clear
            </button>
          )}
        </div>
        {tErr && <p className="tender-conditions__hint">Search failed: {tErr}</p>}
        {tResults && !tBusy && !tErr && tResults.length === 0 && !/^\d{5,}$/.test(tq.trim()) && (
          <p className="tender-conditions__hint">No live tenders match.</p>
        )}
        {tResults && tResults.length > 0 && (
          <div className="tender-conditions__bar">
            <p className="tender-conditions__hint">
              {tTotal > tResults.length ? `Showing the first ${tResults.length} of ${tTotal} tenders — refine the search to narrow them.` : `${tResults.length} tender${tResults.length === 1 ? '' : 's'} found.`}
            </p>
            {(() => {
              const pending = tResults.filter((r) => /^\d+$/.test((r[1] ?? '').trim()) && !(inOffice((r[1] ?? '').trim()) && savedConditions[(r[1] ?? '').trim()]?.conditions.length)).length
              return batch ? (
                <>
                  <span className="tender-conditions__progress" role="status">
                    Saving {batch.done + 1 > batch.total ? batch.total : batch.done + 1} of {batch.total}…
                  </span>
                  <button className="ghost" onClick={() => cancelBatchByOffice.set(currentOfficeKey, true)}>
                    Stop
                  </button>
                </>
              ) : (
                <button className="primary" disabled={manualBusy || pending === 0} onClick={() => void saveAllListed()}>
                  <IconDownload /> {pending === 0 ? 'All saved' : `Save all (${pending})`}
                </button>
              )
            })()}
          </div>
        )}
        {(tResults?.length || /^\d{5,}$/.test(tq.trim())) && (
          <ul className="todo-list tender-conditions__list tender-conditions__results">
            {/^\d{5,}$/.test(tq.trim()) && !(tResults ?? []).some((r) => (r[1] ?? '').trim() === tq.trim()) && (
              <li className="todo-item">
                <div className="todo-body">
                  <span className="todo-text">Tender ID {tq.trim()}</span>
                  <span className="todo-dates">Fetch this tender's conditions straight from the portal</span>
                </div>
                <button className="ghost" disabled={manualBusy || (inOffice(tq.trim()) && !!savedConditions[tq.trim()]?.conditions.length)} onClick={() => void saveManualConditions(tq.trim())}>
                  {inOffice(tq.trim()) && savedConditions[tq.trim()]?.conditions.length ? 'Saved' : 'Fetch & save'}
                </button>
              </li>
            )}
            {(tResults ?? []).map((r, idx) => {
              const id = (r[1] ?? '').trim()
              const saved = inOffice(id) && !!savedConditions[id]?.conditions.length
              return (
                <li key={id || r[2]} className="todo-item">
                  <span className="list-serial" aria-label={`Number ${idx + 1}`}>
                    {idx + 1}
                  </span>
                  <div className="todo-body">
                    <span className="todo-text tender-conditions__work" title={r[4]}>{r[4] || '—'}</span>
                    <span className="todo-dates">ID {id || '—'} · {r[2] || '—'} · closes {r[8] || '—'}</span>
                  </div>
                  <button className="ghost" disabled={manualBusy || !/^\d+$/.test(id) || saved} onClick={() => void saveManualConditions(id, r)}>
                    {savingId === id ? 'Saving…' : saved ? 'Saved' : 'Save conditions'}
                  </button>
                  {rowResult?.id === id && (
                    <span className={`tender-conditions__rowresult ${rowResult.ok ? 'ok' : 'error'}`} role="status">
                      {rowResult.text}
                    </span>
                  )}
                </li>
              )
            })}
          </ul>
        )}
        {mineIds.length > 0 && (
          <>
            <div className="tender-conditions__subbar">
              <h3 className="tender-conditions__sub">Saved tenders ({mineIds.length})</h3>
              {confirmDelAll ? (
                <div className="saved-tender__actions">
                  <span className="saved-tender__confirm">Delete all {mineIds.length} saved tenders?</span>
                  <button
                    className="danger-ghost"
                    onClick={() => {
                      mineIds.forEach((id) => onDeleteConditions(id))
                      setConfirmDelAll(false)
                      setConfirmDelId(null)
                    }}
                  >
                    Delete all
                  </button>
                  <button className="ghost" autoFocus onClick={() => setConfirmDelAll(false)}>
                    Cancel
                  </button>
                </div>
              ) : (
                <div className="saved-tender__actions">
                  <button className="ghost" onClick={() => openShare(mineIds)} title={`Copy these saved tenders to other ${corporation || 'corporation'} users`}>
                    <IconUser /> Share with other users
                  </button>
                  <button className="danger-ghost" onClick={() => setConfirmDelAll(true)}>
                    <IconTrash /> Delete all
                  </button>
                </div>
              )}
            </div>
            <ul className="todo-list tender-conditions__list">
              {mineIds.map((id, idx) => {
                const t = savedConditions[id]
                return (
                  <li key={id} className="saved-tender">
                    <span className="list-serial" aria-label={`Number ${idx + 1}`}>
                      {idx + 1}
                    </span>
                    <div className="saved-tender__main">
                      <div className="saved-tender__top">
                        <span className="saved-tender__id">ID {id}</span>
                        {t.ecv != null && <span className="badge saved-tender__ecv">ECV ₹{t.ecv.toLocaleString('en-IN')}</span>}
                      </div>
                      <p className="saved-tender__work">{t.workName || 'Work name not available — tender no longer on the portal'}</p>
                      <p className="saved-tender__meta">
                        {t.nitNo ? `${t.nitNo} · ` : ''}
                        {t.conditions.length} conditions saved
                      </p>
                    </div>
                    <div className="saved-tender__actions">
                      {confirmDelId === id ? (
                        <>
                          <span className="saved-tender__confirm">Delete this saved tender?</span>
                          <button
                            className="danger-ghost"
                            onClick={() => {
                              onDeleteConditions(id)
                              setConfirmDelId(null)
                            }}
                          >
                            Delete
                          </button>
                          <button className="ghost" autoFocus onClick={() => setConfirmDelId(null)}>
                            Cancel
                          </button>
                        </>
                      ) : (
                        <>
                          <button
                            className="primary"
                            onClick={() => {
                              setDlMsg(null)
                              setDlId(id)
                            }}
                          >
                            <IconDownload /> Download sheet
                          </button>
                          <button className="ghost saved-tender__icon" title="View conditions" aria-label={`View conditions for tender ${id}`} onClick={() => setViewId(id)}>
                            <IconEye />
                          </button>
 {t.sharedBy ? (
                            <button className="ghost" disabled title={`These conditions were shared with you by ${t.sharedBy}`}>
                              <IconUser /> Shared by {t.sharedBy}
                            </button>
                          ) : (
                            <button className="ghost saved-tender__icon" title="Share this tender's conditions with a user" aria-label={`Share tender ${id} with a user`} onClick={() => openShare([id])}>
                              <IconUser />
                            </button>
                          )}
                          <button className="danger-ghost saved-tender__icon" title="Delete saved tender" aria-label={`Delete saved tender ${id}`} onClick={() => setConfirmDelId(id)}>
                            <IconTrash />
                          </button>
                        </>
                      )}
                    </div>
                  </li>
                )
              })}
            </ul>
          </>
        )}
        {dlMsg && (
          <div className={`notice ${dlMsg.ok ? 'ok' : 'error'}`}>
            {dlMsg.ok ? <IconTable /> : <IconWarn />} {dlMsg.text}
          </div>
        )}
      </section>
      {manualMsg && (
        <div className={`notice tender-id-notice ${manualMsg.ok ? 'ok' : 'error'}`}>
          {manualMsg.ok ? <IconTable /> : <IconWarn />} {manualMsg.text}
        </div>
      )}

      {entries.length > 0 && (
        <ul className="estimate-list">
          {entries.map((e) => (
            <li key={e.id} className="estimate-entry">
              <div className="estimate-entry-head">
                <div className="estimate-entry-titles">
                  <span className="estimate-entry-file">{e.sourceName}</span>
                  {e.workLine && <h3 className="estimate-work-name">{e.workLine.replace(/^Name of the work:\s*/, '')}</h3>}
                </div>
                <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                  <IconTrash />
                </button>
              </div>

              {e.error && (
                <p className="estimate-entry-error">
                  <IconWarn /> {e.error}
                </p>
              )}
              {e.saved && <p className="estimate-hint">{e.saved}</p>}
              {e.estimateNote && <p className="estimate-hint">{e.estimateNote}</p>}
              {e.conditionsNote && <p className="estimate-hint">{e.conditionsNote}</p>}
              {e.bidders.length > 0 && (
                <div className="estimate-body">
                  <div className="loa-manual-fields">
                    <label className="field-label">
                      NIT / Tender line
                      <input value={e.nitLine} onChange={(ev) => updateEntry(e.id, { nitLine: ev.target.value })} />
                    </label>
                    <label className="field-label">
                      Name of work line
                      <input value={e.workLine} onChange={(ev) => updateEntry(e.id, { workLine: ev.target.value })} />
                    </label>
                    <div className="loa-manual-grid">
                      <label className="field-label">
                        ECV (rupees)
                        <input value={e.ecv} onChange={(ev) => updateEntry(e.id, { ecv: ev.target.value })} />
                      </label>
                      <label className="field-label">
                        Signature
                        <input
                          value={e.signature.replace(/\n/g, ' — ')}
                          onChange={(ev) => updateEntry(e.id, { signature: ev.target.value.replace(/ — /g, '\n') })}
                        />
                      </label>
                    </div>

                    <div className="ns-group-title">
                      Tender conditions ({e.conditions.length})
                      <button className="ns-add" onClick={() => addCondition(e.id)} title="Add a condition">
                        <IconPlus /> Add
                      </button>
                    </div>
                    <div className="ns-bidders">
                      {e.conditions.map((c, i) => (
                        <div className="ns-bidder-row" key={i} style={{ gridTemplateColumns: '2rem 1fr 2rem' }}>
                          <span className="mb-serial">{i + 1}</span>
                          <textarea
                            value={c}
                            rows={Math.min(5, Math.max(1, Math.ceil(c.length / 70)))}
                            onChange={(ev) => setCondition(e.id, i, ev.target.value)}
                            placeholder="Tender condition"
                            style={{ resize: 'vertical', font: 'inherit' }}
                          />
                          <button className="ns-del" onClick={() => removeCondition(e.id, i)} title="Remove">
                            ✕
                          </button>
                        </div>
                      ))}
                    </div>

                    <div className="ns-group-title">
                      Participating bidders ({e.bidders.length})
                      <button className="ns-add" onClick={() => addBidder(e.id)} title="Add a bidder">
                        <IconPlus /> Add
                      </button>
                    </div>
                    <div className="ns-bidders">
                      {e.bidders.map((b, i) => (
                        <div className="ns-bidder-row" key={i} style={{ gridTemplateColumns: '2rem 1fr 2rem' }}>
                          <span className="mb-serial">{i + 1}</span>
                          <input value={b} onChange={(ev) => setBidder(e.id, i, ev.target.value)} placeholder="Bidder / firm name" />
                          <button className="ns-del" onClick={() => removeBidder(e.id, i)} title="Remove">
                            ✕
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="doc-sheet-footer">
                    <span className="estimate-hint">
                      Fills the header, one column per bidder and the ECV, and wires the 3AN−B bid-capacity formula.
                      Turnover (A), N, B, the eligibility rows and Remarks are left blank to fill in from each bidder's
                      documents.
                    </span>
                    <button className="primary" onClick={() => downloadOne(e)} disabled={busy !== null}>
                      <IconTable /> Download evaluation sheet
                    </button>
                  </div>
                </div>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
