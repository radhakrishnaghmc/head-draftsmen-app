import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, parseAllBidders, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import {
  deriveFields,
  workOrderPlaceholders,
  agreementPlaceholders,
  forwardingSlipPlaceholders,
  civilTenderPlaceholders,
  zonalDocsPlaceholders,
  standaloneRowFromSources,
  circleFromNit
} from '@core/workOrderAgreement'
import { corporationByName, resolveFromDirectory, entriesOf } from '../zoneCircleDirectory'
import { type Office } from '../office'
import { boqToScheduleA, buildBoqFromEstimate, extractWorkNameFromBoq } from '../boqTransform'
import { guessHeaderRow, buildTableFromGrid } from '@core/sheet'
import { extractEstimateItems, extractWorkName } from '@core/estimateExtract'
import { buildScheduleARows, rowsToScheduleAItems, metaFromWorksRow, findWorksRowByName } from '@core/scheduleA'
import { compareWorkNames, workNameMismatchMessage } from '@core/workNameMatch'
import {
  buildNoteSubmittedHtml,
  noteSubmittedFromRow,
  summarizeNonResponsiveness,
  type NoteSubmittedData,
  type NoteBidder
} from '@core/noteSubmitted'
import NoteSubmittedEditor from './NoteSubmittedEditor'
import { mismatchHint } from '../docClassify'
import type { PlaceholderMatch } from '@core/createDocument'
import type { ScheduleAMeta } from '../../electron/ipc-contract'
import type { ExcelTable } from '@core/types'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH, normalizeDocxTextboxes } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconCheck, IconClipboard, IconTable } from './Icons'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
  /**
   * Tools-workspace mode: generate the Work Order / Agreement / Schedule A from
   * ONLY the uploaded L-1, Intimation and estimate — no Works List link, no
   * Zone/Circle or same-work verification (Circle/CNO come from the NIT). The
   * default (false) is the Works-List-driven flow used on the main tab.
   */
  standalone?: boolean
  /**
   * Tools-workspace Schedule-A-only mode: show just the estimate/BOQ upload and
   * its Schedule A — no L1/Intimation uploads, no Work Order/Agreement/date UI.
   * Implies standalone. Used by the Tools "Schedule A" tile.
   */
  scheduleAOnly?: boolean
  /**
   * Tools-workspace single-document mode: ask only for the Online Intimation +
   * L1 selection form (no estimate/BOQ, no non-responsive upload) and show only
   * the one requested output — the Work Order tile passes 'workOrder', the
   * Agreement Bond tile passes 'agreement'. Implies standalone.
   */
  only?: DocKind
  /**
   * Open the estimate/BOQ picker on mount — used by the Tools "Schedule A" tile
   * (scheduleAOnly) so its single upload starts on the tile click, with no
   * second click.
   */
  autoOpen?: boolean
  /**
   * Schedule-A Tools tile only: told whether an estimate/BOQ has been picked
   * yet, so the Tools host gives the panel a full-width grid row only once
   * there's something to show.
   */
  onContent?: (hasContent: boolean) => void
  /**
   * Zone-level login (logged in with a Zone id, not a single Circle). A Zone
   * owns every circle under it, so an uploaded L-1 for any of those circles is
   * legitimate even when its work isn't in the currently-loaded Works List —
   * we don't block on a "no matching row" then, and instead fill the documents
   * from the uploaded L-1 + Intimation (Circle/CNO from the NIT), the same way
   * the standalone Tools flow does.
   */
  zoneLogin?: boolean
  /** The chosen office (Works List page) — supplies Corporation/Circle/CNO for the Forwarding Slip. */
  office?: Office
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

// The uploaded file is a BOQ/estimate, but the saved document is Schedule A —
// drop any "BOQ" wording from its name so the two don't end up side by side.
function stripBoqWord(name: string): string {
  return name
    .replace(/\bboq\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
}

/** "2026-07-15" (a date-input value) -> "15.07.2026", the templates' dd.mm.yyyy date style. */
function isoToDmy(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso)
  return m ? `${m[3]}.${m[2]}.${m[1]}` : iso
}

/** "15.07.2026" (a Works List / LOA date) -> "2026-07-15" to seed a date input. */
function dmyToIso(dmy: string): string {
  const m = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/.exec(dmy.trim())
  if (!m) return ''
  const year = m[3].length === 2 ? `20${m[3]}` : m[3]
  return `${year}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`
}

type DocKind =
  | 'workOrder'
  | 'agreement'
  | 'forwardingSlip'
  | 'civilTender'
  | 'zonalWorkOrder'
  | 'zonalConcludingAgreement'
  | 'zonalMemoEe'
type Output = DocKind | 'scheduleA' | 'note'

// The three Zone-level (SE office) documents, shown in place of the circle-level
// docs when the chosen office is a Zone with no Circle of its own.
const ZONAL_KINDS = ['zonalWorkOrder', 'zonalConcludingAgreement', 'zonalMemoEe'] as const

const DOC_LABEL: Record<Output, string> = {
  forwardingSlip: 'Forwarding Slip',
  civilTender: 'Tender Document',
  workOrder: 'Work Order',
  agreement: 'Agreement Bond',
  zonalWorkOrder: 'Work Order',
  zonalConcludingAgreement: 'Concluding Agreement',
  zonalMemoEe: 'Memo to EE',
  scheduleA: 'Schedule A',
  note: 'Note Submitted'
}

const DOC_KINDS: DocKind[] = [
  'workOrder',
  'agreement',
  'forwardingSlip',
  'civilTender',
  'zonalWorkOrder',
  'zonalConcludingAgreement',
  'zonalMemoEe'
]

/** Whether an Output is one of the docx documents (previewable/fillable), vs Schedule A / Note. */
function isDocKind(o: Output | null): o is DocKind {
  return o != null && (DOC_KINDS as string[]).includes(o)
}

/**
 * Work order and agreement — produces this work's three outputs (Work Order,
 * Agreement Bond, Schedule A) as a gallery of tiles, each showing a live
 * preview. Mirrors Give Intimation's flow: upload the Online Intimation
 * (.html) and the L-1 selection form (.pdf) and the two document tiles appear;
 * upload the technical-sanctioned estimate / BOQ and the Schedule A tile
 * appears. Clicking a tile opens the full preview with Word / PDF / Print
 * (Excel for Schedule A). The L-1 form both fills the documents and updates
 * that work's Works List row. Reuses the docx-preview + export/print pipeline
 * of Give Intimation.
 */
export default function WorkOrderAgreementTab({
  tables,
  onChange,
  standalone: standaloneProp = false,
  scheduleAOnly = false,
  only,
  autoOpen = false,
  onContent,
  zoneLogin = false,
  office
}: Props) {
  const standalone = standaloneProp || scheduleAOnly || !!only
  const table = tables[0] ?? null

  // Zone-level (SE office) mode: a Zone chosen with no Circle of its own. The
  // page then produces the three SE documents (Work Order, Concluding Agreement,
  // Memo to EE) instead of the circle-level Work Order / Agreement / Forwarding
  // Slip / Tender Document. Mirrors GiveIntimationTab's `seMode`. Never in the
  // Tools single-document (`only`) panels.
  const seMode = !only && !!office?.zone?.trim() && !office?.circle?.trim()

  const [workOrderB64, setWorkOrderB64] = useState<string | null>(null)
  const [agreementB64, setAgreementB64] = useState<string | null>(null)
  const [forwardingSlipB64, setForwardingSlipB64] = useState<string | null>(null)
  const [civilTenderB64, setCivilTenderB64] = useState<string | null>(null)
  const [workOrderLabels, setWorkOrderLabels] = useState<string[]>([])
  const [agreementLabels, setAgreementLabels] = useState<string[]>([])
  const [forwardingSlipLabels, setForwardingSlipLabels] = useState<string[]>([])
  const [civilTenderLabels, setCivilTenderLabels] = useState<string[]>([])
  // The three Zone-level (SE) templates + their placeholder labels, keyed by DocKind.
  const [zonalB64, setZonalB64] = useState<Record<string, string>>({})
  const [zonalLabels, setZonalLabels] = useState<Record<string, string[]>>({})
  const [loadError, setLoadError] = useState<string | null>(null)
  // Forwarding Slip fields the office hand-enters (not on the Works List / L-1).
  const [tsNoDate, setTsNoDate] = useState('')
  const [completionMonths, setCompletionMonths] = useState('')
  // Tender Document fields the office hand-enters.
  const [pagesOfAgreement, setPagesOfAgreement] = useState('')
  const [scheduleAItems, setScheduleAItems] = useState('')

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  // Whether the uploaded L1 form's work matched a Works List row. false means
  // it matched none, so the documents would fill from an unrelated row — the
  // tiles are gated until the work is added to the Works List. null = not yet
  // determined (no L1 uploaded, or no works to match against).
  const [worksRowMatched, setWorksRowMatched] = useState<boolean | null>(null)

  // Which output's preview is expanded to the full-size modal, if any.
  const [expanded, setExpanded] = useState<Output | null>(null)
  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [expandedPages, setExpandedPages] = useState(0)

  // Schedule A (from an uploaded technical-sanctioned estimate / BOQ).
  const [boq, setBoq] = useState<ExcelTable | null>(null)
  const [scheduleA, setScheduleA] = useState<ExcelTable | null>(null)
  const [scheduleAError, setScheduleAError] = useState<string | null>(null)
  const [scheduleABusy, setScheduleABusy] = useState(false)

  // One agreement date, shared by the Work Order and the Agreement Bond (both
  // its A.B.No line and the "…day of…" wording). Held as an ISO date-input
  // value; blank until the user sets it (via the field here or the prompt shown
  // when the Agreement Bond tile is opened).
  const [agreementDate, setAgreementDate] = useState('')
  const [datePromptOpen, setDatePromptOpen] = useState(false)
  const [promptDate, setPromptDate] = useState('')
  // Tools/standalone: when the uploaded L-1's name of work / NIT carries no Zone
  // or Circle to derive them from, the office block and Work Order number would
  // come out blank — so the same prompt collects them by hand.
  const [manualCircle, setManualCircle] = useState('')
  const [manualCno, setManualCno] = useState('')
  const [manualZone, setManualZone] = useState('')
  // Which document to open once the shared date has been entered in the prompt.
  const [pendingDoc, setPendingDoc] = useState<DocKind>('agreement')

  // Note Submitted: the full bidder table (from the L-1 PDF) and the editable
  // note data seeded from the picked row + those bidders — the same generator
  // as the Intimation tab, shown here as a fourth output tile.
  const [allBidders, setAllBidders] = useState<NoteBidder[]>([])
  const [noteData, setNoteData] = useState<NoteSubmittedData | null>(null)

  const pdfInputRef = useRef<HTMLInputElement>(null)
  const noticeInputRef = useRef<HTMLInputElement>(null)
  const nonRespInputRef = useRef<HTMLInputElement>(null)
  const woTileRef = useRef<HTMLDivElement>(null)
  const agTileRef = useRef<HTMLDivElement>(null)
  const fsTileRef = useRef<HTMLDivElement>(null)
  const ctTileRef = useRef<HTMLDivElement>(null)
  const zwoTileRef = useRef<HTMLDivElement>(null)
  const zcaTileRef = useRef<HTMLDivElement>(null)
  const zmeTileRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  // Schedule-A-only Tools tile: open the estimate/BOQ picker on mount so the
  // single upload starts on the tile click. Guarded against StrictMode's
  // double-mount.
  const autoOpened = useRef(false)
  useEffect(() => {
    if (autoOpen && scheduleAOnly && !autoOpened.current) {
      autoOpened.current = true
      void uploadBoq()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoOpen, scheduleAOnly])

  // Whether to build the documents' row purely from the uploads instead of a
  // Works List row: always in standalone (Tools) mode, and — for a Zone login —
  // whenever the uploaded L-1's work matched no Works List row (a work from
  // another circle the Zone still owns). Circle/CNO then come from the NIT and
  // the work name from the L-1, exactly as the Tools flow does.
  // The name of work (and agency/amounts) always come from the uploaded L-1 +
  // Online Intimation. When the L-1's work matched no Works List row — in ANY
  // office, not just a Zone login — there's no row to draw supporting details
  // from, so fill everything from the uploads (Circle/CNO from the NIT No). A
  // matched row is still used for the extra Works-List-only columns.
  const deriveFromUploads = standalone || worksRowMatched === false

  const selectedRow = deriveFromUploads
    ? notice || pdfEval
      ? standaloneRowFromSources(pdfEval ?? {}, notice ?? {})
      : null
    : table && table.rows.length > 0
      ? table.rows[Math.min(rowIndex, table.rows.length - 1)]
      : null

  // Load both bundled formats once, and read their placeholders.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const [woB64, agB64, fsB64, ctB64] = await Promise.all([
          api.workOrderTemplate(),
          api.agreementTemplate(),
          api.forwardingSlipTemplate(),
          api.civilTenderTemplate()
        ])
        const [woLabels, agLabels, fsLabels, ctLabels] = await Promise.all([
          api.findPlaceholdersInDocument(woB64),
          api.findPlaceholdersInDocument(agB64),
          api.findPlaceholdersInDocument(fsB64),
          api.findPlaceholdersInDocument(ctB64)
        ])
        if (cancelled) return
        setWorkOrderB64(woB64)
        setAgreementB64(agB64)
        setForwardingSlipB64(fsB64)
        setCivilTenderB64(ctB64)
        setWorkOrderLabels(woLabels)
        setAgreementLabels(agLabels)
        setForwardingSlipLabels(fsLabels)
        setCivilTenderLabels(ctLabels)
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  // Load the three Zone-level (SE) templates once, when the office is zone-level.
  useEffect(() => {
    if (!seMode) return
    let cancelled = false
    void (async () => {
      try {
        const [woB64, caB64, meB64] = await Promise.all([
          api.zonalWorkOrderTemplate(),
          api.zonalConcludingAgreementTemplate(),
          api.zonalMemoEeTemplate()
        ])
        const [woL, caL, meL] = await Promise.all([
          api.findPlaceholdersInDocument(woB64),
          api.findPlaceholdersInDocument(caB64),
          api.findPlaceholdersInDocument(meB64)
        ])
        if (cancelled) return
        setZonalB64({ zonalWorkOrder: woB64, zonalConcludingAgreement: caB64, zonalMemoEe: meB64 })
        setZonalLabels({ zonalWorkOrder: woL, zonalConcludingAgreement: caL, zonalMemoEe: meL })
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [seMode])

  // Everything the two documents print, resolved from the uploaded Online
  // Intimation + L-1 selection form + the matched Works List row.
  const fields = useMemo(() => {
    const f = deriveFields(notice ?? {}, pdfEval ?? {}, selectedRow ?? {})
    // The user-entered agreement date wins for both documents (kept identical);
    // fall back to the LOA/selection date derived from the PDF when unset.
    const dmy = agreementDate ? isoToDmy(agreementDate) : f.agreementDate
    return {
      ...f,
      agreementDate: dmy,
      workOrderDate: dmy,
      // Corporation / Circle / Zone come from the chosen office (Works List page)
      // when set, else the L-1/NIT-derived value, else the hand-entered fallback
      // (Tools mode, when the name of work carried no circle/zone).
      circle: office?.circle || f.circle || manualCircle,
      cno: office?.circleNumber || f.cno || manualCno,
      zone: office?.zone || f.zone || manualZone,
      corporation: office?.corporation ?? '',
      corporationFullName: corporationByName(office?.corporation)?.fullName ?? '',
      tsNoDate,
      completionMonths
    }
  }, [notice, pdfEval, selectedRow, agreementDate, office, tsNoDate, completionMonths, manualCircle, manualCno, manualZone])

  // Tools/standalone: the uploaded L-1 gave neither a Circle nor a Zone (and no
  // office was chosen to supply them), so the prompt must collect them by hand.
  const missingCircleZone = useMemo(() => {
    if (!standalone) return false
    const b = deriveFields(notice ?? {}, pdfEval ?? {}, selectedRow ?? {})
    const circle = (office?.circle || b.circle || '').trim()
    const zone = (office?.zone || b.zone || '').trim()
    return !circle || !zone
  }, [standalone, notice, pdfEval, selectedRow, office])

  // Opening either the Work Order or the Agreement Bond preview requires a date
  // — both documents print the same date, so prompt for it when none has been
  // set yet (seeding the picker with the LOA date if we have it) and remember
  // which document the user was opening so we return to it once the date's in.
  function openDoc(kind: DocKind) {
    // The Forwarding Slip is hand-dated (blank Date line), so it doesn't need
    // the shared agreement date — only the Work Order / Agreement do. The same
    // prompt also collects the Circle/Zone when the L-1 didn't carry them.
    const needsCircleZone = missingCircleZone && (!manualCircle.trim() || !manualZone.trim())
    if ((kind === 'workOrder' || kind === 'agreement') && (!agreementDate || needsCircleZone)) {
      setPendingDoc(kind)
      setPromptDate(dmyToIso(fields.agreementDate))
      setDatePromptOpen(true)
      return
    }
    setExpanded(kind)
  }

  // Write everything derivable about the award — from the L-1 evaluation and,
  // when it's been uploaded, the Online Intimation — into the matching Works
  // List row (Tender ID/Notice/Date, ECV, EMD, ASD, Agency, Address, Tender %,
  // Contract Amount), and select that row so the documents fill from it.
  // Re-run whenever either file changes, so a field only one of them carries
  // (e.g. the address, from the intimation) lands whichever is uploaded second.
  async function syncWorksListRow(ev: TenderEvaluation, noticeVal: IntimationNotice | null): Promise<void> {
    if (!table || !ev.nameOfWork) return
    let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
    const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
    if (nameHeader) {
      try {
        const [rowNameVectors, evalNameVectors] = await Promise.all([
          api.embedTexts(table.rows.map((r) => r[nameHeader] ?? '')),
          api.embedTexts([ev.nameOfWork])
        ])
        embeddings = { rowNameVectors, evalNameVectors }
      } catch {
        embeddings = undefined
      }
    }
    // Match only — the Works List database is updated solely from the Works List
    // page ("Update from L1"), never here. We just find the row so its supporting
    // details fill the documents and select it (an embedding match has no exact
    // name to re-derive from, so without this it would stay on row 0).
    const { matchedCount, matchedRowIndices } = updateWorksListFromEvaluations(
      table,
      [ev],
      embeddings,
      noticeVal ?? undefined
    )
    if (matchedCount > 0) {
      const idx = matchedRowIndices[0]
      if (idx != null && idx >= 0) setRowIndex(idx)
      setWorksRowMatched(true)
      setPdfStatus(`Matched "${ev.nameOfWork}" to a Works List row — its details fill the documents.`)
    } else {
      setWorksRowMatched(false)
      setPdfStatus(`Read the PDF. "${ev.nameOfWork}" isn't in the Works List — the documents fill from the uploaded L1 / Intimation.`)
    }
  }

  // The Online Intimation can be uploaded as either the portal "View
  // Intimation Notice" .html page or the printed Intimation / Letter of
  // Acceptance .pdf — both carry the agency, address, NIT No, ECV and
  // contract value, just in different formats, so parse by file type.
  async function handleNoticeFile(file: File) {
    setActionError(null)
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const parsed = isPdf
        ? parseIntimationNoticeText(await pdfToTextLines(file))
        : parseIntimationNotice(await file.text())
      setNotice(parsed)
      setNoticeName(file.name)
      // The L-1 sheet is often uploaded first; now that the intimation is here,
      // fold its address (and any agency/contract fallback) into the Works List.
      if (pdfEval) await syncWorksListRow(pdfEval, parsed)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  // The L-1 selection / evaluation PDF fills the tender fields (ECV / Tender %/
  // contract / agency) and updates that work's Works List row — matched by
  // Name of Work — so the database reflects the award in one step. The matched
  // work is auto-selected so both documents show it.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    setWorksRowMatched(null)
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)
      setAllBidders(parseAllBidders(lines))

      // Fold the award into the matching Works List row — with the intimation's
      // own fields (address, agency/contract fallback) too when it's already up.
      if (table && ev.nameOfWork) await syncWorksListRow(ev, notice)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function fillDoc(kind: DocKind): Promise<string> {
    const b64 =
      { workOrder: workOrderB64, agreement: agreementB64, forwardingSlip: forwardingSlipB64, civilTender: civilTenderB64 }[
        kind as 'workOrder' | 'agreement' | 'forwardingSlip' | 'civilTender'
      ] ?? zonalB64[kind]
    const labels =
      { workOrder: workOrderLabels, agreement: agreementLabels, forwardingSlip: forwardingSlipLabels, civilTender: civilTenderLabels }[
        kind as 'workOrder' | 'agreement' | 'forwardingSlip' | 'civilTender'
      ] ?? zonalLabels[kind] ?? []
    if (!b64) throw new Error('Format not loaded yet.')
    const values = ZONAL_KINDS.includes(kind as (typeof ZONAL_KINDS)[number])
      ? zonalDocsPlaceholders(fields, notice ?? {}, pdfEval ?? {})
      : kind === 'workOrder'
        ? workOrderPlaceholders(fields)
        : kind === 'agreement'
          ? agreementPlaceholders(fields)
          : kind === 'forwardingSlip'
            ? forwardingSlipPlaceholders(fields)
            : civilTenderPlaceholders(fields, pdfEval ?? {}, { pagesOfAgreement, scheduleAItems })
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(b64, resolved, values)
  }

  async function renderDocInto(kind: DocKind, container: HTMLElement): Promise<number> {
    const filled = await fillDoc(kind)
    container.innerHTML = ''
    await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
    normalizeDocxTextboxes(container)
    return container.querySelectorAll('section.docx').length
  }

  // Guard: the Online Intimation and the L1 selection form must describe the
  // same work (matched by NIT No, or agency name when the NIT No is absent) —
  // otherwise the documents would splice one work's agency onto another's
  // tender.
  const workMatch = useMemo(
    () => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null),
    [notice, pdfEval]
  )
  // Tools mode does no same-work verification — it uses whatever was uploaded.
  const workMismatch = !standalone && workMatch?.status === 'mismatch'

  // The uploaded L1 form's work matched no Works List row, so the selected row
  // (and everything the documents fill from it — name of work, Circle, CNO,
  // estimate…) belongs to a different work. Gate the tiles until the work is
  // added to the Works List.
  // A "no matching row" never blocks — the documents fill their name of work and
  // amounts from the uploaded L-1 / Intimation, and only borrow supporting
  // details from a row when one matched (see deriveFromUploads). It's shown as a
  // soft note so the user knows the Works-List-only columns will be blank.
  const noWorksRowMatch = !standalone && worksRowMatched === false

  // Guard: the uploaded L-1's Circle must match the Circle the documents will
  // fill from. The L-1's Circle comes from its NIT No ("…/EE/Gajularamaram
  // Circle-57/…"), with a fallback to inferring it from the L-1's work name via
  // the directory. We compare it against the office's Circle (Works List page)
  // when set, otherwise the matched Works List row's Circle — so an L-1 from
  // another circle that fuzzy-matched a row here is caught and blocked (it would
  // otherwise issue e.g. a Nizampet work's agreement under Gajularamaram).
  const l1Circle = useMemo(() => {
    const fromNit = circleFromNit(pdfEval?.noticeNo || notice?.nitNo || '').circle
    if (fromNit) return fromNit
    return resolveFromDirectory(pdfEval?.nameOfWork || '', entriesOf(office?.corporation)).circle ?? ''
  }, [pdfEval, notice, office])
  const officeCircle = (office?.circle ?? '').trim()
  const rowCircle = (selectedRow?.['Circle'] ?? '').trim()
  const targetCircle = officeCircle || rowCircle
  const sameCircleId = (a: string, b: string) => {
    const na = a.trim().toLowerCase()
    const nb = b.trim().toLowerCase()
    return !!na && !!nb && (na === nb || na.includes(nb) || nb.includes(na))
  }
  const circleMismatch = !standalone && !!l1Circle && !!targetCircle && !sameCircleId(l1Circle, targetCircle)

  // Only build the documents once BOTH the Online Intimation and the L1
  // selection form are uploaded, they belong to the same work, the L1's work is
  // actually in the Works List, and it's this office's own Circle — no tiles are
  // shown before that.
  const bothUploaded = !!notice && !!pdfEval && !workMismatch && !circleMismatch
  const templatesReady = seMode
    ? ZONAL_KINDS.every((k) => !!zonalB64[k])
    : !!workOrderB64 && !!agreementB64 && !!forwardingSlipB64 && !!civilTenderB64
  const docsReady = templatesReady && bothUploaded

  // Live thumbnails in the document tiles, refreshed whenever the filled values
  // change. The Forwarding Slip shows only on the main tab (not the Tools-mode
  // single-document panels).
  useEffect(() => {
    if (!docsReady) return
    if (seMode) {
      if (zwoTileRef.current) void renderDocInto('zonalWorkOrder', zwoTileRef.current).catch(() => {})
      if (zcaTileRef.current) void renderDocInto('zonalConcludingAgreement', zcaTileRef.current).catch(() => {})
      if (zmeTileRef.current) void renderDocInto('zonalMemoEe', zmeTileRef.current).catch(() => {})
      return
    }
    if (woTileRef.current) void renderDocInto('workOrder', woTileRef.current).catch(() => {})
    if (agTileRef.current) void renderDocInto('agreement', agTileRef.current).catch(() => {})
    if (fsTileRef.current) void renderDocInto('forwardingSlip', fsTileRef.current).catch(() => {})
    if (ctTileRef.current) void renderDocInto('civilTender', ctTileRef.current).catch(() => {})
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [docsReady, seMode, fields, pagesOfAgreement, scheduleAItems])

  // The full-size preview inside the expanded modal (documents only — Schedule
  // A renders its table as JSX below).
  useEffect(() => {
    if (!isDocKind(expanded)) return
    const container = expandedRef.current
    if (!container) return
    setActionError(null)
    void renderDocInto(expanded, container)
      .then((pages) => setExpandedPages(pages))
      .catch((e) => setActionError(e instanceof Error ? e.message : String(e)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [expanded, fields, pagesOfAgreement, scheduleAItems])

  function docName(kind: DocKind): string {
    const agency = fields.agencyName ? ` - ${fields.agencyName}` : ''
    return `${DOC_LABEL[kind]}${agency}`
  }

  async function download(kind: DocKind, formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillDoc(kind)
      const res = await api.exportCreatedDocument(filled, docName(kind), formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function print(kind: DocKind) {
    setBusy('print')
    setActionError(null)
    try {
      const filled = await fillDoc(kind)
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      normalizeDocxTextboxes(container)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // --- Note Submitted (same generator as the Intimation tab) ---

  // Seed the note from the picked row and, once uploaded, the L-1 PDF's full
  // bidder table. Re-seeds on row / PDF change.
  useEffect(() => {
    if (!selectedRow) {
      setNoteData(null)
      return
    }
    const seed = noteSubmittedFromRow(selectedRow, table?.rows[0]?.['Circle'] ?? '')
    if (allBidders.length > 0) {
      seed.bidders = allBidders
      const l1 = allBidders[0]
      seed.l1Name = l1.name
      seed.l1PctText = l1.pct
      seed.l1Tcv = l1.tcv
    }
    // The Intimation date is the L1 sheet's server date (its bottom-right
    // "Server Time: …"), so prefer it over the (often blank here) Works List
    // Intimation Date column.
    if (pdfEval?.serverDate) seed.intimationDate = pdfEval.serverDate
    setNoteData(seed)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowIndex, table, allBidders, pdfEval])

  const notePreviewHtml = useMemo(() => (noteData ? buildNoteSubmittedHtml(noteData) : ''), [noteData])
  // Same gate as the Work Order / Agreement tiles: don't build the Note
  // Submitted while the Online Intimation and L1 selection form disagree on the
  // work/agency, or the L1's work isn't in the Works List (the note is seeded
  // from the selected row) — otherwise it would show the wrong work/agency.
  const noteReady = !!noteData && !!pdfEval && !workMismatch

  // Optional non-responsiveness statement — pre-fills the note's rejection line.
  async function handleNonRespFile(file: File) {
    setActionError(null)
    try {
      const lines = await pdfToTextLines(file)
      const summary = summarizeNonResponsiveness(lines)
      setNoteData((prev) => (prev ? { ...prev, qualificationNote: summary } : prev))
      setPdfStatus(
        summary
          ? `Non-responsiveness read from ${file.name} — check the rejection line in the Note Submitted editor.`
          : `Read ${file.name}, but found no rejected bidders — add the rejection line in the Note Submitted editor if needed.`
      )
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  async function downloadNote(formats: ('docx' | 'pdf')[]) {
    if (!noteData) return
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const b64 = await api.noteSubmittedDocx(notePreviewHtml)
      const name = `Note Submitted${noteData.workName ? ` - ${noteData.workName}` : ''}`
      const res = await api.exportCreatedDocument(b64, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function printNote() {
    if (!noteData) return
    setBusy('print')
    setActionError(null)
    try {
      await api.printCreatedDocument(notePreviewHtml)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // --- Schedule A (uploaded estimate / BOQ) ---

  async function uploadBoq() {
    setScheduleAError(null)
    const g = await api.pickEstimateGrid()
    if (!g) return
    const headerRow = guessHeaderRow(g.grid)
    const t = buildTableFromGrid(g.grid, headerRow, { id: `boq-${Date.now()}`, name: g.name, path: g.path })

    // The work name this file describes — from an estimate's title block or a
    // flat BOQ's header — used both to guard against mixing works and to hop
    // the Works List selection to the matching row.
    const detected = extractWorkName(g.grid, headerRow) ?? extractWorkNameFromBoq(t)

    // The work this workspace is currently building for: the L-1 selection
    // form's Name of Work when it's been uploaded, else the picked Works List
    // row's name. When both it and the uploaded file name a work, they must be
    // the same work — otherwise this file's Schedule A would silently belong to
    // a different work than the Work Order / Agreement. Block the upload and
    // ask for the same work's details.
    const expected = pdfEval?.nameOfWork?.trim() || (selectedRow?.['Name of the work'] ?? '').trim()
    // Tools mode does no work-name verification — it Schedule-A's exactly what
    // was uploaded, for any circle/zone.
    if (!standalone && detected && expected) {
      let embeddings: { aVector: number[]; bVector: number[] } | undefined
      try {
        const [aVector, bVector] = await api.embedTexts([detected, expected])
        embeddings = { aVector, bVector }
      } catch {
        embeddings = undefined
      }
      if (compareWorkNames(detected, expected, embeddings).status === 'mismatch') {
        setBoq(null)
        setScheduleA(null)
        setScheduleAError(workNameMismatchMessage(detected, expected))
        return
      }
    }

    setBoq(t)
    try {
      // A detailed CMC/departmental estimate (multi-row No.s/L/B/D measurement
      // layout) can't be read row-for-row like a flat BOQ — run the estimate
      // extractor first, which resolves each item's real quantity/rate/unit,
      // and only fall back to the flat-BOQ column mapping when it finds no
      // items (i.e. the file really is a plain BOQ). Both yield a Schedule A.
      const items = extractEstimateItems(g.grid, headerRow)
      setScheduleA(items.length > 0 ? boqToScheduleA(buildBoqFromEstimate(items)) : boqToScheduleA(t))

      // If the source names its own work and it matches a Works List row, hop
      // the selection there so the whole tab (documents + Schedule A) lines up.
      if (detected && table) {
        const match = findWorksRowByName(table, detected)
        if (match) {
          const idx = table.rows.indexOf(match.row)
          if (idx >= 0) setRowIndex(idx)
        }
      }
    } catch (e) {
      setScheduleA(null)
      const hint = mismatchHint(t.headers, 'boq')
      setScheduleAError((e instanceof Error ? e.message : String(e)) + (hint ? ` ${hint}` : ''))
    }
  }

  const scheduleAMeta: ScheduleAMeta | undefined = useMemo(() => {
    if (!selectedRow) return undefined
    const base = metaFromWorksRow(selectedRow)
    // The tender % (and contractor / work name) are frequently known only from
    // the uploaded L-1 selection form before the Works List row itself carries
    // them — prefer the resolved `fields` value so Schedule A's "Tender Quoted
    // %" and "Less: (…)% Less" fill from what was uploaded rather than staying
    // blank against a not-yet-updated (or mis-selected) row.
    return {
      ...base,
      // Name of work comes from the uploaded L-1 (via `fields`), not the row.
      nameOfWork: fields.nameOfWork || base.nameOfWork?.trim() || '',
      contractorName: base.contractorName?.trim() || fields.agencyName,
      tenderPercentage: fields.tenderPercent || base.tenderPercentage
    }
  }, [selectedRow, fields])

  const scheduleAPreview = useMemo(
    () => (scheduleA ? buildScheduleARows(rowsToScheduleAItems(scheduleA), scheduleAMeta) : null),
    [scheduleA, scheduleAMeta]
  )

  // Schedule-A Tools tile: report whether an estimate/BOQ has been picked yet —
  // before paint, so the panel gets its full-width row without a one-frame flash.
  useLayoutEffect(() => {
    if (scheduleAOnly) onContent?.(!!boq || !!scheduleAPreview || !!scheduleAError)
  }, [scheduleAOnly, boq, scheduleAPreview, scheduleAError, onContent])

  async function downloadScheduleA() {
    if (!scheduleA) return
    setScheduleABusy(true)
    setScheduleAError(null)
    try {
      const base = boq ? stripBoqWord(stripExt(boq.name)) : ''
      const suggestedName = base ? `${base} Schedule A` : 'Schedule A'
      const savedPath = await api.exportScheduleA(scheduleA, suggestedName, scheduleAMeta)
      if (savedPath) setActionSaved(savedPath ? `Saved: ${savedPath}` : 'Cancelled.')
    } catch (e) {
      setScheduleAError(e instanceof Error ? e.message : String(e))
    } finally {
      setScheduleABusy(false)
    }
  }

  function scheduleATableHtml(): string {
    if (!scheduleAPreview) return ''
    const rows = scheduleAPreview
      .map((row) => {
        const cells = Array.from({ length: 6 }, (_, ci) => String(row[ci] ?? ''))
        const wide = cells.slice(1).every((c) => c === '')
        if (wide) return `<tr><td colspan="6" style="font-weight:600">${cells[0]}</td></tr>`
        return `<tr>${cells.map((c) => `<td>${c}</td>`).join('')}</tr>`
      })
      .join('')
    return `<table style="border-collapse:collapse;width:100%;font-family:'Times New Roman',serif;font-size:12px">
      <style>td{border:1px solid #333;padding:5px 7px;vertical-align:top}</style>${rows}</table>`
  }

  async function printScheduleA() {
    setScheduleABusy(true)
    setScheduleAError(null)
    try {
      await api.printCreatedDocument(scheduleATableHtml())
    } catch (e) {
      setScheduleAError(e instanceof Error ? e.message : String(e))
    } finally {
      setScheduleABusy(false)
    }
  }

  // Tools mode has no Works List by design, so the "add works first" gate and
  // the Schedule-A-disabled state don't apply there.
  // A Zone login can build straight from the uploaded L-1 + Intimation, so an
  // empty (or other-circle) Works List doesn't block it the way it does a
  // Circle login, whose documents must fill from a Works List row.
  const noWorks = !standalone && !zoneLogin && (!table || table.rows.length === 0)
  const anyOutput = docsReady || !!scheduleAPreview || noteReady

  // Schedule-A Tools tile (scheduleAOnly + autoOpen): show nothing until an
  // estimate/BOQ is actually picked (or reading it fails) — the mount effect
  // fires the folder, so clicking the tile opens it directly with no
  // placeholder panel.
  if (autoOpen && scheduleAOnly && !boq && !scheduleAPreview && !scheduleAError) return null

  return (
    <div className={only ? 'wo-compact' : 'card'}>
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className={only ? 'wo-compact-body' : 'empty'}>
        {!only && <IconClipboard />}
        {!only && (
          <p>
            {scheduleAOnly ? (
              <>
                Upload the technical-sanctioned <strong>estimate / BOQ</strong> to generate its Schedule&nbsp;A. The
                output appears below as a tile — click it to preview full size and print / save it.
              </>
            ) : (
              <>
                Upload the <strong>Online Intimation</strong> and the <strong>L1 selection form</strong> to build the
                Work Order and Agreement Bond.{!standalone && " The L1 form also updates that work's row in the Works List."}{' '}
                Upload the technical-sanctioned <strong>estimate / BOQ</strong> to also generate this work's
                Schedule&nbsp;A. Each output appears below as a tile — click one to preview it full size and print / save
                it.
              </>
            )}
          </p>
        )}
        <div className={only ? 'wo-compact-actions' : 'boq-actions boq-actions--grid'}>
          {!scheduleAOnly && (
            <button className="primary upload-btn" onClick={() => noticeInputRef.current?.click()} disabled={!templatesReady}>
              <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
            </button>
          )}
          {!scheduleAOnly && (
            <button className="primary upload-btn" onClick={() => pdfInputRef.current?.click()} disabled={!templatesReady || busy === 'pdf'}>
              <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
            </button>
          )}
          {!only && (
            <button className="primary upload-btn" onClick={uploadBoq} disabled={noWorks}>
              <IconTable /> {boq ? 'Change estimate / BOQ' : 'Upload estimate/BOQ to get schedule A'}
            </button>
          )}
          {!scheduleAOnly && !only && (
            <button className="primary upload-btn" onClick={() => nonRespInputRef.current?.click()} disabled={!pdfEval}>
              <IconFolder /> Upload Non-responsive form
            </button>
          )}
          <input
            ref={noticeInputRef}
            type="file"
            accept=".html,.htm,text/html,.pdf,application/pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNoticeFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={pdfInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handlePdfFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={nonRespInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNonRespFile(file)
              e.target.value = ''
            }}
          />
        </div>
        {!scheduleAOnly && !only && (
          <div className="wo-date-row">
            <label className="wo-date-field">
              <span>Agreement date</span>
              <input type="date" value={agreementDate} onChange={(e) => setAgreementDate(e.target.value)} />
            </label>
            <span className="estimate-hint">Same date fills the Work Order and the Agreement Bond.</span>
          </div>
        )}
        {!scheduleAOnly && !only && (
          <div className="wo-date-row">
            <label className="wo-date-field">
              <span>Technical Sanction No &amp; Date</span>
              <input
                type="text"
                placeholder="11/26-27, Dt: 29.05.2026"
                value={tsNoDate}
                onChange={(e) => setTsNoDate(e.target.value)}
              />
            </label>
            <label className="wo-date-field">
              <span>Period of completion (months)</span>
              <input
                type="text"
                placeholder="02"
                value={completionMonths}
                onChange={(e) => setCompletionMonths(e.target.value)}
              />
            </label>
            <label className="wo-date-field">
              <span>No. of pages of Agreement</span>
              <input
                type="text"
                placeholder="12"
                value={pagesOfAgreement}
                onChange={(e) => setPagesOfAgreement(e.target.value)}
              />
            </label>
            <label className="wo-date-field">
              <span>No. of items in Schedule ‘A’</span>
              <input
                type="text"
                placeholder="8"
                value={scheduleAItems}
                onChange={(e) => setScheduleAItems(e.target.value)}
              />
            </label>
            <span className="estimate-hint">For the Forwarding Slip &amp; Tender Document.</span>
          </div>
        )}
        {!only && noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
        {!only && pdfName && <p className="estimate-hint">Tender details read from {pdfName}</p>}
        {!only && pdfStatus && (
          <div className="notice ok">
            <IconCheck /> {pdfStatus}
          </div>
        )}
      </div>

      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}
      {scheduleAError && (
        <div className="notice error">
          <IconWarn /> {scheduleAError}
        </div>
      )}
      {workMismatch && workMatch && (
        <div className="notice error">
          <IconWarn /> {sameWorkMismatchMessage(workMatch)}
        </div>
      )}
      {circleMismatch && (
        <div className="notice error">
          <IconWarn /> The uploaded L1 / Intimation is for <strong>{l1Circle}</strong> circle, but the documents would be
          issued under <strong>{targetCircle}</strong> circle. Documents are blocked so a work isn’t issued under the wrong
          circle — select the <strong>{l1Circle}</strong> office on the Works List page (and load its Works List), or upload
          the correct circle’s L1.
        </div>
      )}
      {noWorksRowMatch && !circleMismatch && pdfEval && (
        <div className="notice">
          <IconWarn /> “{pdfEval.nameOfWork}” isn’t in your Works List, so the documents fill the name of work, agency and
          amounts from the uploaded L1 / Intimation (Circle/CNO from the NIT No). Add it to the Works List if you want its
          extra columns (Wincode, estimate, TS No/date, …) filled in automatically.
        </div>
      )}
      {actionSaved && !expanded && (
        <div className="notice ok">
          <IconCheck /> {actionSaved}
        </div>
      )}

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the outputs are filled from a work's row.</div>
      ) : anyOutput ? (
        <div className="wo-tiles">
          {docsReady && seMode && (
            <>
              <button className="wo-tile" onClick={() => openDoc('zonalWorkOrder')}>
                <div className="wo-tile-preview">
                  <div ref={zwoTileRef} className="wo-tile-doc" />
                  <span className="wo-tile-open">Click to preview</span>
                </div>
                <div className="wo-tile-foot">{DOC_LABEL.zonalWorkOrder}</div>
              </button>
              <button className="wo-tile" onClick={() => openDoc('zonalConcludingAgreement')}>
                <div className="wo-tile-preview">
                  <div ref={zcaTileRef} className="wo-tile-doc" />
                  <span className="wo-tile-open">Click to preview</span>
                </div>
                <div className="wo-tile-foot">{DOC_LABEL.zonalConcludingAgreement}</div>
              </button>
              <button className="wo-tile" onClick={() => openDoc('zonalMemoEe')}>
                <div className="wo-tile-preview">
                  <div ref={zmeTileRef} className="wo-tile-doc" />
                  <span className="wo-tile-open">Click to preview</span>
                </div>
                <div className="wo-tile-foot">{DOC_LABEL.zonalMemoEe}</div>
              </button>
            </>
          )}
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('forwardingSlip')}>
              <div className="wo-tile-preview">
                <div ref={fsTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.forwardingSlip}</div>
            </button>
          )}
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('civilTender')}>
              <div className="wo-tile-preview">
                <div ref={ctTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.civilTender}</div>
            </button>
          )}
          {docsReady && !seMode && (!only || only === 'agreement') && (
            <button className="wo-tile" onClick={() => openDoc('agreement')}>
              <div className="wo-tile-preview">
                <div ref={agTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.agreement}</div>
            </button>
          )}
          {scheduleAPreview && !only && (
            <button className="wo-tile" onClick={() => setExpanded('scheduleA')}>
              <div className="wo-tile-preview sched">
                <table className="sa-preview-table">
                  <tbody>
                    {scheduleAPreview.slice(0, 14).map((row, ri) => {
                      const cells = Array.from({ length: 6 }, (_, ci) => row[ci] ?? '')
                      const wide = cells.slice(1).every((c) => c === '')
                      return (
                        <tr key={ri}>
                          {wide ? (
                            <td colSpan={6} className="sa-wide">
                              {String(cells[0])}
                            </td>
                          ) : (
                            cells.map((c, ci) => <td key={ci}>{String(c)}</td>)
                          )}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.scheduleA}</div>
            </button>
          )}
          {docsReady && !seMode && (!only || only === 'workOrder') && (
            <button className="wo-tile" onClick={() => openDoc('workOrder')}>
              <div className="wo-tile-preview">
                <div ref={woTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.workOrder}</div>
            </button>
          )}
          {noteReady && !only && (
            <button className="wo-tile" onClick={() => setExpanded('note')}>
              <div className="wo-tile-preview">
                <div className="wo-tile-doc ns-tile-doc" dangerouslySetInnerHTML={{ __html: notePreviewHtml }} />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.note}</div>
            </button>
          )}
        </div>
      ) : null}

      {datePromptOpen && (
        <div className="wo-modal-overlay" onClick={() => setDatePromptOpen(false)}>
          <div className="wo-modal wo-date-modal" onClick={(e) => e.stopPropagation()}>
            <div className="wo-modal-head">
              <span className="wo-modal-title">
                {missingCircleZone ? 'Enter the Circle, Zone and agreement date' : 'Enter the agreement date'}
              </span>
              <button className="wo-modal-close" onClick={() => setDatePromptOpen(false)} title="Close" aria-label="Close">
                ×
              </button>
            </div>
            <div className="wo-modal-body wo-date-body">
              {missingCircleZone && (
                <>
                  <p>
                    The uploaded L-1’s name of work doesn’t name a Circle or Zone, so the office block and the Work Order
                    number would be blank. Enter them here — they fill both documents.
                  </p>
                  <label className="wo-date-field">
                    <span>Circle</span>
                    <input
                      type="text"
                      placeholder="e.g. Gajularamaram"
                      value={manualCircle}
                      onChange={(e) => setManualCircle(e.target.value)}
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Circle number</span>
                    <input
                      type="text"
                      placeholder="e.g. 57"
                      value={manualCno}
                      onChange={(e) => setManualCno(e.target.value)}
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Zone</span>
                    <input
                      type="text"
                      placeholder="e.g. Quthbullapur"
                      value={manualZone}
                      onChange={(e) => setManualZone(e.target.value)}
                    />
                  </label>
                </>
              )}
              <p>
                This date fills the Agreement Bond — the <strong>A.B.No line</strong> (as dd.mm.yyyy) and the{' '}
                <strong>“…day of…”</strong> wording (in words) — and the Work Order date.
              </p>
              <input
                type="date"
                value={promptDate}
                onChange={(e) => setPromptDate(e.target.value)}
                // eslint-disable-next-line jsx-a11y/no-autofocus
                autoFocus
              />
            </div>
            <div className="wo-modal-foot">
              <button
                className="primary"
                disabled={!promptDate || (missingCircleZone && (!manualCircle.trim() || !manualZone.trim()))}
                onClick={() => {
                  setAgreementDate(promptDate)
                  setDatePromptOpen(false)
                  setExpanded(pendingDoc)
                }}
              >
                Continue to preview
              </button>
            </div>
          </div>
        </div>
      )}

      {expanded && (
        <div className="wo-modal-overlay" onClick={() => setExpanded(null)}>
          <div className={`wo-modal ${expanded === 'note' ? 'wide' : ''}`} onClick={(e) => e.stopPropagation()}>
            <div className="wo-modal-head">
              <span className="wo-modal-title">
                {DOC_LABEL[expanded]}
                {isDocKind(expanded) && expandedPages > 1 ? ` — ${expandedPages} pages` : ''}
              </span>
              <button className="wo-modal-close" onClick={() => setExpanded(null)} title="Close" aria-label="Close">
                ×
              </button>
            </div>

            <div className="wo-modal-body">
              {expanded === 'scheduleA' ? (
                <div className="sheet-wrap">
                  <table className="sa-preview-table">
                    <tbody>
                      {scheduleAPreview?.map((row, ri) => {
                        const cells = Array.from({ length: 6 }, (_, ci) => row[ci] ?? '')
                        const wide = cells.slice(1).every((c) => c === '')
                        return (
                          <tr key={ri}>
                            {wide ? (
                              <td colSpan={6} className="sa-wide">
                                {String(cells[0])}
                              </td>
                            ) : (
                              cells.map((c, ci) => <td key={ci}>{String(c)}</td>)
                            )}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              ) : expanded === 'note' ? (
                noteData && <NoteSubmittedEditor data={noteData} onChange={setNoteData} />
              ) : (
                <div ref={expandedRef} className="intimation-docx-preview" />
              )}
            </div>

            <div className="wo-modal-foot">
              {actionError && (
                <div className="notice error" style={{ marginRight: 'auto' }}>
                  <IconWarn /> {actionError}
                </div>
              )}
              {expanded === 'scheduleA' ? (
                <>
                  <button className="primary" onClick={downloadScheduleA} disabled={scheduleABusy}>
                    <IconDownload /> {scheduleABusy ? 'Working…' : 'Excel'}
                  </button>
                  <button className="primary" onClick={printScheduleA} disabled={scheduleABusy}>
                    <IconPrint /> Print
                  </button>
                </>
              ) : expanded === 'note' ? (
                <>
                  <button className="primary" onClick={() => downloadNote(['docx'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                  </button>
                  <button className="primary" onClick={() => downloadNote(['pdf'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                  </button>
                  <button className="primary" onClick={printNote} disabled={busy !== null}>
                    <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                  </button>
                </>
              ) : (
                <>
                  <button className="primary" onClick={() => download(expanded, ['docx'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                  </button>
                  <button className="primary" onClick={() => download(expanded, ['pdf'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                  </button>
                  <button className="primary" onClick={() => print(expanded)} disabled={busy !== null}>
                    <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                  </button>
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
