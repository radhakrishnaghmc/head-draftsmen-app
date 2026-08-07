import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, parseAllBidders, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import {
  deriveFields,
  workOrderPlaceholders,
  fileBackerPlaceholders,
  agreementPlaceholders,
  qccIntimationPlaceholders,
  forwardingSlipPlaceholders,
  civilTenderPlaceholders,
  zonalDocsPlaceholders,
  standaloneRowFromSources,
  circleFromNit,
  indianFinancialYear
} from '@core/workOrderAgreement'
import type { WorkOrderAgreementFields } from '@core/workOrderAgreement'
import { corporationByName, resolveFromDirectory, entriesOf } from '../zoneCircleDirectory'
import { type Office } from '../office'
import { boqToScheduleA, buildBoqFromEstimate, extractWorkNameFromBoq } from '../boqTransform'
import { guessHeaderRow, buildTableFromGrid } from '@core/sheet'
import { extractEstimateItems, extractWorkName } from '@core/estimateExtract'
import { buildScheduleARows, rowsToScheduleAItems, metaFromWorksRow, findWorksRowByName } from '@core/scheduleA'
import { indianDigitGroups } from '@core/worksAmounts'
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
import type { ScheduleAMeta, AgreementBundleFile } from '../../electron/ipc-contract'
import type { ExcelTable } from '@core/types'
import { pdfToTextLines, pdfToPositionedLines } from '../pdfToText'
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

/**
 * The saved Note Submitted file name: "Note Submitted - <work name>", with the
 * work-name part capped at 20 characters (a long work name would otherwise make
 * an unwieldy file name — and some file systems cap the total length).
 */
function noteSubmittedFileName(workName: string | undefined): string {
  const name = (workName ?? '').trim().slice(0, 20).trim()
  return `Note Submitted${name ? ` - ${name}` : ''}`
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
  | 'fileBacker'
  | 'agreement'
  | 'qccIntimation'
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
  fileBacker: 'File Backer',
  forwardingSlip: 'Forwarding Slip',
  civilTender: 'Tender Document',
  workOrder: 'Work Order',
  agreement: 'Agreement Bond',
  qccIntimation: 'QCC Intimation',
  zonalWorkOrder: 'Work Order',
  zonalConcludingAgreement: 'Concluding Agreement',
  zonalMemoEe: 'Memo to EE',
  scheduleA: 'Schedule A',
  note: 'Note Submitted'
}

const DOC_KINDS: DocKind[] = [
  'workOrder',
  'fileBacker',
  'agreement',
  'qccIntimation',
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

// Tools "Fill details manually" mode: every Work Order / Agreement field the
// office types in by hand instead of uploading the L1 + Intimation. Dates are
// held as the date-input's ISO value and converted to dd.mm.yyyy for the docs.
interface ManualEntry {
  nameOfWork: string
  agencyName: string
  address: string
  phone: string
  zone: string
  circle: string
  cno: string
  wincode: string
  corporation: string
  financialYear: string
  estimateLakhs: string
  ecvRupees: string
  tenderPercent: string
  contractRupees: string
  agreementDate: string
  adminSanctionDate: string
  completionMonths: string
  reservation: string
}
const MANUAL_DEFAULTS: ManualEntry = {
  nameOfWork: '',
  agencyName: '',
  address: '',
  phone: '',
  zone: '',
  circle: '',
  cno: '',
  wincode: '',
  corporation: 'CMC',
  financialYear: indianFinancialYear(),
  estimateLakhs: '',
  ecvRupees: '',
  tenderPercent: '',
  contractRupees: '',
  agreementDate: '',
  adminSanctionDate: '',
  completionMonths: '',
  reservation: ''
}

// The fields the two documents fill, built straight from the hand-entered values.
function fieldsFromManual(m: ManualEntry): WorkOrderAgreementFields {
  const dmy = m.agreementDate ? isoToDmy(m.agreementDate) : ''
  return {
    circle: m.circle.trim(),
    cno: m.cno.trim(),
    zone: m.zone.trim(),
    nameOfWork: m.nameOfWork.trim(),
    agencyName: m.agencyName.trim(),
    address: m.address.trim(),
    phone: m.phone.trim(),
    wincode: m.wincode.trim(),
    financialYear: m.financialYear.trim() || indianFinancialYear(),
    estimateLakhs: m.estimateLakhs.trim(),
    ecvRupees: m.ecvRupees.trim(),
    tenderPercent: m.tenderPercent.trim(),
    contractRupees: m.contractRupees.trim(),
    workOrderDate: dmy,
    agreementDate: dmy,
    adminSanctionDate: m.adminSanctionDate ? isoToDmy(m.adminSanctionDate) : '',
    corporation: m.corporation.trim(),
    corporationFullName: corporationByName(m.corporation.trim())?.fullName ?? '',
    tsNoDate: '',
    completionMonths: m.completionMonths.trim(),
    reservation: m.reservation.trim()
  }
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
  const [fileBackerB64, setFileBackerB64] = useState<string | null>(null)
  const [agreementB64, setAgreementB64] = useState<string | null>(null)
  const [qccIntimationB64, setQccIntimationB64] = useState<string | null>(null)
  const [forwardingSlipB64, setForwardingSlipB64] = useState<string | null>(null)
  const [civilTenderB64, setCivilTenderB64] = useState<string | null>(null)
  const [workOrderLabels, setWorkOrderLabels] = useState<string[]>([])
  const [fileBackerLabels, setFileBackerLabels] = useState<string[]>([])
  const [agreementLabels, setAgreementLabels] = useState<string[]>([])
  const [qccIntimationLabels, setQccIntimationLabels] = useState<string[]>([])
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
  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf' | 'bundle'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [expandedPages, setExpandedPages] = useState(0)

  // Schedule A (from an uploaded technical-sanctioned estimate / BOQ).
  const [boq, setBoq] = useState<ExcelTable | null>(null)
  // The work name read from the uploaded estimate/BOQ itself — drives the
  // Tools (scheduleAOnly) Schedule A's name of work and its Works-List lookup.
  const [detectedWorkName, setDetectedWorkName] = useState<string | null>(null)
  // Whether the uploaded file was a detailed estimate (true) or a flat BOQ
  // (false). A BOQ's total is the ECV alone; an estimate fills both figures.
  const [uploadedIsEstimate, setUploadedIsEstimate] = useState(false)
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
  // Tools "Fill details manually": skip the L1/Intimation uploads and type every
  // field in by hand. Only offered in the single-document Tools panels (`only`).
  const [manualMode, setManualMode] = useState(false)
  const [manual, setManual] = useState<ManualEntry>(MANUAL_DEFAULTS)
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
  const fbTileRef = useRef<HTMLDivElement>(null)
  const agTileRef = useRef<HTMLDivElement>(null)
  const qccIntTileRef = useRef<HTMLDivElement>(null)
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
        const [woB64, fbB64, agB64, qiB64, fsB64, ctB64] = await Promise.all([
          api.workOrderTemplate(),
          api.fileBackerTemplate(),
          api.agreementTemplate(),
          api.qccIntimationTemplate(),
          api.forwardingSlipTemplate(),
          api.civilTenderTemplate()
        ])
        const [woLabels, fbLabels, agLabels, qiLabels, fsLabels, ctLabels] = await Promise.all([
          api.findPlaceholdersInDocument(woB64),
          api.findPlaceholdersInDocument(fbB64),
          api.findPlaceholdersInDocument(agB64),
          api.findPlaceholdersInDocument(qiB64),
          api.findPlaceholdersInDocument(fsB64),
          api.findPlaceholdersInDocument(ctB64)
        ])
        if (cancelled) return
        setWorkOrderB64(woB64)
        setFileBackerB64(fbB64)
        setAgreementB64(agB64)
        setQccIntimationB64(qiB64)
        setForwardingSlipB64(fsB64)
        setCivilTenderB64(ctB64)
        setWorkOrderLabels(woLabels)
        setFileBackerLabels(fbLabels)
        setAgreementLabels(agLabels)
        setQccIntimationLabels(qiLabels)
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
    // "Fill details manually": build from the hand-entered values, but let a
    // chosen office supply Circle / Zone / Corporation so they aren't re-asked.
    if (manualMode) {
      const f = fieldsFromManual(manual)
      const corp = office?.corporation || f.corporation
      return {
        ...f,
        circle: office?.circle || f.circle,
        cno: office?.circleNumber || f.cno,
        zone: office?.zone || f.zone,
        corporation: corp,
        corporationFullName: corporationByName(corp)?.fullName ?? f.corporationFullName
      }
    }
    const f = deriveFields(notice ?? {}, pdfEval ?? {}, selectedRow ?? {})
    // The user-entered agreement date fills both documents (kept identical). When
    // it's left unset, the date stays BLANK — the Work Order and Agreement print
    // a "Dt:" line with a ruled blank to hand-write (see DATE_BLANK), rather than
    // silently stamping the tender-notice date. "Download all" without a chosen
    // date therefore yields date-less documents ready to be dated by hand.
    const dmy = agreementDate ? isoToDmy(agreementDate) : ''
    return {
      ...f,
      agreementDate: dmy,
      workOrderDate: dmy,
      // Corporation / Circle / Zone come from the chosen office (Works List page)
      // when set, else the L-1/NIT-derived value, else the hand-entered fallback.
      // In the Tools single-document panels (`only`) the office is ignored for
      // the *uploaded* flow — those are "any circle/zone" and take the L-1's own
      // circle — so an L-1 for a different circle isn't relabelled to the office.
      circle: (only ? '' : office?.circle) || f.circle || manualCircle,
      cno: (only ? '' : office?.circleNumber) || f.cno || manualCno,
      zone: (only ? '' : office?.zone) || f.zone || manualZone,
      corporation: (only ? '' : office?.corporation) ?? '',
      corporationFullName: (only ? '' : corporationByName(office?.corporation)?.fullName) ?? '',
      tsNoDate,
      completionMonths
    }
  }, [manualMode, manual, notice, pdfEval, selectedRow, agreementDate, office, tsNoDate, completionMonths, manualCircle, manualCno, manualZone])

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
    // In manual mode the date (and everything else) is already in the form, so no prompt.
    if (!manualMode && (kind === 'workOrder' || kind === 'agreement') && (!agreementDate || needsCircleZone)) {
      setPendingDoc(kind)
      // Seed the picker with the tender-notice date as a convenient default (the
      // document date now stays blank until explicitly chosen, so fields.agreementDate
      // is empty here) — the user can accept or change it.
      setPromptDate(dmyToIso(pdfEval?.noticeDate ?? ''))
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

  async function fillDoc(kind: DocKind, opts?: { blankDate?: boolean }): Promise<string> {
    const b64 =
      { workOrder: workOrderB64, fileBacker: fileBackerB64, agreement: agreementB64, qccIntimation: qccIntimationB64, forwardingSlip: forwardingSlipB64, civilTender: civilTenderB64 }[
        kind as 'workOrder' | 'fileBacker' | 'agreement' | 'qccIntimation' | 'forwardingSlip' | 'civilTender'
      ] ?? zonalB64[kind]
    const labels =
      { workOrder: workOrderLabels, fileBacker: fileBackerLabels, agreement: agreementLabels, qccIntimation: qccIntimationLabels, forwardingSlip: forwardingSlipLabels, civilTender: civilTenderLabels }[
        kind as 'workOrder' | 'fileBacker' | 'agreement' | 'qccIntimation' | 'forwardingSlip' | 'civilTender'
      ] ?? zonalLabels[kind] ?? []
    if (!b64) throw new Error('Format not loaded yet.')
    // "Download all" asks for the Work Order / Agreement Bond date to be blank
    // (hand-written on the printed copy), regardless of any date picked for a
    // preview — so force the date empty for those two docs here (the placeholder
    // functions then print the ruled "Dt:" fill-in line).
    const f = opts?.blankDate ? { ...fields, agreementDate: '', workOrderDate: '' } : fields
    const values = ZONAL_KINDS.includes(kind as (typeof ZONAL_KINDS)[number])
      ? zonalDocsPlaceholders(f, notice ?? {}, pdfEval ?? {})
      : kind === 'workOrder'
        ? workOrderPlaceholders(f)
        : kind === 'fileBacker'
          ? fileBackerPlaceholders(f)
          : kind === 'agreement'
          ? agreementPlaceholders(f)
          : kind === 'qccIntimation'
            ? qccIntimationPlaceholders(f)
            : kind === 'forwardingSlip'
              ? forwardingSlipPlaceholders(f)
              : civilTenderPlaceholders(f, pdfEval ?? {}, { pagesOfAgreement, scheduleAItems })
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

  // Guard: the uploaded Circle must match the Circle the documents will fill
  // from. The Circle comes from the NIT No ("…/EE/Gajularamaram Circle-57/…") of
  // EITHER upload — the L-1 selection form or the Online Intimation — so the
  // check fires even when only the Online Intimation has been uploaded; it falls
  // back to inferring the circle from the L-1's work name via the directory. We
  // compare it against the office's Circle (Works List page) when set, otherwise
  // the matched Works List row's Circle — so an upload from another circle that
  // fuzzy-matched a row here is caught and blocked (it would otherwise issue e.g.
  // a Nizampet work's agreement under Gajularamaram).
  const l1Circle = useMemo(() => {
    const fromNit = circleFromNit(pdfEval?.noticeNo || '').circle || circleFromNit(notice?.nitNo || '').circle
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
  // "Fill details manually" is ready once the name of work is typed in — the
  // other fields are optional (blank prints blank). Otherwise both files must be
  // uploaded (and pass the same-work / circle checks).
  const bothUploaded = manualMode
    ? !!manual.nameOfWork.trim()
    : !!notice && !!pdfEval && !workMismatch && !circleMismatch
  const templatesReady = seMode
    ? ZONAL_KINDS.every((k) => !!zonalB64[k])
    : !!workOrderB64 && !!fileBackerB64 && !!agreementB64 && !!forwardingSlipB64 && !!civilTenderB64
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
    if (fbTileRef.current) void renderDocInto('fileBacker', fbTileRef.current).catch(() => {})
    if (agTileRef.current) void renderDocInto('agreement', agTileRef.current).catch(() => {})
    if (qccIntTileRef.current) void renderDocInto('qccIntimation', qccIntTileRef.current).catch(() => {})
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

  // Generate every agreement-workspace document at once, each in its preferred
  // format, into one folder the user picks: Tender / QCC Intimation / Work Order
  // as PDF, Schedule A as Excel, and the rest (Forwarding Slip / Agreement Bond
  // / Note Submitted) as Word. Only the outputs that are ready are included.
  async function downloadAll() {
    setBusy('bundle')
    setActionError(null)
    setActionSaved(null)
    try {
      const files: AgreementBundleFile[] = []
      // File Backer — the cover page — as Word, first in the bundle.
      if (fileBackerB64) files.push({ name: docName('fileBacker'), format: 'docx', docxBase64: await fillDoc('fileBacker') })
      if (civilTenderB64) files.push({ name: docName('civilTender'), format: 'pdf', docxBase64: await fillDoc('civilTender') })
      if (scheduleA)
        files.push({
          name: `Schedule A${fields.agencyName ? ` - ${fields.agencyName}` : ''}`,
          format: 'xlsx',
          scheduleATable: scheduleA,
          scheduleAMeta
        })
      if (forwardingSlipB64) files.push({ name: docName('forwardingSlip'), format: 'docx', docxBase64: await fillDoc('forwardingSlip') })
      // Work Order & Agreement Bond print a blank, hand-writable date in the bundle.
      if (agreementB64) files.push({ name: docName('agreement'), format: 'docx', docxBase64: await fillDoc('agreement', { blankDate: true }) })
      if (qccIntimationB64) files.push({ name: docName('qccIntimation'), format: 'pdf', docxBase64: await fillDoc('qccIntimation') })
      if (workOrderB64) files.push({ name: docName('workOrder'), format: 'pdf', docxBase64: await fillDoc('workOrder', { blankDate: true }) })
      if (noteReady)
        files.push({
          name: noteSubmittedFileName(noteData?.workName),
          format: 'docx',
          docxBase64: await api.noteSubmittedDocx(notePreviewHtml)
        })
      if (files.length === 0) {
        setActionError('No documents are ready to download yet.')
        return
      }
      const res = await api.exportAgreementBundle(files)
      if (!res) {
        setActionSaved('Cancelled.')
      } else if (res.failed.length > 0) {
        // Some documents (usually PDFs needing LibreOffice) didn't convert — say
        // which, instead of silently dropping them from the folder.
        setActionSaved(`Saved ${res.written.length} document(s).`)
        setActionError(
          `${res.failed.length} document(s) could not be saved: ${res.failed.join(', ')}. ` +
            `PDFs need LibreOffice installed — install it from libreoffice.org, or download those as Word (.docx) instead.`
        )
      } else {
        setActionSaved(`Saved ${res.written.length} document(s) to the chosen folder.`)
      }
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
    // The tender notice No & date live on the uploaded L1 sheet itself (its
    // "Enquiry/IFB/Tender … Notice Number … Dt:" line). Fall back to them when
    // the Works List row hasn't been updated from L1 yet, so the note reflects
    // the sheet directly rather than sitting blank. NIT No mirrors the notice No.
    if (!seed.tenderNoticeNo && pdfEval?.noticeNo) {
      seed.tenderNoticeNo = pdfEval.noticeNo
      seed.nitNo = pdfEval.noticeNo
    }
    if (!seed.tenderNoticeDate && pdfEval?.noticeDate) {
      seed.tenderNoticeDate = pdfEval.noticeDate
      seed.nitDate = pdfEval.noticeDate
    }
    setNoteData(seed)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowIndex, table, allBidders, pdfEval])

  const notePreviewHtml = useMemo(() => (noteData ? buildNoteSubmittedHtml(noteData) : ''), [noteData])
  // Same gate as the Work Order / Agreement tiles: don't build the Note
  // Submitted while the Online Intimation and L1 selection form disagree on the
  // work/agency, the L1's work isn't in the Works List (the note is seeded
  // from the selected row), or the L1's circle isn't this office's circle —
  // otherwise it would show the wrong work/agency or the wrong circle's note.
  const noteReady = !!noteData && !!pdfEval && !workMismatch && !circleMismatch

  // Optional non-responsiveness statement — pre-fills the note's rejection line.
  async function handleNonRespFile(file: File) {
    setActionError(null)
    try {
      const lines = await pdfToPositionedLines(file)
      const { count, detail } = summarizeNonResponsiveness(lines)
      setNoteData((prev) => (prev ? { ...prev, rejectedCount: count, qualificationNote: detail } : prev))
      setPdfStatus(
        count > 0
          ? `Non-responsiveness read from ${file.name}: (${count}) rejected — check the count & reason in the Note Submitted editor.`
          : `Read ${file.name}, but found no rejected bidders — set the rejected count in the Note Submitted editor if needed.`
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
      const name = noteSubmittedFileName(noteData.workName)
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
    setDetectedWorkName(detected ?? null)

    // The work this workspace is currently building for: the L-1 selection
    // form's Name of Work when it's been uploaded, else the picked Works List
    // row's name. When both it and the uploaded file name a work, they must be
    // the same work — otherwise this file's Schedule A would silently belong to
    // a different work than the Work Order / Agreement. Block the upload and
    // ask for the same work's details.
    // Whether the uploaded file is the right work's estimate is verified
    // reactively (see the scheduleAMismatch effect) so it re-checks whenever the
    // L-1 / selected work changes too — not only at upload time.
    setBoq(t)
    try {
      // A detailed CMC/departmental estimate (multi-row No.s/L/B/D measurement
      // layout) can't be read row-for-row like a flat BOQ — run the estimate
      // extractor first, which resolves each item's real quantity/rate/unit,
      // and only fall back to the flat-BOQ column mapping when it finds no
      // items (i.e. the file really is a plain BOQ). Both yield a Schedule A.
      const items = extractEstimateItems(g.grid, headerRow)
      setUploadedIsEstimate(items.length > 0)
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
    // Tools Schedule A (scheduleAOnly): no picked row — works for any zone/circle.
    // The name of work comes from the uploaded estimate/BOQ itself; the estimate
    // amount (and the other figures) are looked up from the Works List by that
    // name when it's present there, and left blank otherwise — never shown as the
    // ECV or the BOQ total. Returning a (non-undefined) meta object is what keeps
    // the blanks blank rather than falling back to the item total.
    if (scheduleAOnly) {
      const name = (detectedWorkName ?? '').trim()
      if (!name && !boq) return undefined
      // If this work is in the loaded Works List, its own figures win.
      const match = name && table ? findWorksRowByName(table, name) : undefined
      if (match) return { ...metaFromWorksRow(match.row), nameOfWork: name }
      // Otherwise the amounts come from the uploaded file itself: a flat BOQ's
      // total is the ECV alone; a detailed estimate fills both the Estimate
      // Amount and the ECV. With neither a Works List match nor a total, both
      // stay blank.
      const total = scheduleA
        ? rowsToScheduleAItems(scheduleA).reduce(
            (sum, it) => sum + (Number(String(it.amount).replace(/,/g, '')) || 0),
            0
          )
        : 0
      const amt = total > 0 ? `${indianDigitGroups(total)}/-` : undefined
      return {
        nameOfWork: name,
        estimateAmount: uploadedIsEstimate ? amt : undefined,
        ecvAmount: amt
      }
    }
    if (!selectedRow) return undefined
    const base = metaFromWorksRow(selectedRow)
    // The tender % (and contractor / work name / amounts) are frequently known
    // only from the uploaded L-1 selection form before the Works List row itself
    // carries them — prefer the resolved `fields` values so Schedule A's ECV,
    // Contract Amount, "Tender Quoted %" and "Less: (…)% Less" fill from what was
    // uploaded rather than staying blank against a not-yet-updated (or
    // mis-selected) row. `fields.ecvRupees`/`contractRupees` are the ECV/contract
    // (never the estimate), so this never violates "blank ECV stays blank".
    const ecvNum = fields.ecvRupees.trim() ? Number(fields.ecvRupees.replace(/,/g, '')) : null
    const contractNum = fields.contractRupees.trim() ? Number(fields.contractRupees.replace(/,/g, '')) : null
    return {
      ...base,
      // Name of work comes from the uploaded L-1 (via `fields`), not the row.
      nameOfWork: fields.nameOfWork || base.nameOfWork?.trim() || '',
      contractorName: base.contractorName?.trim() || fields.agencyName,
      tenderPercentage: fields.tenderPercent || base.tenderPercentage,
      ecvAmount: ecvNum != null && Number.isFinite(ecvNum) ? `${indianDigitGroups(ecvNum)}/-` : base.ecvAmount,
      contractAmount:
        contractNum != null && Number.isFinite(contractNum) ? `${indianDigitGroups(contractNum)}/-` : base.contractAmount
    }
  }, [scheduleAOnly, detectedWorkName, boq, scheduleA, uploadedIsEstimate, table, selectedRow, fields])

  const scheduleAPreview = useMemo(
    () => (scheduleA ? buildScheduleARows(rowsToScheduleAItems(scheduleA), scheduleAMeta) : null),
    [scheduleA, scheduleAMeta]
  )

  // Guard the Schedule A against a WRONG estimate: the uploaded file's own work
  // name must match the work this page is building for (the L-1's / selected
  // row's name). Runs reactively — so it re-checks when the L-1 is uploaded after
  // the estimate — and uses a stricter threshold than the lenient shared 0.5,
  // because two different road works are worded almost identically and would
  // otherwise pass. On a mismatch the message is set and the Schedule A tile is
  // withheld, so the wrong estimate's items never masquerade under this work.
  const [scheduleAMismatch, setScheduleAMismatch] = useState<string | null>(null)
  useEffect(() => {
    if (standalone || !boq) {
      setScheduleAMismatch(null)
      return
    }
    const expected = (pdfEval?.nameOfWork ?? '').trim() || (selectedRow?.['Name of the work'] ?? '').trim()
    if (!expected) {
      setScheduleAMismatch(null)
      return
    }
    const det = (detectedWorkName ?? '').trim()
    if (!det) {
      setScheduleAMismatch(
        `Couldn't read a work name from the uploaded estimate to confirm it's for “${expected}”. ` +
          `Please check you uploaded the correct estimate for this work.`
      )
      return
    }
    let cancelled = false
    void (async () => {
      let embeddings: { aVector: number[]; bVector: number[] } | undefined
      try {
        const [aVector, bVector] = await api.embedTexts([det, expected])
        embeddings = { aVector, bVector }
      } catch {
        embeddings = undefined
      }
      const cmp = compareWorkNames(det, expected, embeddings)
      // With the embedding model available, use a stricter line than
      // compareWorkNames' own 0.5 — a different-but-similarly-worded road work
      // still scores well above 0.5. Without embeddings, fall back to its normal
      // token-overlap verdict (0.82 overlap would false-flag legit wording drift).
      const ok = embeddings ? (cmp.score ?? 1) >= 0.82 : cmp.status !== 'mismatch'
      if (!cancelled) setScheduleAMismatch(ok ? null : workNameMismatchMessage(det, expected))
    })()
    return () => {
      cancelled = true
    }
  }, [standalone, boq, detectedWorkName, pdfEval, selectedRow])

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

      <div className={only ? 'wo-compact-body' : 'empty empty--tight'}>
        {!only && <IconClipboard />}
        <div className={only ? 'wo-compact-actions' : 'boq-actions boq-actions--grid'}>
          {only && !scheduleAOnly && (
            <button className="ghost upload-btn" onClick={() => setManualMode((m) => !m)}>
              <IconClipboard /> {manualMode ? 'Upload L1 + Intimation instead' : 'Fill details manually'}
            </button>
          )}
          {!scheduleAOnly && !(only && manualMode) && (
            <button className="primary upload-btn" onClick={() => noticeInputRef.current?.click()} disabled={!templatesReady}>
              <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
            </button>
          )}
          {!scheduleAOnly && !(only && manualMode) && (
            <button className="primary upload-btn" onClick={() => pdfInputRef.current?.click()} disabled={!templatesReady || busy === 'pdf'}>
              <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
            </button>
          )}
          {!only && (
            <button className="primary upload-btn" onClick={uploadBoq} disabled={noWorks}>
              <IconTable /> {boq ? 'Change estimate / BOQ' : 'Upload estimate / BOQ'}
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
        {only && manualMode && (
          <div className="wo-manual-form">
            <label className="wo-date-field wo-manual-wide">
              <span>Name of the work *</span>
              <input
                type="text"
                value={manual.nameOfWork}
                placeholder="Laying of CC road from … in ward no … , … Circle-…, … Zone, CMC"
                onChange={(e) => setManual((p) => ({ ...p, nameOfWork: e.target.value }))}
              />
            </label>
            <label className="wo-date-field">
              <span>Name of the agency</span>
              <input type="text" value={manual.agencyName} onChange={(e) => setManual((p) => ({ ...p, agencyName: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Agency phone</span>
              <input type="text" value={manual.phone} onChange={(e) => setManual((p) => ({ ...p, phone: e.target.value }))} />
            </label>
            <label className="wo-date-field wo-manual-wide">
              <span>Agency address</span>
              <input type="text" value={manual.address} onChange={(e) => setManual((p) => ({ ...p, address: e.target.value }))} />
            </label>
            {(office?.corporation || office?.zone || office?.circle) && (
              <span className="estimate-hint wo-manual-wide">
                Using office:{' '}
                {[office?.corporation, office?.zone && `${office.zone} Zone`, office?.circle && `${office.circle} Circle${office?.circleNumber ? `-${office.circleNumber}` : ''}`]
                  .filter(Boolean)
                  .join(' · ')}{' '}
                — Circle / Zone aren’t asked below.
              </span>
            )}
            {!(office?.zone ?? '').trim() && (
              <label className="wo-date-field">
                <span>Zone</span>
                <input type="text" placeholder="Quthbullapur" value={manual.zone} onChange={(e) => setManual((p) => ({ ...p, zone: e.target.value }))} />
              </label>
            )}
            {!(office?.circle ?? '').trim() && (
              <label className="wo-date-field">
                <span>Circle</span>
                <input type="text" placeholder="Gajularamaram" value={manual.circle} onChange={(e) => setManual((p) => ({ ...p, circle: e.target.value }))} />
              </label>
            )}
            {!(office?.circle ?? '').trim() && (
              <label className="wo-date-field">
                <span>Circle number</span>
                <input type="text" placeholder="57" value={manual.cno} onChange={(e) => setManual((p) => ({ ...p, cno: e.target.value }))} />
              </label>
            )}
            {!(office?.corporation ?? '').trim() && (
              <label className="wo-date-field">
                <span>Corporation</span>
                <input type="text" placeholder="CMC" value={manual.corporation} onChange={(e) => setManual((p) => ({ ...p, corporation: e.target.value }))} />
              </label>
            )}
            <label className="wo-date-field">
              <span>Financial year</span>
              <input type="text" placeholder="2026-27" value={manual.financialYear} onChange={(e) => setManual((p) => ({ ...p, financialYear: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Wincode</span>
              <input type="text" value={manual.wincode} onChange={(e) => setManual((p) => ({ ...p, wincode: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Amount of estimate (Lakhs)</span>
              <input type="text" placeholder="45.00" value={manual.estimateLakhs} onChange={(e) => setManual((p) => ({ ...p, estimateLakhs: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>ECV (₹)</span>
              <input type="text" placeholder="4149409" value={manual.ecvRupees} onChange={(e) => setManual((p) => ({ ...p, ecvRupees: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Contract amount (₹)</span>
              <input type="text" placeholder="3114647.29" value={manual.contractRupees} onChange={(e) => setManual((p) => ({ ...p, contractRupees: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Tender %</span>
              <input type="text" placeholder="25.01" value={manual.tenderPercent} onChange={(e) => setManual((p) => ({ ...p, tenderPercent: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Agreement date</span>
              <input type="date" value={manual.agreementDate} onChange={(e) => setManual((p) => ({ ...p, agreementDate: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Admin sanction date</span>
              <input type="date" value={manual.adminSanctionDate} onChange={(e) => setManual((p) => ({ ...p, adminSanctionDate: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Period of completion (months)</span>
              <input type="text" placeholder="03" value={manual.completionMonths} onChange={(e) => setManual((p) => ({ ...p, completionMonths: e.target.value }))} />
            </label>
            <label className="wo-date-field">
              <span>Reservation</span>
              <select value={manual.reservation} onChange={(e) => setManual((p) => ({ ...p, reservation: e.target.value }))}>
                <option value="">None</option>
                <option value="SC">SC</option>
                <option value="ST">ST</option>
                <option value="BC">BC</option>
              </select>
            </label>
            <span className="estimate-hint wo-manual-wide">
              Enter the details, then click the {only === 'agreement' ? 'Agreement Bond' : 'Work Order'} tile below to preview. Blank
              fields print blank.
            </span>
          </div>
        )}
        {!scheduleAOnly && !only && (
          <div className="wo-date-row">
            <label className="wo-date-field">
              <span>Agreement date</span>
              <input type="date" value={agreementDate} onChange={(e) => setAgreementDate(e.target.value)} />
            </label>
            <span className="estimate-hint">
              Same date fills the Work Order and the Agreement Bond. Leave blank to print a ruled date line to fill in by hand.
            </span>
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
        {/* The "Matched … its details fill the documents" success line is
            suppressed on a circle mismatch — there, the L-1 belongs to another
            circle, nothing fills, and the red circle-mismatch error below is the
            only message that should show (otherwise the two contradict). */}
        {!only && pdfStatus && !circleMismatch && (
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
      {scheduleAMismatch && (
        <div className="notice error">
          <IconWarn /> {scheduleAMismatch}
        </div>
      )}
      {workMismatch && workMatch && (
        <div className="notice error">
          <IconWarn /> {sameWorkMismatchMessage(workMatch)}
        </div>
      )}
      {circleMismatch && (
        <div className="notice error">
          <IconWarn /> Office selected and L1 sheet circle does not match — upload the L1 sheet of{' '}
          <strong>{targetCircle}</strong> circle only.
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
      {actionError && !expanded && (
        <div className="notice error">
          <IconWarn /> {actionError}
        </div>
      )}

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the outputs are filled from a work's row.</div>
      ) : anyOutput ? (
        <>
          {docsReady && !seMode && !only && (
            <div className="wo-download-all">
              <button className="primary" disabled={busy === 'bundle'} onClick={downloadAll}>
                <IconDownload /> {busy === 'bundle' ? 'Preparing documents…' : 'Download all documents'}
              </button>
              <span className="wo-download-all-note">
                Tender, QCC Intimation &amp; Work Order as PDF · Schedule A as Excel · Forwarding Slip, Agreement Bond &amp;
                Note Submitted as Word — all into one folder you choose.
              </span>
              {busy === 'bundle' && <span className="wo-download-all-note">This can take a moment (PDF conversion).</span>}
            </div>
          )}
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
          {/* Order (main tab): File Backer, Note Submitted, Forwarding Slip,
              Agreement Bond, Schedule A, Work Order, Tender Document, QCC
              Intimation. The Tools single-document panels (`only`) still show
              just their one tile via the `only ===` guards. */}
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('fileBacker')}>
              <div className="wo-tile-preview">
                <div ref={fbTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.fileBacker}</div>
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
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('forwardingSlip')}>
              <div className="wo-tile-preview">
                <div ref={fsTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.forwardingSlip}</div>
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
          {scheduleAPreview && !only && !scheduleAMismatch && (
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
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('civilTender')}>
              <div className="wo-tile-preview">
                <div ref={ctTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.civilTender}</div>
            </button>
          )}
          {docsReady && !only && !seMode && (
            <button className="wo-tile" onClick={() => openDoc('qccIntimation')}>
              <div className="wo-tile-preview">
                <div ref={qccIntTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.qccIntimation}</div>
            </button>
          )}
          </div>
        </>
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
              {/* Skip the date: preview with the date left blank, so the document
                  prints its ruled "Dt: ____" line to be hand-dated. Circle/Zone,
                  when the L-1 didn't carry them, are still required. */}
              <button
                className="ghost"
                disabled={missingCircleZone && (!manualCircle.trim() || !manualZone.trim())}
                onClick={() => {
                  setAgreementDate('')
                  setDatePromptOpen(false)
                  setExpanded(pendingDoc)
                }}
              >
                Continue without date
              </button>
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
