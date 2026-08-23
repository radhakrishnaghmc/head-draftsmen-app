import { useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, parseAllBidders, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { parseBalanceEmdReceipt } from '@core/balanceEmdReceipt'
import { parseBankGuarantees, extractBankName } from '@core/bankGuaranteePdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import {
  deriveFields,
  workOrderPlaceholders,
  fileBackerPlaceholders,
  agreementPlaceholders,
  seAgreementBondPlaceholders,
  zonalDocsPlaceholders,
  qccIntimationPlaceholders,
  forwardingSlipPlaceholders,
  civilTenderPlaceholders,
  standaloneRowFromSources,
  circleFromNit,
  indianFinancialYear
} from '@core/workOrderAgreement'
import type { WorkOrderAgreementFields } from '@core/workOrderAgreement'
import { extractItemNo, stripItemNoTag } from '@core/bidDocument'
import { corporationByName, resolveFromDirectory, entriesOf } from '../zoneCircleDirectory'
import { type Office, officeScopedKey, TEMPLATE_KEYS } from '../office'
import { boqToScheduleA, buildBoqFromEstimate, extractWorkNameFromBoq } from '../boqTransform'
import { guessHeaderRow, buildTableFromGrid } from '@core/sheet'
import { extractEstimateItems, extractWorkName } from '@core/estimateExtract'
import { buildScheduleARows, rowsToScheduleAItems, metaFromWorksRow, findWorksRowByName } from '@core/scheduleA'
import { indianDigitGroups } from '@core/worksAmounts'
import { compareWorkNames, workNameMismatchMessage, normWorkName } from '@core/workNameMatch'
import {
  buildNoteSubmittedHtml,
  noteSubmittedFromRow,
  summarizeNonResponsiveness,
  tenderPctMagnitude,
  type NoteSubmittedData,
  type NoteBidder
} from '@core/noteSubmitted'
import NoteSubmittedEditor from './NoteSubmittedEditor'
import { mismatchHint } from '../docClassify'
import type { PlaceholderMatch } from '@core/createDocument'
import type { ScheduleAMeta, AgreementBundleFile } from '../../electron/ipc-contract'
import type { ExcelTable } from '@core/types'
import { pdfToTextLines, pdfToPositionedLines } from '../pdfToText'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview } from './docPage'
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
  /**
   * DOM node (rendered by the page header, App.tsx) to portal the "Download
   * all documents" button into — puts it in the page-head-action slot next
   * to the title, matching Calendar's "Issue tender notice" placement,
   * instead of buried in the body below several upload/field sections.
   */
  headerActionRef?: RefObject<HTMLDivElement | null>
}

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(file)
  })
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
  | 'seAgreementBond'
  | 'zonalWorkOrder'
  | 'zonalConcludingAgreement'
  | 'zonalMemoEe'
  | 'seAgreementNote'
  | 'contractDeed'
type Output = DocKind | 'scheduleA' | 'note'

const DOC_LABEL: Record<Output, string> = {
  fileBacker: 'File Backer',
  forwardingSlip: 'Forwarding Slip',
  civilTender: 'Tender Document',
  workOrder: 'Work Order',
  agreement: 'Agreement Bond',
  qccIntimation: 'QCC Intimation',
  seAgreementBond: 'Agreement Bond',
  zonalWorkOrder: 'Work Order',
  zonalConcludingAgreement: 'Concluding Agreement',
  zonalMemoEe: 'Memo to EE',
  seAgreementNote: 'Agreement Put-up Note',
  contractDeed: 'Contract Deed',
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
  'seAgreementBond',
  'zonalWorkOrder',
  'zonalConcludingAgreement',
  'zonalMemoEe',
  'seAgreementNote',
  'contractDeed'
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
    nameOfWork: stripItemNoTag(m.nameOfWork.trim()),
    itemNo: extractItemNo(m.nameOfWork) ?? '',
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
    ceLetterNoDate: '',
    completionMonths: m.completionMonths.trim(),
    reservation: m.reservation.trim(),
    emdDetails: '',
    // Not collected by the manual-entry form — same precedent as tsNoDate/
    // ceLetterNoDate/emdDetails above, left blank until a real L1/Intimation
    // upload (or Works List match) supplies them.
    noticeNo: '',
    tenderId: '',
    noticeDate: '',
    intimationDate: ''
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
  office,
  headerActionRef
}: Props) {
  const standalone = standaloneProp || scheduleAOnly || !!only
  const table = tables[0] ?? null

  // Whether headerActionRef's DOM node is actually mounted yet — reading
  // headerActionRef.current directly during render is unreliable (refs
  // attach during commit, after render runs, so the very first render
  // always sees it as null and never gets another render to correct that
  // unless something else also changes). This effect re-renders once the
  // node is there, so the "Download all documents" button reliably ends up
  // in the header instead of silently staying in its (now-removed) inline
  // fallback forever.
  const [headerMounted, setHeaderMounted] = useState(false)
  useLayoutEffect(() => {
    setHeaderMounted(!!headerActionRef?.current)
  })

  // Zone-level (SE office) mode: a Zone chosen with no Circle of its own. The
  // page then produces the SE documents (currently just the Agreement Bond
  // paper, being rebuilt one at a time) instead of the circle-level Work
  // Order / Agreement / Forwarding Slip / Tender Document. Mirrors
  // GiveIntimationTab's `seMode`. Never in the Tools single-document (`only`)
  // panels.
  const seMode = !only && !!office?.zone?.trim() && !office?.circle?.trim()

  const [workOrderB64, setWorkOrderB64] = useState<string | null>(null)
  const [fileBackerB64, setFileBackerB64] = useState<string | null>(null)
  const [agreementB64, setAgreementB64] = useState<string | null>(null)
  const [qccIntimationB64, setQccIntimationB64] = useState<string | null>(null)
  const [forwardingSlipB64, setForwardingSlipB64] = useState<string | null>(null)
  const [civilTenderB64, setCivilTenderB64] = useState<string | null>(null)
  const [seAgreementBondB64, setSeAgreementBondB64] = useState<string | null>(null)
  const [zonalWorkOrderB64, setZonalWorkOrderB64] = useState<string | null>(null)
  const [zonalConcludingAgreementB64, setZonalConcludingAgreementB64] = useState<string | null>(null)
  const [zonalMemoEeB64, setZonalMemoEeB64] = useState<string | null>(null)
  const [seAgreementNoteB64, setSeAgreementNoteB64] = useState<string | null>(null)
  const [contractDeedB64, setContractDeedB64] = useState<string | null>(null)
  const [workOrderLabels, setWorkOrderLabels] = useState<string[]>([])
  const [fileBackerLabels, setFileBackerLabels] = useState<string[]>([])
  const [agreementLabels, setAgreementLabels] = useState<string[]>([])
  const [qccIntimationLabels, setQccIntimationLabels] = useState<string[]>([])
  const [forwardingSlipLabels, setForwardingSlipLabels] = useState<string[]>([])
  const [civilTenderLabels, setCivilTenderLabels] = useState<string[]>([])
  const [seAgreementBondLabels, setSeAgreementBondLabels] = useState<string[]>([])
  const [zonalWorkOrderLabels, setZonalWorkOrderLabels] = useState<string[]>([])
  const [zonalConcludingAgreementLabels, setZonalConcludingAgreementLabels] = useState<string[]>([])
  const [zonalMemoEeLabels, setZonalMemoEeLabels] = useState<string[]>([])
  const [seAgreementNoteLabels, setSeAgreementNoteLabels] = useState<string[]>([])
  const [contractDeedLabels, setContractDeedLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)
  // Forwarding Slip fields the office hand-enters (not on the Works List / L-1).
  const [tsNoDate, setTsNoDate] = useState('')
  // Chief Engineer's administrative-sanction letter No & Date — the SE zonal
  // Work Order / Agreement Put-up Note still print this ({{CE Letter No Date}}),
  // but the app no longer offers a way to enter it (removed from this
  // workspace on request), so it always prints blank now.
  const ceLetterNoDate = ''
  // EMD/ASD payment details for the Concluding Agreement's item-5 row,
  // hand-entered — free text (online split payment vs. Bank Guarantee, with
  // receipt/payment/UTR numbers), never derivable from the Works List.
  const [emdDetails, setEmdDetails] = useState('')
  // The same "Upload Balance EMD Payment" receipt's Online Receipt No / date —
  // Note Submitted's note-6 EMD clause prints these ("vide Online Receipt No:
  // ___ Dt: ___") but noteSubmittedFromRow always seeds them blank (nothing in
  // the Works List/L1/Intimation carries them). Kept as its own state (like
  // allBidders below) rather than written straight into noteData, so it
  // survives noteData's re-seed effect instead of being wiped the next time
  // the row/upload changes — see that effect's dependency array.
  const [emdReceiptInfo, setEmdReceiptInfo] = useState<{ receiptNo: string; receiptDate: string } | null>(null)
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
  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf' | 'bundle' | 'emdReceipt'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  // Which format "Download all documents" saves every document as (Schedule
  // A is always Excel regardless — the one format it actually has). Defaults
  // to Word since PDF needs LibreOffice installed.
  const [bundleFormat, setBundleFormat] = useState<'docx' | 'pdf'>('docx')
  // Live "Download all documents" progress — filling each document locally,
  // then (a separate, often slower phase, especially for PDF: LibreOffice per
  // file) the main process writing/converting them — so the button reads
  // "Preparing 3 of 8…" / "Saving 3 of 8…" instead of a single static
  // "Preparing…" that looks frozen on a big bundle.
  const [bundleProgress, setBundleProgress] = useState<{ phase: 'preparing' | 'saving'; done: number; total: number } | null>(null)
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

  // "No. of items in Schedule 'A'" is just the uploaded estimate's own item
  // count — auto-fill it from there instead of asking the office to count and
  // type it in by hand. Still a plain editable field (not derived inline at
  // render time) so it can be corrected if the estimate ever needs it, and so
  // it stays blank — not "0" — until an estimate is actually uploaded.
  useEffect(() => {
    if (scheduleA) setScheduleAItems(String(scheduleA.rows.length))
  }, [scheduleA])

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
  const emdReceiptInputRef = useRef<HTMLInputElement>(null)
  const woTileRef = useRef<HTMLDivElement>(null)
  const fbTileRef = useRef<HTMLDivElement>(null)
  const agTileRef = useRef<HTMLDivElement>(null)
  const qccIntTileRef = useRef<HTMLDivElement>(null)
  const fsTileRef = useRef<HTMLDivElement>(null)
  const ctTileRef = useRef<HTMLDivElement>(null)
  const seBondTileRef = useRef<HTMLDivElement>(null)
  const zwoTileRef = useRef<HTMLDivElement>(null)
  const zcaTileRef = useRef<HTMLDivElement>(null)
  const zmeTileRef = useRef<HTMLDivElement>(null)
  const seNoteTileRef = useRef<HTMLDivElement>(null)
  const contractDeedTileRef = useRef<HTMLDivElement>(null)
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

  // Memoized: standaloneRowFromSources builds a brand-new plain object every
  // call, so leaving this as a plain `const` (recomputed on every render, not
  // just when its inputs change) meant `selectedRow` never held a stable
  // reference in standalone/Tools mode or whenever the L-1's work matched no
  // Works List row — which broke `fields`' memoization below it, which broke
  // the live-preview effect's memoization, which re-filled and re-rendered
  // every document tile (6) via docx-preview on EVERY
  // render of this component, not just on real data changes. That's what
  // made the whole page feel sluggish after uploading: any click anywhere
  // that re-rendered this component re-triggered the full document re-render
  // burst, whether or not the underlying data had actually changed.
  const selectedRow = useMemo(
    () =>
      deriveFromUploads
        ? notice || pdfEval
          ? standaloneRowFromSources(pdfEval ?? {}, notice ?? {})
          : null
        : !notice && !pdfEval
          ? null // Neither uploaded yet — stay blank rather than leaking row 0's
            // unrelated work (SE mode shows its document catalog before either
            // upload; see docsReady).
          : worksRowMatched === null
            ? null // An L-1/Notice upload just landed and its Works List match
              // (syncWorksListRow's embedding check) hasn't resolved yet — stay
              // blank for this brief async window rather than leaking whatever
              // row[rowIndex] currently points to (often row 0's unrelated work).
              // Note Submitted's own seeding effect watches selectedRow, so this
              // is what actually stops it momentarily showing the wrong work.
            : table && table.rows.length > 0
              ? table.rows[Math.min(rowIndex, table.rows.length - 1)]
              : null,
    [deriveFromUploads, notice, pdfEval, worksRowMatched, table, rowIndex]
  )

  // Load both bundled formats once, and read their placeholders.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        // The office's chosen Work Order / Agreement Bond template variants
        // (Settings' Document Templates section — see
        // core/workOrderTemplateVariants.ts), read fresh on each mount the
        // same way the office's other per-office settings (CONTACT_KEYS)
        // already are — falls back to the default when this office hasn't
        // picked one.
        const workOrderVariant = localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.workOrder, office)) ?? undefined
        const agreementVariant = localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.agreement, office)) ?? undefined
        const [woB64, fbB64, agB64, qiB64, fsB64, ctB64] = await Promise.all([
          api.workOrderTemplate(workOrderVariant),
          api.fileBackerTemplate(),
          api.agreementTemplate(agreementVariant),
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

  // Zone-level (SE office) templates: loaded separately, only when relevant,
  // since a Circle office never needs them.
  useEffect(() => {
    if (!seMode) return
    let cancelled = false
    void (async () => {
      try {
        const [bondB64, woB64, caB64, meB64, noteB64, deedB64] = await Promise.all([
          api.seAgreementBondTemplate(),
          api.zonalWorkOrderTemplate(),
          api.zonalConcludingAgreementTemplate(),
          api.zonalMemoEeTemplate(),
          api.seAgreementNoteTemplate(),
          api.contractDeedTemplate()
        ])
        const [bondLabels, woLabels, caLabels, meLabels, noteLabels, deedLabels] = await Promise.all([
          api.findPlaceholdersInDocument(bondB64),
          api.findPlaceholdersInDocument(woB64),
          api.findPlaceholdersInDocument(caB64),
          api.findPlaceholdersInDocument(meB64),
          api.findPlaceholdersInDocument(noteB64),
          api.findPlaceholdersInDocument(deedB64)
        ])
        if (cancelled) return
        setSeAgreementBondB64(bondB64)
        setZonalWorkOrderB64(woB64)
        setZonalConcludingAgreementB64(caB64)
        setZonalMemoEeB64(meB64)
        setSeAgreementNoteB64(noteB64)
        setContractDeedB64(deedB64)
        setSeAgreementBondLabels(bondLabels)
        setZonalWorkOrderLabels(woLabels)
        setZonalConcludingAgreementLabels(caLabels)
        setZonalMemoEeLabels(meLabels)
        setSeAgreementNoteLabels(noteLabels)
        setContractDeedLabels(deedLabels)
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [seMode])

  // Circle/CNO for the work, read off the NIT No ("…/EE/Gajularamaram
  // Circle-57/…") of either upload, falling back to a directory scan of the
  // L-1's work name — the same resolution the circle-mismatch guard below
  // uses. Serves two purposes: that guard's comparison value, and (for a Zone
  // office, which has no circle of its own) the fallback that fills the
  // documents' Circle/CNO fields when no Works List row matched.
  const l1CircleInfo = useMemo(() => {
    const fromEval = circleFromNit(pdfEval?.noticeNo || '')
    if (fromEval.circle) return fromEval
    const fromNotice = circleFromNit(notice?.nitNo || '')
    if (fromNotice.circle) return fromNotice
    const d = resolveFromDirectory(pdfEval?.nameOfWork || '', entriesOf(office?.corporation))
    return { circle: d.circle ?? '', cno: d.cno ?? '' }
  }, [pdfEval, notice, office])

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
      circle: (only ? '' : office?.circle) || f.circle || l1CircleInfo.circle || manualCircle,
      cno: (only ? '' : office?.circleNumber) || f.cno || l1CircleInfo.cno || manualCno,
      zone: (only ? '' : office?.zone) || f.zone || manualZone,
      corporation: (only ? '' : office?.corporation) ?? '',
      corporationFullName: (only ? '' : corporationByName(office?.corporation)?.fullName) ?? '',
      tsNoDate,
      ceLetterNoDate,
      completionMonths,
      emdDetails
    }
  }, [
    manualMode,
    manual,
    notice,
    pdfEval,
    selectedRow,
    agreementDate,
    office,
    tsNoDate,
    ceLetterNoDate,
    completionMonths,
    emdDetails,
    manualCircle,
    manualCno,
    manualZone,
    l1CircleInfo
  ])

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
    if (
      !manualMode &&
      (kind === 'workOrder' || kind === 'agreement' || kind === 'seAgreementBond') &&
      (!agreementDate || needsCircleZone)
    ) {
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

  // Balance EMD 1.5%/2.5% and ASD are sometimes furnished as a Bank
  // Guarantee instead of an online payment — the certificate has no receipt
  // number to key off, and different banks lay theirs out completely
  // differently, so this can't be matched by label. Instead: compute what
  // this work's EMD/ASD SHOULD be from its ECV (same formula as
  // worksAmounts.computeWorkAmounts) and match each parsed BG's amount to
  // whichever expected figure it's closest to — a BG's actual amount can be
  // off by a rupee or two from the formula (bank rounding), so nearest-match
  // rather than an exact-equality check. Returns undefined (falls back to
  // the "couldn't read that" error) when no BG or no usable ECV is found.
  function composeBankGuaranteeSummary(lines: string[]): string | undefined {
    const bgs = parseBankGuarantees(lines).filter((b) => b.amountRupees != null)
    if (bgs.length === 0) return undefined

    const ecv = fields.ecvRupees.trim() ? Number(fields.ecvRupees.replace(/,/g, '')) : null
    if (ecv == null || !Number.isFinite(ecv)) return undefined
    const tenderPercent = fields.tenderPercent.trim() ? Number(fields.tenderPercent) : null

    const emdPct = fields.reservation ? '2.5%' : '1.5%'
    const expectedEmd = Math.round(ecv * (fields.reservation ? 0.025 : 0.015))
    const expectedAsd =
      tenderPercent != null && Number.isFinite(tenderPercent) && tenderPercent > 25
        ? Math.round(ecv * ((tenderPercent - 25) / 100))
        : 0

    type Slot = { label: 'EMD' | 'ASD'; expected: number; bg?: (typeof bgs)[number] }
    const slots: Slot[] = [{ label: 'EMD', expected: expectedEmd }]
    if (expectedAsd > 0) slots.push({ label: 'ASD', expected: expectedAsd })

    for (const bg of bgs) {
      let best: Slot | undefined
      let bestDiff = Infinity
      for (const s of slots) {
        if (s.bg) continue
        const diff = Math.abs((bg.amountRupees as number) - s.expected)
        if (diff < bestDiff) {
          bestDiff = diff
          best = s
        }
      }
      if (best) best.bg = bg
    }

    const amt2 = (n: number) => `${indianDigitGroups(n)}.00`
    const normDate = (d: string) => d.replace(/[/\-]/g, '.')
    const emdSlot = slots.find((s) => s.label === 'EMD')
    const asdSlot = slots.find((s) => s.label === 'ASD')

    const parts: string[] = []
    if (emdSlot?.bg) {
      const bg = emdSlot.bg
      parts.push(
        `EMD ${emdPct} of Rs.${amt2(bg.amountRupees as number)} BG Payment` +
          (bg.bgNo ? ` with BG NO:${bg.bgNo}` : '') +
          (bg.issueDate ? `, dated:${normDate(bg.issueDate)}` : '')
      )
    }
    if (asdSlot?.bg) {
      const bg = asdSlot.bg
      parts.push(
        `ASD of Rs.${amt2(bg.amountRupees as number)}` +
          (bg.bgNo ? ` with BG No:${bg.bgNo}` : '') +
          (bg.issueDate ? ` dated:${normDate(bg.issueDate)}` : '')
      )
    }
    if (parts.length === 0) return undefined

    const bankName = extractBankName(lines)
    return parts.join(', ') + (bankName ? ` of ${bankName}` : '')
  }

  // The CURE portal's "Balance EMD payment Receipt" — the confirmation printed
  // after the L-1 agency pays the balance 1.5%/2.5% EMD online — either a saved
  // PDF or a phone photo of the same receipt (offices often only have a photo).
  // Composes a one-line summary for the Concluding Agreement's item-5 "EMD
  // details" row; the field stays free-text and editable so a Bank-Guarantee
  // payment (no receipt at all) or any correction can still be typed by hand.
  async function handleEmdReceiptFile(file: File) {
    setBusy('emdReceipt')
    setActionError(null)
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const lines = isPdf
        ? await pdfToTextLines(file)
        : await api.ocrPhotosToLines([await readAsDataUrl(file)])
      const r = parseBalanceEmdReceipt(lines)
      if (r.balanceEmdRupees == null && !r.receiptNo) {
        const summary = composeBankGuaranteeSummary(lines)
        if (!summary) {
          throw new Error(
            "Couldn't read that receipt — is it the CURE portal's Balance EMD payment Receipt or a Bank Guarantee? " +
              "If it's a Bank Guarantee, make sure this work's ECV and Tender Percentage are filled in first " +
              '(the amounts on the certificate are matched to EMD/ASD by the work\'s expected figures) — ' +
              'otherwise type the BG details into EMD Details by hand.'
          )
        }
        setEmdDetails((prev) => (prev.trim() ? `${prev}\n${summary}` : summary))
        return
      }
      // The office's own note always commits to one specific rate (never the
      // receipt's own ambiguous "1.5% / 2.5%" label) — 2.5% for a reserved
      // (SC/ST) work, 1.5% otherwise, matching the reservation already on
      // this work. A reserved work's Balance EMD is exempted (ASD alone
      // covers it) — the receipt shows Rs.0 for it, so say "exempted"
      // outright rather than a confusing "Rs.0" or silently dropping it.
      const emdPct = fields.reservation ? '2.5%' : '1.5%'
      // The office's note always prints these two amounts to 2 decimals
      // ("Rs.3,90,288.00"), even though the receipt's own rupee figures are
      // always whole numbers.
      const amt2 = (n: number) => `${indianDigitGroups(n)}.00`
      const amounts = [
        r.balanceEmdRupees
          ? `EMD ${emdPct} Online payment Rs.${amt2(r.balanceEmdRupees)}`
          : fields.reservation
            ? `EMD ${emdPct} exempted`
            : '',
        r.asdRupees ? `ASD Rs.${amt2(r.asdRupees)}` : ''
      ].filter(Boolean)
      const parts = [
        amounts.join(' & '),
        r.receiptNo ? `, Receipt No:${r.receiptNo}` : '',
        r.paymentDate ? ` Dt.${r.paymentDate.replace(/-/g, '.')}` : ''
      ]
      const summary = parts.join(' ').replace(/\s+,/g, ',').replace(/\s+/g, ' ').trim()
      setEmdDetails((prev) => (prev.trim() ? `${prev}\n${summary}` : summary))
      // Also feeds Note Submitted's note-6 EMD clause ("vide Online Receipt
      // No: ___ Dt: ___"), which noteSubmittedFromRow always seeds blank —
      // this is the one upload that actually carries a receipt number/date.
      const hasReceiptInfo = !!(r.receiptNo || r.paymentDate)
      if (hasReceiptInfo) {
        setEmdReceiptInfo({ receiptNo: r.receiptNo ?? '', receiptDate: (r.paymentDate ?? '').replace(/-/g, '.') })
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function fillDoc(kind: DocKind, opts?: { blankDate?: boolean }): Promise<string> {
    const b64 =
      {
        workOrder: workOrderB64, fileBacker: fileBackerB64, agreement: agreementB64, qccIntimation: qccIntimationB64, forwardingSlip: forwardingSlipB64, civilTender: civilTenderB64, seAgreementBond: seAgreementBondB64,
        zonalWorkOrder: zonalWorkOrderB64, zonalConcludingAgreement: zonalConcludingAgreementB64, zonalMemoEe: zonalMemoEeB64, seAgreementNote: seAgreementNoteB64, contractDeed: contractDeedB64
      }[kind]
    const labels =
      {
        workOrder: workOrderLabels, fileBacker: fileBackerLabels, agreement: agreementLabels, qccIntimation: qccIntimationLabels, forwardingSlip: forwardingSlipLabels, civilTender: civilTenderLabels, seAgreementBond: seAgreementBondLabels,
        zonalWorkOrder: zonalWorkOrderLabels, zonalConcludingAgreement: zonalConcludingAgreementLabels, zonalMemoEe: zonalMemoEeLabels, seAgreementNote: seAgreementNoteLabels, contractDeed: contractDeedLabels
      }[kind] ?? []
    if (!b64) throw new Error('Format not loaded yet.')
    // "Download all" asks for the Work Order / Agreement Bond date to be blank
    // (hand-written on the printed copy), regardless of any date picked for a
    // preview — so force the date empty for those two docs here (the placeholder
    // functions then print the ruled "Dt:" fill-in line).
    const f = opts?.blankDate ? { ...fields, agreementDate: '', workOrderDate: '' } : fields
    const values =
      kind === 'workOrder'
        ? workOrderPlaceholders(f)
        : kind === 'fileBacker'
          ? fileBackerPlaceholders(f)
          : kind === 'agreement'
          ? agreementPlaceholders(f)
          : kind === 'qccIntimation'
            ? qccIntimationPlaceholders(f)
            : kind === 'forwardingSlip'
              ? forwardingSlipPlaceholders(f)
              : kind === 'seAgreementBond'
                ? seAgreementBondPlaceholders(f)
                : kind === 'zonalWorkOrder' ||
                    kind === 'zonalConcludingAgreement' ||
                    kind === 'zonalMemoEe' ||
                    kind === 'seAgreementNote' ||
                    kind === 'contractDeed'
                  ? zonalDocsPlaceholders(f, notice ?? {}, pdfEval ?? {})
                  : civilTenderPlaceholders(f, pdfEval ?? {}, { pagesOfAgreement, scheduleAItems })
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(b64, resolved, values)
  }

  // The full-size, accurate (LibreOffice-backed) render — used only for the
  // expanded preview modal and for print, where a user is actually looking
  // at (or about to print) one specific document.
  async function renderDocInto(kind: DocKind, container: HTMLElement): Promise<number> {
    const filled = await fillDoc(kind)
    const { pageCount } = await renderDocPreview(base64ToUint8(filled), container)
    return pageCount
  }

  // The small, live thumbnail shown inside each tile on the catalog grid — up
  // to 11 of these render at once. Uses the same accurate, LibreOffice-backed
  // render as the expanded modal (renderDocInto, above) — the fast docx-
  // preview.js render this used to take was visibly distorted for some of
  // these templates (font-substitution/layout quirks docx-preview.js can't
  // reproduce; see docxToPdf.ts's own note on why LibreOffice's own
  // rasterizer is trusted over anything else for this). The batched-page
  // conversion in core/docxToPdf.ts (docxToPageImages) made a single
  // document's accurate render fast enough that 11 of them is no longer the
  // problem it once was; the effect below still debounces the trigger so
  // typing doesn't fire 11 LibreOffice conversions per keystroke.
  async function renderTileThumbnail(kind: DocKind, container: HTMLElement): Promise<void> {
    const filled = await fillDoc(kind)
    await renderDocPreview(base64ToUint8(filled), container)
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
  // from — compared against the office's Circle (Works List page) when set,
  // otherwise the matched Works List row's Circle — so an upload from another
  // circle that fuzzy-matched a row here is caught and blocked (it would
  // otherwise issue e.g. a Nizampet work's agreement under Gajularamaram).
  const l1Circle = l1CircleInfo.circle
  const officeCircle = (office?.circle ?? '').trim()
  const rowCircle = (selectedRow?.['Circle'] ?? '').trim()
  const targetCircle = officeCircle || rowCircle
  const sameCircleId = (a: string, b: string) => {
    const na = a.trim().toLowerCase()
    const nb = b.trim().toLowerCase()
    return !!na && !!nb && (na === nb || na.includes(nb) || nb.includes(na))
  }
  const circleMismatch = !standalone && !!l1Circle && !!targetCircle && !sameCircleId(l1Circle, targetCircle)

  // A Zone-level (SE) office's 5 rebuilt Work Order/Agreement documents.
  // Schedule A remains a separate output, driven entirely by the uploaded
  // estimate/BOQ, not a template.
  const templatesReady = seMode
    ? !!seAgreementBondB64 &&
      !!zonalWorkOrderB64 &&
      !!zonalConcludingAgreementB64 &&
      !!zonalMemoEeB64 &&
      !!seAgreementNoteB64 &&
      !!contractDeedB64
    : !!workOrderB64 && !!fileBackerB64 && !!agreementB64 && !!forwardingSlipB64 && !!civilTenderB64
  // The document catalog always shows as soon as its templates load,
  // regardless of office (EE or SE) or whether anything's been uploaded yet
  // — blank fields fill in as uploads land. EE no longer needs an upload
  // just to see its tiles (matches SE, which always worked this way).
  const docsReady = templatesReady

  // Live thumbnails in the document tiles, refreshed whenever the filled values
  // change. The Forwarding Slip shows only on the main tab (not the Tools-mode
  // single-document panels). Debounced — `fields` changes on every keystroke,
  // and each tile is now a real LibreOffice conversion (see renderTileThumbnail's
  // own note); waiting for typing to pause before firing keeps this from
  // running 11 conversions per character typed.
  useEffect(() => {
    if (!docsReady) return
    const timer = setTimeout(() => {
      // Schedule A (the SE catalog's other remaining output) renders its own
      // JSX table below, not through this docx-preview effect.
      if (seMode) {
        if (seBondTileRef.current) void renderTileThumbnail('seAgreementBond', seBondTileRef.current).catch(() => {})
        if (zwoTileRef.current) void renderTileThumbnail('zonalWorkOrder', zwoTileRef.current).catch(() => {})
        if (zcaTileRef.current) void renderTileThumbnail('zonalConcludingAgreement', zcaTileRef.current).catch(() => {})
        if (zmeTileRef.current) void renderTileThumbnail('zonalMemoEe', zmeTileRef.current).catch(() => {})
        if (seNoteTileRef.current) void renderTileThumbnail('seAgreementNote', seNoteTileRef.current).catch(() => {})
        if (contractDeedTileRef.current) void renderTileThumbnail('contractDeed', contractDeedTileRef.current).catch(() => {})
        return
      }
      if (woTileRef.current) void renderTileThumbnail('workOrder', woTileRef.current).catch(() => {})
      if (fbTileRef.current) void renderTileThumbnail('fileBacker', fbTileRef.current).catch(() => {})
      if (agTileRef.current) void renderTileThumbnail('agreement', agTileRef.current).catch(() => {})
      if (qccIntTileRef.current) void renderTileThumbnail('qccIntimation', qccIntTileRef.current).catch(() => {})
      if (fsTileRef.current) void renderTileThumbnail('forwardingSlip', fsTileRef.current).catch(() => {})
      if (ctTileRef.current) void renderTileThumbnail('civilTender', ctTileRef.current).catch(() => {})
    }, 600)
    return () => clearTimeout(timer)
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
      await renderDocPreview(base64ToUint8(filled), container)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Generate every agreement-workspace document at once, into one folder the
  // user picks — each in the format chosen by the toggle next to the button
  // (Word or PDF). Schedule A is always Excel, the one format it actually
  // has. Only the outputs that are ready are included.
  async function downloadAll() {
    setBusy('bundle')
    setActionError(null)
    setActionSaved(null)
    // A Zone-level (SE) office's 6 rebuilt documents — Schedule A (pushed
    // below, outside this count) is its other output.
    const totalPrep = seMode
      ? [seAgreementBondB64, zonalWorkOrderB64, zonalConcludingAgreementB64, zonalMemoEeB64, seAgreementNoteB64, contractDeedB64].filter(
          Boolean
        ).length
      : [fileBackerB64, civilTenderB64, forwardingSlipB64, agreementB64, qccIntimationB64, workOrderB64, noteReady].filter(
          Boolean
        ).length
    let prepDone = 0
    setBundleProgress({ phase: 'preparing', done: 0, total: totalPrep })
    const unsubscribe = api.onAgreementBundleProgress(({ done, total }) =>
      setBundleProgress({ phase: 'saving', done, total })
    )
    try {
      const files: AgreementBundleFile[] = []
      // Safety net: a template placeholder that never got a value stays as
      // literal "{{Label}}" text in the filled document — catch that here,
      // before the office downloads a document with a visible unfilled blank,
      // instead of only finding out from a printed copy.
      const unresolvedWarnings: string[] = []
      const add = async (name: string, docxBase64: string) => {
        files.push({ name, format: bundleFormat, docxBase64 })
        prepDone += 1
        setBundleProgress({ phase: 'preparing', done: prepDone, total: totalPrep })
        const leftover = await api.findPlaceholdersInDocument(docxBase64)
        if (leftover.length > 0) unresolvedWarnings.push(`${name} (${leftover.join(', ')})`)
      }
      if (seMode) {
        // File Backer-equivalent ordering: Work Order first, then the Agreement
        // Bond, then its two forwarding memos and the put-up note.
        if (zonalWorkOrderB64) await add(docName('zonalWorkOrder'), await fillDoc('zonalWorkOrder'))
        // Same "blank, hand-writable date" policy as the EE Work Order / Agreement Bond bundle.
        if (seAgreementBondB64) await add(docName('seAgreementBond'), await fillDoc('seAgreementBond', { blankDate: true }))
        if (seAgreementNoteB64) await add(docName('seAgreementNote'), await fillDoc('seAgreementNote'))
        if (zonalConcludingAgreementB64) await add(docName('zonalConcludingAgreement'), await fillDoc('zonalConcludingAgreement'))
        if (zonalMemoEeB64) await add(docName('zonalMemoEe'), await fillDoc('zonalMemoEe'))
        if (contractDeedB64) await add(docName('contractDeed'), await fillDoc('contractDeed'))
      } else {
        // File Backer — the cover page — first in the bundle.
        if (fileBackerB64) await add(docName('fileBacker'), await fillDoc('fileBacker'))
        if (civilTenderB64) await add(docName('civilTender'), await fillDoc('civilTender'))
        if (forwardingSlipB64) await add(docName('forwardingSlip'), await fillDoc('forwardingSlip'))
        // Work Order & Agreement Bond print a blank, hand-writable date in the bundle.
        if (agreementB64) await add(docName('agreement'), await fillDoc('agreement', { blankDate: true }))
        if (qccIntimationB64) await add(docName('qccIntimation'), await fillDoc('qccIntimation'))
        if (workOrderB64) await add(docName('workOrder'), await fillDoc('workOrder', { blankDate: true }))
        if (noteReady) await add(noteSubmittedFileName(noteData?.workName), await api.noteSubmittedDocx(notePreviewHtml))
      }
      if (scheduleA)
        files.push({
          name: `Schedule A${fields.agencyName ? ` - ${fields.agencyName}` : ''}`,
          format: 'xlsx',
          scheduleATable: scheduleA,
          scheduleAMeta
        })
      if (files.length === 0) {
        setActionError('No documents are ready to download yet.')
        return
      }
      setBundleProgress({ phase: 'saving', done: 0, total: files.length })
      const res = await api.exportAgreementBundle(files)
      const unresolvedNote =
        unresolvedWarnings.length > 0
          ? ` Warning: ${unresolvedWarnings.length} document(s) have an unfilled field — ${unresolvedWarnings.join('; ')}.`
          : ''
      if (!res) {
        setActionSaved('Cancelled.')
      } else if (res.failed.length > 0) {
        // Some documents (usually PDFs needing LibreOffice) didn't convert — say
        // which, instead of silently dropping them from the folder.
        setActionSaved(`Saved ${res.written.length} document(s).`)
        setActionError(
          `${res.failed.length} document(s) could not be saved: ${res.failed.join(', ')}. ` +
            `PDFs need LibreOffice installed — install it from libreoffice.org, or download those as Word (.docx) instead.` +
            unresolvedNote
        )
      } else {
        setActionSaved(`Saved ${res.written.length} document(s) to the chosen folder.`)
        if (unresolvedNote) setActionError(unresolvedNote.trim())
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      unsubscribe()
      setBundleProgress(null)
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
    // noteSubmittedFromRow now takes the uploads (pdf/notice) directly and
    // prefers them over the matched row's own columns for every identifying
    // fact (name, ECV, tender %, contract value, agency, NIT No/date) — a
    // Works List match is name-similarity based (falls back to embeddings
    // when there's no exact text match) and can land on a different, merely
    // similar-sounding work, so the row is only trusted for what the uploads
    // don't carry (Circle, Financial Year, estimate). Same precedence as
    // every other document in this workspace (see deriveFields).
    const seed = noteSubmittedFromRow(selectedRow, pdfEval ?? {}, notice ?? {}, table?.rows[0]?.['Circle'] ?? '')
    if (allBidders.length > 0) {
      seed.bidders = allBidders
      const l1 = allBidders[0]
      seed.l1Name = l1.name
      seed.l1PctText = l1.pct
      seed.l1Tcv = l1.tcv
      // Drive the EMD/ASD clause off the L1 sheet's own quoted percentage (the
      // source of truth), not just the Works List column — so ASD reflects even
      // when the row's Tender Percentage is blank or formatted (e.g. "32 % Less").
      const mag = tenderPctMagnitude(l1.pct)
      if (mag != null) seed.l1PctNumber = mag
    }
    // The uploaded Balance EMD receipt's Online Receipt No / date, when one's
    // been read — noteSubmittedFromRow itself always seeds these blank, since
    // neither the Works List, the L1 sheet, nor the Intimation ever carries
    // them (see handleEmdReceiptFile).
    if (emdReceiptInfo) {
      if (emdReceiptInfo.receiptNo) seed.receiptNo = emdReceiptInfo.receiptNo
      if (emdReceiptInfo.receiptDate) seed.receiptDate = emdReceiptInfo.receiptDate
    }
    setNoteData(seed)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowIndex, table, allBidders, pdfEval, notice, deriveFromUploads, selectedRow, emdReceiptInfo])

  const notePreviewHtml = useMemo(() => (noteData ? buildNoteSubmittedHtml(noteData) : ''), [noteData])
  // Same gate as the Work Order / Agreement tiles: don't build the Note
  // Submitted while the Online Intimation and L1 selection form disagree on the
  // work/agency, the L1's work isn't in the Works List (the note is seeded
  // from the selected row), or the L1's circle isn't this office's circle —
  // otherwise it would show the wrong work/agency or the wrong circle's note.
  const noteReady = !!noteData && !!pdfEval && !workMismatch && !circleMismatch

  // Defence-in-depth: Note Submitted is seeded independently of every other
  // document (its own effect, above) from the same upload/matched row — if it
  // ever falls out of sync (e.g. a missed re-seed after the match changes),
  // this catches it before the office downloads mismatched documents instead
  // of silently shipping them, the way today's "wrong work in Note Submitted"
  // bug did.
  const noteWorkMismatch =
    !!noteData?.workName && !!fields.nameOfWork && normWorkName(noteData.workName) !== normWorkName(fields.nameOfWork)
  const noteEstimateMismatch =
    !!noteData?.estimateLakhs && !!fields.estimateLakhs && noteData.estimateLakhs.trim() !== fields.estimateLakhs.trim()

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
      // Only consumed by the SE (Zone-office) Schedule A's signature block —
      // the EE template's replaceOfficeText never reads it.
      corporation: fields.corporation,
      ecvAmount: ecvNum != null && Number.isFinite(ecvNum) ? `${indianDigitGroups(ecvNum)}/-` : base.ecvAmount,
      contractAmount:
        contractNum != null && Number.isFinite(contractNum) ? `${indianDigitGroups(contractNum)}/-` : base.contractAmount
    }
  }, [scheduleAOnly, detectedWorkName, boq, scheduleA, uploadedIsEstimate, table, selectedRow, fields])

  // Always built, even with no estimate uploaded yet — an empty item table
  // under the filled preamble/meta (Name of work, Estimate/ECV/Contract
  // Amount, Contractor) so the tile shows automatically like the other
  // documents, instead of waiting on its own separate upload. Uploading the
  // estimate later fills in the item rows.
  const scheduleAPreview = useMemo(
    () =>
      buildScheduleARows(
        scheduleA ? rowsToScheduleAItems(scheduleA) : [],
        scheduleAMeta,
        seMode ? 'Superintending Engineer' : 'Executive Engineer'
      ),
    [scheduleA, scheduleAMeta, seMode]
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
    if (scheduleAOnly) onContent?.(!!boq || !!scheduleAError)
  }, [scheduleAOnly, boq, scheduleAError, onContent])

  async function downloadScheduleA() {
    if (!scheduleA) return
    setScheduleABusy(true)
    setScheduleAError(null)
    try {
      const base = boq ? stripBoqWord(stripExt(boq.name)) : ''
      const suggestedName = base ? `${base} Schedule A` : 'Schedule A'
      // A Zone-level (SE) office must never save the EE-signed Schedule A — its
      // own template names the Superintending Engineer's Zone, not the work's
      // Executive Engineer / Circle.
      const savedPath = seMode
        ? await api.exportSeScheduleA(scheduleA, suggestedName, scheduleAMeta)
        : await api.exportScheduleA(scheduleA, suggestedName, scheduleAMeta)
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
  // scheduleAPreview is now always non-null (it shows an empty framework
  // before any estimate is uploaded), so `scheduleA` — the actual uploaded
  // BOQ — is what signals real Schedule-A progress here.
  const anyOutput = docsReady || !!scheduleA || noteReady

  // Schedule-A Tools tile (scheduleAOnly + autoOpen): show nothing until an
  // estimate/BOQ is actually picked (or reading it fails) — the mount effect
  // fires the folder, so clicking the tile opens it directly with no
  // placeholder panel.
  if (autoOpen && scheduleAOnly && !boq && !scheduleAError) return null

  const bundleProgressLabel = bundleProgress
    ? `${bundleProgress.phase === 'preparing' ? 'Preparing' : 'Saving'} ${bundleProgress.done} of ${bundleProgress.total}…`
    : 'Preparing…'

  const downloadAllButton = docsReady && !only && (
    <div className="download-all-group">
      <div className="format-toggle" role="group" aria-label="Download format">
        <button
          type="button"
          className={bundleFormat === 'docx' ? 'on' : ''}
          disabled={busy === 'bundle'}
          onClick={() => setBundleFormat('docx')}
        >
          Word
        </button>
        <button
          type="button"
          className={bundleFormat === 'pdf' ? 'on' : ''}
          disabled={busy === 'bundle'}
          onClick={() => setBundleFormat('pdf')}
        >
          PDF
        </button>
      </div>
      <button
        className="primary"
        disabled={busy === 'bundle'}
        onClick={downloadAll}
        title={`Every document as ${bundleFormat === 'pdf' ? 'PDF' : 'Word'}, into one folder you choose`}
      >
        <IconDownload /> {busy === 'bundle' ? bundleProgressLabel : 'Download all documents'}
      </button>
    </div>
  )

  return (
    <div className={only ? 'wo-compact' : 'card'}>
      {headerMounted && headerActionRef?.current && createPortal(downloadAllButton, headerActionRef.current)}
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
              <IconTable /> {boq ? 'Change Technical sanctioned Estimate' : 'Upload Technical sanctioned Estimate'}
            </button>
          )}
          {!scheduleAOnly && !only && (
            <button
              className="primary upload-btn"
              onClick={() => emdReceiptInputRef.current?.click()}
              disabled={busy === 'emdReceipt'}
            >
              <IconFolder /> {busy === 'emdReceipt' ? 'Reading receipt…' : 'Upload Balance EMD Payment (PDF or photo)'}
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
          <input
            ref={emdReceiptInputRef}
            type="file"
            accept="application/pdf,.pdf,image/*"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleEmdReceiptFile(file)
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
          <div className="wo-manual-form">
            <label className="wo-date-field">
              <span>Agreement date</span>
              <input type="date" value={agreementDate} onChange={(e) => setAgreementDate(e.target.value)} />
            </label>
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
            <span className="estimate-hint wo-manual-wide">For the Forwarding Slip &amp; Tender Document.</span>
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
      {(noteWorkMismatch || noteEstimateMismatch) && (
        <div className="notice error">
          <IconWarn /> Note Submitted doesn’t match the other documents in this workspace
          {noteWorkMismatch && (
            <>
              {' '}
              — Name of Work: “{noteData?.workName}” here vs “{fields.nameOfWork}” everywhere else
            </>
          )}
          {noteEstimateMismatch && (
            <>
              {' '}
              — Amount of Estimate: {noteData?.estimateLakhs} Lakhs here vs {fields.estimateLakhs} Lakhs everywhere else
            </>
          )}
          . Re-pick the work or re-upload the L1 sheet before downloading.
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

      {anyOutput ? (
        <>
          {noWorks && (
            <div className="notice">
              Works List is empty — supporting details (Wincode, estimate, TS No/date, …) will be blank until you add
              the work there; the documents below still fill from your uploads.
            </div>
          )}
          {!headerMounted && downloadAllButton}
          <div className="wo-tiles">
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
          {noteReady && !only && !seMode && (
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
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('zonalWorkOrder')}>
              <div className="wo-tile-preview">
                <div ref={zwoTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.zonalWorkOrder}</div>
            </button>
          )}
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('zonalConcludingAgreement')}>
              <div className="wo-tile-preview">
                <div ref={zcaTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.zonalConcludingAgreement}</div>
            </button>
          )}
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('zonalMemoEe')}>
              <div className="wo-tile-preview">
                <div ref={zmeTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.zonalMemoEe}</div>
            </button>
          )}
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('seAgreementNote')}>
              <div className="wo-tile-preview">
                <div ref={seNoteTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.seAgreementNote}</div>
            </button>
          )}
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('seAgreementBond')}>
              <div className="wo-tile-preview">
                <div ref={seBondTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.seAgreementBond}</div>
            </button>
          )}
          {docsReady && seMode && (
            <button className="wo-tile" onClick={() => openDoc('contractDeed')}>
              <div className="wo-tile-preview">
                <div ref={contractDeedTileRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.contractDeed}</div>
            </button>
          )}
          {!only && !scheduleAMismatch && (
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
      ) : noWorks ? (
        <div className="notice">
          Add works to the Works List first, or upload the L1 selection form / Online Intimation above — either fills
          the outputs.
        </div>
      ) : null}

      {datePromptOpen &&
        createPortal(
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
          </div>,
          document.body
        )}

      {expanded &&
        createPortal(
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
                  {!scheduleA && (
                    <span className="estimate-hint" style={{ marginRight: 'auto' }}>
                      Upload the Technical sanctioned Estimate to fill in the item rows.
                    </span>
                  )}
                  <button className="primary" onClick={downloadScheduleA} disabled={scheduleABusy || !scheduleA}>
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
          </div>,
          document.body
        )}
    </div>
  )
}
