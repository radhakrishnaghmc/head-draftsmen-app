// Shared IPC contract between the Electron main process and the renderer.
import type {
  ExcelTable,
  TenderQuery,
  TenderResult,
  PersistedState
} from '../core/types'
import type { SheetGrid } from '../core/sheet'
import type { SheetPreview } from '../core/excel'
import type { CalendarData } from '../core/calendar'
import type { ScheduleAMeta } from '../core/scheduleA'
import type { TenderNoticeInput } from '../core/tenderNotice'
import type { CellEdit } from '../core/technicalSanction'
import type { BidDocumentInput } from '../core/bidDocument'
import type { LoginResult } from '../core/auth'
import type { DeviationItem, DeviationMeta } from '../core/deviationTemplate'
import type { DocBlock } from '../core/docxBuilder'
import type { OcrPage } from '../core/ocrReconstruct'
import type { EvaluationSheetInput } from '../core/evaluationSheet'
import type { PlaceholderMatch } from '../core/createDocument'
import type { EstimateWorkItem } from '../core/estimateExtract'
import type { DetailedEstimateMeta } from '../core/estimateTemplate'
import type { MaterialTotals } from '../core/materialEstimate'
import type { MaterialEstimateMeta } from '../core/materialTemplate'

export type ManualCheckResult = 'update-available' | 'up-to-date' | 'error' | 'dev-mode'

export interface UpdateProgress {
  percent: number
  transferred: number
  total: number
  bytesPerSecond: number
}

export interface SplitProgress {
  /** Sheets finished so far. */
  done: number
  /** Total sheets in the workbook. */
  total: number
  /** Name of the sheet currently being written. */
  sheet: string
}

export interface AgreementBundleProgress {
  /** Files finished so far (written or failed). */
  done: number
  /** Total files in the bundle. */
  total: number
  /** Name of the file currently being written/converted. */
  name: string
}

export const IPC = {
  pickExcels: 'dialog:pickExcels',
  pickExcelGrids: 'dialog:pickExcelGrids',
  pickEstimateGrid: 'dialog:pickEstimateGrid',
  pickDataSheet: 'dialog:pickDataSheet',
  ocrEstimatePhotos: 'data:ocrEstimatePhotos',
  openPath: 'shell:openPath',
  revealItem: 'shell:revealItem',
  defaultDir: 'shell:defaultDir',
  getAppVersion: 'app:getVersion',
  fetchCalendar: 'calendar:fetch',
  login: 'auth:login',
  logout: 'auth:logout',
  importFromLink: 'data:importFromLink',
  importAllSheetsFromLink: 'data:importAllSheetsFromLink',
  exportTable: 'data:exportTable',
  exportScheduleA: 'data:exportScheduleA',
  exportEvaluationSheet: 'data:exportEvaluationSheet',
  exportAgreementBundle: 'doc:exportAgreementBundle',
  agreementBundleProgress: 'doc:agreementBundleProgress',
  exportBoq: 'data:exportBoq',
  exportBoqBatch: 'data:exportBoqBatch',
  pickWorkbookForSplit: 'tools:pickWorkbookForSplit',
  splitWorkbook: 'tools:splitWorkbook',
  splitProgress: 'tools:splitProgress',
  pickPdfsForMerge: 'tools:pickPdfsForMerge',
  mergePdfs: 'tools:mergePdfs',
  pickPdfForSplit: 'tools:pickPdfForSplit',
  splitPdf: 'tools:splitPdf',
  savePdf: 'tools:savePdf',
  savePdfsToFolder: 'tools:savePdfsToFolder',
  docxToPdf: 'tools:docxToPdf',
  mergeDocx: 'tools:mergeDocx',
  splitDocxSections: 'tools:splitDocxSections',
  saveDocxsToFolder: 'tools:saveDocxsToFolder',
  ocrGpsOverlay: 'tools:ocrGpsOverlay',
  ocrPhotosToLines: 'tools:ocrPhotosToLines',
  savePhotosAsWord: 'tools:savePhotosAsWord',
  savePhotosAsExcel: 'tools:savePhotosAsExcel',
  savePdfAsWord: 'tools:savePdfAsWord',
  saveWordDoc: 'tools:saveWordDoc',
  ocrPhotosToLayout: 'tools:ocrPhotosToLayout',
  saveRowsAsExcel: 'tools:saveRowsAsExcel',
  exportDeviation: 'data:exportDeviation',
  exportDetailedEstimate: 'data:exportDetailedEstimate',
  exportMaterialEstimate: 'data:exportMaterialEstimate',
  generateTenderNotice: 'data:generateTenderNotice',
  previewTenderNotice: 'data:previewTenderNotice',
  generateBidDocument: 'data:generateBidDocument',
  generateBidDocumentBatch: 'data:generateBidDocumentBatch',
  previewBidDocument: 'data:previewBidDocument',
  generateTechnicalSanction: 'data:generateTechnicalSanction',
  searchTenders: 'tenders:search',
  embedTexts: 'ai:embedTexts',
  listDocumentParagraphs: 'doc:listDocumentParagraphs',
  saveDocumentEdits: 'doc:saveDocumentEdits',
  findPlaceholdersInDocument: 'doc:findPlaceholdersInDocument',
  fillPlaceholdersInDocument: 'doc:fillPlaceholdersInDocument',
  bakeFixedPlaceholdersInDocument: 'doc:bakeFixedPlaceholdersInDocument',
  exportCreatedDocument: 'doc:exportCreatedDocument',
  printCreatedDocument: 'doc:printCreatedDocument',
  noteSubmittedDocx: 'doc:noteSubmittedDocx',
  intimationTemplate: 'doc:intimationTemplate',
  workOrderTemplate: 'doc:workOrderTemplate',
  fileBackerTemplate: 'doc:fileBackerTemplate',
  agreementTemplate: 'doc:agreementTemplate',
  qccIntimationTemplate: 'doc:qccIntimationTemplate',
  forwardingSlipTemplate: 'doc:forwardingSlipTemplate',
  civilTenderTemplate: 'doc:civilTenderTemplate',
  zonalWorkOrderTemplate: 'doc:zonalWorkOrderTemplate',
  zonalConcludingAgreementTemplate: 'doc:zonalConcludingAgreementTemplate',
  zonalMemoEeTemplate: 'doc:zonalMemoEeTemplate',
  seAgreementNoteTemplate: 'doc:seAgreementNoteTemplate',
  seAgreementBondTemplate: 'doc:seAgreementBondTemplate',
  seContractDeedTemplate: 'doc:seContractDeedTemplate',
  exportSeScheduleA: 'data:exportSeScheduleA',
  loaSeTemplate: 'doc:loaSeTemplate',
  tsNoteTemplate: 'doc:tsNoteTemplate',
  eligibilityCriteriaTemplate: 'doc:eligibilityCriteriaTemplate',
  loadState: 'state:load',
  saveState: 'state:save',
  remoteStateUpdate: 'state:remoteUpdate',
  updateDownloaded: 'update:downloaded',
  updateProgress: 'update:progress',
  restartToUpdate: 'update:restart',
  updateInstallError: 'update:installError',
  checkForUpdates: 'update:check'
} as const

/** The typed API exposed to the renderer via contextBridge (window.docugen). */
export interface DocuGenApi {
  pickExcels(): Promise<ExcelTable[]>
  pickExcelGrids(): Promise<SheetGrid[]>
  pickEstimateGrid(): Promise<SheetGrid | null>
  pickDataSheet(): Promise<SheetGrid[] | null>
  /** Runs local OCR on photos of a paper estimate (in page order) and reconstructs one combined grid, best-effort — always review before exporting. */
  ocrEstimatePhotos(dataUrls: string[]): Promise<SheetGrid>
  openPath(target: string): Promise<void>
  revealItem(target: string): Promise<void>
  defaultDir(): Promise<string>
  getAppVersion(): Promise<string>
  fetchCalendar(force?: boolean): Promise<CalendarData>
  login(loginId: string, password: string): Promise<LoginResult>
  /** Releases this device's session slot so another device can sign in. */
  logout(): Promise<void>
  importFromLink(url: string): Promise<ExcelTable>
  /** Downloads every sheet of a monitoring-format workbook link (one sheet per circle). */
  importAllSheetsFromLink(url: string): Promise<SheetGrid[]>
  exportTable(table: ExcelTable, suggestedName: string): Promise<string | null>
  exportScheduleA(table: ExcelTable, suggestedName: string, meta?: ScheduleAMeta): Promise<string | null>
  /** Build the Bid Capacity Evaluation Sheet (one column per participating bidder) and save it as .xlsx. Returns the written path, or null if cancelled. */
  exportEvaluationSheet(input: EvaluationSheetInput, suggestedName: string): Promise<string | null>
  /** Save several agreement-workspace documents into ONE chosen folder, each in its given format (docx/pdf/xlsx). Returns the written paths plus any that failed (e.g. PDF conversion), or null if cancelled. */
  exportAgreementBundle(files: AgreementBundleFile[]): Promise<{ written: string[]; failed: string[] } | null>
  /** Fires as each file in an exportAgreementBundle call is written/converted, so the UI can show live progress instead of a single static "Preparing…". Returns an unsubscribe function. */
  onAgreementBundleProgress(callback: (progress: AgreementBundleProgress) => void): () => void
  exportBoq(table: ExcelTable, suggestedName: string, workName?: string): Promise<string | null>
  exportBoqBatch(
    entries: { table: ExcelTable; suggestedName: string; workName?: string }[]
  ): Promise<string[] | null>
  /** Tool: pick a multi-sheet workbook to split; returns its path and a lightweight per-sheet preview (name + a corner of cells, for the sheet tiles), or null if cancelled. */
  pickWorkbookForSplit(): Promise<{ path: string; name: string; sheets: SheetPreview[] } | null>
  /** Tool: split the chosen workbook into one .xlsx per sheet (named after the tab) in a folder the user picks. `sheetNames` limits it to those sheets; null/empty separates all. Returns the folder and saved file paths, or null if cancelled. */
  splitWorkbook(srcPath: string, sheetNames: string[] | null): Promise<{ dir: string; files: string[] } | null>
  /** Fires as each sheet is written while splitWorkbook runs, so the UI can show a progress bar. Returns an unsubscribe function. */
  onSplitProgress(callback: (progress: SplitProgress) => void): () => void
  /** Tool: pick two or more PDFs to merge; returns each one's path, name and page count (in the order picked), or null if cancelled. */
  pickPdfsForMerge(): Promise<{ path: string; name: string; pages: number }[] | null>
  /** Tool: merge the given PDFs (in this order) into one file the user names/saves. Returns the saved file path, or null if cancelled. */
  mergePdfs(srcPaths: string[]): Promise<{ file: string } | null>
  /** Tool: pick a PDF to separate; returns its path, name and page count, or null if cancelled. */
  pickPdfForSplit(): Promise<{ path: string; name: string; pages: number } | null>
  /** Tool: split a PDF into one file per range (1-based inclusive) in a folder the user picks; null `ranges` = every page separately. Returns the folder and saved files, or null if cancelled. */
  splitPdf(srcPath: string, ranges: [number, number][] | null): Promise<{ dir: string; files: string[] } | null>
  /** Tool (PDF workspace): save already-built PDF bytes to a file the user names/picks. Returns the written path, or null if cancelled. */
  savePdf(bytes: Uint8Array, suggestedName: string): Promise<{ file: string } | null>
  /** Tool (PDF workspace): save several already-built PDFs (name + bytes) into ONE folder the user picks. Returns the folder and written paths, or null if cancelled. */
  savePdfsToFolder(files: { name: string; bytes: Uint8Array }[]): Promise<{ dir: string; files: string[] } | null>
  /** Tool (Word workspace): convert one .docx (raw bytes) to PDF via LibreOffice, for a page-level preview. Returns the PDF bytes; throws a clear error if LibreOffice isn't installed. */
  docxToPdf(docxBytes: Uint8Array): Promise<Uint8Array>
  /** Tool (Word workspace): merge several .docx files (raw bytes, in order) into one .docx and save it via a dialog. Returns the saved path, or null if cancelled. */
  mergeDocx(docxBytesList: Uint8Array[]): Promise<{ file: string } | null>
  /** Tool (Word workspace): split one .docx (raw bytes) into one .docx per page at its page breaks; each keeps full formatting. Returns the section .docx byte arrays (a single item when there are no page breaks). */
  splitDocxSections(docxBytes: Uint8Array): Promise<Uint8Array[]>
  /** Tool (Word workspace): save several .docx files (name + bytes) into ONE folder the user picks. Returns the folder and written paths, or null if cancelled. */
  saveDocxsToFolder(files: { name: string; bytes: Uint8Array }[]): Promise<{ dir: string; files: string[] } | null>
  /** Tool (GPS Photos): OCR the GPS overlay stamped on a photo (multi-threshold passes) and return its text lines, for parsing coordinates out of. Used only when the photo has no EXIF GPS. */
  ocrGpsOverlay(imageBytes: Uint8Array): Promise<string[]>
  /** Tool (Photos/PDF → Word/Excel): OCR each page image (data URLs, in order) and return all recognised lines in reading order, blank-line-separated between pages. Best-effort — always reviewed/edited before export. */
  ocrPhotosToLines(dataUrls: string[]): Promise<string[]>
  /** Tool (Photos/PDF → Word/Excel): save the reviewed text as a .docx (one paragraph per line). Returns the written path, or null if cancelled. */
  savePhotosAsWord(text: string, suggestedName: string): Promise<string | null>
  /** Tool (Photos/PDF → Word/Excel): save the reviewed text as an .xlsx (one row per line, split into columns on wide gaps). Returns the written path, or null if cancelled. */
  savePhotosAsExcel(text: string, suggestedName: string): Promise<string | null>
  /** Tool (Photos/PDF → Word): convert uploaded PDF(s) to a LAYOUT-PRESERVING .docx via LibreOffice (writer_pdf_import) — tables/borders/positioning kept as editable Word content — merging multiple PDFs in order. Returns the written path, or null if cancelled. Needs LibreOffice installed. */
  savePdfAsWord(pdfs: { name: string; bytes: Uint8Array }[], suggestedName: string): Promise<string | null>
  /** Tool (Photos/PDF → Word): save a reconstructed doc-model (real paragraphs/tables rebuilt from a text-PDF's geometry, or plain OCR lines) as a directly-built, Word-valid .docx (core/docxBuilder — not html-to-docx). Returns the written path, or null if cancelled. */
  saveWordDoc(blocks: DocBlock[], suggestedName: string): Promise<string | null>
  /** Tool (Photos/PDF → Word/Excel): OCR each page image and return its text lines WITH bounding boxes, per page — the input to offline image-based table reconstruction (core/ocrReconstruct). */
  ocrPhotosToLayout(dataUrls: string[]): Promise<OcrPage[]>
  /** Tool (Photos/PDF → Excel): save a reconstructed 2-D grid of cells as an .xlsx. Returns the written path, or null if cancelled. */
  saveRowsAsExcel(rows: string[][], suggestedName: string): Promise<string | null>
  exportDeviation(items: DeviationItem[], meta: DeviationMeta, suggestedName: string): Promise<string | null>
  exportDeviation(items: DeviationItem[], meta: DeviationMeta, suggestedName: string): Promise<string | null>
  /** Builds and saves a full "Detailed and Abstract Estimate" workbook (letterhead, item table, standard surcharge cascade, signature block) from scratch — not a bare Sl No/Description/Qty/Rate/Amount table. */
  exportDetailedEstimate(
    items: EstimateWorkItem[],
    meta: DetailedEstimateMeta,
    suggestedName: string
  ): Promise<string | null>
  /** Fills the bundled Material Estimation Template (Work Name/Department/District/ECV header + the 7 standard material rows) from computed material totals — see core/materialEstimate.ts. */
  exportMaterialEstimate(
    totals: MaterialTotals,
    meta: MaterialEstimateMeta,
    suggestedName: string
  ): Promise<string | null>
  generateTenderNotice(input: TenderNoticeInput, suggestedName?: string): Promise<string | null>
  previewTenderNotice(input: TenderNoticeInput): Promise<string>
  /** Save one filled Bid Document (one save dialog per call — call once per work). */
  generateBidDocument(input: BidDocumentInput, suggestedName: string): Promise<string | null>
  /** Save every work's Bid Document from one batch into a single chosen folder — one folder picker for the whole batch instead of a save dialog per work. */
  generateBidDocumentBatch(
    entries: { input: BidDocumentInput; suggestedName: string }[]
  ): Promise<string[] | null>
  previewBidDocument(input: BidDocumentInput): Promise<string>
  generateTechnicalSanction(
    estimatePath: string,
    sheetName: string,
    edits: CellEdit[],
    suggestedName: string,
    rateAnalysisRows?: (string | number)[][]
  ): Promise<string | null>
  searchTenders(query: TenderQuery): Promise<TenderResult>
  embedTexts(texts: string[]): Promise<number[][]>
  /** Plain text of every paragraph in a base64 .docx, in document order — used to diff a user's typed edits back against the original before saving. */
  listDocumentParagraphs(docxBase64: string): Promise<string[]>
  /** Rewrites only the paragraphs that actually changed, preserving every other run's formatting — returns the updated base64 .docx. */
  saveDocumentEdits(docxBase64: string, edits: { index: number; text: string }[]): Promise<string>
  /** Every distinct {{Label}} found in a base64 .docx's body paragraphs. */
  findPlaceholdersInDocument(docxBase64: string): Promise<string[]>
  /** Replaces every {{Label}} occurrence with its resolved row value (blank if unresolved) — returns the filled base64 .docx. */
  fillPlaceholdersInDocument(
    docxBase64: string,
    resolved: PlaceholderMatch[],
    row: Record<string, string>
  ): Promise<string>
  /** Bakes only the given labels (Zone/Circle/CNO) into a base64 .docx, leaving every other placeholder for later per-row resolution. */
  bakeFixedPlaceholdersInDocument(docxBase64: string, values: Record<string, string>): Promise<string>
  /** Save the filled document as Word and/or PDF (one save dialog per chosen format). */
  exportCreatedDocument(
    docxBase64: string,
    suggestedName: string,
    formats: ('docx' | 'pdf')[]
  ): Promise<{ file: string; format: 'docx' | 'pdf' }[] | null>
  /** Open the OS print dialog directly against the already-rendered document HTML (docx-preview's own output, captured by the caller). */
  printCreatedDocument(renderedHtml: string): Promise<void>
  /** Converts built Note Submitted HTML into a base64 .docx, for export via exportCreatedDocument. */
  noteSubmittedDocx(html: string): Promise<string>
  /** Reads the bundled Intimation format (.docx) and returns it base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  intimationTemplate(): Promise<string>
  /** Reads the bundled Work Order format (.docx) and returns it base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  workOrderTemplate(): Promise<string>
  /** Reads the bundled File Backer format (.docx) — the file's cover page — base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  fileBackerTemplate(): Promise<string>
  /** Reads the bundled Agreement format (.docx) and returns it base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  agreementTemplate(): Promise<string>
  /** Reads the bundled QCC Intimation letter (.docx) — the Dy.EE's request to Quality Control to inspect a starting work — base64-encoded for filling its {{placeholders}}. */
  qccIntimationTemplate(): Promise<string>
  /** Reads the bundled Forwarding Slip format (.docx) and returns it base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  forwardingSlipTemplate(): Promise<string>
  /** Reads the bundled full Civil Tender Document (.docx) — the 41-page NIT/tender document whose page 1 is the Forwarding Slip — returns it base64-encoded for filling its {{placeholders}}. */
  civilTenderTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Work Order format (.docx), for the Work Order/Agreement page when the office is a Zone with no Circle. Base64-encoded, for filling its {{placeholders}}. */
  zonalWorkOrderTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Memo Concluding Agreement format (.docx). Base64-encoded, for filling its {{placeholders}}. */
  zonalConcludingAgreementTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Memo forwarding the Agreement Bond to the EE (.docx). Base64-encoded, for filling its {{placeholders}}. */
  zonalMemoEeTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Agreement Put-up Note (.docx) — the cover note requesting sign-off on the agreement + memo to EE. Base64-encoded, for filling its {{placeholders}}. */
  seAgreementNoteTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Agreement Bond paper (.docx) — the "A G R E E M E N T" cover page signed by the Superintending Engineer. Base64-encoded, for filling its {{placeholders}}. */
  seAgreementBondTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Contract Deed (.docx) — the Form No.6 legal deed. Base64-encoded, for filling its {{placeholders}}. */
  seContractDeedTemplate(): Promise<string>
  /** Fills and saves the Zone-level (SE office) Schedule A / BOQ workbook — same item table as exportScheduleA, but its preamble and signature name the Superintending Engineer's Zone office. Prompts for a save location; returns the saved path, or null if cancelled. */
  exportSeScheduleA(table: ExcelTable, suggestedName: string, meta?: ScheduleAMeta): Promise<string | null>
  /** Reads the bundled Superintending-Engineer LOA format (.docx) — used for the Give Intimation letter when the office is zone-level (a Zone with no Circle). `reserved` picks the SC/ST-reserved variant (no EMD balance item). Returns it base64-encoded, for filling its {{placeholders}}. */
  loaSeTemplate(reserved: boolean): Promise<string>
  /** Reads the bundled Zone-level (SE office) TS (Technical Sanction) Note (.docx) — issued alongside the LOA. Base64-encoded, for filling its {{placeholders}}. */
  tsNoteTemplate(): Promise<string>
  /** Reads the bundled Zone-level (SE office) Eligibility Criteria note (.docx) — issued alongside the LOA. Base64-encoded, for filling its {{placeholders}}. */
  eligibilityCriteriaTemplate(): Promise<string>
  loadState(): Promise<PersistedState | null>
  /** Persist the workspace. `skipCloud` writes only to local disk and does NOT push to the cloud — used when the change being saved *came from* a remote sync, so it isn't echoed straight back and made to ping-pong between concurrent sessions. */
  saveState(state: PersistedState, skipCloud?: boolean): Promise<void>
  /** Fires when the other signed-in device changes the workspace, so this one can merge it in live. Returns an unsubscribe function. */
  onRemoteStateUpdate(callback: (partial: Partial<PersistedState>) => void): () => void
  /** Fires once a new version has finished downloading in the background and is ready to install. Returns an unsubscribe function. */
  onUpdateDownloaded(callback: () => void): () => void
  /** Fires repeatedly while a new version downloads in the background. Returns an unsubscribe function. */
  onUpdateProgress(callback: (progress: UpdateProgress) => void): () => void
  /** Quits and installs the already-downloaded update, then relaunches. */
  restartToUpdate(): void
  /** Fires if restartToUpdate's install attempt fails right away (e.g. an unsigned build on macOS) instead of the app just silently not restarting. Returns an unsubscribe function. */
  onUpdateInstallError(callback: (message: string) => void): () => void
  /** Checks for an update on demand (the small update icon in the sidebar) and reports the outcome, rather than checking silently. */
  checkForUpdates(): Promise<ManualCheckResult>
}

export type { ExcelTable }
export type { SheetGrid }
export type { SheetPreview }
export type { TenderQuery, TenderResult, PersistedState }
export type { ScheduleAMeta }

/** One file in an agreement-workspace batch export (see exportAgreementBundle). */
export interface AgreementBundleFile {
  /** Base file name (no extension) — the format's extension is appended. */
  name: string
  format: 'docx' | 'pdf' | 'xlsx'
  /** Filled .docx (base64) — required for 'docx' and 'pdf' (converted to PDF in main). */
  docxBase64?: string
  /** Schedule A source — required for 'xlsx' (the workbook is built in main). */
  scheduleATable?: ExcelTable
  scheduleAMeta?: ScheduleAMeta
}
export type { TenderNoticeInput }
export type { CellEdit }
export type { BidDocumentInput }
export type { LoginResult }
export type { PlaceholderMatch }
export type { DeviationItem, DeviationMeta }
export type { EstimateWorkItem }
export type { DetailedEstimateMeta }
export type { MaterialTotals }
export type { MaterialEstimateMeta }