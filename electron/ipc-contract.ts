// Shared IPC contract between the Electron main process and the renderer.
import type {
  ExcelTable,
  TenderQuery,
  TenderResult,
  PersistedState
} from '../core/types'
import type { SheetGrid } from '../core/sheet'
import type { CalendarData } from '../core/calendar'
import type { ScheduleAMeta } from '../core/scheduleA'
import type { TenderNoticeInput } from '../core/tenderNotice'
import type { CellEdit } from '../core/technicalSanction'
import type { BidDocumentInput } from '../core/bidDocument'
import type { LoginResult } from '../core/auth'
import type { DeviationItem, DeviationMeta } from '../core/deviationTemplate'
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

export const IPC = {
  pickExcels: 'dialog:pickExcels',
  pickExcelGrids: 'dialog:pickExcelGrids',
  pickEstimateGrid: 'dialog:pickEstimateGrid',
  pickDataSheet: 'dialog:pickDataSheet',
  pickTenderEvalFolder: 'dialog:pickTenderEvalFolder',
  readFileBase64: 'fs:readFileBase64',
  listDriveFolderPdfs: 'drive:listFolderPdfs',
  downloadDriveFile: 'drive:downloadFile',
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
  exportBoq: 'data:exportBoq',
  exportBoqBatch: 'data:exportBoqBatch',
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
  createDocumentFromClipboard: 'doc:createDocumentFromClipboard',
  listDocumentParagraphs: 'doc:listDocumentParagraphs',
  saveDocumentEdits: 'doc:saveDocumentEdits',
  findPlaceholdersInDocument: 'doc:findPlaceholdersInDocument',
  fillPlaceholdersInDocument: 'doc:fillPlaceholdersInDocument',
  bakeFixedPlaceholdersInDocument: 'doc:bakeFixedPlaceholdersInDocument',
  exportCreatedDocument: 'doc:exportCreatedDocument',
  printCreatedDocument: 'doc:printCreatedDocument',
  intimationTemplate: 'doc:intimationTemplate',
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
  /** Prompts for a folder of tender-evaluation PDFs and returns the folder path plus its .pdf file paths (top level), or null if cancelled. */
  pickTenderEvalFolder(): Promise<{ dir: string; files: string[] } | null>
  /** Reads a .pdf file at an absolute path and returns it base64-encoded — used to feed folder PDFs into the renderer's pdf.js text extractor. */
  readFileBase64(filePath: string): Promise<string>
  /** Lists the PDF files (id + name) in a public Google Drive folder link. */
  listDriveFolderPdfs(link: string): Promise<{ id: string; name: string }[]>
  /** Downloads one Google Drive file (by id) and returns it base64-encoded. */
  downloadDriveFile(id: string): Promise<string>
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
  exportBoq(table: ExcelTable, suggestedName: string, workName?: string): Promise<string | null>
  exportBoqBatch(
    entries: { table: ExcelTable; suggestedName: string; workName?: string }[]
  ): Promise<string[] | null>
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
  /** Reads Word's own RTF clipboard flavor and converts it to a real .docx via a local LibreOffice install — base64-encoded, or null if the clipboard has no rich content. */
  createDocumentFromClipboard(): Promise<string | null>
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
  /** Reads the bundled Intimation format (.docx) and returns it base64-encoded, for filling its {{placeholders}} via fillPlaceholdersInDocument. */
  intimationTemplate(): Promise<string>
  loadState(): Promise<PersistedState | null>
  saveState(state: PersistedState): Promise<void>
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
export type { TenderQuery, TenderResult, PersistedState }
export type { ScheduleAMeta }
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