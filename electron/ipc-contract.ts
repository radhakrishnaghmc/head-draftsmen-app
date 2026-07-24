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
  generateTenderNotice: 'data:generateTenderNotice',
  previewTenderNotice: 'data:previewTenderNotice',
  generateBidDocument: 'data:generateBidDocument',
  previewBidDocument: 'data:previewBidDocument',
  generateTechnicalSanction: 'data:generateTechnicalSanction',
  searchTenders: 'tenders:search',
  embedTexts: 'ai:embedTexts',
  exportCreatedDocument: 'doc:exportCreatedDocument',
  printCreatedDocument: 'doc:printCreatedDocument',
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
  generateTenderNotice(input: TenderNoticeInput, suggestedName?: string): Promise<string | null>
  previewTenderNotice(input: TenderNoticeInput): Promise<string>
  /** Save one filled Bid Document (one save dialog per call — call once per work). */
  generateBidDocument(input: BidDocumentInput, suggestedName: string): Promise<string | null>
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
  /** Save the filled document as Word and/or PDF (one save dialog per chosen format). */
  exportCreatedDocument(
    html: string,
    suggestedName: string,
    formats: ('docx' | 'pdf')[]
  ): Promise<{ file: string; format: 'docx' | 'pdf' }[] | null>
  /** Open the OS print dialog directly against the filled document. */
  printCreatedDocument(html: string): Promise<void>
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
export type { DeviationItem, DeviationMeta }