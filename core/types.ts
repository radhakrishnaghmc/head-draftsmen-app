// Framework-agnostic shared types for DocuGen core logic.
// No Electron/DOM imports here so this module can be reused in a future web app.

/** A query against the Telangana e-procurement tender listing. */
export interface TenderQuery {
  /** 0-based row offset for server-side pagination. */
  start: number
  /** Page size. */
  length: number
  /** Listing bucket: 'current' (live), etc. */
  type: string
  /** Server-side keyword (Tender ID / IFB No / Name of Work). Empty = all. */
  search?: string
}

/** DataTables-style response, one tender per row (raw 10-column arrays). */
export interface TenderResult {
  data: string[][]
  total: number
}


export interface ExcelColumn {
  name: string
  /** Source file name the column came from. */
  source: string
}

export interface ExcelTable {
  id: string
  name: string
  path: string
  headers: string[]
  rows: Record<string, string>[]
}

/** A column name that appears in more than one uploaded Excel file. */
export interface Collision {
  column: string
  sources: string[]
}

/** How to resolve a collision: which source file wins for a given column. */
export type CollisionResolution = Record<string, string> // column -> winning source file

export interface MergedDataset {
  columns: ExcelColumn[]
  rows: Record<string, string>[]
  collisions: Collision[]
}

export interface TodoItem {
  id: string
  text: string
  /** ISO date (YYYY-MM-DD) the item was created — defaults to the day it was added. */
  createdDate: string
  /** ISO date (YYYY-MM-DD) the work should be completed by. */
  targetDate: string
  done: boolean
  /** ISO date (YYYY-MM-DD) the item was marked done — it stays visible through that day, then rolls off. */
  completedDate?: string
}

/** One point noted during scrutiny — can be ticked off once attended to. */
export interface MBScrutinyRemark {
  text: string
  done: boolean
}

/** One Measurement Book received for scrutiny — the record persists whether scrutiny is pending or done. */
export interface MBScrutinyItem {
  id: string
  /** Running register number, assigned in receipt order when the entry is added — fixed thereafter. */
  serialNo: number
  mbNo: string
  agencyName: string
  /** ISO date (YYYY-MM-DD) the MB was received for scrutiny. */
  receivedDate: string
  /** ISO date (YYYY-MM-DD) scrutiny is targeted to be completed by. */
  targetDate?: string
  done: boolean
  /** ISO date (YYYY-MM-DD) scrutiny was completed — set (and editable) when marked done. */
  completedDate?: string
  /** Objections/notes from scrutiny, one entry per point (not a single paragraph). */
  remarks?: MBScrutinyRemark[]
}

export interface TenderReminderItem {
  workName?: string
  tenderId?: string
  bidClosing?: string
}

export interface TenderReminder {
  id: string
  nitNo: string
  /** One entry per work found on the portal under this NIT number — usually one, but a single NIT can cover several works/lots. */
  items: TenderReminderItem[]
  status: 'pending' | 'found' | 'not-found'
  /** ISO date (YYYY-MM-DD) the reminder was created. */
  createdDate: string
}

/**
 * A saved document template (with {{Placeholder}} markers). `docx` is a
 * base64-encoded real .docx (OOXML) buffer — the source of truth for every
 * "Issue Document" run that fills it against a Works List row.
 */
export interface CreatedDocument {
  id: string
  name: string
  docx: string
  createdDate: string
}

/** One work item within a Bid Document batch — mirrors core/bidDocument.ts's BidDocumentWorkItem. */
export interface BidDocumentWork {
  /** Position within the tender notice's item table (1-based) — labels the row "BID Document N". */
  serial: number
  /** Win Code this work was selected by — used to name the saved file "Bid Document <winCode>". */
  winCode?: string
  name: string
  /** Estimate amount as entered on the Works List, in Lakhs. */
  amount: string
  /** ECV, in Lakhs — EMD @ 1% is computed from this. Left blank when ECV isn't available yet. */
  ecv?: string
  zone?: string
  circle?: string
  completionPeriod?: string
}

/**
 * The set of Bid Documents to generate for one issued tender notice — one
 * entry per work in the notice's item table. Created automatically by the
 * Calendar section's "Issue tender notice" flow.
 */
export interface BidDocumentBatch {
  id: string
  nitNo: string
  /** DD.MM.YYYY */
  dated: string
  downloadStartDate: string
  downloadEndDate: string
  works: BidDocumentWork[]
  /** ISO date (YYYY-MM-DD) the batch was created. */
  createdDate: string
}

/** Snapshot of the workspace persisted to disk so it survives app restarts. */
export interface PersistedState {
  version: number
  tables: ExcelTable[]
  resolution: CollisionResolution
  todos?: TodoItem[]
  /** The last Google Sheets/Drive link used to fill the Works List, so it can be refreshed later. */
  lastGoogleLink?: string
  tenderReminders?: TenderReminder[]
  createdDocuments?: CreatedDocument[]
  bidDocumentBatches?: BidDocumentBatch[]
  mbScrutiny?: MBScrutinyItem[]
}
