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
  /**
   * When the table was imported from a sheet: that sheet's own columns, in the
   * sheet's original order (each mapped to its standard app name). A download of
   * the Works List reproduces exactly these columns in this order, so the file
   * mirrors the source Google sheet's arrangement rather than the app's internal
   * schema order. Absent for the built-in database (which exports its headers).
   */
  sourceHeaders?: string[]
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
  /**
   * The office (Corporation|Zone|Circle key) this task belongs to, so the To Do
   * list is per-office like the Works List — a task added in one circle isn't
   * shown in another. Absent on tasks created before per-office scoping; those
   * are adopted into the office in view on first load (see App.tsx).
   */
  officeKey?: string
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
  /**
   * The office (Corporation|Zone|Circle key) this record belongs to, so the MB
   * register is per-office (each circle keeps its own running serial number).
   * Absent on records created before per-office scoping; those are adopted into
   * the office in view on first load (see App.tsx).
   */
  officeKey?: string
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
  /**
   * Restricts which office this document is offered for on the Issue Documents
   * tab. Undefined = shown for every office. 'zonal' = only a Zone-level office
   * (no circle picked, e.g. the Superintending Engineer's office); 'circle' =
   * only a Circle-level (Executive Engineer's) office. Display-only filter — the
   * document still lives in the synced list regardless of the current office.
   */
  officeScope?: 'zonal' | 'circle'
}

/**
 * Ids of the original hand-made example documents that early builds shipped in
 * seed-state.json (and that never carried a stable, template-driven id). They
 * predate the version-gated default-document injection and have all since been
 * superseded by the bundled, template-refreshed defaults (doc_public_participation,
 * doc_action_taken_report, …). They must be actively removed — from local state,
 * from the seed, and from the cloud — because injection only ever adds/refreshes
 * and would otherwise leave these 8 lingering on the Issue Documents tab forever.
 * See pruneLegacyDocuments and injectDefaultDocuments (electron/main.ts).
 */
export const LEGACY_SEED_DOC_IDS: readonly string[] = [
  'doc_mrxmtf35', // Intimation
  'doc_mrxragvl', // Fowarding Slip for agreement
  'doc_mrx6ybmq', // agreement
  'doc_mrxcr44j', // work order
  'doc_mrxmve28', // EOT
  'doc_mrxqx32n', // Completion report
  'doc_mrxr0i1s', // Action Taken Report
  'doc_mrylcnwp' // note submitted
]

/** Drop any legacy example documents (see LEGACY_SEED_DOC_IDS) from a doc list. */
export function pruneLegacyDocuments<T extends { id: string }>(docs: T[]): T[] {
  return docs.filter((d) => !LEGACY_SEED_DOC_IDS.includes(d.id))
}

/** One outside QC agency's contact block (the "To" of a 3rd/4th-party letter). */
export interface QcPartyDetails {
  name: string
  address: string
  phone: string
}

/**
 * The 3rd-party (QC college) and 4th-party (testing lab) agencies for one
 * office. Entered once on the Works List page and reused by the 3rd/4th-party
 * QC letters, since these agencies change year to year but not work to work.
 */
export interface QcOfficeParties {
  third: QcPartyDetails
  fourth: QcPartyDetails
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
  /** Corporation abbreviation, e.g. "CMC" — was hard-coded in the EE template's submission-location line ("…Circle, SLPZ, CMC…"), now the work's actual corporation. */
  corporation?: string
  /** Corporation full name, e.g. "Cyberabad Municipal Corporation" — its upper-cased form fills the EE template's letterhead (was hard-coded "CYBERABAD MUNICIPAL CORPORATION"). */
  corporationFullName?: string
  completionPeriod?: string
  /** SE-office (zone-only, no circle) template fields — see core/bidDocument.ts's BidDocumentWorkItem for what each becomes on the document. Unused for an EE (circle) work. */
  itemNo?: string
  tsNo?: string
  tsDate?: string
  asAuthority?: 'zonal' | 'commissioner'
  asDate?: string
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
  /** The office (key: "Corporation|Zone|Circle") this batch was issued under — like the Works List/To Do, so switching offices doesn't show another office's batches. Absent on batches created before per-office scoping. */
  officeKey?: string
}

/** Snapshot of the workspace persisted to disk so it survives app restarts. */
export interface PersistedState {
  version: number
  tables: ExcelTable[]
  /**
   * The Works List data kept per office (key: "Corporation|Zone|Circle"), so
   * switching offices preserves each office's own database instead of erasing it
   * and re-importing from the link. Each office's entry syncs independently, so
   * concurrent sessions working on different offices never overwrite each other.
   */
  tablesByOffice?: Record<string, ExcelTable[]>
  /** Transient (not stored on disk): which office `tables` belongs to, so the cloud sync updates only that office's entry. */
  currentOfficeKey?: string
  resolution: CollisionResolution
  todos?: TodoItem[]
  /** The last Google Sheets/Drive link used to fill the Works List, so it can be refreshed later. */
  lastGoogleLink?: string
  /**
   * The Works List link remembered per office (key: "Corporation|Zone|Circle"),
   * so switching back to a circle reloads its own database automatically — and,
   * because this is part of the synced state, the links follow the user to any
   * system they log in from.
   */
  worksListLinks?: Record<string, string>
  /**
   * The 3rd/4th-party QC agencies remembered per office (key:
   * "Corporation|Zone|Circle"), entered once on the Works List page and reused
   * by the 3rd/4th-party QC letters. Synced like worksListLinks.
   */
  qcParties?: Record<string, QcOfficeParties>
  /**
   * The Head Draughtsman's chosen office (Corporation/Zone/Circle/CNO). Persisted
   * here — not just in the renderer's localStorage — so it's remembered on the
   * next login and restored on any system the user signs in from, and never
   * re-asked once picked. localStorage still holds it for an instant read at
   * startup; this synced copy is the fallback whenever that's been lost (a fresh
   * install, a cleared cache, or a different machine).
   */
  office?: { corporation?: string; zone?: string; circle?: string; circleNumber?: string }
  tenderReminders?: TenderReminder[]
  createdDocuments?: CreatedDocument[]
  /**
   * Highest version of the bundled default documents already merged into this
   * state. New standard documents (e.g. the Public Participation Log Book) are
   * injected once when this trails the current version — see
   * injectDefaultDocuments in electron/main.ts. Persisted so the injection runs
   * exactly once and a document the user later deletes is never re-added.
   */
  seededDocVersion?: number
  bidDocumentBatches?: BidDocumentBatch[]
  mbScrutiny?: MBScrutinyItem[]
}
