import { api } from './ipc'
import type { ExcelTable } from '@core/types'
import type { EstimateWorkItem } from '@core/estimateExtract'
import { buildBoqFromEstimate, boqToScheduleA } from './boqTransform'
import { findWorksRowByName, metaFromWorksRow } from '@core/scheduleA'
import type { WorksRowMatch } from '@core/scheduleA'
import { computeMaterialTotals } from '@core/materialEstimate'
import { matchMaterialRag } from '@core/materialQuantityRag'
import type { MaterialRag } from '@core/materialQuantityRag'
import materialRagData from '@core/data/materialQuantityRag.json'
import type { DeviationItem } from '@core/deviationTemplate'

/**
 * Matches a work name against the Works List by exact name first, falling
 * back to the local embedding model when that fails (an estimate's
 * title-block wording very often differs slightly from the Works List's own
 * entry) — shared by every "download from an estimate" action below, and
 * previously duplicated near-identically across the tabs this module
 * replaces (EstimateToBoqTab, UploadPhotosTab).
 */
export async function matchWorksRow(name: string, table: ExcelTable): Promise<WorksRowMatch | undefined> {
  const exact = findWorksRowByName(table, name)
  if (exact) return exact
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  if (!nameHeader) return undefined
  try {
    const rowNames = table.rows.map((r) => r[nameHeader] ?? '')
    const vectors = await api.embedTexts([name, ...rowNames])
    return findWorksRowByName(table, name, { workNameVector: vectors[0], rowNameVectors: vectors.slice(1) })
  } catch {
    return undefined
  }
}

export interface WorksRowEmbeddingCache {
  rowNameVectors: number[][]
}

/**
 * Embeds every Works List row name ONCE, up front — for matchWorksRowCached
 * below to reuse across a whole batch of uploaded estimates. matchWorksRow
 * re-embeds all rows on every call; uploading N estimates that each miss an
 * exact match turned one embedding pass over the list into N of them, which
 * is what made bulk estimate uploads slow. Returns null when there's no
 * Works List, no "Name of the work" column, or the embedding model itself is
 * unavailable — callers should fall back to exact-match-only in that case.
 */
export async function buildWorksRowEmbeddingCache(table: ExcelTable): Promise<WorksRowEmbeddingCache | null> {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  if (!nameHeader) return null
  try {
    const rowNames = table.rows.map((r) => r[nameHeader] ?? '')
    const rowNameVectors = await api.embedTexts(rowNames)
    return { rowNameVectors }
  } catch {
    return null
  }
}

/**
 * Same as matchWorksRow, but reuses a WorksRowEmbeddingCache instead of
 * re-embedding every Works List row — only the one work name being looked up
 * still needs embedding, so a batch of uploads costs one row-embedding pass
 * total plus one small per-name call each, instead of a full pass per file.
 */
export async function matchWorksRowCached(
  name: string,
  table: ExcelTable,
  cache: WorksRowEmbeddingCache | null
): Promise<WorksRowMatch | undefined> {
  const exact = findWorksRowByName(table, name)
  if (exact) return exact
  if (!cache) return undefined
  try {
    const [workNameVector] = await api.embedTexts([name])
    return findWorksRowByName(table, name, { workNameVector, rowNameVectors: cache.rowNameVectors })
  } catch {
    return undefined
  }
}

export function downloadBoqFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  suggestedNameBase: string,
  // When given, used verbatim as the suggested file name (e.g. the estimate
  // flow's "Boq <sheet> <lakhs>"); otherwise the base gets the " BOQ" suffix.
  exactName?: string
): Promise<string | null> {
  const boq = buildBoqFromEstimate(items)
  return api.exportBoq(boq, exactName ?? `${suggestedNameBase} BOQ`, workName)
}

export async function downloadScheduleAFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  worksTable: ExcelTable | null,
  suggestedNameBase: string
): Promise<string | null> {
  const scheduleARows = boqToScheduleA(buildBoqFromEstimate(items))
  const match = workName && worksTable ? await matchWorksRow(workName, worksTable) : undefined
  const meta = match ? metaFromWorksRow(match.row) : undefined
  return api.exportScheduleA(scheduleARows, `${suggestedNameBase} Schedule A`, meta)
}

export async function downloadDeviationFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  agencyName: string,
  estimateAmountLakhs: number,
  worksTable: ExcelTable | null,
  suggestedNameBase: string
): Promise<string | null> {
  const deviationItems: DeviationItem[] = items.map((it) => ({
    description: it.description,
    unit: it.unit,
    quantity: it.quantity,
    rate: it.rate
  }))
  const match = workName && worksTable ? await matchWorksRow(workName, worksTable) : undefined
  return api.exportDeviation(
    deviationItems,
    { circle: match?.row['Circle'], nameOfWork: workName, agencyName: agencyName.trim(), estimateAmountLakhs },
    `${suggestedNameBase} Deviation`
  )
}

const MATERIAL_RAG = materialRagData as unknown as MaterialRag

/**
 * Material Quantity totals for an estimate: the bundled material-quantity index
 * (built from the SE office Data Sheet — see scripts/buildMaterialRag.ts) supplies
 * each matched item's per-unit material quantities, and the standard rules in
 * core/materialEstimate.ts cover whatever it doesn't match.
 */
export function materialTotalsForItems(items: EstimateWorkItem[]): ReturnType<typeof computeMaterialTotals> {
  return computeMaterialTotals(items, undefined, matchMaterialRag(items, MATERIAL_RAG))
}

export function downloadMaterialFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  ecvRupees: number,
  departmentName: string,
  district: string,
  suggestedNameBase: string
): Promise<string | null> {
  const { totals } = materialTotalsForItems(items)
  return api.exportMaterialEstimate(
    totals,
    {
      workName,
      departmentName: departmentName.trim() || undefined,
      district: district.trim() || undefined,
      ecvRupees
    },
    `${suggestedNameBase} Material Quantity`
  )
}
