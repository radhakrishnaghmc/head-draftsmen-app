import { api } from './ipc'
import type { ExcelTable } from '@core/types'
import type { EstimateWorkItem } from '@core/estimateExtract'
import { buildBoqFromEstimate, boqToScheduleA } from './boqTransform'
import { findWorksRowByName, metaFromWorksRow } from '@core/scheduleA'
import type { WorksRowMatch } from '@core/scheduleA'
import { applyEcvFromBoq } from '@core/worksAmounts'
import { computeMaterialTotals } from '@core/materialEstimate'
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

/**
 * Writes an estimate's computed ECV back to its matching Works List row
 * (name match, with the same embedding fallback as matchWorksRow) — returns
 * the updated table, or null if no matching row was found so the caller
 * leaves the Works List untouched.
 */
export async function saveEcvToWorksList(
  workName: string,
  ecvRupees: number,
  worksTable: ExcelTable
): Promise<ExcelTable | null> {
  let result = applyEcvFromBoq(worksTable, workName, ecvRupees)
  if (!result.matched) {
    const nameHeader = worksTable.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
    if (nameHeader) {
      try {
        const rowNames = worksTable.rows.map((r) => r[nameHeader] ?? '')
        const vectors = await api.embedTexts([workName, ...rowNames])
        result = applyEcvFromBoq(worksTable, workName, ecvRupees, {
          workNameVector: vectors[0],
          rowNameVectors: vectors.slice(1)
        })
      } catch {
        // Embeddings unavailable — leave the Works List untouched.
      }
    }
  }
  return result.matched ? result.table : null
}

export function downloadBoqFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  suggestedNameBase: string
): Promise<string | null> {
  const boq = buildBoqFromEstimate(items)
  return api.exportBoq(boq, `${suggestedNameBase} BOQ`, workName)
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

export function downloadMaterialFromItems(
  items: EstimateWorkItem[],
  workName: string | undefined,
  ecvRupees: number,
  departmentName: string,
  district: string,
  suggestedNameBase: string
): Promise<string | null> {
  const { totals } = computeMaterialTotals(items)
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
