import type { RateEntry } from './rateDatabase'
import type { EstimateWorkItem } from './estimateExtract'
import { matchEstimateItems } from './technicalSanction'
import { breakdownMaterialLines } from './materialEstimate'
import type { MaterialKey, MaterialLine } from './materialEstimate'

/**
 * One Data Sheet item and the materials it consumes, per 1 unit of the item.
 * Rates are deliberately not kept: they change every year, but the quantities
 * of material an item needs do not.
 */
export interface MaterialRagEntry {
  /** The item's description (what estimate items are matched against). */
  description: string
  /** The Data Sheet's index code for the item, when it has one (e.g. "RBR-EECD-8(c)"). */
  code?: string
  /** The item's measuring unit (cum/sqm/rmt/…), when the sheet states it. */
  unit?: string
  sheet: string
  materials: MaterialLine[]
}

/** The built index: every material-bearing Data Sheet item, plus per-material lists of the entries that use it. */
export interface MaterialRag {
  /** Data Sheet the index was built from, e.g. "Data 2026-27 SE office data for the month July 2026". */
  source: string
  entries: MaterialRagEntry[]
  /** For each material (cement, steel, sand, …), the indices into `entries` of every item that uses it. */
  byMaterial: Record<MaterialKey, number[]>
}

const LABOUR_ITEM_RE = /\blabou?r\b/i

const MATERIAL_KEYS: MaterialKey[] = ['cement', 'steel', 'sand', 'aggregate', 'gravel', 'granite', 'napa']

/** Build the index from a parsed Data Sheet's rate entries (see core/rateDatabase.ts's buildRateIndex). */
export function buildMaterialRag(rateEntries: RateEntry[], source: string): MaterialRag {
  const entries: MaterialRagEntry[] = []
  const seen = new Set<string>()
  for (const e of rateEntries) {
    // Only worked-out buildup blocks carry per-unit material quantities; a
    // one-row entry is a flat rate (e.g. "granite 2789 per sqm"), whose number
    // is a price, not a quantity.
    if ((e.breakdown?.length ?? 0) < 3) continue
    // A labour-only item ("Labour for cutting, bending, tying…") carries just the
    // binding wire it uses, which is not steel to be supplied: no steel from it.
    const labourItem = LABOUR_ITEM_RE.test(e.description)
    const materials = breakdownMaterialLines(e.breakdown ?? []).filter((m) => !(labourItem && m.material === 'steel'))
    if (materials.length === 0) continue
    // The same item repeats across sheets/rate variants; one entry per (code or description, materials).
    const key = `${e.code ?? e.description}|${materials.map((m) => `${m.material}:${m.unit}:${m.amount.toPrecision(6)}`).join(',')}`
    if (seen.has(key)) continue
    seen.add(key)
    entries.push({ description: e.description, code: e.code, unit: e.unit, sheet: e.sheet, materials })
  }
  const byMaterial = Object.fromEntries(MATERIAL_KEYS.map((k) => [k, [] as number[]])) as Record<MaterialKey, number[]>
  entries.forEach((entry, i) => {
    for (const k of new Set(entry.materials.map((m) => m.material))) byMaterial[k].push(i)
  })
  return { source, entries, byMaterial }
}

const entryKey = (e: { code?: string; description: string }): string => (e.code ?? e.description).trim().toLowerCase()

/**
 * Minimum similarity for taking an index entry's quantities for an estimate
 * item. Lower than Give Technical Sanction's auto-fill bar (0.7): a wrong rate
 * lands on an official document, whereas a quantity here only replaces the
 * standard-coefficient estimate, and estimate wording is rarely identical to
 * the Data Sheet's (a right match typically scores 0.4–0.65). Exactly one
 * candidate, found by real keyword overlap, is still required.
 */
export const MATERIAL_MATCH_THRESHOLD = 0.4

/**
 * Look each estimate item up in the index (exact Data Sheet code first, then
 * keyword similarity — the same matcher Give Technical Sanction uses, which
 * also rejects a unit-incompatible entry) and return that item's material
 * lines, same order/length as `items`. An item with no confident match is
 * undefined, so the caller falls back to its standard rules for it.
 */
export function matchMaterialRag(items: EstimateWorkItem[], rag: MaterialRag): (MaterialLine[] | undefined)[] {
  if (items.length === 0 || rag.entries.length === 0) return items.map(() => undefined)
  const byKey = new Map<string, MaterialRagEntry>()
  for (const e of rag.entries) if (!byKey.has(entryKey(e))) byKey.set(entryKey(e), e)
  const index: RateEntry[] = rag.entries.map((e) => ({
    description: e.description,
    rate: '',
    sheet: e.sheet,
    breakdown: [],
    unit: e.unit,
    code: e.code
  }))
  return matchEstimateItems(items, index).map((m) => {
    if (m.candidates.length !== 1 || m.semanticOnly || m.bestScore < MATERIAL_MATCH_THRESHOLD) return undefined
    return byKey.get(entryKey(m.candidates[0]))?.materials
  })
}
