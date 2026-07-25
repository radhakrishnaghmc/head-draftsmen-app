import type { EstimateWorkItem } from './estimateExtract'

/**
 * Converts an estimate's work items into the 7 material totals the bundled
 * Material Estimation Template asks for, using standard SSR/CPWD "Analysis
 * of Rates" material constants (nominal-mix concrete ratios, the 1.54
 * dry-volume factor, standard bulk densities, ...). These are the same
 * general-purpose coefficients used across Telangana/AP PWD rate analyses,
 * not this department's own analysis of rates — treat the output as a
 * first-pass estimate to be checked against the department's own figures
 * before it's used for actual procurement.
 */
export interface MaterialTotals {
  stoneAggregatesMt: number
  sandMt: number
  gravelMt: number
  graniteSqft: number
  napaSqft: number
  cementMt: number
  steelMt: number
}

export interface MaterialEstimateResult {
  totals: MaterialTotals
  /** How many of the estimate's items matched a known material rule — the rest had no recognizable concrete/masonry/plaster/flooring/steel wording and were skipped rather than guessed at. */
  matchedItemCount: number
  totalItemCount: number
}

// Standard bulk/dry densities (tonnes per cum) used to turn a computed
// volume into the template's required MT figure.
const CEMENT_DENSITY_MT_PER_CUM = 1.44
const SAND_DENSITY_MT_PER_CUM = 1.6
const AGGREGATE_DENSITY_MT_PER_CUM = 1.55
const GRAVEL_DENSITY_MT_PER_CUM = 1.75

// Dry volume of cement+sand+aggregate needed to produce 1 cum of *finished*
// (wet, compacted) concrete or mortar — the standard CPWD/SSR constant that
// accounts for the voids between the dry ingredients closing up on mixing.
const DRY_VOLUME_FACTOR = 1.54
// Same idea for a thin mortar layer (plaster/pointing/brick-joint mortar),
// where the standard allowance is smaller since it's not being compacted
// the same way a concrete pour is.
const MORTAR_BULKING_FACTOR = 1.27
// Mortar consumed per cum of brick/block masonry, net of the brick/block
// volume itself — the standard rule-of-thumb allowance (SSR analysis of
// rates) for a typical modular brick with 10mm joints.
const MASONRY_MORTAR_FRACTION = 0.3

const SQM_PER_SQFT = 1 / 10.7639

// Nominal mixes (cement:sand:aggregate) for the standard, pre-M25 concrete
// grades — the traditional fixed ratios these grades were originally
// defined by. M25 and above are design mixes with no fixed ratio, so those
// fall back to DESIGN_MIX_DEFAULT below instead.
const NOMINAL_MIX_BY_GRADE: Record<number, [number, number, number]> = {
  5: [1, 5, 10],
  7.5: [1, 4, 8],
  10: [1, 3, 6],
  15: [1, 2, 4],
  20: [1, 1.5, 3]
}
// Typical design-mix material content for M25+ concrete, per cum — an
// approximation (real design mixes vary by mix design), used only when the
// item names a grade but not a ratio.
const DESIGN_MIX_DEFAULT = { cementMtPerCum: 0.34, sandCumPerCum: 0.44, aggregateCumPerCum: 0.88 }

function toNumber(s: string | undefined): number {
  const n = Number(String(s ?? '').replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

function normUnit(u: string): string {
  return String(u ?? '').toLowerCase().replace(/[^a-z]/g, '')
}
function isCum(u: string): boolean {
  return /^cum/.test(normUnit(u))
}
function isSqm(u: string): boolean {
  return /^sq?m/.test(normUnit(u)) && !/^sqft/.test(normUnit(u))
}
function isSqft(u: string): boolean {
  return /^s?qft/.test(normUnit(u))
}
function isKg(u: string): boolean {
  return /^kgs?$/.test(normUnit(u))
}
function isMt(u: string): boolean {
  return /^(mt|tonn?es?|tons?)$/.test(normUnit(u))
}
function isQuintal(u: string): boolean {
  return /^q(tl|uintal)/.test(normUnit(u))
}
function isBag(u: string): boolean {
  return /^bags?/.test(normUnit(u))
}

/** Area in Sq.Ft from a Sqm or Sq.Ft item — undefined for any other unit. */
function areaSqft(qty: number, unit: string): number | undefined {
  if (isSqft(unit)) return qty
  if (isSqm(unit)) return qty / SQM_PER_SQFT
  return undefined
}

/** Weight in MT from a Kg/MT/Quintal item — undefined for any other unit. */
function weightMt(qty: number, unit: string): number | undefined {
  if (isMt(unit)) return qty
  if (isKg(unit)) return qty / 1000
  if (isQuintal(unit)) return qty * 0.1
  return undefined
}

const RATIO3_RE = /(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)/
const RATIO2_RE = /(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)(?!\s*:)/
// "M20", "M 20", "M-20" — not "20MM" (see the \b/no-preceding-digit reasoning below).
const GRADE_RE = /\bM[\s-]?(\d{1,2}(?:\.\d+)?)\b/i
const THICKNESS_MM_RE = /(\d+(?:\.\d+)?)\s*mm/i

const STEEL_RE = /\b(reinforce?ment|tmt|hysd|fe[\s-]?(415|500|550|550d)|mild\s+steel|structural\s+steel|steel\s+bars?|binding\s+wire)\b/i
// Plain "concrete/PCC/RCC" (cement-bound) — checked after BITUMINOUS_RE so a
// "bituminous concrete" road-surface item (bitumen-bound, no cement at all)
// is routed there instead of here.
const CONCRETE_RE = /\b(concrete|pcc|rcc)\b/i
// Bituminous road layers (base/binder/surface courses) — overwhelmingly
// crushed stone aggregate premixed with a bitumen binder (MoRTH specs put
// the binder at ~4-5.5% by weight, the rest aggregate+filler), not cement
// concrete despite "Bituminous Concrete" sharing the word "concrete".
const BITUMINOUS_RE =
  /\bbituminous\s+(macadam|concrete)\b|\bdense\s+graded\s+bituminous\b|\bsemi[\s-]?dense\s+bituminous\b|\bpremix\s+carpet\b|\bbuilt[\s-]?up\s+spray\s+grout\b|\bDBM\b|\bSDBC\b/i
const MASONRY_RE = /\b(brick\s*work|brickwork|block\s*work|blockwork|masonry)\b/i
const PLASTER_RE = /\b(plaster(ing)?|pointing|rendering)\b/i
const GRANITE_RE = /\bgranite\b/i
// Napa/Kadapa/Shahabad slabs are regional names for the same limestone
// flooring slab, commonly used interchangeably in Telangana estimates.
const NAPA_RE = /\b(napa|kadapa|shahabad)\b/i
// WBM (water bound macadam) and WMM (wet mix macadam) are both graded-
// aggregate road base/sub-base courses — bound with water/moisture only, no
// cement or bitumen, matching the template's "Gravel" (base-course
// aggregate) bucket rather than "Stone Aggregates" (concrete aggregate).
const GRAVEL_RE =
  /\b(gravel|wbm|wmm|water\s*bound\s*macadam|wet\s*mix\s*macadam|g\.?s\.?b\.?|granular\s+sub[\s-]?base|moorum|murum)\b/i
const AGGREGATE_SUPPLY_RE = /\b(stone\s+aggregate|coarse\s+aggregate|graded\s+aggregate|metal\s*(20|40|12)\s*mm)\b/i
const SAND_RE = /\bsand\b/i
const CEMENT_RE = /\bcement\b/i

// Typical compacted density of a hot-mix bituminous layer (DBM/BC/SDBC),
// times the aggregate+filler share of that mix (bitumen binder is ~4-5.5%
// by weight per MoRTH specs, so aggregate is ~95%) — an approximation, since
// the real figure depends on the specific mix design's binder content and
// gradation.
const BITUMINOUS_AGGREGATE_MT_PER_CUM = 2.35 * 0.95

function addConcreteOrMortar(
  totals: MaterialTotals,
  dryVolumeCum: number,
  ratio: [number, number]
): void {
  const [a, b] = ratio
  const sum = a + b
  if (sum <= 0) return
  totals.cementMt += (dryVolumeCum * a) / sum * CEMENT_DENSITY_MT_PER_CUM
  totals.sandMt += (dryVolumeCum * b) / sum * SAND_DENSITY_MT_PER_CUM
}

function addConcreteWithAggregate(totals: MaterialTotals, dryVolumeCum: number, ratio: [number, number, number]): void {
  const sum = ratio[0] + ratio[1] + ratio[2]
  if (sum <= 0) return
  totals.cementMt += (dryVolumeCum * ratio[0]) / sum * CEMENT_DENSITY_MT_PER_CUM
  totals.sandMt += (dryVolumeCum * ratio[1]) / sum * SAND_DENSITY_MT_PER_CUM
  totals.stoneAggregatesMt += (dryVolumeCum * ratio[2]) / sum * AGGREGATE_DENSITY_MT_PER_CUM
}

/**
 * Classify and accumulate one estimate item's material contribution.
 * Returns true if the item matched a known material rule (concrete,
 * masonry, plaster, flooring slab, gravel/WBM/WMM, a bituminous road layer,
 * steel, or a standalone aggregate/sand/cement supply item) — false if
 * nothing recognizable was found, in which case nothing is added (fails
 * closed, same philosophy as the rest of this app's estimate-parsing code:
 * never guess a figure). Rules are checked in order of specificity and the
 * first match wins, so a single item (e.g. "granite flooring with cement
 * mortar bedding") is only counted once, toward its primary material.
 */
function classifyItem(item: EstimateWorkItem, totals: MaterialTotals): boolean {
  const desc = item.description
  const unit = item.unit
  const qty = toNumber(item.quantity)
  if (qty <= 0) return false

  if (STEEL_RE.test(desc)) {
    const mt = weightMt(qty, unit)
    if (mt !== undefined) {
      totals.steelMt += mt
      return true
    }
  }

  if (BITUMINOUS_RE.test(desc)) {
    if (isCum(unit)) {
      totals.stoneAggregatesMt += qty * BITUMINOUS_AGGREGATE_MT_PER_CUM
      return true
    }
    const mt = weightMt(qty, unit)
    if (mt !== undefined) {
      totals.stoneAggregatesMt += mt * 0.95
      return true
    }
    // A bituminous item quoted in Sqm (a thin surface dressing/seal coat with
    // no meaningful volume) has no reliable aggregate quantity to derive —
    // left unmatched rather than guessed.
  } else if (CONCRETE_RE.test(desc) && isCum(unit)) {
    const dryVol = qty * DRY_VOLUME_FACTOR
    const m3 = RATIO3_RE.exec(desc)
    if (m3) {
      addConcreteWithAggregate(totals, dryVol, [Number(m3[1]), Number(m3[2]), Number(m3[3])])
      return true
    }
    const grade = GRADE_RE.exec(desc)
    if (grade) {
      const nominal = NOMINAL_MIX_BY_GRADE[Number(grade[1])]
      if (nominal) {
        addConcreteWithAggregate(totals, dryVol, nominal)
      } else {
        totals.cementMt += qty * DESIGN_MIX_DEFAULT.cementMtPerCum
        totals.sandMt += qty * DESIGN_MIX_DEFAULT.sandCumPerCum * SAND_DENSITY_MT_PER_CUM
        totals.stoneAggregatesMt += qty * DESIGN_MIX_DEFAULT.aggregateCumPerCum * AGGREGATE_DENSITY_MT_PER_CUM
      }
      return true
    }
    // Concrete mentioned with neither a ratio nor a grade — too ambiguous
    // to assume a mix for, so this item is left unmatched rather than
    // defaulting to a guessed ratio.
  }

  if (MASONRY_RE.test(desc) && isCum(unit)) {
    const ratio = RATIO2_RE.exec(desc)
    if (ratio) {
      const mortarVol = qty * MASONRY_MORTAR_FRACTION * MORTAR_BULKING_FACTOR
      addConcreteOrMortar(totals, mortarVol, [Number(ratio[1]), Number(ratio[2])])
      return true
    }
  }

  if (PLASTER_RE.test(desc)) {
    const area = areaSqft(qty, unit)
    const ratio = RATIO2_RE.exec(desc)
    if (area !== undefined && ratio) {
      const thicknessMatch = THICKNESS_MM_RE.exec(desc)
      const thicknessMm = thicknessMatch ? Number(thicknessMatch[1]) : /pointing/i.test(desc) ? 6 : 12
      const areaSqm = area * SQM_PER_SQFT
      const mortarVol = areaSqm * (thicknessMm / 1000) * MORTAR_BULKING_FACTOR
      addConcreteOrMortar(totals, mortarVol, [Number(ratio[1]), Number(ratio[2])])
      return true
    }
  }

  if (GRANITE_RE.test(desc)) {
    const area = areaSqft(qty, unit)
    if (area !== undefined) {
      totals.graniteSqft += area
      return true
    }
  }

  if (NAPA_RE.test(desc)) {
    const area = areaSqft(qty, unit)
    if (area !== undefined) {
      totals.napaSqft += area
      return true
    }
  }

  if (GRAVEL_RE.test(desc) && isCum(unit)) {
    totals.gravelMt += qty * GRAVEL_DENSITY_MT_PER_CUM
    return true
  }

  if (AGGREGATE_SUPPLY_RE.test(desc) && isCum(unit)) {
    totals.stoneAggregatesMt += qty * AGGREGATE_DENSITY_MT_PER_CUM
    return true
  }

  if (SAND_RE.test(desc) && isCum(unit)) {
    totals.sandMt += qty * SAND_DENSITY_MT_PER_CUM
    return true
  }

  if (CEMENT_RE.test(desc)) {
    if (isBag(unit)) {
      totals.cementMt += qty * 0.05
      return true
    }
    const mt = weightMt(qty, unit)
    if (mt !== undefined) {
      totals.cementMt += mt
      return true
    }
  }

  return false
}

// A role name on a breakdown row (e.g. "Cement Mason", "Coolie") marks it as
// a labour line, not a material supply line — excluded even if the role
// name happens to contain a material word, so a mason's wage never gets
// mistaken for a bag of cement.
const LABOUR_ROLE_RE = /\b(mason|coolie|labour|labor|mazdoor|beldar|bhisti|carpenter|fitter|welder|operator|driver)\b/i

type MaterialKey = 'steel' | 'cement' | 'sand' | 'aggregate' | 'gravel' | 'granite' | 'napa'

const FINE_AGGREGATE_RE = /\bfine\s+aggregate\b/i
const COARSE_AGGREGATE_RE = /\bcoarse\s+aggregate\b|\baggregate\b|\bmetal\b/i

function materialKeyFromText(text: string): MaterialKey | null {
  if (STEEL_RE.test(text)) return 'steel'
  if (CEMENT_RE.test(text)) return 'cement'
  if (FINE_AGGREGATE_RE.test(text) || SAND_RE.test(text)) return 'sand'
  if (GRAVEL_RE.test(text)) return 'gravel'
  if (COARSE_AGGREGATE_RE.test(text)) return 'aggregate'
  if (GRANITE_RE.test(text)) return 'granite'
  if (NAPA_RE.test(text)) return 'napa'
  return null
}

function isNumericCell(s: string): boolean {
  const t = s.trim().replace(/,/g, '')
  return t !== '' && Number.isFinite(Number(t))
}

/** Add `perUnitQty` amount (in physical unit `unit`) of `material`, scaled by the item's own quantity, to `totals` — converting through the same density constants classifyItem uses. Returns false for a (material, unit) combination that doesn't make physical sense (e.g. granite slabs measured in Cum), left unmatched rather than guessed. */
function addBreakdownMaterial(totals: MaterialTotals, material: MaterialKey, amount: number, unit: string): boolean {
  if (isCum(unit)) {
    switch (material) {
      case 'cement':
        totals.cementMt += amount * CEMENT_DENSITY_MT_PER_CUM
        return true
      case 'sand':
        totals.sandMt += amount * SAND_DENSITY_MT_PER_CUM
        return true
      case 'aggregate':
        totals.stoneAggregatesMt += amount * AGGREGATE_DENSITY_MT_PER_CUM
        return true
      case 'gravel':
        totals.gravelMt += amount * GRAVEL_DENSITY_MT_PER_CUM
        return true
      default:
        return false
    }
  }
  if (isBag(unit) && material === 'cement') {
    totals.cementMt += amount * 0.05
    return true
  }
  const mt = weightMt(amount, unit)
  if (mt !== undefined) {
    switch (material) {
      case 'cement':
        totals.cementMt += mt
        return true
      case 'sand':
        totals.sandMt += mt
        return true
      case 'aggregate':
        totals.stoneAggregatesMt += mt
        return true
      case 'gravel':
        totals.gravelMt += mt
        return true
      case 'steel':
        totals.steelMt += mt
        return true
      default:
        return false
    }
  }
  const sqft = areaSqft(amount, unit)
  if (sqft !== undefined && (material === 'granite' || material === 'napa')) {
    if (material === 'granite') totals.graniteSqft += sqft
    else totals.napaSqft += sqft
    return true
  }
  return false
}

/**
 * Pull material coefficients straight out of a Data Sheet rate-analysis
 * breakdown (the raw material/labour/machinery buildup rows behind one
 * item's rate — see core/rateDatabase.ts) and apply them, scaled by this
 * estimate item's own quantity, to `totals`. Each breakdown row is the
 * department's own stated "per 1 <item unit>" coefficient (e.g. "Cement
 * 0.22 Cum" per Cum of PCC 1:2:4), the same convention the rate analysis
 * itself is built on — so no ratio/grade guessing is needed here, unlike
 * classifyItem's fallback path.
 *
 * A row counts as a material line only when it names a recognized material
 * and carries a recognized physical unit (Cum/Sqm/Sqft/Kg/MT/Bag/Quintal) —
 * a labour day-rate or a machinery hour-rate row has neither and is safely
 * skipped, as is any row naming a labour role (LABOUR_ROLE_RE) even if it
 * happens to mention a material word (e.g. "Cement Mason").
 */
function applyBreakdownMaterials(itemQty: number, breakdown: string[][], totals: MaterialTotals): boolean {
  let matched = false
  for (const row of breakdown) {
    const cells = row.map((c) => String(c ?? '').trim())
    const rowText = cells.join(' ')
    if (LABOUR_ROLE_RE.test(rowText)) continue
    const material = materialKeyFromText(rowText)
    if (!material) continue

    const unitIdx = cells.findIndex(
      (c) => isCum(c) || isSqm(c) || isSqft(c) || isKg(c) || isMt(c) || isQuintal(c) || isBag(c)
    )
    if (unitIdx === -1) continue
    let qtyIdx = -1
    for (let i = unitIdx - 1; i >= 0; i--) {
      if (isNumericCell(cells[i])) {
        qtyIdx = i
        break
      }
    }
    if (qtyIdx === -1) continue

    const coefficient = Number(cells[qtyIdx].replace(/,/g, ''))
    if (addBreakdownMaterial(totals, material, coefficient * itemQty, cells[unitIdx])) matched = true
  }
  return matched
}

export function computeMaterialTotals(
  items: EstimateWorkItem[],
  /** The Data Sheet's rate-analysis breakdown for each item that was resolved against it (see core/technicalSanction.ts's ItemMatch), same order/length as `items` — undefined entries (no Data Sheet, or that item wasn't resolved) fall back to classifyItem's own standard-coefficient rules. */
  breakdowns?: (string[][] | undefined)[]
): MaterialEstimateResult {
  const totals: MaterialTotals = {
    stoneAggregatesMt: 0,
    sandMt: 0,
    gravelMt: 0,
    graniteSqft: 0,
    napaSqft: 0,
    cementMt: 0,
    steelMt: 0
  }
  let matchedItemCount = 0
  items.forEach((item, i) => {
    const breakdown = breakdowns?.[i]
    const matchedFromDataSheet =
      breakdown && breakdown.length > 0 ? applyBreakdownMaterials(toNumber(item.quantity), breakdown, totals) : false
    const matched = matchedFromDataSheet || classifyItem(item, totals)
    if (matched) matchedItemCount += 1
  })
  return { totals, matchedItemCount, totalItemCount: items.length }
}
