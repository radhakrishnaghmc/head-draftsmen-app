import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'
import { checkSameWork } from './sameWorkCheck'
import type { ParsedDoc } from './massNotices'

/**
 * Bulk Intimations: pairs each L-1 selection sheet with its LOA (Online Intimation) and only accepts the pair
 * when the LOA agrees with the L-1 on all three of NIT number, agency name and contract amount. Anything less is
 * reported with the exact reason instead of producing a letter from mismatched sources.
 */

export interface LoaCheck {
  ok: boolean
  /** One message per failed check ("NIT number differs: …"). Empty when ok. */
  problems: string[]
}

const near = (a: number, b: number) => Math.abs(a - b) <= 1

/** NIT No, agency name and contract amount of an LOA against the L-1 sheet — all three must be present on both sides and agree. */
export function checkLoaAgainstL1(loa: IntimationNotice, l1: TenderEvaluation): LoaCheck {
  const problems: string[] = []
  // checkSameWork compares agency first and only falls back to the NIT when no agency is given, so run it once per field.
  const nit = checkSameWork({ nitNo: loa.nitNo }, l1)
  if (nit.status === 'mismatch') problems.push(`NIT number differs — LOA “${loa.nitNo}” vs L-1 “${l1.noticeNo}”`)
  else if (nit.status === 'unknown') problems.push('NIT number missing on the LOA or the L-1 sheet')
  const agency = checkSameWork({ agencyName: loa.agencyName }, l1)
  if (agency.status === 'mismatch') problems.push(`Agency name differs — LOA “${loa.agencyName}” vs L-1 “${l1.l1AgencyName}”`)
  else if (agency.status === 'unknown') problems.push('Agency name missing on the LOA or the L-1 sheet')
  if (loa.contractRupees == null || l1.contractRupees == null) problems.push('Contract amount missing on the LOA or the L-1 sheet')
  else if (!near(loa.contractRupees, l1.contractRupees))
    problems.push(`Contract amount differs — LOA Rs.${loa.contractRupees.toFixed(2)} vs L-1 Rs.${l1.contractRupees.toFixed(2)}`)
  return { ok: problems.length === 0, problems }
}

export interface IntimationPair {
  l1: TenderEvaluation
  loa: IntimationNotice
  l1File: string
  loaFile: string
}
export interface IntimationRejection {
  l1: TenderEvaluation
  l1File: string
  loaFile?: string
  reason: string
}

/**
 * Each L-1 takes the LOA that passes all three checks; when none does, the closest LOA (same agency or same NIT)
 * is named in the rejection with its failed checks, and an L-1 with no related LOA at all says so. An LOA is used
 * for at most one L-1.
 */
export function pairForIntimations(
  l1s: ParsedDoc<TenderEvaluation>[],
  loas: ParsedDoc<IntimationNotice>[]
): { pairs: IntimationPair[]; rejected: IntimationRejection[] } {
  const used = new Set<number>()
  const pairs: IntimationPair[] = []
  const rejected: IntimationRejection[] = []
  for (const { doc: l1, file: l1File } of l1s) {
    const related = loas
      .map((l, i) => ({ l, i, check: checkLoaAgainstL1(l.doc, l1) }))
      .filter(({ i }) => !used.has(i))
      .filter(
        ({ l }) =>
          checkSameWork({ nitNo: l.doc.nitNo }, l1).status === 'match' ||
          checkSameWork({ agencyName: l.doc.agencyName }, l1).status === 'match'
      )
    const good = related.find((r) => r.check.ok)
    if (good) {
      used.add(good.i)
      pairs.push({ l1, loa: good.l.doc, l1File, loaFile: good.l.file })
    } else if (related.length > 0) {
      rejected.push({ l1, l1File, loaFile: related[0].l.file, reason: related[0].check.problems.join('; ') })
    } else {
      rejected.push({ l1, l1File, reason: 'No LOA in the folder with this NIT number or agency' })
    }
  }
  return { pairs, rejected }
}
